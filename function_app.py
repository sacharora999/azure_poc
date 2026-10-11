"""Azure Functions Python v2 Durable orchestration; no HTTP request held during analysis."""
import os,io,json,uuid,re,logging
from collections import defaultdict
from datetime import timedelta
import azure.functions as func
import azure.durable_functions as df
from pypdf import PdfReader,PdfWriter
from processor import PDF_AMOUNT_FIELDS, _num, parse_two_excels,reconcile_two_inputs,build_output
from storage import get,put,stream_lines,sha256
from pdf_extract import analyze,extract_rows
from postgres_persistence import persist

app=df.DFApp(http_auth_level=func.AuthLevel.FUNCTION)
INPUT=os.getenv('RECON_INPUT_CONTAINER','input-recon')
WORK=os.getenv('RECON_WORK_CONTAINER','recon-work')
OUTPUT=os.getenv('RECON_OUTPUT_CONTAINER','output-recon')
MAX_PAGES=int(os.getenv('PDF_CHUNK_PAGES','200'))


def clean_name(name):
    if not isinstance(name,str) or not name or name!=name.split('/')[-1] or '\\' in name or '..' in name:
        raise ValueError('Invalid blob filename')
    return name

def validate_request(body):
    if not isinstance(body,dict):raise ValueError('JSON object expected')
    pdfs=body.get('pdfFiles')
    if not isinstance(pdfs,list) or not 1<=len(pdfs)<=10 or len(pdfs)!=len(set(pdfs)):
        raise ValueError('pdfFiles must contain 1–10 unique PDF filenames')
    for n in pdfs:clean_name(n)
    for k in ('mainAccountFile','smwFile'):clean_name(body[k])
    raw=body.get('states') or body.get('state')
    states=raw.split(',') if isinstance(raw,str) else raw
    if not isinstance(states,list) or not states:raise ValueError('states required')
    states=[str(x).strip().upper() for x in states if str(x).strip()]
    if not states or any(not re.fullmatch(r'[A-Z]{2}',s) for s in states):raise ValueError('Invalid state codes')
    body=dict(body,states=list(dict.fromkeys(states)),state=','.join(dict.fromkeys(states)))
    body.setdefault('period','July 2026');body.setdefault('requestedBy','unknown')
    return body

@app.route(route='state-recon/start',methods=['POST'])
@app.durable_client_input(client_name='client')
async def start(req:func.HttpRequest,client:df.DurableOrchestrationClient):
    try:
        body=validate_request(req.get_json())
        # Avoid logging user data or signed URLs.
        run_id=str(uuid.uuid4())
        await client.start_new('state_recon_orchestrator',instance_id=run_id,client_input={'runId':run_id,'body':body})
        return client.create_check_status_response(req,run_id)
    except (ValueError,KeyError,TypeError) as exc:
        return func.HttpResponse(json.dumps({'error':str(exc)}),status_code=400,mimetype='application/json')

@app.orchestration_trigger(context_name='context')
def state_recon_orchestrator(context:df.DurableOrchestrationContext):
    payload=context.get_input();body=payload['body'];run_id=payload['runId']
    preparations=yield context.task_all([context.call_activity('prepare_pdf',{'runId':run_id,'filename':name,'fileIndex':i}) for i,name in enumerate(body['pdfFiles'])])
    jobs=[job for group in preparations for job in group]
    # Each batch runs concurrently. The configured cap is per orchestration, not a global DI quota.
    parallel=max(1,min(10,int(os.getenv('DI_MAX_CONCURRENCY','5'))))
    results=[]
    for start_index in range(0,len(jobs),parallel):
        batch=jobs[start_index:start_index+parallel]
        tasks=[context.call_activity_with_retry('analyze_pdf_chunk',df.RetryOptions(5000,3),job) for job in batch]
        results.extend((yield context.task_all(tasks)))
    outcome=yield context.call_activity('finalize_recon',{'runId':run_id,'body':body,'chunks':results})
    return outcome

@app.activity_trigger(input_name='job')
def prepare_pdf(job:dict):
    filename=clean_name(job['filename']); data=get(INPUT,filename)
    reader=PdfReader(io.BytesIO(data));pages=len(reader.pages)
    if pages==0:raise ValueError(f'{filename}: empty PDF')
    limit=max(1,min(200,int(os.getenv('PDF_CHUNK_PAGES','200'))))
    tasks=[]
    if pages<=limit:
        tasks.append({'runId':job['runId'],'source':filename,'fileIndex':job['fileIndex'],'offset':0,'pages':pages,'blobContainer':INPUT,'blobName':filename,'hash':sha256(data)})
    else:
        for offset in range(0,pages,limit):
            writer=PdfWriter()
            for p in range(offset,min(offset+limit,pages)):writer.add_page(reader.pages[p])
            out=io.BytesIO();writer.write(out);chunk=out.getvalue()
            key=f"{job['runId']}/chunks/{job['fileIndex']:02d}_{offset:06d}.pdf"
            put(WORK,key,chunk,'application/pdf')
            tasks.append({'runId':job['runId'],'source':filename,'fileIndex':job['fileIndex'],'offset':offset,'pages':min(limit,pages-offset),'blobContainer':WORK,'blobName':key,'hash':sha256(chunk)})
    return tasks

@app.activity_trigger(input_name='job')
def analyze_pdf_chunk(job:dict):
    # Deterministic output key makes activity retries idempotent.
    key=f"{job['runId']}/extracted/{job['fileIndex']:02d}_{job['offset']:06d}.jsonl"
    metadata_key=key+'.meta.json'
    from storage import service
    blob=service().get_blob_client(WORK,key)
    if blob.exists():
        return {**job,'resultKey':key,'metaKey':metadata_key,'cached':True}
    cache_key=f"di/{os.getenv('DOCUMENT_INTELLIGENCE_MODEL','prebuilt-layout')}/{os.getenv('DOCUMENT_INTELLIGENCE_API_VERSION','2024-11-30')}/{job['hash']}.json"
    cache_blob=service().get_blob_client(WORK,cache_key)
    if os.getenv('DI_CACHE_ENABLED','true').lower()=='true' and cache_blob.exists():
        result=json.loads(get(WORK,cache_key))
    else:
        result=analyze(get(job['blobContainer'],job['blobName']))
        if os.getenv('DI_CACHE_ENABLED','true').lower()=='true':put(WORK,cache_key,json.dumps(result).encode(),'application/json')
    rows,meta=extract_rows(result,job['source'],job['offset'])
    # Preserve source PDF/page offsets; chunk-local row IDs normalized in finalizer.
    put(WORK,key,b''.join((json.dumps(r,separators=(',',':'))+'\n').encode() for r in rows),'application/x-ndjson')
    put(WORK,metadata_key,json.dumps(meta).encode(),'application/json')
    return {**job,'resultKey':key,'metaKey':metadata_key,'cached':False}

@app.activity_trigger(input_name='job')
def finalize_recon(job:dict):
    run_id=job['runId'];body=job['body']
    chunks=sorted(job['chunks'],key=lambda c:(c['fileIndex'],c['offset']))
    sums=defaultdict(lambda:defaultdict(float));counts=defaultdict(int);rows=0;spools=[]
    # Canonicalize source order and row numbering, bounded by a single chunk's rows.
    for chunk in chunks:
        key=f"{run_id}/normalized/{chunk['fileIndex']:02d}_{chunk['offset']:06d}.jsonl"
        lines=[]
        for row in stream_lines(WORK,chunk['resultKey']):
            source=chunk['source'];counts[source]+=1;rows+=1
            row['Source Row']=counts[source]
            state=str(row.get('State') or '').upper();entity=str(row.get('Legal Entity') or '').upper()
            if state in body['states']:
                bucket=sums[(state,entity)]
                for f in PDF_AMOUNT_FIELDS:
                    if f in row:bucket[f]+=_num(row[f])
            lines.append((json.dumps(row,separators=(',',':'))+'\n').encode())
        put(WORK,key,b''.join(lines),'application/x-ndjson');spools.append(key)
    summary={'__aggregated__':True,'entities':[{'state':s,'entity':e,'amounts':dict(v)} for (s,e),v in sorted(sums.items())]}
    raw=parse_two_excels(get(INPUT,body['mainAccountFile']),get(INPUT,body['smwFile']))
    main,details=reconcile_two_inputs(raw,summary,body)
    if not main:raise ValueError('No Main Account Recon rows for selected states')
    output=build_output(main,details)
    output_key=f'{run_id}/state_recon_{run_id}.xlsx'
    # Stage workbook first; publish final blob only after successful database commit.
    staged_key=f'{run_id}/staging/workbook.xlsx'
    put(WORK,staged_key,output,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    persisted=persist(run_id,body,output_key,main,details,spools,rows)
    put(OUTPUT,output_key,output,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    return {'runId':run_id,'status':'COMPLETED','outputContainer':OUTPUT,'outputBlob':output_key,'pdfRows':rows,'rows':len(main),'entities':len(details),'pdfChunks':len(chunks),'inputPdfs':len(body['pdfFiles'])}
