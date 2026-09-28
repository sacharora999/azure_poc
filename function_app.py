import azure.functions as func
import os, json, io, uuid, logging
from azure.storage.blob import BlobServiceClient
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient
import psycopg
from processor import parse_excel,reconcile,build_output

app=func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

def blob_service(): return BlobServiceClient.from_connection_string(os.environ['AzureWebJobsStorage'])
def download(container,name): return blob_service().get_blob_client(container,name).download_blob().readall()
def upload(container,name,data,ctype):
    bc=blob_service().get_blob_client(container,name); bc.upload_blob(data,overwrite=True,content_type=ctype); return bc.url

def extract_pdf(pdf_bytes):
    client=DocumentIntelligenceClient(os.environ['DOCINTEL_ENDPOINT'],AzureKeyCredential(os.environ['DOCINTEL_KEY']))
    result=client.begin_analyze_document('prebuilt-layout',body=pdf_bytes).result()
    page_entity={}
    for page in result.pages or []:
        text=' '.join(x.content for x in (page.lines or []))
        page_entity[page.page_number]=next((e for e in ['ABC','XYZ','LMN'] if f'Legal Entity: {e}' in text),None)
    out=[]
    for table in result.tables or []:
        grid={}
        for c in table.cells: grid[(c.row_index,c.column_index)]=c.content
        headers=[grid.get((0,j),'').strip() for j in range(table.column_count)]
        page_no=(table.bounding_regions[0].page_number if table.bounding_regions else None)
        ent=page_entity.get(page_no)
        for i in range(1,table.row_count):
            row={headers[j]:grid.get((i,j),'') for j in range(table.column_count) if headers[j]}
            if row.get('GL Account'):
                row['Legal Entity']=ent
                out.append(row)
    return out

def persist(run_id,user_inputs,main,details,output_url):
    dsn=os.environ.get('POSTGRES_DSN')
    if not dsn: return
    with psycopg.connect(dsn) as conn:
        with conn.cursor() as cur:
            cur.execute('INSERT INTO recon_run(run_id,state,period,status,output_url,request_json) VALUES(%s,%s,%s,%s,%s,%s::jsonb)',(run_id,user_inputs.get('state'),user_inputs.get('period'),'COMPLETED',output_url,json.dumps(user_inputs)))
            for r in main:
                cur.execute('INSERT INTO recon_account(run_id,account_no,legal_entity,state,ending_balance,payload) VALUES(%s,%s,%s,%s,%s,%s::jsonb)',(run_id,str(r.get('Account #')),r.get('LegEnt'),r.get('State'),r.get('Ending Balance'),json.dumps(r,default=str)))
        conn.commit()

@app.route(route='state-recon',methods=['POST'])
def state_recon(req: func.HttpRequest) -> func.HttpResponse:
    try:
        body=req.get_json(); run_id=str(uuid.uuid4())
        container=body.get('container','recon-input')
        main_name=body['mainAccountFile']; id_name=body['idSampleFile']; pdf_name=body['pdfFile']
        user_inputs={'state':body.get('state','ID'),'period':body.get('period','July 2026'),'requestedBy':body.get('requestedBy','PowerApps')}
        raw=parse_excel(download(container,main_name),download(container,id_name))
        pdf_rows=extract_pdf(download(container,pdf_name))
        main,details=reconcile(raw,pdf_rows,user_inputs)
        output=build_output(main,details); out_name=f"state_recon_{run_id}.xlsx"
        output_url=upload(os.environ.get('OUTPUT_CONTAINER','recon-output'),out_name,output,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        persist(run_id,user_inputs,main,details,output_url)
        return func.HttpResponse(json.dumps({'runId':run_id,'status':'COMPLETED','outputFile':out_name,'outputUrl':output_url,'rows':len(main)}),mimetype='application/json')
    except Exception as e:
        logging.exception('state-recon failed')
        return func.HttpResponse(json.dumps({'status':'FAILED','error':str(e)}),status_code=500,mimetype='application/json')
