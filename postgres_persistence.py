"""Transactional PostgreSQL writes with password authentication and idempotent run keys."""
import os,json
from decimal import Decimal
import psycopg
from storage import stream_lines

def connection():
    password=os.environ['POSTGRES_PASSWORD']
    return psycopg.connect(host=os.environ['POSTGRES_HOST'],port=int(os.getenv('POSTGRES_PORT','5432')),
        dbname=os.environ['POSTGRES_DB'],user=os.environ['POSTGRES_USER'],password=password,
        sslmode='require',connect_timeout=20)

def persist(run_id,body,output_key,main,details,spool_keys,row_count):
    batch_size=max(100,min(5000,int(os.getenv('POSTGRES_BATCH_SIZE','1000'))))
    def batch_insert(cur,sql,items):
        batch=[]
        for item in items:
            batch.append(item)
            if len(batch)>=batch_size:
                cur.executemany(sql,batch);batch.clear()
        if batch:cur.executemany(sql,batch)
    with connection() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT pg_advisory_xact_lock(hashtext(%s))',(run_id,))
            for table in ('recon_pdf_detail','recon_report','recon_entity_summary'):
                cur.execute(f'DELETE FROM {table} WHERE run_id=%s',(run_id,))
            cur.execute('''INSERT INTO recon_run(run_id,state,period,requested_by,status,output_blob,row_count,entity_count,request_json)
                VALUES(%s,%s,%s,%s,'COMPLETED',%s,%s,%s,%s::jsonb)
                ON CONFLICT(run_id) DO UPDATE SET status=EXCLUDED.status,output_blob=EXCLUDED.output_blob,
                row_count=EXCLUDED.row_count,entity_count=EXCLUDED.entity_count,request_json=EXCLUDED.request_json,
                updated_at=now()''',(run_id,','.join(body['states']),body['period'],body.get('requestedBy'),output_key,len(main),len(details),json.dumps(body)))
            def entities():
                for entity,d in details.items():
                    m=d.get('metrics',{})
                    yield (run_id,entity,d.get('state'),body['period'],m.get('Extract Net Sales',0),m.get('SMW Sales',0),m.get('SMW Tax',0),m.get('Vertex Sales',0),m.get('Vertex Use Tax Sales',0),m.get('Vertex Tax',0))
            batch_insert(cur,'''INSERT INTO recon_entity_summary(run_id,legal_entity,state,period,extract_sales,smw_sales,smw_tax,vertex_sales,vertex_use_tax_sales,vertex_tax)
                VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)''',entities())
            keys=['Beginning Balance','Lag','JEs & Payments','Open Item','Add.Entity Activity','Actual Open Items','Extract','Lag Ledger Activity','Ledger Activity','AE Tracking','Ending Balance']
            def reports():
                for r in main:
                    yield (run_id,r.get('State'),body['period'],str(r.get('LegEnt') or ''),r.get('Account Description'),str(r.get('Account #') or ''),*[float(r.get(k) or 0) for k in keys])
            batch_insert(cur,'''INSERT INTO recon_report(run_id,state,period,legal_entity,account_description,account_number,beginning_balance,lag,jes_payments,open_item,additional_entity_activity,actual_open_items,extract,lag_ledger_activity,ledger_activity,ae_tracking,ending_balance)
                VALUES ('''+','.join(['%s']*17)+')',reports())
            def pdf_details():
                for key in spool_keys:
                    for r in stream_lines(os.getenv('RECON_WORK_CONTAINER','recon-work'),key):
                        yield (run_id,r.get('Source PDF'),r.get('Source Page'),r.get('Source Row'),r.get('Legal Entity'),r.get('State'),r.get('Net Sales',0),r.get('Sales Tax',0),r.get('Use Tax',0),r.get('Net Invoice Amount',0),r.get('Invoice Amount',0),json.dumps({k:v for k,v in r.items() if k not in ('Source PDF','Source Page','Source Row','Legal Entity','State')}))
            batch_insert(cur,'''INSERT INTO recon_pdf_detail(run_id,source_pdf,source_page,source_row,legal_entity,state,net_sales,sales_tax,use_tax,net_invoice_amount,invoice_amount,amounts_json)
                VALUES('''+','.join(['%s']*12)+'''::jsonb)''',pdf_details())
    return {'rows':len(main),'entities':len(details),'pdfRows':row_count}
