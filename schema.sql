-- Apply once via controlled migration from a private-network administration host.
CREATE TABLE IF NOT EXISTS recon_run (
 run_id uuid PRIMARY KEY, state text, period text, requested_by text, status text NOT NULL,
 output_blob text, row_count integer, entity_count integer, error_message text,
 request_json jsonb, created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now());
CREATE TABLE IF NOT EXISTS recon_entity_summary (
 run_id uuid NOT NULL REFERENCES recon_run(run_id), legal_entity text NOT NULL, state text, period text,
 extract_sales numeric(20,2), smw_sales numeric(20,2), smw_tax numeric(20,2),
 vertex_sales numeric(20,2), vertex_use_tax_sales numeric(20,2), vertex_tax numeric(20,2),
 PRIMARY KEY(run_id,legal_entity));
CREATE TABLE IF NOT EXISTS recon_report (
 id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY, run_id uuid NOT NULL REFERENCES recon_run(run_id),
 state text,period text,legal_entity text,account_description text,account_number text,
 beginning_balance numeric(20,2),lag numeric(20,2),jes_payments numeric(20,2),open_item numeric(20,2),
 additional_entity_activity numeric(20,2),actual_open_items numeric(20,2),extract numeric(20,2),
 lag_ledger_activity numeric(20,2),ledger_activity numeric(20,2),ae_tracking numeric(20,2),ending_balance numeric(20,2));
CREATE INDEX IF NOT EXISTS idx_recon_report_period_state_entity ON recon_report(period,state,legal_entity);
CREATE INDEX IF NOT EXISTS idx_recon_report_run ON recon_report(run_id);
CREATE TABLE IF NOT EXISTS recon_pdf_detail (
 id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY, run_id uuid NOT NULL REFERENCES recon_run(run_id),
 source_pdf text,source_page integer,source_row integer,legal_entity text,state text,
 net_sales numeric(20,2),sales_tax numeric(20,2),use_tax numeric(20,2),
 net_invoice_amount numeric(20,2),invoice_amount numeric(20,2),amounts_json jsonb);
CREATE INDEX IF NOT EXISTS idx_recon_pdf_run_entity_state ON recon_pdf_detail(run_id,legal_entity,state);
CREATE INDEX IF NOT EXISTS idx_recon_pdf_run_source_page ON recon_pdf_detail(run_id,source_pdf,source_page);
