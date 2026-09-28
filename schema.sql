CREATE TABLE IF NOT EXISTS recon_run (
 run_id uuid PRIMARY KEY, created_at timestamptz NOT NULL DEFAULT now(), state varchar(10), period varchar(50),
 status varchar(20) NOT NULL, output_url text, request_json jsonb NOT NULL
);
CREATE TABLE IF NOT EXISTS recon_account (
 id bigserial PRIMARY KEY, run_id uuid NOT NULL REFERENCES recon_run(run_id), account_no varchar(50),
 legal_entity varchar(30), state varchar(10), ending_balance numeric(18,2), payload jsonb NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_recon_account_run ON recon_account(run_id);
CREATE INDEX IF NOT EXISTS ix_recon_account_entity ON recon_account(legal_entity,state);
