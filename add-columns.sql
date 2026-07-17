-- OrderFlow — new fields (run once in Supabase ▸ SQL Editor). Safe to re-run.
-- 1) Gate Entry: driver mobile number
alter table "GateEntry"  add column if not exists "MobileNo" text;

-- 2) Gate Out: weighment verification
alter table "GateOut"    add column if not exists "WeighmentDone" text;

-- 3+4) Collection: deduction, actual received, and weights
alter table "Collection" add column if not exists "DeductionAmount" numeric;
alter table "Collection" add column if not exists "ActualReceived"  numeric;
alter table "Collection" add column if not exists "GrossWeight"     numeric;
alter table "Collection" add column if not exists "NetWeight"       numeric;
alter table "Collection" add column if not exists "PartyNetWeight"  numeric;

-- Tell the API about the new columns (important!)
NOTIFY pgrst, 'reload schema';
