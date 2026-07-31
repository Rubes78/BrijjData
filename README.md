# Brijj TPM Sync

Small internal tool for onboarding Brijj POS customers off the rcsaero API. For a selected company, pulls live item data from `rcsaero`'s `/api/items` endpoint and generates fresh, import-ready TPM spreadsheets:

- **Department_Sample_Sheet.xlsx** — one row per distinct `category_code`
- **Category_Sample_Sheet.xlsx** — one row per distinct `(category_code, subcategory_code)` pair
- **Pricing_Import_Template.xlsx** — `Price Lists by Department` sheet, one row per distinct `(department, price_1)` pair

Multi-tenant: each company's rcsaero API credentials are saved once and reused, so onboarding a new customer is just "add company → click download" rather than re-running scripts by hand.

## Stack

- Flask + openpyxl + requests
- No database — companies (name + rcsaero API ID/Key + base URL) are stored as plaintext JSON on a mounted volume

## Running locally

```bash
pip install -r requirements.txt
PORT=5000 python app.py
```

Companies are stored at `/data/companies.json` — set that path to somewhere writable locally, or just `mkdir /data` if running as root.

## Production

```bash
docker compose up -d --build
```

Deployed on HoloDev (192.168.1.215:5601). Data volume: `/Quarks/Data/Dockers/brijj-tpm-sync/companies.json`.

## Field mapping (rcsaero → Brijj)

| rcsaero | Brijj |
|---|---|
| `category_code` / `category_code_description` | Department |
| `subcategory_code` / `subcategory_code_description` | Category |
| `price_1` | Price |

Notes:
- rcsaero auth is two static headers, `x-api-id` and `x-api-key` — not a bearer token / sign-in flow.
- `profile_code_5` on rcsaero items is a **color** field, not quality/condition — don't repurpose it.
- The Category and Pricing sheets only include what rcsaero actually has data for — no Quality/Condition/Store Group fields exist in the source system, so those columns are left blank rather than guessed.
- Downloads never overwrite the original templates on disk — every sync generates a fresh in-memory file.
