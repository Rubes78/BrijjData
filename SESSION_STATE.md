# Session State — resume point (2026-07-31)

Working notes for picking this project back up from a different machine/session. Written because Claude's cross-session memory lives in `~/.claude` on the machine where it ran, not in this repo — this file is the portable version of that.

## Status: deployed, working, user-confirmed good

- Live at `http://192.168.1.215:5601` (HoloDev LXC), container name `brijj-tpm-sync`, port 5601 → container 5000.
- Source of truth: `/Quarks/Data/DevOps/claude/BrijjTPMSync/` on HoloDev. Git repo, pushed to `git@github.com:Rubes78/BrijjData.git`, branch `main`, currently at commit `00c412e`.
- Data volume (company credentials + all persisted config): `/Quarks/Data/Dockers/brijj-tpm-sync/companies.json` — NOT in git, lives outside the container.
- A real company, `DataStreamDemo`, is already configured in that companies.json and in active use.

## To redeploy / verify after a fresh clone or pull

```bash
cd /Quarks/Data/DevOps/claude/BrijjTPMSync
docker compose up -d --build
docker ps --filter name=brijj-tpm-sync   # should show Up, 0.0.0.0:5601->5000/tcp
```

No `.env` or secrets needed in the repo — companies (including their rcsaero API credentials) are entered through the page itself and persist in the data volume, which is untouched by rebuilds.

## What this tool does

Flask app that pulls live item data from a customer's rcsaero API (`/api/items`, auth via `x-api-id`/`x-api-key` headers) and generates fresh, import-ready TPM spreadsheets for Brijj POS onboarding. Multi-tenant — each customer ("company" in the UI) has its own saved credentials.

Three buttons on the page:
1. **Download Department Data** — one row per rcsaero `category_code` → Brijj Department.
2. **Download Category Data** — one row per `(category_code, subcategory_code)` pair → Brijj Category. Only produces rows for items that actually have both set.
3. **Download Pricing Data** — the interesting one, see below.

## Pricing sync architecture (the bulk of this session's work)

One download covers all pricing, not three. Every department AND every category gets assigned exactly one pricing model via a `<select>` in the UI: **None / List Pricing / GBB / Quality & Condition**. The single downloaded workbook has all 6 pricing sheets populated at once, each entity's rows landing in exactly one sheet:

| Level | Sheet used for List | Sheet used for GBB | Sheet used for Quality & Condition |
|---|---|---|---|
| Department | Price Lists by Department | GBB by Department | Quality and Condition by Depart |
| Category | Price Lists by Category and Sub | GBB by Category and SubCategory | Quality and Condition by Catego |

Storage (in `companies.json`, per company):
- `department_pricing_models`: `{category_code: "list"|"gbb"|"quality_condition"}`
- `category_pricing_models`: `{"{category_code}-{subcategory_code}": "list"|"gbb"|"quality_condition"}`
- `gbb_labels`: ordered tier names, e.g. `["Good","Better","Best"]` — user-defined, not just a count. Doubles as the Quality axis for Quality & Condition.
- `condition_labels`: ordered tier names, e.g. `["Distressed","Average","Like New"]`.

Price derivation (no manual per-department price entry anywhere):
- **List**: every distinct `price_1` value seen for that entity, ascending.
- **GBB**: entity's prices split into `len(gbb_labels)` equal-count buckets (quantiles, not equal-dollar-width), median of each bucket = that tier's price.
- **Quality & Condition**: prices split into Quality buckets first, then *each Quality bucket* is further split into Condition buckets (nested quantile split) — mirrors the ascending Quality-then-Condition price progression already used in the hand-authored template sheet.

Sparse entities (fewer distinct prices than tier count) gracefully produce fewer rows instead of erroring.

Key files: `app.py` (routes), `sheet_builder.py` (all the xlsx-building logic — `build_combined_pricing_workbook` is the core), `companies.py` (JSON persistence), `rcsaero_client.py` (paginated API fetch).

## Known gaps (not bugs, just not asked for yet)

- Category pricing will look sparse until the rcsaero test data ("Whiskered Thrifter" / `DataStreamDemo` company) has more items with `subcategory_code` populated — currently only 2 of 410 items have one set. User said they'd fix their test data.
- No cross-level exclusivity: a department and one of its own categories could theoretically both get a pricing model assigned. Only same-level exclusivity ("a department can only be in one pricing model") was actually requested and built.
- Company credentials are stored as plaintext JSON on the data volume. Acceptable for this LAN-only homelab tool; not hardened further.

## A bug that already happened once — worth knowing about

Editing the inline `<script>` in `templates/index.html` once introduced a JS temporal-dead-zone `ReferenceError` (calling a function that used a `const` declared later in the file) that silently killed the *entire* script — every button stopped working, not just the new one, and backend curl tests didn't catch it (they all passed). If something in the UI stops responding after an edit, check load order: the final `renderCompanies();` call at the bottom of the script must come after every `const`/function it (transitively) depends on. A quick way to verify a change didn't break this: curl the rendered page, extract the `<script>` body, and run it under Node with a minimal stubbed `document`/`fetch` — a real ReferenceError will throw immediately.
