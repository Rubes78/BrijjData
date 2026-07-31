import json
import uuid
from pathlib import Path

DATA_PATH = Path("/data/companies.json")


def _load():
    if not DATA_PATH.exists():
        return []
    return json.loads(DATA_PATH.read_text())


def _save(companies):
    DATA_PATH.parent.mkdir(parents=True, exist_ok=True)
    DATA_PATH.write_text(json.dumps(companies, indent=2))


def list_companies():
    return _load()


def get_company(company_id):
    for c in _load():
        if c["id"] == company_id:
            return c
    return None


def add_company(name, api_id, api_key, base_url):
    companies = _load()
    company = {
        "id": uuid.uuid4().hex[:12],
        "name": name,
        "api_id": api_id,
        "api_key": api_key,
        "base_url": base_url,
    }
    companies.append(company)
    _save(companies)
    return company


def delete_company(company_id):
    companies = [c for c in _load() if c["id"] != company_id]
    _save(companies)


def set_list_pricing_departments(company_id, department_codes):
    """Persist which department codes are included in the List Pricing sync.
    None means "all departments" (the default until the user narrows it down)."""
    companies = _load()
    for c in companies:
        if c["id"] == company_id:
            c["list_pricing_departments"] = department_codes
            break
    _save(companies)
