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


def set_department_pricing_models(company_id, department_models):
    """department_models: {category_code: "list" | "gbb" | "quality_condition"}.
    A department not present here gets no pricing model at all — it's excluded
    entirely from the combined pricing sync. Each department maps to exactly one
    model; there's no separate per-model inclusion list anymore."""
    companies = _load()
    for c in companies:
        if c["id"] == company_id:
            c["department_pricing_models"] = department_models
            break
    _save(companies)


def set_category_pricing_models(company_id, category_models):
    """category_models: {"{category_code}-{subcategory_code}": "list" | "gbb" | "quality_condition"}.
    Same shape and semantics as department_pricing_models, one level finer-grained."""
    companies = _load()
    for c in companies:
        if c["id"] == company_id:
            c["category_pricing_models"] = category_models
            break
    _save(companies)


DEFAULT_GBB_LABELS = ["Good", "Better", "Best"]


def set_gbb_labels(company_id, labels):
    """labels: ordered list of tier names, lowest price first. Its length is
    the number of GBB levels — there's no separate "levels" count to keep in sync.
    Also used as the Quality axis for Quality and Condition by Department, so
    there's one place to name Good/Better/Best-style tiers, not two."""
    companies = _load()
    for c in companies:
        if c["id"] == company_id:
            c["gbb_labels"] = labels
            break
    _save(companies)


DEFAULT_CONDITION_LABELS = ["Distressed", "Average", "Like New"]


def set_condition_labels(company_id, labels):
    companies = _load()
    for c in companies:
        if c["id"] == company_id:
            c["condition_labels"] = labels
            break
    _save(companies)
