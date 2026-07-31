import io
import os
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, send_file, jsonify

import companies
from rcsaero_client import fetch_all_items, RcsAeroError
from sheet_builder import (
    build_department_workbook,
    build_category_workbook,
    build_combined_pricing_workbook,
    list_departments,
    NoDataError,
    VALID_MODELS,
)

app = Flask(__name__)

DEPT_TEMPLATE_PATH = Path(__file__).parent / "Department_Template.xlsx"
CAT_TEMPLATE_PATH = Path(__file__).parent / "Category_Template.xlsx"
PRICING_TEMPLATE_PATH = Path(__file__).parent / "Pricing_Template.xlsx"


@app.route("/")
def index():
    return render_template("index.html", companies=companies.list_companies())


@app.route("/companies", methods=["GET"])
def list_companies():
    return jsonify(companies.list_companies())


@app.route("/companies", methods=["POST"])
def add_company():
    data = request.get_json(force=True)
    name = (data.get("name") or "").strip()
    api_id = (data.get("api_id") or "").strip()
    api_key = (data.get("api_key") or "").strip()
    base_url = (data.get("base_url") or "https://api.rcsaero.com").strip()

    if not name or not api_id or not api_key:
        return jsonify({"error": "Name, API ID, and API Key are all required."}), 400

    company = companies.add_company(name, api_id, api_key, base_url)
    return jsonify({k: v for k, v in company.items() if k != "api_key"}), 201


@app.route("/companies/<company_id>", methods=["DELETE"])
def remove_company(company_id):
    companies.delete_company(company_id)
    return "", 204


def _effective_department_models(company):
    """{category_code: model}. If the company has never configured this, every
    department defaults to List Pricing (the original all-List behavior) —
    except when an old-style list_pricing_departments boolean list exists from
    before per-department models existed, which is honored as a one-time
    migration default (those codes -> "list", everything else -> excluded)."""
    stored = company.get("department_pricing_models")
    if stored is not None:
        return stored

    legacy_included = company.get("list_pricing_departments")
    if legacy_included is not None:
        return {code: "list" for code in legacy_included}

    return None  # signals "default every department to list" to callers


@app.route("/companies/<company_id>/departments", methods=["GET"])
def get_departments(company_id):
    company = companies.get_company(company_id)
    if not company:
        return jsonify({"error": "Unknown company."}), 404

    try:
        items = fetch_all_items(company["api_id"], company["api_key"], company["base_url"])
    except RcsAeroError as e:
        return jsonify({"error": str(e)}), 502

    depts = list_departments(items)
    models = _effective_department_models(company)
    for d in depts:
        d["model"] = "list" if models is None else models.get(d["code"], "none")

    return jsonify(depts)


@app.route("/companies/<company_id>/department-models", methods=["POST"])
def set_department_models(company_id):
    company = companies.get_company(company_id)
    if not company:
        return jsonify({"error": "Unknown company."}), 404

    data = request.get_json(force=True)
    models = data.get("models")
    if not isinstance(models, dict) or not all(
        isinstance(k, str) and v in VALID_MODELS for k, v in models.items()
    ):
        return jsonify({
            "error": f"'models' must be a dict of department code -> one of {sorted(VALID_MODELS)}."
        }), 400

    companies.set_department_pricing_models(company_id, models)
    return "", 204


@app.route("/companies/<company_id>/gbb-labels", methods=["GET"])
def get_gbb_labels(company_id):
    company = companies.get_company(company_id)
    if not company:
        return jsonify({"error": "Unknown company."}), 404
    return jsonify({"labels": company.get("gbb_labels", companies.DEFAULT_GBB_LABELS)})


@app.route("/companies/<company_id>/gbb-labels", methods=["POST"])
def set_gbb_labels(company_id):
    company = companies.get_company(company_id)
    if not company:
        return jsonify({"error": "Unknown company."}), 404

    data = request.get_json(force=True)
    labels = data.get("labels")
    if not isinstance(labels, list) or len(labels) == 0 or not all(
        isinstance(l, str) and l.strip() for l in labels
    ):
        return jsonify({"error": "'labels' must be a non-empty list of non-blank tier names."}), 400

    companies.set_gbb_labels(company_id, [l.strip() for l in labels])
    return "", 204


@app.route("/companies/<company_id>/condition-labels", methods=["GET"])
def get_condition_labels(company_id):
    company = companies.get_company(company_id)
    if not company:
        return jsonify({"error": "Unknown company."}), 404
    return jsonify({"labels": company.get("condition_labels", companies.DEFAULT_CONDITION_LABELS)})


@app.route("/companies/<company_id>/condition-labels", methods=["POST"])
def set_condition_labels(company_id):
    company = companies.get_company(company_id)
    if not company:
        return jsonify({"error": "Unknown company."}), 404

    data = request.get_json(force=True)
    labels = data.get("labels")
    if not isinstance(labels, list) or len(labels) == 0 or not all(
        isinstance(l, str) and l.strip() for l in labels
    ):
        return jsonify({"error": "'labels' must be a non-empty list of non-blank tier names."}), 400

    companies.set_condition_labels(company_id, [l.strip() for l in labels])
    return "", 204


def _sync(company_id, kind):
    company = companies.get_company(company_id)
    if not company:
        return jsonify({"error": "Unknown company."}), 404

    try:
        items = fetch_all_items(company["api_id"], company["api_key"], company["base_url"])
    except RcsAeroError as e:
        return jsonify({"error": str(e)}), 502

    if kind == "department":
        template_bytes = DEPT_TEMPLATE_PATH.read_bytes()
        build_fn = build_department_workbook
        extra_kwargs = {}
        label = "Department"
    elif kind == "category":
        template_bytes = CAT_TEMPLATE_PATH.read_bytes()
        build_fn = build_category_workbook
        extra_kwargs = {}
        label = "Category"
    else:
        template_bytes = PRICING_TEMPLATE_PATH.read_bytes()
        build_fn = build_combined_pricing_workbook
        models = _effective_department_models(company)
        if models is None:
            depts = list_departments(items)
            models = {d["code"]: "list" for d in depts}
        extra_kwargs = {
            "department_models": models,
            "gbb_labels": company.get("gbb_labels", companies.DEFAULT_GBB_LABELS),
            "condition_labels": company.get("condition_labels", companies.DEFAULT_CONDITION_LABELS),
        }
        label = "Pricing"

    try:
        out_bytes, row_count = build_fn(template_bytes, items, **extra_kwargs)
    except NoDataError as e:
        return jsonify({"error": str(e)}), 422

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = "".join(c if c.isalnum() else "_" for c in company["name"])
    filename = f"{safe_name}_{label}_Sample_Sheet_{ts}.xlsx"

    return send_file(
        io.BytesIO(out_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    ), 200, {"X-Row-Count": str(row_count)}


@app.route("/sync/department/<company_id>", methods=["POST"])
def sync_department(company_id):
    return _sync(company_id, "department")


@app.route("/sync/category/<company_id>", methods=["POST"])
def sync_category(company_id):
    return _sync(company_id, "category")


@app.route("/sync/pricing/<company_id>", methods=["POST"])
def sync_pricing(company_id):
    return _sync(company_id, "pricing")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
