import io
import statistics
from collections import OrderedDict

import openpyxl


class NoDataError(Exception):
    pass


def _clear_data_rows(ws):
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.value = None


def build_department_workbook(template_bytes, items):
    depts = OrderedDict()
    for it in items:
        code = (it.get("category_code") or "").strip()
        if not code:
            continue
        desc = (it.get("category_code_description") or "").strip()
        if code not in depts or (desc and not depts[code]):
            depts[code] = desc

    if not depts:
        raise NoDataError("No items with a category_code were returned — nothing to sync.")

    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = wb.active
    _clear_data_rows(ws)

    row_num = 2
    sort_order = 10
    for code in sorted(depts):
        desc = depts[code]
        name = desc if desc else code.title()
        ws.cell(row=row_num, column=1).value = name
        ws.cell(row=row_num, column=2).value = desc or name
        ws.cell(row=row_num, column=3).value = code
        ws.cell(row=row_num, column=4).value = sort_order
        ws.cell(row=row_num, column=5).value = None  # IconFileName — not sourced from rcsaero
        ws.cell(row=row_num, column=6).value = True  # TPM
        ws.cell(row=row_num, column=7).value = True  # POS
        ws.cell(row=row_num, column=8).value = True  # Lister
        ws.cell(row=row_num, column=9).value = True  # Donation
        row_num += 1
        sort_order += 10

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), len(depts)


def build_category_workbook(template_bytes, items):
    cats = OrderedDict()  # (dept_code, sub_code) -> (sub_desc, dept_desc)
    for it in items:
        dept_code = (it.get("category_code") or "").strip()
        sub_code = (it.get("subcategory_code") or "").strip()
        if not dept_code or not sub_code:
            continue
        sub_desc = (it.get("subcategory_code_description") or "").strip()
        dept_desc = (it.get("category_code_description") or "").strip()
        key = (dept_code, sub_code)
        if key not in cats:
            cats[key] = (sub_desc, dept_desc)

    if not cats:
        raise NoDataError(
            "No items had both a category_code and subcategory_code set — "
            "there is no subcategory data in this company's rcsaero items to sync."
        )

    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = wb.active
    _clear_data_rows(ws)

    row_num = 2
    for (dept_code, sub_code), (sub_desc, dept_desc) in sorted(cats.items()):
        name = sub_desc if sub_desc else sub_code.title()
        product_department = dept_desc if dept_desc else dept_code.title()
        ws.cell(row=row_num, column=1).value = name
        ws.cell(row=row_num, column=2).value = sub_desc or name
        ws.cell(row=row_num, column=3).value = f"{dept_code}-{sub_code}"
        ws.cell(row=row_num, column=4).value = product_department
        ws.cell(row=row_num, column=5).value = None  # Category_Type — not sourced from rcsaero
        ws.cell(row=row_num, column=6).value = None  # Parent_Category
        ws.cell(row=row_num, column=7).value = "ZebraTag"
        ws.cell(row=row_num, column=8).value = None  # IconFileName — not sourced from rcsaero
        ws.cell(row=row_num, column=9).value = "Y"  # TPM
        ws.cell(row=row_num, column=10).value = "Y"  # POS
        ws.cell(row=row_num, column=11).value = "N"  # DONATION
        ws.cell(row=row_num, column=12).value = "N"  # LISTER
        row_num += 1

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), len(cats)


def list_departments(items):
    """Distinct (code, description) departments present in a set of rcsaero items,
    sorted by code. Used to drive the pricing-department selection UI."""
    depts = OrderedDict()
    for it in items:
        code = (it.get("category_code") or "").strip()
        if not code:
            continue
        desc = (it.get("category_code_description") or "").strip()
        if code not in depts or (desc and not depts[code]):
            depts[code] = desc
    return [{"code": code, "description": depts[code] or code.title()} for code in sorted(depts)]


PRICE_LIST_SHEET_NAME = "Price Lists by Department"
GBB_SHEET_NAME = "GBB by Department"
QUALITY_CONDITION_SHEET_NAME = "Quality and Condition by Depart"
DEFAULT_STORE_GROUP = "Standard"

MODEL_LIST = "list"
MODEL_GBB = "gbb"
MODEL_QUALITY_CONDITION = "quality_condition"
VALID_MODELS = {MODEL_LIST, MODEL_GBB, MODEL_QUALITY_CONDITION}


def _collect_department_prices(items):
    """Returns {category_code: {"name": dept_name, "prices": [price, ...]}}
    for items with a price_1 set."""
    result = OrderedDict()
    for it in items:
        code = (it.get("category_code") or "").strip()
        if not code:
            continue
        price = it.get("price_1")
        if price is None:
            continue
        desc = (it.get("category_code_description") or "").strip()
        name = desc if desc else code.title()
        entry = result.setdefault(code, {"name": name, "prices": []})
        entry["prices"].append(float(price))
    return result


def _bucket_groups(prices, k):
    """Split a sorted list of prices into k equal-count buckets (quantile-style)
    and return the k price-lists in ascending order. Buckets can be empty if
    there are fewer observations than k."""
    prices = sorted(prices)
    n = len(prices)
    sizes = [n // k + (1 if i < n % k else 0) for i in range(k)]

    groups = []
    idx = 0
    for size in sizes:
        groups.append(prices[idx:idx + size])
        idx += size
    return groups


def _bucket_medians(prices, levels):
    """Split a sorted list of prices into `levels` equal-count buckets and
    return the median of each bucket, ascending. Empty buckets (fewer price
    observations than levels) come back as None."""
    return [round(statistics.median(g), 2) if g else None for g in _bucket_groups(prices, levels)]


def _list_pricing_prices(prices):
    """Every distinct price point, ascending — the List Pricing model has no tiers."""
    return sorted({round(p, 2) for p in prices})


def _gbb_tiers(prices, labels):
    """[(label, price), ...] — one row per non-empty equal-count bucket."""
    medians = _bucket_medians(prices, len(labels))
    return [(label, price) for label, price in zip(labels, medians) if price is not None]


def _quality_condition_tiers(prices, quality_labels, condition_labels):
    """[(quality_label, condition_label, price), ...]. Prices are bucketed into
    Quality tiers first, then each Quality bucket is further bucketed into
    Condition tiers — mirroring the ascending Quality-then-Condition price
    progression already used in the hand-built Quality and Condition by Category sheet."""
    rows = []
    for quality_label, quality_group in zip(quality_labels, _bucket_groups(prices, len(quality_labels))):
        if not quality_group:
            continue
        for condition_label, price in zip(condition_labels, _bucket_medians(quality_group, len(condition_labels))):
            if price is not None:
                rows.append((quality_label, condition_label, price))
    return rows


def build_combined_pricing_workbook(
    template_bytes, items, department_models, gbb_labels, condition_labels
):
    """department_models: {category_code: "list" | "gbb" | "quality_condition"}.
    Each department gets exactly one pricing model — its rows go into exactly one
    of the three sheets below. Departments absent from department_models (or mapped
    to anything else) are skipped entirely, on every sheet."""
    dept_data = _collect_department_prices(items)

    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    sheets = {}
    for name in (PRICE_LIST_SHEET_NAME, GBB_SHEET_NAME, QUALITY_CONDITION_SHEET_NAME):
        if name not in wb.sheetnames:
            raise NoDataError(f'Pricing template is missing the "{name}" sheet.')
        sheets[name] = wb[name]
        _clear_data_rows(sheets[name])

    next_row = {PRICE_LIST_SHEET_NAME: 2, GBB_SHEET_NAME: 2, QUALITY_CONDITION_SHEET_NAME: 2}
    row_count = 0

    for code in sorted(dept_data):
        model = department_models.get(code)
        if model not in VALID_MODELS:
            continue
        name = dept_data[code]["name"]
        prices = dept_data[code]["prices"]

        if model == MODEL_LIST:
            ws = sheets[PRICE_LIST_SHEET_NAME]
            for price in _list_pricing_prices(prices):
                r = next_row[PRICE_LIST_SHEET_NAME]
                ws.cell(row=r, column=1).value = DEFAULT_STORE_GROUP
                ws.cell(row=r, column=2).value = name
                ws.cell(row=r, column=3).value = price
                next_row[PRICE_LIST_SHEET_NAME] += 1
                row_count += 1

        elif model == MODEL_GBB:
            ws = sheets[GBB_SHEET_NAME]
            for label, price in _gbb_tiers(prices, gbb_labels):
                r = next_row[GBB_SHEET_NAME]
                ws.cell(row=r, column=1).value = DEFAULT_STORE_GROUP
                ws.cell(row=r, column=2).value = name
                ws.cell(row=r, column=3).value = label
                ws.cell(row=r, column=4).value = price
                next_row[GBB_SHEET_NAME] += 1
                row_count += 1

        else:  # MODEL_QUALITY_CONDITION
            ws = sheets[QUALITY_CONDITION_SHEET_NAME]
            for quality_label, condition_label, price in _quality_condition_tiers(
                prices, gbb_labels, condition_labels
            ):
                r = next_row[QUALITY_CONDITION_SHEET_NAME]
                ws.cell(row=r, column=1).value = DEFAULT_STORE_GROUP
                ws.cell(row=r, column=2).value = name
                ws.cell(row=r, column=3).value = quality_label
                ws.cell(row=r, column=4).value = condition_label
                ws.cell(row=r, column=5).value = price
                next_row[QUALITY_CONDITION_SHEET_NAME] += 1
                row_count += 1

    if row_count == 0:
        raise NoDataError(
            "No departments were assigned a pricing model (or none had enough price data) — nothing to sync."
        )

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), row_count
