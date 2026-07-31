import io
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
DEFAULT_STORE_GROUP = "Standard"


def build_pricing_workbook(template_bytes, items, included_department_codes=None):
    """included_department_codes: None means include every department; otherwise
    only items whose category_code is in this set/list are synced."""
    included = set(included_department_codes) if included_department_codes is not None else None

    prices = set()  # (dept_name, price)
    for it in items:
        code = (it.get("category_code") or "").strip()
        if not code:
            continue
        if included is not None and code not in included:
            continue
        desc = (it.get("category_code_description") or "").strip()
        dept_name = desc if desc else code.title()
        price = it.get("price_1")
        if price is None:
            continue
        prices.add((dept_name, round(float(price), 2)))

    if not prices:
        raise NoDataError(
            "No items matched the selected departments with a price_1 set — nothing to sync."
        )

    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    if PRICE_LIST_SHEET_NAME not in wb.sheetnames:
        raise NoDataError(f'Pricing template is missing the "{PRICE_LIST_SHEET_NAME}" sheet.')
    ws = wb[PRICE_LIST_SHEET_NAME]
    _clear_data_rows(ws)

    row_num = 2
    for dept_name, price in sorted(prices):
        ws.cell(row=row_num, column=1).value = DEFAULT_STORE_GROUP
        ws.cell(row=row_num, column=2).value = dept_name
        ws.cell(row=row_num, column=3).value = price
        row_num += 1

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), len(prices)
