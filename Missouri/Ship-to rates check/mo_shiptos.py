"""Step two reference: check the ship-to list against the state's files.

Each active ship-to is placed by street and city (the list has no zip),
Address 2 first, then Address 1, for this quarter and last quarter. Its Tax
(the rate the order system charges now) is compared with the state's use
rate. The rules, set Oct 1, 2026:

- NULL means empty, in every column.
- Inactive (customer or ship-to status) is skipped and counted.
- Exempt (TaxExemption not empty) is still checked, listed last.
- Both address lines are tried every time, Address 2 first. One places: use
  it. Both place on the same code: fine. Both place on different codes: map.
  Neither places: map.
- Tax is compared with the state's use rate for the code.
- Placed means street-exact (EXACT or CLOSE). A street-only guess (the house
  number isn't in the state's ranges) goes to the map, its likely code in
  the note, as the plan said for anything not placed street-exact.

Defaults applied, not yet okayed: an active, non-exempt ship-to with Tax empty
is flagged as no rate set up; Tax matching the sales rate is said so; the
list's county is a cross-check (a different county goes to the map).

This Python copy is the test oracle for the Access copy (test_access.py).
"""
import csv
import openpyxl
from mo_match import county_code

COLUMNS = {"acct": "acct", "shipto": "ship_to", "customername": "customer", "shiptoname": "name",
           "address1": "address1", "address2": "address2", "city": "city", "state": "state",
           "county": "county", "tax": "tax", "taxexemption": "exemption", "producttypes": "products",
           "customerstatus": "customer_status", "shiptostatus": "ship_to_status"}
NEEDED = ("acct", "ship_to", "address1", "address2", "city")
PLACED = ("EXACT", "CLOSE")  # street-exact; a street-only guess goes to the map


def header_key(h):
    """'Ship-To #' -> 'shipto'."""
    return "".join(c for c in str(h or "").lower() if "a" <= c <= "z" or "0" <= c <= "9")


def cell(v):
    """A cell as text; NULL means empty."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return "" if s.upper() == "NULL" else s


def tax_value(v):
    """0.08475, '0.08475', '8.475%' or 8.475 -> 0.08475; empty -> None."""
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        if s == "" or s.upper() == "NULL":
            return None
        pct = s.endswith("%")
        try:
            x = float(s.rstrip("%"))
        except ValueError:
            return None
        if pct:
            x /= 100
    else:
        x = float(v)
    if x > 0.5:  # written as a percent number (8.475)
        x /= 100
    return round(x, 6)


def read_list(path):
    """The ship-to list (.xlsx or .csv) -> list of dicts. The header row is the
    first of the top ten rows that names both City and Address 1."""
    path = str(path)
    if path.lower().endswith(".csv"):
        with open(path, newline="", encoding="latin-1") as f:
            grid = [row for row in csv.reader(f)]
    else:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        grid = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    head = None
    for i, row in enumerate(grid[:10]):
        keys = [header_key(h) for h in row]
        if "city" in keys and "address1" in keys:
            head = i
            break
    if head is None:
        raise ValueError("no header row with City and Address 1")
    cols = {}
    for j, h in enumerate(grid[head]):
        k = header_key(h)
        if k in COLUMNS and COLUMNS[k] not in cols:
            cols[COLUMNS[k]] = j
    out = []
    for row in grid[head + 1:]:
        rec = {}
        for name, j in cols.items():
            v = row[j] if j < len(row) else None
            rec[name] = tax_value(v) if name == "tax" else cell(v)
        for name in COLUMNS.values():
            rec.setdefault(name, None if name == "tax" else "")
        if rec["acct"] == "" and rec["ship_to"] == "":
            continue
        out.append(rec)
    return out


def combine(b, rec):
    """Both address lines, Address 2 first -> (how, code, note, where, line)."""
    tried = []
    for line_name, line in (("Address 2", rec["address2"]), ("Address 1", rec["address1"])):
        if line.strip():
            tried.append((line_name,) + b.match_ship_to(line, rec["city"], rec["county"]))
    placed = [t for t in tried if t[1] in PLACED]
    if len(placed) == 2:
        if placed[0][2] == placed[1][2]:
            ln, how, code, note, where = placed[0]
            return how, code, note, where, "both"
        return ("MIXED", "", f"Address 2 is in {placed[0][2]} but Address 1 is in {placed[1][2]}", "", "")
    if len(placed) == 1:
        ln, how, code, note, where = placed[0]
        return how, code, note, where, ln
    if not tried:
        return "NONE", "", "no address", "", ""
    mixed = [t for t in tried if t[1] == "MIXED"]
    ln, how, code, note, where = (mixed or tried)[0]
    return how, "", note, where, ""


def check(rec, b_now, b_last, rates_now, rates_last, q_start="Oct 1"):
    """One ship-to -> its result. rates_* map code -> (sales, use)."""
    out = {"result": "", "why": "", "code_now": "", "how_now": "", "where_now": "", "line_now": "",
           "use_now": None, "code_last": "", "use_last": None,
           "exempt": rec["exemption"] != ""}
    if rec["customer_status"].upper() == "INACTIVE" or rec["ship_to_status"].upper() == "INACTIVE":
        out.update(result="Skipped", why="inactive")
        return out
    if rec["state"] and rec["state"].upper() not in ("MO", "MISSOURI"):
        out.update(result="Skipped", why="not a Missouri address")
        return out
    how, code, note, where, line = combine(b_now, rec)
    out.update(how_now=how, where_now=where, line_now=line)
    pinned = rec.get("pinned", "")
    if pinned:
        if how in PLACED and code != pinned:
            out.update(result="Map", code_now=code,
                       why=f"you typed {pinned}, but the state's file now puts it in {code}")
            return out
        code, out["how_now"], out["line_now"] = pinned, "PINNED", "typed"
    if not code:
        out.update(result="Map", why=note)
        return out
    out["code_now"] = code
    if code not in rates_now or rates_now[code][1] is None:
        out.update(result="Map", why=f"the state's rate file has no use rate for {code}")
        return out
    sales_now, use_now = rates_now[code]
    out["use_now"] = use_now
    if b_last is not None:
        if pinned:
            code_last = pinned
        else:
            h2, c2, _, _, _ = combine(b_last, rec)
            code_last = c2 if h2 in PLACED else ""
        out["code_last"] = code_last
        if code_last and code_last in rates_last:
            out["use_last"] = rates_last[code_last][1]
    tax = rec["tax"]
    use_last = out["use_last"]
    if tax is None:
        if out["exempt"]:
            out.update(result="OK", why="exempt, no rate set up")
        else:
            out.update(result="Fix", why="no rate set up")
    elif abs(tax - use_now) < 1e-6:
        out.update(result="OK", why="")
    elif sales_now is not None and abs(tax - sales_now) < 1e-6:
        out.update(result="Fix", why="has the sales rate, not the use rate")
    elif use_last is not None and abs(tax - use_last) < 1e-6:
        if out["code_last"] and out["code_last"] != code:
            out.update(result="Fix", why=f"moved from {out['code_last']} to {code} on {q_start}")
        else:
            out.update(result="Fix", why=f"new rate on {q_start}")
    else:
        out.update(result="Fix", why="doesn't match the state's rate")
    return out
