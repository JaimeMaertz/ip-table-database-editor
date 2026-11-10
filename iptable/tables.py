"""Finding device tables in an IP-table workbook and reading their rows.

A sheet can hold several tables, each under a header row, with title rows, totals and empty sub-tables
around them. A header row is one where a cell names the MAC column and at least two other cells name
columns this tool knows. A table runs from below its header to the next header-like row or the sheet's end.
A sheet with no header row gets its MAC column, and nothing else, guessed from cell contents.
"""
import os
import re
import warnings
from dataclasses import dataclass, field

import openpyxl

from . import macs

# Column role for a header, matched after lower-casing and dropping everything but letters and digits.
HEADERS = {
    "mac": ["mac", "macaddress", "macaddr", "macid", "hwaddress", "physicaladdress"],
    "ip": ["ip", "ipaddress", "ipaddr", "ipv4", "address"],
    "hostname": ["hostname", "host", "dnsname", "networkname"],
    "make_model": ["makemodel", "manufacturermodel"],
    "make": ["manufacturer", "make", "mfr", "mfg", "brand", "vendor"],
    "model": ["model", "modelnumber", "partnumber", "part"],
    "description": ["description", "device", "devicetype", "function", "equipment", "type"],
    "room": ["roomserved", "room", "physicallocation", "location", "space"],
    "location": ["equipmentlocation", "mountinglocation", "rack"],
    "serial": ["serial", "serialnumber", "sn", "serialno"],
    "notes": ["notes", "note", "comments", "comment"],
}
_ROLE_OF = {name: role for role, names in HEADERS.items() for name in names}


def _key(text):
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def header_roles(cells):
    """{column index: role} for the cells of one row; the first cell to claim a role keeps it."""
    roles = {}
    for col, value in cells:
        role = _ROLE_OF.get(_key(value)) if isinstance(value, str) else None
        if role and role not in roles.values():
            roles[col] = role
    return roles


@dataclass
class Row:
    sheet: str
    number: int                     # the spreadsheet's row number
    values: dict                    # role -> cell text, stripped; only non-empty cells
    mac_cell: str = ""              # what the MAC column held, as typed
    macs: list = field(default_factory=list)          # MACs in the MAC column
    stray_macs: list = field(default_factory=list)    # (header or column letter, MAC) found in other columns


@dataclass
class Table:
    sheet: str
    header_row: int                 # 0 when the columns were guessed
    columns: dict                   # column index -> role
    rows: list
    titles: dict                    # label -> value from the cells above the header ("PROJECT:" -> "...")
    sweep: object = None            # the sweep.Sweep, when the rows came from pinger rather than a workbook


def _text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _sheet_cells(ws):
    grid = {}
    for row in ws.iter_rows():
        for c in row:
            t = _text(c.value)
            if t:
                grid[(c.row, c.column)] = t
    return grid


def _titles(grid, above):
    """Label: value pairs from cells above a header: a cell ending in ':' and the next filled cell to its right."""
    out = {}
    for (r, c), t in grid.items():
        if r < above and t.endswith(":"):
            nxt = [grid[(r, cc)] for cc in range(c + 1, c + 4) if (r, cc) in grid]
            if nxt and not nxt[0].endswith(":"):
                out[t.rstrip(":").strip().lower()] = nxt[0]
    return out


def read_workbook(path):
    """Every device table in the workbook, sheet by sheet."""
    with warnings.catch_warnings():
        # Page headers and footers Excel wrote in a form openpyxl can't parse; they hold no devices.
        warnings.filterwarnings("ignore", message="Cannot parse header or footer")
        wb = openpyxl.load_workbook(path, data_only=True)
    tables = []
    for ws in wb.worksheets:
        tables += read_sheet(ws.title, _sheet_cells(ws))
    return tables


def read_sheet(name, grid):
    if not grid:
        return []
    last = max(r for r, _ in grid)
    by_row = {}
    for (r, c), t in grid.items():
        by_row.setdefault(r, []).append((c, t))

    headers = []
    for r in sorted(by_row):
        roles = header_roles(sorted(by_row[r]))
        if len(roles) >= 3:
            headers.append((r, roles))

    tables = []
    for i, (r, roles) in enumerate(headers):
        if "mac" not in roles.values():
            continue
        end = headers[i + 1][0] if i + 1 < len(headers) else last + 1
        rows = [_row(name, rr, by_row[rr], roles) for rr in range(r + 1, end) if rr in by_row]
        tables.append(Table(name, r, roles, [x for x in rows if x], _titles(grid, r)))
    if tables or headers:
        return tables
    return _guessed(name, grid, by_row)


def _row(sheet, number, cells, roles):
    values = {}
    stray = []
    mac_cell = ""
    found = []
    for col, t in cells:
        role = roles.get(col)
        if role == "mac":
            mac_cell = t
            found = macs.find(t)
        else:
            for m in macs.find(t, bare=False):
                stray.append((role or openpyxl.utils.get_column_letter(col), m))
        if role and role != "mac":
            values[role] = t
    # A row that is only a line number, a total or a title is not a device.
    if not (values or mac_cell) or (not mac_cell and set(values) <= {"notes"}):
        return None
    if not mac_cell and not any(k in values for k in ("make", "model", "make_model", "hostname", "ip", "description")):
        return None
    return Row(sheet, number, values, mac_cell, found, stray)


def _guessed(name, grid, by_row):
    """No header row: the column where most filled cells hold a MAC is the MAC column."""
    counts = {}
    filled = {}
    for (r, c), t in grid.items():
        filled[c] = filled.get(c, 0) + 1
        if macs.find(t):
            counts[c] = counts.get(c, 0) + 1
    if not counts:
        return []
    col = max(counts, key=counts.get)
    if counts[col] < 3 or counts[col] < filled[col] / 2:
        return []
    roles = {col: "mac"}
    rows = [_row(name, r, by_row[r], roles) for r in sorted(by_row) if any(c == col for c, _ in by_row[r])]
    return [Table(name, 0, roles, [x for x in rows if x and x.macs], {})]


def split_make_model(text):
    """'Crestron / CP4N M200000001' -> ('Crestron', 'CP4N'). The make is before the first slash; a trailing
    Crestron part number (M + 9 digits) is dropped. Without a slash the whole text is the model."""
    make, sep, model = str(text).partition("/")
    if not sep:
        make, model = "", make
    model = re.sub(r"\s+M\d{9}\b", "", model)
    return make.strip(), re.sub(r"\s+", " ", model).strip()


def clean_model(model, make=""):
    """The model without ordering codes: a variant after '=' (Shure's MXWAPX8=Z10, ULXD4Q=G50) and, on Shure
    gear, a frequency band written after a space (ULXD4Q G50). The band names the radio's range, not the model."""
    model = re.sub(r"\s*=\S*$", "", str(model)).strip()
    if re.search(r"(?i)\bshure\b", str(make)):
        model = re.sub(r"\s+[A-Z]\d{2}[A-Z]?$", "", model)
    return model


def workbook_name(path):
    stem = os.path.splitext(os.path.basename(path))[0]
    return re.sub(r"[_\s]+", " ", stem).strip()
