"""Tests for the IP-table reader and import plan.

The fixtures (tests/fixtures/, written by make_fixtures.py) are made-up workbooks in two common IP-table
layouts, each mistake the checks look for planted on a known row, and a device database named by hand. The
oracles are kept outside the code under test where they can be:

  * the MAC column counts are taken straight from the sheet with openpyxl, not through the reader;
  * the manufacturer answers come from tests/fixtures/oui-subset.txt, a fixed copy of pinger's blocks;
  * pinger itself must accept a database the importer wrote (when a pinger checkout sits beside this one).

    python3 tests/iptable_tests.py      exit 0 when every test passes, 1 on a failure
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import warnings

HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURES = os.path.join(HERE, "fixtures")
os.environ["IPTABLE_OUI"] = os.path.join(FIXTURES, "oui-subset.txt")
sys.path.insert(0, os.path.dirname(HERE))
warnings.filterwarnings("ignore")
import openpyxl  # noqa: E402
from iptable import macs, sweep, tables  # noqa: E402
from iptable.__main__ import main  # noqa: E402
from iptable.db import load as load_db  # noqa: E402
from iptable.plan import Planner, same_maker  # noqa: E402

SHEET = os.path.join(FIXTURES, "example-network-sheet.xlsx")
SUMMARY = os.path.join(FIXTURES, "example-network-summary.xlsx")
SINGLE = os.path.join(FIXTURES, "example-single-site.xlsx")
DB = os.path.join(FIXTURES, "devices.json")
SWEEP = os.path.join(FIXTURES, "sweep.json")             # pinger -json's shape (pinger's jsonHost)
PINGER = os.path.normpath(os.path.join(HERE, "..", "..", "pinger", "pinger.exe" if os.name == "nt" else "pinger"))
RESULTS = []


def check(name, ok, detail=""):
    RESULTS.append(ok)
    print(("PASS  " if ok else "FAIL  ") + name + ("" if ok else "   " + str(detail)[:300]))


def mac_tests():
    want = "00:10:7F:AA:0B:CC"
    for text in ["00:10:7f:aa:0b:cc", "00-10-7F-AA-0B-CC", "00.10.7F.AA.0B.CC", "0010.7faa.0bcc", "00107FAA0BCC",
                 " 00107faa0bcc ", "MAC 00:10:7F:AA:0B:CC (LAN)"]:
        check("find %r" % text, macs.find(text) == [want], macs.find(text))
    check("second MAC in a VLAN cell", macs.find("00:0E:DD:00:12:34 99 (Dante)") == ["00:0E:DD:00:12:34"])
    check("bare=False skips twelve digits", macs.find("120000012345", bare=False) == [])
    for text in ["172.22.0.042", "1234JBH05678", "00:10:7f:aa:bb", "00:10:7f:aa:bb:cc:dd", "01234567"]:
        check("not a MAC: %r" % text, macs.find(text) == [], macs.find(text))
    check("lookalike letter", macs.diagnose("00:1O:7F:AA:0B:CC") == ("a letter typed for a digit", want))
    check("wrong length", macs.diagnose("00:10:7F:AA:0B")[0] == "10 hex digits, not 12")
    check("blank words", all(macs.is_blank(x) for x in ["", "NA", " - ", "?", "TBD"]))
    check("locally administered", macs.is_local("DA:A1:19:00:00:01") and not macs.is_local(want))


def split_tests():
    cases = {
        "Crestron / CP4N M200000001": ("Crestron", "CP4N"),
        "Crestron/ TSW-770 M200000002": ("Crestron", "TSW-770"),
        " Sony / XR-65X90L Display": ("Sony", "XR-65X90L Display"),
        "Epson / V11H000000/PowerLite L630U Projector": ("Epson", "V11H000000/PowerLite L630U Projector"),
        "BrightSign XT1144": ("", "BrightSign XT1144"),
    }
    for text, want in cases.items():
        check("split %r" % text, tables.split_make_model(text) == want, tables.split_make_model(text))
    for (model, make), want in {("MXWAPX8=Z10", "Shure"), ("ULXD4Q G50", "Shure"), ("ULXD4Q=-G50", "Shure"),
                                ("CV 605", "Marshall"), ("TSW-770", "Crestron"), ("MXA920", "Shure")} and {
            ("MXWAPX8=Z10", "Shure"): "MXWAPX8", ("ULXD4Q G50", "Shure"): "ULXD4Q", ("ULXD4Q=-G50", "Shure"): "ULXD4Q",
            ("CV 605", "Marshall"): "CV 605", ("TSW-770", "Crestron"): "TSW-770", ("MXA920", "Shure"): "MXA920"}.items():
        check("clean model %r (%s)" % (model, make), tables.clean_model(model, make) == want, tables.clean_model(model, make))
    check("same maker across registry spellings", same_maker("Shure", "SHURE INCORPORATED")
          and same_maker("Biamp", "Biamp Systems") and not same_maker("Crestron", "Dell Inc."))


def name_tests():
    from iptable import names
    check("room from a room cell", names.room_token("Training Room C214") == "214")
    check("room from a description", names.room_token("Event Hall", "Touch Panel (240)") == "240")
    check("room words when there is no number", names.room_token("Network Lab Room") == "Network-Lab")
    check("role from a description", names.role("Ceiling Microphone (Dante) (250)", "") == "Ceiling Mic")
    check("role from the model", names.role("", "DM-NVX-E30") == "Encoder" and names.role("", "TSW-570P") == "Touch Panel")
    check("punctuation stays out of names", names.join("B2", "Pat's Conference", "Processor/Matrix (Rear End)")
          == "B2-Pats-Conference-Processor-Matrix-Rear-End")
    got = names.assign([
        {"room": "C310", "location": "310 Wall", "description": "Encoder"},
        {"room": "C310", "location": "Rack", "description": "Encoder"},
        {"room": "C310", "location": "310 Ceiling", "description": "Microphone"},
        {"room": "C310", "location": "310 Ceiling", "description": "Microphone", "hostname": "B1-C310-MIC-Dante-0B11"},
        {"room": "C310", "location": "310 Wall", "description": "Touch Panel"},
    ], names.prefix_for("B1"))
    check("assign: location separates two encoders, mics always carry it, Dante port named",
          got == ["B1-310-Wall-Encoder", "B1-310-Encoder", "B1-310-Ceiling-Mic", "B1-310-Ceiling-Mic-Dante",
                  "B1-310-Touch-Panel"], got)
    got = names.assign([
        {"room": "115 Classroom", "model": "MXWAPX8", "hostname": "EX-MXWAP-2-4A1C"},
        {"room": "115 Classroom", "model": "MXWAPX8", "hostname": "EX-MXWAP-2-D-4A1D"},
        {"room": "Second Floor: 260 Boardroom", "description": "Mic Access Point Control"},
        {"room": "Second Floor: 260 Boardroom", "description": "Mic Access Point Dante"},
    ])
    check("Dante port from a -D- hostname; a trailing Control is the main port",
          got == ["115-Mic-Access-Point", "115-Mic-Access-Point-Dante", "260-Mic-Access-Point", "260-Mic-Access-Point-Dante"], got)
    check("assign: equal names numbered", names.assign([{"room": "C320", "description": "Decoder"}] * 2)
          == ["320-Decoder", "320-Decoder-2"])


def column_count(path, sheet, header_row, header):
    """MAC cells under a header, read straight from the sheet: the oracle for the reader's row count."""
    ws = openpyxl.load_workbook(path, data_only=True)[sheet]
    col = next(c.column for c in ws[header_row] if str(c.value or "").strip().lower() == header)
    return sum(1 for r in range(header_row + 1, ws.max_row + 1)
               if len(re.sub(r"[^0-9A-Fa-f]", "", str(ws.cell(r, col).value or ""))) == 12)


def by_where(proposals):
    out = {}
    for p in proposals:
        out.setdefault(p.where, []).append(p)
    return out


def from_mac_column(props):
    return [p for p in props if not any("column, not the MAC column" in m for m in p.maybe)]


def network_sheet_tests(planner):
    found, props, unreadable, _ = planner.plan(SHEET)
    check("network sheet: one table, under header row 6", [(t.sheet, t.header_row) for t in found] == [("Customer Network", 6)],
          [(t.sheet, t.header_row) for t in found])
    n = column_count(SHEET, "Customer Network", 6, "mac")
    check("network sheet: every MAC-column MAC read", len(from_mac_column(props)) == n, (len(props), n))
    w = by_where(props)
    r = lambda n: w["Customer Network r%d" % n]          # noqa: E731
    check("make and model split, part number dropped", (r(7)[0].make, r(7)[0].model) == ("Crestron", "CP4N"),
          (r(7)[0].make, r(7)[0].model))
    check("jobsite from the PROJECT: cell", r(7)[0].jobsite == "Example Project", r(7)[0].jobsite)
    check("unregistered block flagged", any("no registered block" in f for f in r(8)[0].flags), r(8)[0].flags)
    check("Dell MAC on a Crestron codec accepted", not any("MAC block" in f for f in r(9)[0].flags), r(9)[0].flags)
    check("service laptop is a maybe", not r(10)[0].ticked and any("AV catalog" in m for m in r(10)[0].maybe), r(10)[0].maybe)
    check("hostname tail typo flagged", any("-C0A2" in f for f in r(11)[0].flags), r(11)[0].flags)
    check("a MAC in the VLAN cell is a maybe", len(r(12)) == 2 and not r(12)[1].ticked
          and any("L column" in m for m in r(12)[1].maybe), [p.maybe for p in r(12)])
    check("an entry with no jobsite differs from a row that has one", r(13)[0].status == "differs", r(13)[0].status)
    check("credentials are never read", not any("password" in v for p in props for v in p.row.values.values()))
    check("no unreadable MACs", not unreadable, unreadable)


def network_summary_tests(planner):
    found, props, unreadable, _ = planner.plan(SUMMARY)
    check("summary: one table per building sheet, header row 4",
          [(t.sheet, t.header_row) for t in found] == [("B1", 4), ("B2", 4), ("Event Hall", 4)],
          [(t.sheet, t.header_row) for t in found])
    n = column_count(SUMMARY, "B1", 4, "mac address")
    check("summary: every MAC-column MAC read", len([p for p in from_mac_column(props) if p.where.startswith("B1 ")]) == n, n)
    w = by_where(props)
    r = lambda n: w["B1 r%d" % n][0]                     # noqa: E731
    check("dotted MAC read", r(6).mac == "C4:42:68:00:20:01", r(6).mac)
    check("Audinate MAC on a Biamp accepted", not any("MAC block" in f for f in r(8).flags), r(8).flags)
    for row, mic in ((13, 11), (14, 12)):
        check("camera r%d hostname copied from r%d" % (row, mic), any("B1 r%d" % mic in f for f in r(row).flags), r(row).flags)
    check("a Dante port named after its control port is not flagged", not any("hostname" in f for f in r(17).flags), r(17).flags)
    check("the same MAC on two rows", any("B1 r19" in f for f in r(18).flags) and any("B1 r18" in f for f in r(19).flags))
    check("a letter typed for a digit, with the likely MAC", [(u.where, u.guess) for u in unreadable] == [("B1 r20", "00:10:7F:00:90:02")],
          [(u.where, u.why, u.guess) for u in unreadable])
    check("a spare with no IP or room is a maybe", not r(22).ticked and any("no IP" in m for m in r(22).maybe), r(22).maybe)
    check("a MAC in Notes is a maybe", len(w["B1 r24"]) == 2 and not w["B1 r24"][1].ticked)
    check("hostname tail matching no MAC flagged", any("-B099 but the MAC ends B001" in f for f in r(26).flags), r(26).flags)
    db_macs = {macs.find(e["mac"])[0] for e in load_db(DB)}
    check("every database MAC on the sheet is recognized", sum(p.mac in db_macs for p in props) == 8
          and all(p.status != "new" for p in props if p.mac in db_macs))
    check("same entry is In database, a formatting difference Differs", r(11).status == "same" and r(15).status == "differs")
    check("jobsite from the database", r(5).jobsite == "Example Campus B1", r(5).jobsite)
    check("jobsite from the workbook and sheet when nothing else says", w["B2 r5"][0].jobsite == "example B2", w["B2 r5"][0].jobsite)

    names_of = {p.where: p.name for p in props}
    check("names by building, room and role", [names_of["B1 r%d" % i] for i in (6, 9, 10, 11, 12)] ==
          ["B1-101-Processor", "B1-101-Wall-Encoder", "B1-101-Encoder", "B1-101-Ceiling-Mic", "B1-101-Ceiling-Mic-Dante"],
          [names_of["B1 r%d" % i] for i in (6, 9, 10, 11, 12)])
    check("hand names that carry the table's room come back from the table",
          all(p.name == p.existing["name"] for p in props if p.where in ("B1 r9", "B1 r10", "B1 r11", "B1 r15")))
    check("rack devices named for the rooms they serve are the known gap",
          [p.existing["name"] for p in props if p.existing and "10X" in p.existing["name"]] ==
          ["B1-10X-Switch", "B1-10X-Processor", "B1-10X-Tesira"])
    check("equal names numbered", [names_of["B2 r%d" % i] for i in (5, 6, 7, 8)] ==
          ["B2-201-Decoder", "B2-201-Decoder-2", "B2-201-Touch-Panel", "B2-201-Touch-Panel-2"])
    check("a row with no model or description is named for its maker", [names_of["B2 r9"], names_of["B2 r10"]] ==
          ["B2-201-Shure", "B2-201-Shure-2"])
    check("room from a description, else the room's words",
          (names_of["Event Hall r5"], names_of["Event Hall r6"]) == ("240-Touch-Panel", "Event-Hall-DSP"))


def single_site_tests(planner):
    found, props, _, _ = planner.plan(SINGLE)
    check("single site: one table, header row 3; a network sheet with no MACs is no table",
          [(t.sheet, t.header_row) for t in found] == [("Riverside", 3)], [(t.sheet, t.header_row) for t in found])
    check("single site: every MAC read", len(props) == column_count(SINGLE, "Riverside", 3, "mac address"))
    check("jobsite names the sheet", props[0].jobsite.endswith("Riverside"), props[0].jobsite)
    check("WilliamsAV's Williams Sound block accepted",
          not any("MAC block" in f for p in props if p.make == "WilliamsAV" for f in p.flags))


def table_tests():
    planner = Planner(load_db(DB))
    network_sheet_tests(planner)
    network_summary_tests(planner)
    single_site_tests(planner)


def write_tests():
    tmp = tempfile.mkdtemp()
    try:
        db = os.path.join(tmp, "devices.json")
        shutil.copy(DB, db)
        before = load_db(db)
        main([SUMMARY, "--db", db, "--write"])
        after = load_db(db)
        check("write keeps every existing entry unchanged", after[:len(before)] == before)
        added = after[len(before):]
        check("write adds only new MACs", not ({e["mac"] for e in added} & {macs.find(e["mac"])[0] for e in before}))
        # B1: r8 r12 r14 r16 r17 r18 (r19 is its MAC again) r21 r24 r25 r26; B2: all six; Event Hall: both.
        check("write adds the ticked rows, a MAC on two rows once", len(added) == 18, len(added))
        check("write leaves a backup", any(f.endswith(".bak") for f in os.listdir(tmp)))
        if os.path.exists(PINGER):
            r = subprocess.run([PINGER, "-db", db, "-no-clear", "-timeout", "200ms", "-count", "1", "192.0.2.1"],
                               capture_output=True, text=True, timeout=30)
            check("pinger accepts the written database", "not a MAC" not in r.stderr and r.returncode == 0, r.stderr)
        else:
            print("SKIP  pinger accepts the written database: no built pinger at %s" % PINGER)
    finally:
        shutil.rmtree(tmp)



def sweep_parse_tests():
    raw = open(SWEEP, "rb").read()
    sw = sweep.parse(raw)
    check("a sweep reads", len(sw.hosts) == 10 and sw.targets == 254 and sw.arp_cleared and sw.skipped == ["172.22.0.200"])
    check("a PowerShell 5.1 redirect (UTF-16) reads", len(sweep.parse(raw.decode().encode("utf-16")).hosts) == 10)
    check("a UTF-8 BOM reads", len(sweep.parse(b"\xef\xbb\xbf" + raw).hosts) == 10)
    for name, data, says in (("a device database is refused", b'[{"mac": "00:10:7F:00:00:01"}]', "device database"),
                             ("empty input is refused", b"  \n", "nothing came in"),
                             ("pinger's table output is refused", b"IP  TIME  MAC\n", "-json")):
        try:
            sweep.parse(data)
            check(name, False, "accepted")
        except sweep.NotASweep as e:
            check(name, says in str(e), e)


def sweep_plan_tests():
    found, ps, unreadable, no_mac = Planner(load_db(DB)).plan_sweep(sweep.read(SWEEP))
    by_ip = {p.where: p for p in ps}
    check("one row per host with a MAC, routed hosts counted apart", len(ps) == 9 and no_mac == ["10.1.1.5"], (len(ps), no_mac))
    check("rows in IP order", [p.where for p in ps][:3] == ["172.22.0.9", "172.22.0.10", "172.22.0.31"], [p.where for p in ps][:3])
    known = by_ip["172.22.0.10"]
    check("a known MAC is In database with its entry", known.status == "same" and known.name == "B1-10X-Processor"
          and known.model == "RMC4" and not known.ticked)
    new = by_ip["172.22.0.31"]
    check("a new row has no name and is not ticked", new.status == "new" and new.name == "" and not new.ticked
          and new.maybe == [Planner.SWEEP_MAYBE], (new.name, new.maybe))
    check("the jobsite is the one the database gives the answering MACs", new.jobsite == "Example Campus B1", new.jobsite)
    check("non-AV gear says so", any("not in the AV catalog" in m for m in by_ip["172.22.0.61"].maybe), by_ip["172.22.0.61"].maybe)
    check("one MAC on two IPs is flagged", any("also on 172.22.0.62" in f for f in by_ip["172.22.0.61"].flags), by_ip["172.22.0.61"].flags)
    check("a private address is flagged", any("locally administered" in f for f in by_ip["172.22.0.70"].flags))
    check("a host that answered ARP only says so", by_ip["172.22.0.50"].row.values["reply"].startswith("no reply"))
    check("the source cells are what the search reads", by_ip["172.22.0.31"].row.values["ip"] == "172.22.0.31")


def sweep_cli_tests():
    tmp = tempfile.mkdtemp()
    try:
        db = os.path.join(tmp, "devices.json")
        shutil.copy(DB, db)
        before = load_db(db)
        main([SWEEP, "--db", db, "--write", "--maybe"])
        check("the command line never writes a nameless sweep row", load_db(db) == before)
        env = dict(os.environ, PYTHONPATH=os.path.dirname(HERE))
        r = subprocess.run([sys.executable, "-m", "iptable", "-", "--db", db], input=open(SWEEP, "rb").read(),
                           capture_output=True, env=env, timeout=60)
        out = r.stdout.decode()
        check("a sweep piped in is read", r.returncode == 0 and "10 of 254 hosts answered" in out
              and "pinger sweep from stdin" in out, out[-400:] + r.stderr.decode()[-300:])
        r = subprocess.run([sys.executable, "-m", "iptable", "-", "--db", db], input=b"[]", capture_output=True,
                           env=env, timeout=60)
        check("a database piped in by mistake stops with why", r.returncode != 0 and b"device database" in r.stderr, r.stderr)
    finally:
        shutil.rmtree(tmp)


if __name__ == "__main__":
    mac_tests()
    split_tests()
    name_tests()
    sweep_parse_tests()
    sweep_plan_tests()
    sweep_cli_tests()
    table_tests()
    write_tests()
    print("%d/%d passed" % (sum(RESULTS), len(RESULTS)))
    sys.exit(0 if all(RESULTS) else 1)
