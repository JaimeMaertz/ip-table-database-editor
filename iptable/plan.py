"""What an import would do: one proposal per MAC found, compared against pinger's device database, with
the checks that catch a mistyped table.

A proposal is ticked (imported unless untouched) only when it is new to the database, its manufacturer is
known AV gear, and the row says where the device is (an IP, a hostname or a room). Everything else is a
"maybe" with its reasons, so client devices and placeholder rows stay visible without being imported.
"""
import json
import os
import re
from collections import Counter
from dataclasses import dataclass, field

from . import macs, names, tables
from .oui import Registry

CATALOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "catalog.json")

_COMPANY_WORDS = {"inc", "incorporated", "corp", "corporation", "co", "ltd", "llc", "l", "pty", "gmbh",
                  "electronics", "systems", "technologies", "technology", "the", "company", "division"}


def maker_key(name):
    words = [w for w in re.split(r"[^a-z0-9]+", str(name).lower()) if w and w not in _COMPANY_WORDS]
    return "".join(words)


def _ip_key(ip):
    parts = re.findall(r"\d+", str(ip))
    return [int(x) for x in parts] if len(parts) == 4 else [1 << 30]


def same_maker(a, b):
    ka, kb = maker_key(a), maker_key(b)
    return bool(ka and kb) and (ka.startswith(kb) or kb.startswith(ka) or ka in kb or kb in ka)


@dataclass
class Proposal:
    mac: str
    name: str
    model: str
    jobsite: str
    where: str                      # "Sheet r12"
    make: str = ""
    oui: str = ""
    status: str = "new"             # new | same | differs
    existing: dict = None           # the database entry, when the MAC is already there
    ticked: bool = False
    maybe: list = field(default_factory=list)   # why it is not ticked
    flags: list = field(default_factory=list)   # what looks wrong, ticked or not
    parts: dict = field(default_factory=dict)   # what the name was built from (names.compose)
    row: object = None                          # the tables.Row it came from
    table: object = None                        # the tables.Table it came from

    def entry(self):
        e = {"mac": self.mac, "name": self.name}
        if self.model:
            e["model"] = self.model
        if self.jobsite:
            e["jobsite"] = self.jobsite
        return e


@dataclass
class Unreadable:
    where: str
    cell: str
    why: str
    guess: str = None


class Planner:
    def __init__(self, db_entries, registry=None, catalog_path=CATALOG):
        self.db = {}
        for e in db_entries:
            found = macs.find(e.get("mac", ""))
            if found:
                self.db[found[0]] = e
        self.registry = registry or Registry()
        with open(catalog_path, encoding="utf-8") as f:
            cat = json.load(f)
        self.av = cat["av_manufacturers"]
        self.oem = cat["oem_ouis"]

    def jobsite_for(self, table, path):
        """The jobsite to suggest for a table: the one the database already gives most of the table's MACs,
        else the sheet's own PROJECT cell, else the workbook's name plus the sheet's (a sheet is usually a
        building or site), leaving out a sheet name that says nothing ("Sheet1", "Main")."""
        known = Counter(self.db[m].get("jobsite") for r in table.rows for m in r.macs
                        if m in self.db and self.db[m].get("jobsite"))
        if known:
            return known.most_common(1)[0][0]
        for label in ("project", "project name", "job name", "jobsite", "site"):
            if table.titles.get(label):
                return table.titles[label]
        name = re.sub(r"(?i)\b(network|summary|sheet|ip|table|working|copy|final)\b|^\d{6}\b|-", " ",
                      tables.workbook_name(path))
        name = " ".join(w.capitalize() if w.isupper() and len(w) > 3 else w for w in name.split())
        generic = re.fullmatch(r"(?i)(sheet\s*\d*|main|data|network|devices?|ip\s*table|summary|av network|customer network)",
                               table.sheet.strip())
        return name if generic else (name + " " + table.sheet).strip()

    def plan(self, path, jobsite=None):
        found = tables.read_workbook(path)
        proposals, unreadable, no_mac = [], [], []
        for t in found:
            site = jobsite or self.jobsite_for(t, path)
            items = []                  # (mac, row, where, stray column or None)
            for row in t.rows:
                where = "%s r%d" % (row.sheet, row.number)
                if row.mac_cell and not row.macs and not macs.is_blank(row.mac_cell):
                    why, guess = macs.diagnose(row.mac_cell)
                    unreadable.append(Unreadable(where, row.mac_cell, why, guess))
                if not row.macs and not row.stray_macs:
                    no_mac.append(where)
                items += [(mac, row, where, None) for mac in row.macs]
                items += [(mac, row, where, col) for col, mac in row.stray_macs if mac not in row.macs]
            suggested = names.suggest([self._name_fields(row) for _, row, _, _ in items], names.prefix_for(t.sheet))
            for (mac, row, where, stray), (name, parts) in zip(items, suggested):
                p = self._propose(mac, row, where, site, name)
                p.parts, p.row, p.table = parts, row, t
                if stray:
                    p.maybe.append("found in the %s column, not the MAC column: whose MAC is it?" % stray)
                    p.ticked = False
                proposals.append(p)
        self._duplicates(proposals)
        self._hostname_tails(proposals)
        return found, proposals, unreadable, no_mac

    SWEEP_MAYBE = "from a sweep, which gives no name: set Room and Role, or type one, then tick it"

    def plan_sweep(self, sw, label="Sweep", jobsite=None):
        """Proposals for a pinger sweep, the same shape plan() returns. One per host with a MAC, in IP order,
        with no name and unticked; a MAC already in the database shows its entry and is In database. The
        suggested jobsite is the one the database gives most of the answering MACs."""
        rows, no_mac = [], []
        for n, h in enumerate(sorted(sw.hosts, key=lambda h: _ip_key(h.get("ip", ""))), start=1):
            ip = h.get("ip", "")
            found = macs.find(h.get("mac", ""))
            if not found:
                no_mac.append(ip)
                continue
            v = {"ip": ip, "reply": "%g ms" % h["rtt_ms"] if h.get("replied") and h.get("rtt_ms") is not None
                 else ("answered" if h.get("replied") else "no reply: answered ARP, not the ping")}
            rows.append(tables.Row(label, n, v, h.get("mac", ""), [found[0]]))
        t = tables.Table(label, 0, {}, rows, {}, sweep=sw)
        if jobsite is None:
            known = Counter(self.db[r.macs[0]].get("jobsite") for r in rows
                            if r.macs[0] in self.db and self.db[r.macs[0]].get("jobsite"))
            jobsite = known.most_common(1)[0][0] if known else ""
        proposals = []
        for r in rows:
            mac = r.macs[0]
            p = Proposal(mac, "", "", jobsite, r.values["ip"], oui=self.registry.lookup(mac), row=r, table=t)
            self._check(p, r, "")
            if mac in self.db:
                e = p.existing = self.db[mac]
                p.name, p.model, p.jobsite, p.status = e.get("name", ""), e.get("model", ""), e.get("jobsite", ""), "same"
            else:
                p.maybe.append(self.SWEEP_MAYBE)
                if not self._is_av("", p.oui):
                    p.maybe.append("manufacturer %r is not in the AV catalog" % (p.oui or "unknown"))
            proposals.append(p)
        self._duplicates(proposals)
        return [t], proposals, [], no_mac

    def _make_model(self, v):
        make, model = v.get("make", ""), v.get("model", "")
        if "make_model" in v:
            make2, model2 = tables.split_make_model(v["make_model"])
            make, model = make or make2, model or model2
        if not make and model:
            first = model.split()[0]
            if any(same_maker(first, m) for m in self.av):
                make, model = first, model[len(first):].strip()
        return make, tables.clean_model(model, make)

    def _name_fields(self, row):
        fields = dict(row.values)
        fields["make"], fields["model"] = self._make_model(row.values)
        return fields

    def _propose(self, mac, row, where, site, name):
        v = row.values
        make, model = self._make_model(v)
        hostname = v.get("hostname", "")
        place = v.get("room") or v.get("location") or ""
        p = Proposal(mac, name, model, site, where, make=make, oui=self.registry.lookup(mac))

        if mac in self.db:
            p.existing = self.db[mac]
            e = p.existing
            same = (e.get("name", "") == p.name and e.get("model", "") == p.model and e.get("jobsite", "") == p.jobsite)
            p.status = "same" if same else "differs"

        self._check(p, row, hostname)
        if p.status == "new":
            if not (v.get("ip") or hostname or place):
                p.maybe.append("no IP, hostname or room on the row")
            if not self._is_av(make, p.oui):
                p.maybe.append("manufacturer %r is not in the AV catalog" % (make or p.oui or "unknown"))
            p.ticked = not p.maybe
        return p

    def _is_av(self, make, oui):
        return any(same_maker(x, m) for x in (make, oui) if x for m in self.av)

    def _check(self, p, row, hostname):
        mac = p.mac
        if macs.is_multicast(mac):
            p.flags.append("a multicast address, which no device has as its own")
        elif macs.is_local(mac):
            p.flags.append("a locally administered address (made up by the device; it can change)")
        elif not p.oui:
            p.flags.append("its first half is in no registered block: check the first six digits")
        elif p.make and not same_maker(p.make, p.oui):
            accepted = next((alts for k, alts in self.oem.items() if same_maker(k, p.make)), [])
            if not any(same_maker(a, p.oui) for a in accepted):
                p.flags.append("the table says %s, the MAC block is %s's (a typo, or an OEM part to add to the catalog)"
                               % (p.make, p.oui))
        p.hostname_tail = ""
        tail = hostname.rsplit("-", 1)[-1] if "-" in hostname else ""
        if len(tail) in (4, 6) and re.fullmatch(r"[0-9A-Fa-f]+", tail) and re.search(r"\d", tail):
            if not mac.replace(":", "").upper().endswith(tail.upper()):
                p.hostname_tail = tail.upper()

    @staticmethod
    def _hostname_tails(proposals):
        """Flag a hostname that ends in another MAC's digits. A two-port device's second port is often named
        after the first (a ULXD4Q's Dante port carries its control MAC), so a tail that matches a MAC on a
        row with the same model is not flagged; one that matches a different model's MAC is a copied row."""
        by_tail = {}
        for p in proposals:
            for n in (4, 6):
                by_tail.setdefault(p.mac.replace(":", "")[-n:], []).append(p)
        for p in proposals:
            tail = getattr(p, "hostname_tail", "")
            if not tail:
                continue
            owners = [o for o in by_tail.get(tail, []) if o is not p]
            if any(o.model and o.model == p.model for o in owners):
                continue
            mine = p.mac.replace(":", "")[-len(tail):]
            if owners:
                p.flags.append("hostname ends -%s, which is the MAC of %s (%s), not this one (%s): copied from that row?"
                               % (tail, owners[0].where, owners[0].model or owners[0].name, mine))
            else:
                p.flags.append("hostname ends -%s but the MAC ends %s" % (tail, mine))

    @staticmethod
    def _duplicates(proposals):
        seen = {}
        for p in proposals:
            seen.setdefault(p.mac, []).append(p)
        for mac, ps in seen.items():
            if len(ps) > 1:
                for p in ps:
                    others = ", ".join(o.where for o in ps if o is not p)
                    p.flags.append("the same MAC is also on %s" % others)

