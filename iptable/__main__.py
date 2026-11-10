"""Read IP-table workbooks and show what importing them into pinger's device database would do.

    python3 -m iptable TABLE.xlsx [TABLE.xlsx ...] [--db devices.json] [--jobsite NAME] [-v]
    python3 -m iptable TABLE.xlsx --write            add the ticked rows (a backup is written first)
    python3 -m iptable TABLE.xlsx --write --maybe    also add the maybes
    pinger -json 172.22.0.1/24 | python3 -m iptable -     a sweep from pinger: what answered, new or known
    python3 -m iptable sweep.json                      the same, from a saved pinger -json file

Nothing is written without --write. Rows already in the database are never changed.
"""
import argparse
import os
import sys

from . import db, sweep
from .plan import Planner

DEFAULT_DB = db.DEFAULT

MARK = {"new": "+", "same": "=", "differs": "~"}


def show(path, found, proposals, unreadable, no_mac, verbose, out=sys.stdout):
    print("%s" % ("pinger sweep from stdin" if path == "-" else os.path.basename(path)), file=out)
    for t in found:
        if t.sweep is not None:
            sw = t.sweep
            print("  %d of %d hosts answered in %.1fs, %d with a MAC" % (len(sw.hosts), sw.targets, sw.seconds, len(t.rows)), file=out)
            if not sw.arp_cleared:
                print("  ! the ARP cache was not cleared (run pinger from an Administrator prompt): a MAC can be one "
                      "that has since left", file=out)
            continue
        if not t.rows:
            continue
        where = "header row %d" % t.header_row if t.header_row else "no header row, MAC column guessed"
        print("  %s (%s): %d device rows" % (t.sheet, where, len(t.rows)), file=out)
    for p in proposals:
        if p.status == "same" and not verbose and not p.flags:
            continue
        mark = MARK[p.status] if p.status != "new" else ("+" if p.ticked else "?")
        print("  %s %-14s %s  %-28s %-22s %s" % (mark, p.where, p.mac, p.name[:28], p.model[:22], p.jobsite), file=out)
        if p.status == "differs":
            e = p.existing
            print("        in the database as: %s" % ", ".join("%s=%r" % (k, e[k]) for k in ("name", "model", "jobsite") if e.get(k)), file=out)
        for m in p.maybe:
            print("        maybe: %s" % m, file=out)
        for f in p.flags:
            print("        ! %s" % f, file=out)
    for u in unreadable:
        print("  ! %-14s MAC cell %r: %s%s" % (u.where, u.cell, u.why, "; meant %s?" % u.guess if u.guess else ""), file=out)
    n = {k: sum(1 for p in proposals if (p.status, p.ticked) == k) for k in [("new", True), ("new", False)]}
    print("  %d to add, %d maybe, %d already in the database (%d with other details), %d rows with no MAC%s"
          % (n[("new", True)], n[("new", False)], sum(p.status != "new" for p in proposals),
             sum(p.status == "differs" for p in proposals), len(no_mac),
             (": " + ", ".join(no_mac)) if verbose and no_mac else ""), file=out)


def write(db_path, entries, additions):
    backup = db.write(db_path, db.merge(entries, additions))
    if backup:
        print("backup: %s" % backup)
    print("added %d devices to %s" % (len(additions), db_path))


def main(argv=None):
    ap = argparse.ArgumentParser(prog="iptable", description=__doc__.split("\n\n")[0])
    ap.add_argument("tables", nargs="+")
    ap.add_argument("--db", default=DEFAULT_DB, help="pinger device database (default: %(default)s)")
    ap.add_argument("--jobsite", help="jobsite for every row, instead of the suggested one")
    ap.add_argument("--write", action="store_true", help="add the ticked rows to the database")
    ap.add_argument("--maybe", action="store_true", help="with --write, also add the maybes")
    ap.add_argument("-v", "--verbose", action="store_true", help="also list rows already in the database")
    args = ap.parse_args(argv)

    entries = db.load(args.db)
    try:
        planner = Planner(entries)
    except FileNotFoundError as e:
        sys.exit("iptable: %s" % e)
    additions, added, unnamed = [], set(), 0
    for path in args.tables:
        if sweep.is_sweep_path(path):
            try:
                sw = sweep.parse(sys.stdin.buffer.read()) if path == "-" else sweep.read(path)
            except sweep.NotASweep as e:
                sys.exit("%s: %s" % (path, e))
            found, proposals, unreadable, no_mac = planner.plan_sweep(sw, "Sweep", args.jobsite)
        else:
            found, proposals, unreadable, no_mac = planner.plan(path, args.jobsite)
        show(path, found, proposals, unreadable, no_mac, args.verbose)
        for p in proposals:
            if p.status == "new" and (p.ticked or args.maybe) and p.mac not in added:
                if not p.name:
                    unnamed += 1
                    continue
                additions.append(p.entry())
                added.add(p.mac)
        print()
    if unnamed:
        print("%d rows have no name and are left out: name them in the window (ip-table-database-editor.cmd)" % unnamed)
    if args.write:
        write(args.db, entries, additions)
    else:
        print("%d devices would be added to %s (run with --write to add them)" % (len(additions), args.db))


if __name__ == "__main__":
    main()
