"""A pinger sweep (`pinger -json`) as a source of import rows, beside the IP tables.

    pinger.exe -json 172.22.0.1/24 > sweep.json        then open or drop sweep.json
    pinger.exe -json 172.22.0.1/24 | ip-table-database-editor.cmd

A sweep says only what answered: IP, MAC, manufacturer. Its rows therefore start with no name and unticked;
the name comes from Room and Role, or is typed. Hosts with no MAC (routed, off the machine's own subnet) can't
go in the database, which is keyed by MAC, and are counted like a table's rows with no MAC.
"""
import json
from dataclasses import dataclass, field


class NotASweep(ValueError):
    pass


@dataclass
class Sweep:
    hosts: list                     # pinger's jsonHost objects: ip, replied, rtt_ms, mac, manufacturer, ...
    targets: int = 0
    seconds: float = 0.0
    arp_cleared: bool = False
    skipped: list = field(default_factory=list)     # the sweeping machine's own addresses


def decode(data):
    """Text from a file's or pipe's bytes: UTF-8 (with or without a BOM), or the UTF-16 Windows PowerShell 5.1
    writes when its output is redirected to a file."""
    if isinstance(data, str):
        return data
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return data.decode("utf-16")
    return data.decode("utf-8-sig", errors="replace")


def parse(data):
    """A Sweep from pinger -json output (bytes or text)."""
    text = decode(data).strip()
    if not text:
        raise NotASweep("nothing came in: no pinger -json output")
    try:
        obj = json.loads(text)
    except ValueError as e:
        raise NotASweep("not JSON (pinger needs -json, before the addresses): %s" % e)
    if isinstance(obj, list):
        raise NotASweep("this is a device database, not a sweep: open it with Database… or --db")
    if not isinstance(obj, dict) or not isinstance(obj.get("hosts"), list):
        raise NotASweep("not a pinger -json sweep: there is no hosts list")
    return Sweep(obj["hosts"], obj.get("targets", 0), obj.get("seconds", 0.0), bool(obj.get("arp_cleared")),
                 obj.get("skipped") or [])


def read(path):
    with open(path, "rb") as f:
        return parse(f.read())


def is_sweep_path(path):
    return path == "-" or str(path).lower().endswith(".json")
