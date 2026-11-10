"""pinger's device database: a JSON list of {"mac", "name", "model", "jobsite"} entries keyed by MAC.

Writes keep every entry not being replaced exactly as it was, in its place, add new entries at the end, and
copy the old file to devices.json.<time>.bak first.
"""
import json
import os
import shutil
import time

from . import macs
from .paths import pinger_file

DEFAULT = pinger_file("devices.json", "IPTABLE_DB")


def load(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def key(entry):
    found = macs.find(entry.get("mac", ""))
    return found[0] if found else None


def merge(entries, additions, replacements=()):
    """The database after an import: replacements (entries whose MAC is already there) take their old entry's
    place, additions go at the end. Additions whose MAC is already in the database are left out."""
    by_mac = {key(e): e for e in replacements}
    out, present = [], set()
    for e in entries:
        k = key(e)
        out.append(by_mac.pop(k, e))
        present.add(k)
    out += [e for e in additions if key(e) not in present]
    return out


def write(path, entries):
    """Write the database, backing up the old file first. Returns the backup's path, or None."""
    backup = None
    if os.path.exists(path):
        backup = "%s.%s.bak" % (path, time.strftime("%Y%m%d-%H%M%S"))
        shutil.copy2(path, backup)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="\n") as f:
        json.dump(entries, f, indent=4)
        f.write("\n")
    os.replace(tmp, path)
    return backup
