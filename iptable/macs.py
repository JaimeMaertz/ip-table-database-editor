"""MAC addresses as they are typed into IP tables.

A MAC is written AA:BB:CC:DD:EE:FF here, the form pinger's devices.json uses. Tables write them with
colons, dashes, dots (C4.42.68.12.34.56 or Cisco-style c442.6812.3456) or no separators at all.
"""
import re

# Words a table puts in an empty MAC cell.
BLANKS = {"", "-", "--", "?", "na", "n/a", "tbd", "none", "unknown", "x"}

_SEPARATED = re.compile(
    r"(?<![0-9A-Fa-f:.-])"
    r"([0-9A-Fa-f]{2}([:.-])[0-9A-Fa-f]{2}(?:\2[0-9A-Fa-f]{2}){4}"
    r"|[0-9A-Fa-f]{4}\.[0-9A-Fa-f]{4}\.[0-9A-Fa-f]{4})"
    r"(?![0-9A-Fa-f:.-])"
)
_BARE = re.compile(r"(?<![0-9A-Za-z])([0-9A-Fa-f]{12})(?![0-9A-Za-z])")

# Letters typed for the digit they look like.
_LOOKALIKES = str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "L": "1", "S": "5", "s": "5"})


def normalize(hex12):
    h = re.sub(r"[^0-9A-Fa-f]", "", hex12).upper()
    return ":".join(h[i:i + 2] for i in range(0, 12, 2))


def find(text, bare=True):
    """Every MAC written in a cell, normalized, in order. Twelve hex digits with no separators count only
    when bare is true: a serial number column can hold twelve digits that are not a MAC."""
    text = str(text)
    found = [normalize(m.group(1)) for m in _SEPARATED.finditer(text)]
    if bare:
        found += [normalize(m.group(1)) for m in _BARE.finditer(text)]
    out = []
    for mac in found:
        if mac not in out:
            out.append(mac)
    return out


def is_blank(text):
    return str(text).strip().lower() in BLANKS


def diagnose(text):
    """Why a MAC-column cell that holds no MAC is not one, and the MAC it was probably meant to be (or None)."""
    raw = str(text).strip()
    fixed = raw.translate(_LOOKALIKES)
    if fixed != raw and find(fixed):
        return "a letter typed for a digit", find(fixed)[0]
    digits = re.sub(r"[\s:.\-]", "", raw)
    if re.fullmatch(r"[0-9A-Fa-f]+", digits):
        return "%d hex digits, not 12" % len(digits), None
    if re.fullmatch(r"[0-9A-Za-z]{12}", digits):
        return "holds letters that are not hex (%s)" % "".join(sorted(set(c for c in digits.upper() if c not in "0123456789ABCDEF"))), None
    return "not a MAC address", None


def is_local(mac):
    """A locally administered address: made up by the device (private Wi-Fi, virtual machines), so its first
    half names no manufacturer and it may change."""
    return bool(int(mac[:2], 16) & 0x02)


def is_multicast(mac):
    return bool(int(mac[:2], 16) & 0x01)
