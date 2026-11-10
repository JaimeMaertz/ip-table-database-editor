"""Role-style device names: building, room, role, and the port when a device has two.

    B1-310-Ceiling-Mic   B1-310-Ceiling-Mic-Dante   B1-312-Projector-Decoder   240-Touch-Panel

The room is the number in the row's room cell (C310 -> 310), else the room number a description carries
("Touch Panel (240)"), else the room's own words. The role is the row's description, or, when the table has
none, the role the catalog gives the model. A location word (Ceiling, Wall, Projector) goes in front of a
mic's or speaker's role always, and in front of any other role when two devices in the same room share it,
so it tells them apart. Names still equal after that get -2, -3.

What a table can't say is which room a rack device serves: the processor in the room-311 rack serves 310 and
312, so the name has room 311 until someone changes it.
"""
import json
import os
import re

CATALOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "catalog.json")

# Words of a description, rewritten to the short form used in names.
WORDS = [
    (r"\bdigital sound processor\b", "DSP"),
    (r"\bcontrol processor\b", "Processor"),
    (r"\bnetwork (audio )?switch\b", "Switch"),
    (r"\btouch ?panel\b|\btouchscreen\b", "Touch Panel"),
    (r"\bmicrophone\b", "Mic"),
    (r"\bwireless mic(rophone)? (receiver|rx)\b|\bwmic rx\b", "Wireless Mic Receiver"),
    (r"\bvideo conference codec\b", "Codec"),
    (r"\bhdmi (encoder|decoder)\b", r"\1"),
]
# Location words that can tell two same-role devices apart; others ("Rack", "Behind Display") can't.
PLACES = ["Ceiling", "Wall", "Projector", "Podium", "Lectern", "Table", "Floor", "Display", "Desk", "Credenza"]
# Port words: Dante is kept in the name, Control is the device's main port and dropped.
DANTE = re.compile(r"\b(dante|dan)\b", re.I)


def _load_roles(path=CATALOG):
    with open(path, encoding="utf-8") as f:
        return [(re.compile(rx, re.I), role) for rx, role in json.load(f).get("roles", [])]


ROLES = _load_roles()


def room_token(room, description=""):
    """'Training Room C214' -> '214', 'Storage C320' -> '320', 'Event Hall' + '(240)' -> '240'."""
    for text in (room, description):
        nums = re.findall(r"\b[A-Za-z]?(\d{3,4}[A-Za-z]?)\b", str(text))
        if nums:
            return nums[-1] if text is room else nums[0]
    words = [w for w in re.split(r"[^A-Za-z0-9]+", str(room)) if w and w.lower() not in ("room", "the")]
    return "-".join(w.capitalize() if w.islower() else w for w in words)


def role(description, model):
    text = str(description or "")
    if text:
        text = re.sub(r"\((?:[^)]*\d[^)]*|control|dante|lan|control subnet|.*input|.*display|.*codec.*)\)", " ", text, flags=re.I)
        for rx, short in WORDS:
            text = re.sub(rx, short, text, flags=re.I)
        text = DANTE.sub(" ", text)
        text = re.sub(r"[()+]", " ", text)
        text = re.sub(r"(?i)\s+control$", "", text.strip())    # the main port: "Mic Access Point Control"
    else:
        text = next((r for rx, r in ROLES if rx.search(str(model))), str(model))
    return " ".join(text.split())


def place(location):
    words = str(location).replace("-", " ").split()
    return next((p for p in PLACES if any(w.lower() == p.lower() for w in words)), "")


# A hostname that marks the Dante port with a lone D before the MAC digits: EX-MXWAP-2-D-4A1D.
DANTE_HOSTNAME = re.compile(r"-D-[0-9A-Fa-f]{4,6}$")


def is_dante(*texts, hostname=""):
    return any(DANTE.search(str(t)) for t in texts if t) or bool(DANTE_HOSTNAME.search(str(hostname)))


def _word(w):
    return w if (w.isupper() or any(c.isdigit() for c in w)) else w[:1].upper() + w[1:]


def join(*parts):
    words = []
    for p in parts:
        p = re.sub(r"['’]", "", str(p))                 # Pat's -> Pats
        p = re.sub(r"[()\[\]/,:;&+]", " ", p)                # (PC End), Processor/Matrix
        words += [_word(w) for w in re.split(r"[\s_]+|(?<!\d)-(?!\d)", p) if w]
    return "-".join(words)


def prefix_for(sheet):
    """A building code for a sheet named like one ('B1', 'B2'), else nothing."""
    s = sheet.strip()
    return s if re.fullmatch(r"[A-Za-z]{1,2}\d{1,2}", s) else ""


def compose(part):
    """The name for one device's parts: {"prefix", "room", "role", "dante"}."""
    return join(part.get("prefix", ""), part.get("room", ""), part.get("role", ""), "Dante" if part.get("dante") else "")


def assign(items, prefix=""):
    """Names for items (dicts with room, description, model, location, hostname, make_model, notes)."""
    return [n for n, _ in suggest(items, prefix)]


def suggest(items, prefix=""):
    """(name, parts) for each item. The parts are what the name was built from, so an editor can change the
    room or role and build the name again with compose(); the -2, -3 of equal names is not in the parts."""
    parts = []
    for it in items:
        parts.append({
            "prefix": prefix,
            "room": room_token(it.get("room", ""), it.get("description", "")),
            "role": role(it.get("description", ""), it.get("model", "")) or it.get("make", ""),
            "place": place(it.get("location", "")),
            "dante": is_dante(it.get("description"), it.get("hostname"), it.get("make_model"), it.get("notes"),
                              hostname=it.get("hostname", "")),
        })
    # Location only where it separates two devices with the same room and role.
    groups = {}
    for p in parts:
        groups.setdefault((p["room"], p["role"].lower(), p["dante"]), []).append(p)
    for g in groups.values():
        separates = len(g) > 1 and len({p["place"] for p in g}) > 1
        for p in g:
            # A mic's or speaker's mounting is part of what it is (a ceiling mic, a table mic), so it is
            # always named; for other roles the location only separates two of a kind.
            always = re.search(r"\b(mic|speaker)\b", p["role"], re.I)
            if (separates or always) and p["place"] and p["place"].lower() not in p["role"].lower():
                p["role"] = p["place"] + " " + p["role"]
    names = [compose(p) for p in parts]
    seen = {}
    for i, n in enumerate(names):
        seen.setdefault(n.lower(), []).append(i)
    for idx in seen.values():
        if len(idx) > 1:
            for k, i in enumerate(idx[1:], start=2):
                names[i] = "%s-%d" % (names[i], k)
    return [(n, {k: p[k] for k in ("prefix", "room", "role", "dante")}) for n, p in zip(names, parts)]
