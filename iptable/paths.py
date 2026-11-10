"""Where pinger's files are: the device database and the manufacturer list sit beside pinger.

pinger reads devices.json from beside its own executable, so that is the database this tool edits by default.
A file is looked for beside the pinger on PATH, then in a pinger folder beside this checkout
(../pinger, the layout when both repositories are cloned side by side). IPTABLE_DB and IPTABLE_OUI name a
file directly and win over both.
"""
import os
import shutil

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def pinger_file(name, env=None):
    if env and os.environ.get(env):
        return os.environ[env]
    candidates = []
    exe = shutil.which("pinger") or shutil.which("pinger.exe")
    if exe:
        candidates.append(os.path.join(os.path.dirname(os.path.realpath(exe)), name))
    sibling = os.path.join(os.path.dirname(HERE), "pinger", name)
    candidates.append(sibling)
    return next((p for p in candidates if os.path.exists(p)), candidates[0])
