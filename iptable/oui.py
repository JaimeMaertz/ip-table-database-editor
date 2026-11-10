"""Manufacturer of a MAC address from pinger's copy of the IEEE registry (oui.txt, beside pinger: see paths.py).

Each line is a hex prefix of 6, 7 or 9 digits (the 24, 28 and 36-bit blocks), a tab, and the organization.
The longest matching prefix wins, as in pinger.
"""
import os

from .paths import pinger_file

DEFAULT = pinger_file("oui.txt", "IPTABLE_OUI")


class Registry:
    def __init__(self, path=DEFAULT):
        self.blocks = {}
        if not os.path.exists(path):
            raise FileNotFoundError("the manufacturer list %s is missing: put pinger's oui.txt beside pinger, "
                                    "or name it with IPTABLE_OUI" % path)
        with open(path, encoding="utf-8") as f:
            for line in f:
                prefix, _, org = line.rstrip("\n").partition("\t")
                if org:
                    self.blocks[prefix.upper()] = org.strip()

    def lookup(self, mac):
        """The organization for a normalized MAC, or "" when the MAC is in no registered block."""
        h = mac.replace(":", "")
        for n in (9, 7, 6):
            org = self.blocks.get(h[:n])
            if org and not (n == 6 and org == "IEEE Registration Authority"):
                return org
        return ""
