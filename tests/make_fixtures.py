"""Writes the test fixtures: three made-up IP-table workbooks, a hand-named device database and a pinger sweep.

Every MAC, name, password and address here is invented. The workbooks copy the shape of two common IP-table
layouts and plant the mistakes the checks exist to catch, each on a known row, so the tests can name them:

  example-network-sheet.xlsx    the "network sheet" layout: a PROJECT: cell, a header on row 6, MAKE / MODEL in
                                one cell, a VLAN column
  example-network-summary.xlsx  the "network summary" layout: one sheet per building, a header on row 4,
                                Manufacturer / Model / Description columns, usernames and passwords
  example-single-site.xlsx      the summary layout with its header on row 3 and a network sheet with no MACs

    python3 tests/make_fixtures.py      rewrites tests/fixtures/; the tests read what is committed there
"""
import json
import os

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "fixtures")

SUMMARY_HEADER = ["Room Served", "Equipment Location", "Manufacturer", "Model", "Description", "MAC Address",
                  "IP Address", "DHCP/ Static", "Hostname", "Wired/ Wireless", "Port", "Username", "Password", "Notes"]

# Building B1, header on row 4, so the first device is row 5. The comment names what each row is for.
B1 = [
    ["Training Room 101", "Rack", "Netgear", "M4250-26G4XF-POE+", "Network Switch", "28:94:01:00:10:01", "172.22.0.4", "Static", "B1-101-SW-1001"],      # r5   in the database
    ["Training Room 101", "Rack", "Crestron", "RMC4", "Processor", "C4.42.68.00.20.01", "172.22.0.5", "Static", "B1-101-PRO-2001"],               # r6   dotted MAC; database names it for 10X
    ["Training Room 101", "Rack", "Biamp", "Tesira Server-IO", "Digital Sound Processor", "78:45:01:00:30:01", "172.22.0.6", "Static", "B1-101-DSP-3001"],  # r7   database: other name and model
    ["Training Room 101", "Rack", "Biamp", "Tesira Server-IO", "Digital Sound Processor (Dante)", "00:1D:C1:00:30:03", "172.22.0.7", "Static", "B1-101-DSP-D-3003"],  # r8  Audinate MAC on Biamp: accepted
    ["Training Room 101", "101 Wall", "Crestron", "DM-NVX-E30", "Encoder", "C4:42:68:00:40:01", "172.22.0.21", "Static", "B1-101-ENC-4001"],     # r9   two encoders in one room,
    ["Training Room 101", "Rack", "Crestron", "DM-NVX-E30", "Encoder", "C4:42:68:00:40:02", "172.22.0.22", "Static", "B1-101-ENC-4002"],        # r10  told apart by location
    ["Training Room 101", "101 Ceiling", "Shure", "MXA920", "Microphone", "00:0E:DD:00:50:01", "172.22.0.31", "Static", "B1-101-MIC-5001"],      # r11
    ["Training Room 101", "101 Ceiling", "Shure", "MXA920", "Microphone (Dante)", "00:0E:DD:00:50:02", "172.22.0.32", "Static", "B1-101-MIC-D-5002"],  # r12
    ["Training Room 101", "101 Ceiling", "Marshall", "CV605", "Camera", "00:23:A8:00:60:01", "172.22.0.41", "Static", "B1-101-CAM-5001"],        # r13  hostname copied from r11
    ["Training Room 103", "103 Ceiling", "Marshall", "CV605", "Camera", "00:23:A8:00:60:02", "172.22.0.42", "Static", "B1-103-CAM-5002"],        # r14  hostname copied from r12
    ["Training Room 103", "103 Wall", "Crestron", "TSW-770", "Touch Panel", "C4:42:68:00:70:01", "172.22.0.51", "Static", "B1-103-TP-7001"],     # r15  database: "TSW 770"
    ["Training Room 103", "Rack", "Shure", "ULXD4Q=-G50", "Wireless Mic Receiver", "00:0E:DD:00:80:01", "172.22.0.61", "Static", "B1-103-RX-8001"],  # r16
    ["Training Room 103", "Rack", "Shure", "ULXD4Q=-G50", "Wireless Mic Receiver (Dante)", "00:0E:DD:00:80:02", "172.22.0.62", "Static", "B1-103-RX-D-8001"],  # r17 named after its control port
    ["Training Room 103", "Rack", "Crestron", "CP4N", "Processor", "00:10:7F:00:90:01", "172.22.0.70", "Static", "B1-103-PRO-9001"],             # r18  same MAC as r19
    ["Training Room 103", "Rack", "Crestron", "CP4N", "Processor", "00:10:7F:00:90:01", "10.1.1.70", "Static", "B1-103-PRO-9001"],               # r19
    ["Training Room 103", "Rack", "Crestron", "CP4N", "Processor", "00:10:7F:00:9O:02", "172.22.0.71", "Static", "B1-103-PRO-9002"],             # r20  letter O for a zero
    ["Training Room 103", "Behind Display", "LG", "98UH5J-H", "Display", "UNREGISTERED", "172.22.0.81", "Static", "B1-103-DISP-0001"],          # r21  block in no registry
    ["", "", "Crestron", "TSW-770", "Touch Panel", "C4:42:68:00:70:09", "", "", ""],                                                          # r22  a spare: no IP or room
    ["IT Office", "Desk", "Dell", "Latitude 5440", "Service Laptop", "00:0B:DB:00:00:01", "", "DHCP", ""],                                     # r23  not AV gear
    ["Training Room 101", "Ceiling", "Shure", "MXWAPX8=Z10", "Mic Access Point", "00:0E:DD:00:A0:01", "172.22.0.35", "Static", "B1-101-AP-A001", "Wired", "", "admin", "example-password", "second port 00:0E:DD:00:A0:02"],  # r24 a MAC in Notes
    ["Training Room 101", "Rack", "Crestron", "UC-CX100-T", "Video Conference Codec", "B0:4F:13:00:00:01", "172.22.0.36", "Static", "B1-101-UC-0001"],  # r25 Dell MAC on a Crestron: accepted
    ["Training Room 101", "101 Wall", "Crestron", "TSW-770", "Touch Panel", "C4:42:68:00:B0:01", "172.22.0.52", "Static", "B1-101-TP-B099"],     # r26  hostname tail matches no MAC
]

B2 = [
    ["Room 201", "Rack", "Crestron", "DM-NVX-D30", "Decoder", "C4:42:68:00:D0:01", "172.22.1.11", "Static", "B2-201-DEC-D001"],
    ["Room 201", "Rack", "Crestron", "DM-NVX-D30", "Decoder", "C4:42:68:00:D0:02", "172.22.1.12", "Static", "B2-201-DEC-D002"],
    ["Room 201", "Wall", "Crestron", "TSW-770", "Touch Panel", "C4:42:68:00:D0:03", "172.22.1.13", "Static", "B2-201-TP-D003"],
    ["Room 201", "Wall", "Crestron", "TSW-770", "Touch Panel", "C4:42:68:00:D0:04", "172.22.1.14", "Static", "B2-201-TP-D004"],
    ["Room 201", "Table", "Shure", "", "", "00:0E:DD:00:D0:05", "172.22.1.15", "Static", ""],
    ["Room 201", "Table", "Shure", "", "", "00:0E:DD:00:D0:06", "172.22.1.16", "Static", ""],
]

CONFERENCE = [
    ["Event Hall", "Closet", "Crestron", "TSW-1070", "Touch Panel (240)", "C4:42:68:00:E0:01", "172.22.2.11", "Static", "EH-TP-E001"],
    ["Event Hall", "Closet", "Biamp", "Tesira Server-IO", "Digital Sound Processor", "78:45:01:00:E0:02", "172.22.2.12", "Static", "EH-DSP-E002"],
]

# Network-sheet layout: header on row 6.
SHEET_HEADER = ["PHYSICAL LOCATION", "SWITCH", "SWITCH PORT", "MAKE / MODEL", "STATIC/DHCP/PRIVATE", "SERIAL", "MAC",
                "HOSTNAME", "IP", "SUBNET MASK", "GATEWAY", "VLAN", "POE BUDGET"]
NETWORK_SHEET = [
    ["Lobby", "SW1", "1", "Crestron / CP4N M200000001", "Static", "EXAMPLE0001", "00:10:7F:00:C0:01", "EX-CP4N-C001", "172.22.0.10", "255.255.255.0", "172.22.0.1", "Control", ""],   # r7
    ["Lobby", "SW1", "2", "Barco / ClickShare CX-50", "Static", "EXAMPLE0002", "UNREGISTERED2", "EX-CX50-0002", "172.22.0.11", "255.255.255.0", "172.22.0.1", "Control", ""],          # r8  block in no registry
    ["Lobby", "SW1", "3", "Crestron / UC-CX100-T", "Static", "EXAMPLE0003", "B0:4F:13:00:00:03", "EX-UC-0003", "172.22.0.12", "255.255.255.0", "172.22.0.1", "Control", ""],           # r9  Dell MAC on a Crestron
    ["", "", "", "Dell / Latitude 5440 Service Laptop", "DHCP", "", "00:0B:DB:00:00:04", "", "", "", "", "", ""],                                                                   # r10 not AV gear
    ["Lobby", "SW1", "4", "Shure / MXWAPX8=Z10", "Static", "EXAMPLE0005", "00:0E:DD:00:C0:A1", "EX-MXWAP-1-C0A2", "172.22.0.13", "255.255.255.0", "172.22.0.1", "Control", "PoE+"],   # r11 hostname tail typo
    ["Lobby", "SW1", "5", "Shure / MXA920", "Static", "EXAMPLE0006", "00:0E:DD:00:C0:B1", "EX-MXA-C0B1", "172.22.0.14", "255.255.255.0", "172.22.0.1", "00:0E:DD:00:C0:B2 99 (Dante)", "PoE"],  # r12 MAC in VLAN
    ["Training Room 104", "SW1", "6", "Crestron / TSW-770 M200000002", "Static", "EXAMPLE0007", "C4:42:68:00:C0:07", "EX-TP-C007", "172.22.0.15", "255.255.255.0", "172.22.0.1", "Control", "PoE"],  # r13
]

SINGLE_SITE = [
    ["Boardroom 260", "Credenza", "Crestron", "TSW-1070", "Touch Panel", "C4:42:68:00:F0:01", "172.22.3.11", "Static", "BR-TP-F001"],
    ["Boardroom 260", "Credenza", "Crestron", "TSW-1070", "Touch Panel", "C4:42:68:00:F0:02", "172.22.3.12", "Static", "BR-TP-F002"],
    ["Boardroom 260", "Ceiling", "Shure", "", "", "00:0E:DD:00:F0:03", "172.22.3.13", "Static", ""],
    ["Boardroom 260", "Ceiling", "Shure", "", "", "00:0E:DD:00:F0:04", "172.22.3.14", "Static", ""],
    ["Boardroom 260", "Rack", "WilliamsAV", "WF T5", "Assistive Listening", "70:B3:D5:58:B0:05", "172.22.3.15", "Static", "BR-ALS-B005"],  # WilliamsAV: accepted OEM
]

DATABASE = [
    {"mac": "28:94:01:00:10:01", "name": "B1-10X-Switch", "model": "M4250-26G4XF-POE+", "jobsite": "Example Campus B1"},
    {"mac": "C4:42:68:00:20:01", "name": "B1-10X-Processor", "model": "RMC4", "jobsite": "Example Campus B1"},
    {"mac": "78:45:01:00:30:01", "name": "B1-10X-Tesira", "model": "X400", "jobsite": "Example Campus B1"},
    {"mac": "C4:42:68:00:40:01", "name": "B1-101-Wall-Encoder", "model": "NVX E30", "jobsite": "Example Campus B1"},
    {"mac": "C4:42:68:00:40:02", "name": "B1-101-Encoder", "model": "NVX E30", "jobsite": "Example Campus B1"},
    {"mac": "00:0E:DD:00:50:01", "name": "B1-101-Ceiling-Mic", "model": "MXA920", "jobsite": "Example Campus B1"},
    {"mac": "00:23:A8:00:60:01", "name": "B1-101-PTZ-Camera", "model": "CV605", "jobsite": "Example Campus B1"},
    {"mac": "C4:42:68:00:70:01", "name": "B1-103-Touch-Panel", "model": "TSW 770", "jobsite": "Example Campus B1"},
    {"mac": "C4:42:68:00:C0:07", "name": "104-Touch-Panel", "model": "TSW-770"},
    {"mac": "C4:42:68:00:FE:47", "name": "Lab Touch Panel", "model": "TSW-770", "jobsite": "Example Lab"},
    {"mac": "00:10:7F:AA:00:01", "name": "Lab Processor", "model": "CP4N", "jobsite": "Example Lab"},
]

SWEEP = {
    "hosts": [
        {"ip": "172.22.0.10", "replied": True, "rtt_ms": 0.6, "mac": "C4:42:68:00:20:01", "manufacturer": "Crestron Electronics, Inc.",
         "name": "B1-10X-Processor", "model": "RMC4", "jobsite": "Example Campus B1"},
        {"ip": "172.22.0.9", "replied": True, "rtt_ms": 1.2, "mac": "00:23:A8:00:60:01", "manufacturer": "Marvell Technology Group Ltd.",
         "name": "B1-101-PTZ-Camera", "model": "CV605", "jobsite": "Example Campus B1"},
        {"ip": "172.22.0.40", "replied": True, "rtt_ms": 0.4, "mac": "C4:42:68:00:70:01", "manufacturer": "Crestron Electronics, Inc."},
        {"ip": "172.22.0.31", "replied": True, "rtt_ms": 0.8, "mac": "00:10:7F:12:34:56", "manufacturer": "CRESTRON ELECTRONICS, INC."},
        {"ip": "172.22.0.32", "replied": True, "rtt_ms": 0.9, "mac": "00:10:7F:12:34:57", "manufacturer": "CRESTRON ELECTRONICS, INC."},
        {"ip": "172.22.0.50", "replied": False, "mac": "00:0E:DD:AA:BB:01", "manufacturer": "SHURE INCORPORATED"},
        {"ip": "172.22.0.61", "replied": True, "rtt_ms": 2.1, "mac": "F8:B4:6A:12:34:56", "manufacturer": "Hewlett Packard"},
        {"ip": "172.22.0.62", "replied": True, "rtt_ms": 2.0, "mac": "F8:B4:6A:12:34:56", "manufacturer": "Hewlett Packard"},
        {"ip": "172.22.0.70", "replied": True, "rtt_ms": 3.3, "mac": "DA:A1:19:00:00:01", "manufacturer": "(private address)"},
        {"ip": "10.1.1.5", "replied": True, "rtt_ms": 9.5},
    ],
    "targets": 254,
    "answered": 10,
    "arp_cleared": True,
    "skipped": ["172.22.0.200"],
}


def unregistered(n):
    """A MAC whose first three bytes are in no IEEE block in oui.txt (a universally administered one)."""
    from iptable.oui import Registry
    reg = Registry()
    found = []
    for a in range(0x10, 0xFF, 4):              # even and not locally administered
        for b in range(0, 256, 37):
            mac = "%02X:%02X:00:00:00:%02X" % (a, b, len(found) + 1)
            if not reg.lookup(mac):
                found.append(mac)
                if len(found) == n:
                    return found
    raise SystemExit("no unregistered block found")


def summary_sheet(wb, title, rows, header_row=4, info=True):
    ws = wb.create_sheet(title)
    ws.cell(1, 1, "Network Summary")
    if info and header_row >= 4:
        for c, v in enumerate(["Subnet Mask:", "255.255.255.0", "Gateway:", "172.22.0.1", "Primary DNS:", "", "VLAN ID:"], 1):
            ws.cell(2, c, v or None)
        for c, v in enumerate(["SSID1:", "", "Password:", "example-wifi-password"], 1):
            ws.cell(3, c, v or None)
    for c, v in enumerate(SUMMARY_HEADER, 1):
        ws.cell(header_row, c, v)
    for r, row in enumerate(rows, header_row + 1):
        full = row + ["Wired", "", "admin", "example-password"][len(row) - 9:] if len(row) < 13 else row
        for c, v in enumerate(full, 1):
            if v != "":
                ws.cell(r, c, v)
    return ws


def main():
    import sys
    sys.path.insert(0, os.path.dirname(HERE))
    os.makedirs(OUT, exist_ok=True)
    bad = unregistered(2)
    rows = [[bad[0] if v == "UNREGISTERED" else v for v in r] for r in B1]
    sheet_rows = [[bad[1] if v == "UNREGISTERED2" else v for v in r] for r in NETWORK_SHEET]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    summary_sheet(wb, "B1", rows)
    summary_sheet(wb, "B2", B2)
    summary_sheet(wb, "Event Hall", CONFERENCE)
    wb.save(os.path.join(OUT, "example-network-summary.xlsx"))

    wb = openpyxl.Workbook()
    info = wb.active
    info.title = "Project Information"
    for r, (k, v) in enumerate([("CUSTOMER:", "Example Customer"), ("PROJECT:", "Example Project"), ("JOB NUMBER:", "0000")], 2):
        info.cell(r, 1, k)
        info.cell(r, 2, v)
    ws = wb.create_sheet("Customer Network")
    ws.cell(1, 1, "DEVICE INFO")
    for r, (k, v) in enumerate([("CUSTOMER:", "Example Customer"), ("PROJECT:", "Example Project"), ("JOB NUMBER:", "0000")], 2):
        ws.cell(r, 1, k)
        ws.cell(r, 2, v)
    ws.cell(2, 6, "SITE PASSWORD:")
    ws.cell(2, 7, "example-site-password")
    ws.cell(5, 1, "IP INFO")
    for c, v in enumerate(SHEET_HEADER, 1):
        ws.cell(6, c, v)
    for r, row in enumerate(sheet_rows, 7):
        for c, v in enumerate(row, 1):
            if v != "":
                ws.cell(r, c, v)
    wb.save(os.path.join(OUT, "example-network-sheet.xlsx"))

    wb = openpyxl.Workbook()
    main_ws = wb.active
    main_ws.title = "Main"
    main_ws.cell(1, 1, "Network Summary")
    for c, v in enumerate(["Location", "VLAN Name", "VLAN ID", "Gateway", "Subnet Mask", "DNS1", "Main SSID", "Main Password", "Notes"], 1):
        main_ws.cell(3, c, v)
    for c, v in enumerate(["Riverside", "AV", "20", "172.22.3.1", "255.255.255.0", "8.8.8.8", "Example", "example-password", ""], 1):
        if v:
            main_ws.cell(4, c, v)
    summary_sheet(wb, "Riverside", SINGLE_SITE, header_row=3, info=False)
    wb.save(os.path.join(OUT, "example-single-site.xlsx"))

    with open(os.path.join(OUT, "devices.json"), "w", encoding="utf-8", newline="\n") as f:
        json.dump(DATABASE, f, indent=2)
        f.write("\n")
    with open(os.path.join(OUT, "sweep.json"), "w", encoding="utf-8", newline="\n") as f:
        json.dump(SWEEP, f, indent=2)
        f.write("\n")
    write_oui_subset()
    print("wrote", OUT, "unregistered blocks:", bad)


def write_oui_subset():
    """fixtures/oui-subset.txt: the blocks of pinger's oui.txt that the fixtures' and tests' MACs fall in, so
    the tests run without pinger and give the same answers whatever pinger's list has since gained."""
    import glob
    import re
    from iptable import macs
    from iptable.oui import DEFAULT
    text = "".join(open(p, encoding="utf-8").read() for p in glob.glob(os.path.join(HERE, "*.py")))
    text += "".join(open(p, encoding="utf-8").read() for p in glob.glob(os.path.join(OUT, "*.json")))
    found = set()
    for m in re.findall(r"(?i)\b(?:[0-9a-f]{2}[:.\-]){5}[0-9a-f]{2}\b|\b[0-9a-f]{4}\.[0-9a-f]{4}\.[0-9a-f]{4}\b", text):
        found.update(macs.find(m))
    full = {}
    with open(DEFAULT, encoding="utf-8") as f:
        for line in f:
            k, _, v = line.rstrip("\n").partition("\t")
            full[k.upper()] = v
    keep = {h[:n]: full[h[:n]] for h in (m.replace(":", "") for m in found) for n in (6, 7, 9) if h[:n] in full}
    with open(os.path.join(OUT, "oui-subset.txt"), "w", encoding="utf-8", newline="\n") as f:
        f.writelines("%s\t%s\n" % kv for kv in sorted(keep.items()))


if __name__ == "__main__":
    main()
