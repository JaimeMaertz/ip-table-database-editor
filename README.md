# ip-table-database-editor

Turns a job's IP table (an `.xlsx` workbook) or a [pinger](https://github.com/JaimeMaertz/pinger)
sweep into named entries in pinger's device database, checks the table for typing mistakes on
the way, and edits that database directly. It has a PySide6 window with Import and Database
tabs, and a command line. Written in Python, for Windows and Linux.

Use it when a job's IP table should become device names in your network sweeps, or to check
a table for MAC and hostname typos before it goes to the client.

```
pinger.exe -json 172.22.0.1/24 | ip-table-database-editor.cmd      the window, with the sweep loaded
ip-table-database-editor.cmd "Network Summary.xlsx"                 the window, with a table loaded
```

## Install

Needs Python 3.10 or later. On Windows, the `.cmd` launchers use `python` if it's on PATH,
else the `py` launcher.

```
git clone https://github.com/JaimeMaertz/ip-table-database-editor
cd ip-table-database-editor
pip install -r requirements.txt
```

On Linux distributions that refuse a system-wide `pip install` (Debian, Ubuntu), make a
virtual environment first: `python3 -m venv .venv && . .venv/bin/activate`.

To run the `.cmd` launchers from any folder, add this folder to your PATH. The
`python -m iptable` forms work from this folder.

It reads two of pinger's files: **`devices.json`**, the database it edits, and **`oui.txt`**,
the IEEE manufacturer list. It looks for both beside the `pinger` on your PATH, then in a
`pinger` folder beside this one (clone both repositories side by side). pinger's release zip
has `oui.txt` beside the exe; keep them together. Which database is used, first match wins:

1. `--db PATH`
2. the `IPTABLE_DB` environment variable (`IPTABLE_OUI` names the manufacturer list the same way)
3. in the window only, the database used last time
4. `devices.json` beside pinger

If something it needs is missing (PySide6, openpyxl or `oui.txt`), the launcher says so and
exits instead of opening the window.

## The window

```
ip-table-database-editor.cmd [TABLE.xlsx | SWEEP.json | -] [--db devices.json]     Windows
python -m iptable.gui [TABLE.xlsx | SWEEP.json | -] [--db devices.json]           from this folder, any OS
```

Open or drop a workbook or a sweep, review each row, then write. The tool suggests and you
decide: every row starts with a suggested name and tick, and nothing is written until the
review dialog's **Write**.

- **One row per MAC.** Each row shows a tick, its status (New / Maybe / In database /
  Differs), the sheet row, MAC, IP, Name, Room, Role, Model, Jobsite, and the row's checks.
  The IP is shown for reference only: the database has no IP field. Columns sort; the
  filters show All / Ticked / Maybes / Flagged / In database, per table or all of them; and
  the search matches any cell of the source row.
- **Names are built from Room and Role.** Editing either rebuilds the name. Typing over the
  name keeps what you typed (shown bold) until **Reset names**. Equal names in a table are
  numbered `-2`, `-3` in row order. A name already used by another ticked row, or by a
  database entry on another MAC, is flagged rather than renumbered.
- **Bulk edits on the selection:**
  - Tick / Untick (Space toggles)
  - **Set room…**: select a rack's rows and type the rooms it serves (`10X` for rooms 101 to 109)
  - Set role…, Set jobsite…, Reset names
  - **Rows like this** (same model and description)

  These are also on the right-click menu.
- **The right pane** shows the selected row's source cells, the MAC's registered
  manufacturer, every check in full, and the database entry the row would replace.
- **The Database tab** (Ctrl+2; Ctrl+1 goes back to Import) is the database file, editable:
  - MAC, Name, Model and Jobsite edit in place.
  - Delete marks rows, and Revert undoes a row's edits.
  - Edited cells are bold (the old value is in the tooltip), deleted rows are struck
    through, and the tab title counts unsaved changes.
  - **Save changes…** reviews the edits, backs up the file and writes. If the file changed on
    disk since it was loaded, it asks before saving over that change.
- **On a row already in the database, each cell that differs from the entry is
  highlighted**: orange for a real difference, and pale for a difference only of spacing,
  dashes or case (`TSW 770` / `TSW-770`). Right-click a highlighted cell for **Use database
  value**.
- **A row already in the database is replaced only when you tick it.** The review dialog
  lists ticked rows with checks first, then every add and every replacement (old above new).
  Write backs the file up first.

## Command line

```
python -m iptable TABLE.xlsx [TABLE.xlsx ...]          show what an import would do; writes nothing
python -m iptable TABLE.xlsx --write                   add the ticked rows (backup written first)
python -m iptable TABLE.xlsx --write --maybe           also add the maybes
python -m iptable SWEEP.json                           a pinger -json sweep: what answered, new or known
pinger -json 172.22.0.1/24 | python -m iptable -       the same, piped in
    --db PATH      database to compare and write
    --jobsite X    jobsite for every row instead of the suggested one
    -v             also list rows already in the database, and rows with no MAC
```

On Windows, `ip-table-database-editor-cli.cmd` runs the same command line from any folder.

## How a table is read

- **Tables are found by their header row**: a row naming the MAC column and at least two
  other known columns (`iptable/tables.py` → `HEADERS` lists the spellings). A sheet can
  hold several tables, each running to the next header-like row. On a sheet with no header
  row, only the MAC column is guessed, from its contents.
- **Two common layouts read as they are**: a network sheet (`MAKE / MODEL` in one cell, a
  `PROJECT:` cell above the header) and a network summary (`Manufacturer`, `Model` and
  `Description` columns, one sheet per building). `tests/fixtures/` has an example of each.
  Username and password columns are never read.
- **A MAC is found in any spelling**: colons, dashes, dots, Cisco dots, or none. A MAC in
  another column (typed into a VLAN or Notes cell) becomes a maybe, since whose it is can't
  be told from the row.
- **Rows with no MAC are counted and skipped**, because the database is keyed by MAC.

## How a sweep is read

A sweep is a separate source, like opening another table.

- **One row per host with a MAC**, in IP order. Hosts with no MAC (routed, off the sweeping
  machine's subnet) can't go in the database and are listed in the right pane.
- **A new MAC starts with no name and unticked**, since a sweep gives only IP, MAC and
  manufacturer. Set Room and Role, or type a name, then tick it. Nothing is written while a
  ticked row has no name.
- **A MAC already in the database** shows its entry.
- **Checks**: one MAC on two IPs, a private (locally administered) address, a block in no
  registered range. A sweep taken without clearing the ARP cache is flagged, since a MAC can
  then belong to a device that has left that IP.
- Reads pinger's `-json` output as UTF-8, or the UTF-16 that Windows PowerShell 5.1 writes
  for `> sweep.json`.

## What gets ticked

A row is imported by default only when all three of these hold:

- its MAC is new to the database
- its manufacturer (the make column, or the MAC's registered owner) is in the AV list in
  `iptable/catalog.json`
- the row has an IP, a hostname or a room

Everything else is a **maybe**, shown with its reasons: spares not yet installed, service
laptops, client equipment. Edit `catalog.json` to fit your trade.

The suggested jobsite is the one the database already gives most of the table's MACs, else
the sheet's `PROJECT:` cell, else the workbook and sheet names.

## Names

Names say what a device is and where it is, not its hostname (`iptable/names.py`):
`B1-310-Ceiling-Mic`, `B1-310-Ceiling-Mic-Dante`, `B1-312-Touch-Panel`, `240-Camera`.

- **Building:** the sheet name when it's a building code (`B1`); otherwise nothing.
- **Room:** the number in the row's room cell (`C310` → `310`), else a room number in the
  description (`Touch Panel (240)`), else the room's words (`Event-Hall`).
- **Role:** the description, shortened (`Microphone` → `Mic`, `Digital Sound Processor` →
  `DSP`). With no description, the role `catalog.json` → `roles` gives the model
  (`DM-NVX-E30` → `Encoder`). With neither, the manufacturer.
- **Location:** always for mics and speakers (`Ceiling-Mic`). For anything else, only when
  two devices in one room share a role (`Wall-Encoder` beside the rack's `Encoder`).
- **Port:** `-Dante` when the description, hostname or model says Dante.
- Names still equal after that get `-2`, `-3`.

A rack device is named for the rack's room, because no column says which rooms a rack
serves. Use **Set room…** on the rack's rows.

## The checks

| Flag                                      | What it catches                                                                                                          |
| ----------------------------------------- | ------------------------------------------------------------------------------------------------------------------------ |
| hostname ends in another row's MAC digits | a row copied and not fully edited (a same-model match is allowed: a Dante port named after its control port)             |
| hostname ends in digits matching no MAC   | a mistyped MAC or hostname                                                                                               |
| the same MAC on two rows                  | across every sheet in the run                                                                                            |
| MAC block in no registered range          | a mistyped first half (or a block newer than your `oui.txt`)                                                             |
| table's make ≠ the MAC block's owner      | a typo, or an OEM part. Often correct: Dante cards carry Audinate MACs. Accepted pairs go in `catalog.json` → `oem_ouis` |
| unreadable MAC cell                       | wrong length, non-hex letters, a letter typed for a digit (with the likely MAC)                                          |
| locally administered or multicast         | not a fixed hardware address                                                                                             |

## Tests

```
python tests/iptable_tests.py      the reader, the plan, the command line
python tests/gui_tests.py          the window, offscreen (--shot DIR saves screenshots)
python tests/make_fixtures.py      rewrites tests/fixtures/ (needs pinger's oui.txt)
```

The tests run on the made-up workbooks in `tests/fixtures/`, each mistake planted on a known
row, and on `oui-subset.txt`, a fixed copy of the manufacturer blocks they use. They need
nothing outside this repository.

## Source layout

| File                   | What it holds                                                           |
| ---------------------- | ----------------------------------------------------------------------- |
| `iptable/tables.py`    | finding tables and header roles, reading rows, splitting make and model |
| `iptable/macs.py`      | finding, normalizing and diagnosing MACs                                |
| `iptable/oui.py`       | manufacturer lookup from pinger's `oui.txt`                             |
| `iptable/paths.py`     | where pinger's files are                                                |
| `iptable/plan.py`      | proposals, ticking, checks, the jobsite suggestion, sweeps              |
| `iptable/sweep.py`     | reading pinger `-json` output                                           |
| `iptable/names.py`     | role-style names                                                        |
| `iptable/catalog.json` | AV manufacturers, model → role patterns, accepted OEM MAC blocks        |
| `iptable/db.py`        | loading, merging and writing the database (backup first)                |
| `iptable/gui.py`       | the window                                                              |
| `iptable/launch.py`    | reads a piped sweep and starts the window detached                      |
| `iptable/__main__.py`  | the command line                                                        |

## Related tools

- [pinger](https://github.com/JaimeMaertz/pinger): the subnet sweep and the database's reader
- [netdiagram](https://github.com/JaimeMaertz/netdiagram): draws switches and devices, named from the same database
- [ip-switcher](https://github.com/JaimeMaertz/ip-switcher): a tray app for switching your adapter between DHCP and static presets

## License

Source-available: an MIT-based license with a no-resale condition (see [LICENSE](LICENSE)),
not an OSI open-source license. You're free to use it and modify it,
including at work and on paid jobs, to share it free of charge, and to charge for installing,
training on or supporting it. Keep the copyright notice and the license with any copy.

What needs permission is selling it, or selling a product or hosted service built on it. For
that, contact JaimeMaertz@gmail.com.

Contributions are welcome: see [CONTRIBUTING.md](CONTRIBUTING.md).
