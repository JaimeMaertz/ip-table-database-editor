"""Tests for the importer window, offscreen, on the made-up workbooks in tests/fixtures/ and a copy of the
fixtures' hand-named devices.json (never a real database).

    python3 tests/gui_tests.py [--shot DIR]    exit 0 when every test passes, 1 on a failure, 3 without PySide6;
                                               --shot saves screenshots of the window and the review dialog
"""
import json
import os
import shutil
import sys
import tempfile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURES = os.path.join(HERE, "fixtures")
os.environ["IPTABLE_OUI"] = os.path.join(FIXTURES, "oui-subset.txt")
sys.path.insert(0, os.path.dirname(HERE))
try:
    from PySide6.QtCore import QItemSelectionModel, QSettings, Qt
    from PySide6.QtWidgets import QApplication
except ImportError:
    print("PySide6 is not installed")
    sys.exit(3)

SETTINGS = tempfile.mkdtemp(prefix="iptable-gui-tests-")    # never the user's own window settings
QSettings.setDefaultFormat(QSettings.IniFormat)
QSettings.setPath(QSettings.IniFormat, QSettings.UserScope, SETTINGS)
from iptable import db  # noqa: E402
from iptable.gui import (CHECKS, IP, JOBSITE, MODEL, NAME, ROOM, TICK, MainWindow, ReviewDialog, category)  # noqa: E402

SUMMARY = os.path.join(FIXTURES, "example-network-summary.xlsx")
SHEET = os.path.join(FIXTURES, "example-network-sheet.xlsx")
SINGLE = os.path.join(FIXTURES, "example-single-site.xlsx")
DB = os.path.join(FIXTURES, "devices.json")
RESULTS = []
SHOT = None
SWEEP = os.path.join(FIXTURES, "sweep.json")


def check(name, ok, detail=""):
    RESULTS.append(ok)
    print(("PASS  " if ok else "FAIL  ") + name + ("" if ok else "   " + str(detail)[:300]))


def row_of(win, where):
    return next(i for i, p in enumerate(win.model.rows) if p.where == where)


def select(win, rows):
    sel = win.view.selectionModel()
    sel.clearSelection()
    for r in rows:
        sel.select(win.proxy.mapFromSource(win.model.index(r, 0)), QItemSelectionModel.Select | QItemSelectionModel.Rows)


def db_row(win, mac):
    return next(i for i, r in enumerate(win.db_model.rows) if db.key(r["entry"]) == mac)


def database_edit_tests(win, dbfile):
    dm = win.db_model
    before = db.load(dbfile)
    cam = db_row(win, "00:23:A8:00:60:01")                  # B1-101-PTZ-Camera, also B1 r13 in the open table
    check("database tab starts clean", not dm.dirty() and not win.db_save_btn.isEnabled())
    dm.setData(dm.index(cam, 1), "B1-101-Cam")
    check("an edited cell is bold and the tab says unsaved", dm.data(dm.index(cam, 1), Qt.FontRole).bold()
          and "1 unsaved" in win.tabs.tabText(1), win.tabs.tabText(1))
    check("nothing is written before Save", db.load(dbfile) == before)
    check("an invalid MAC is refused", not dm.setData(dm.index(cam, 0), "00:23:A8:00:60") and db.key(dm.rows[cam]["entry"]) == "00:23:A8:00:60:01")
    check("a MAC already in the database is refused", not dm.setData(dm.index(cam, 0), "C4:42:68:00:20:01"))
    tp = db_row(win, "C4:42:68:00:FE:47")                   # Lab Touch Panel
    check("a MAC typed any way is normalized", dm.setData(dm.index(tp, 0), "c442.6800.fe48")
          and dm.rows[tp]["entry"]["mac"] == "C4:42:68:00:FE:48")
    dm.revert([tp])
    check("revert puts the row back", dm.rows[tp]["entry"] == dm.rows[tp]["orig"])
    ic = [i for i, r in enumerate(dm.rows) if r["entry"].get("jobsite") == "Example Lab"]
    dm.set_field(ic, "jobsite", "Example Lab 2")
    check("set jobsite for a selection", len(ic) == 2 and all(dm.rows[i]["entry"]["jobsite"] == "Example Lab 2" for i in ic)
          and win.db_jobsite.findData("Example Lab 2") >= 0)
    tesira = db_row(win, "78:45:01:00:30:01")               # B1-10X-Tesira, B1 r7 in the open table
    dm.delete([tesira])
    check("a deleted row is struck through", dm.data(dm.index(tesira, 1), Qt.FontRole).strikeOut())
    text = win.review_text()
    check("an import review lists the Database tab's edits too", "Also written, from the Database tab:" in text
          and "Delete 1 database entries:" in text, text[-600:])

    if SHOT:
        win.tabs.setCurrentIndex(1)
        win.db_view.selectRow(win.db_proxy.mapFromSource(dm.index(cam, 0)).row())
        win.grab().save(os.path.join(SHOT, "database-edits.png"))
        win.tabs.setCurrentIndex(0)

    # Something else changes the file: a save asks first, and No writes nothing.
    os.utime(dbfile, (win.db_mtime + 10, win.db_mtime + 10))
    from PySide6.QtWidgets import QMessageBox
    QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.No)
    check("a file changed on disk is not overwritten without a yes", win.save_database() is None and db.load(dbfile) == before)
    QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.Yes)
    backup = win.save_database()
    after = db.load(dbfile)
    check("save writes the edits", backup and next(e for e in after if db.key(e) == "00:23:A8:00:60:01")["name"] == "B1-101-Cam"
          and not any(db.key(e) == "78:45:01:00:30:01" for e in after) and len(after) == len(before) - 1)
    check("everything else unchanged", [e for e in after if db.key(e) not in ("00:23:A8:00:60:01",) and e.get("jobsite") != "Example Lab 2"]
          == [e for e in before if db.key(e) not in ("00:23:A8:00:60:01", "78:45:01:00:30:01") and e.get("jobsite") != "Example Lab"])
    check("the tab is clean after save", not dm.dirty() and "unsaved" not in win.tabs.tabText(1))
    r27 = row_of(win, "B1 r7")
    check("an import row whose entry was deleted is new again", win.model.rows[r27].status == "new")
    check("an import row whose entry was renamed differs", win.model.rows[row_of(win, "B1 r13")].status == "differs")


def sweep_tests(win, dbfile):
    from iptable.gui import ROLE
    m = win.model
    check("opens a sweep file", win.open_workbook(SWEEP) and len(m.rows) == 9, len(m.rows))
    check("the table chooser names the sweep", win.table_box.itemText(1) == "sweep.json (9 MACs)", win.table_box.itemText(1))
    check("the notes count what answered and what was left out", "10 of 254 hosts" in win.details.toPlainText()
          and "10.1.1.5" in win.details.toPlainText(), win.details.toPlainText()[:300])
    a, b = row_of(win, "172.22.0.31"), row_of(win, "172.22.0.32")
    known = row_of(win, "172.22.0.10")
    m.set_part([a], "room", "310")
    m.setData(m.index(a, ROLE), "Encoder")
    check("Room and Role name a sweep row", m.rows[a].name == "310-Encoder", m.rows[a].name)
    in_db = {db.key(e): e for e in db.load(dbfile)}["C4:42:68:00:20:01"]["name"]
    check("the other sweep rows keep their names", m.rows[b].name == "" and m.rows[known].name == in_db,
          (m.rows[b].name, m.rows[known].name, in_db))
    m.set_ticked([a, b], True)
    check("a ticked row with no name is flagged", any("no name" in c for c in m.checks(m.rows[b])), m.checks(m.rows[b]))
    before = db.load(dbfile)
    check("nothing is written while a ticked row has no name", win.unnamed() and win.write() is None and db.load(dbfile) == before)
    m.set_ticked([b], False)
    win.write()
    after = {db.key(e): e for e in db.load(dbfile)}
    check("a named sweep row is written", after.get("00:10:7F:12:34:56") == {"mac": "00:10:7F:12:34:56", "name": "310-Encoder",
          "jobsite": "Example Campus B1"}, after.get("00:10:7F:12:34:56"))
    check("a sweep piped in opens", win.open_workbook("-", open(SWEEP, "rb").read()) and len(m.rows) == 9
          and win.file_label.text() == "Sweep, piped in")
    from PySide6.QtWidgets import QMessageBox
    warned = []
    QMessageBox.warning = staticmethod(lambda *a, **k: warned.append(a[2]))
    check("a device database opened as a sweep is refused", not win.open_workbook(dbfile) and "device database" in warned[-1],
          warned)
    if SHOT:
        win.open_workbook(SWEEP)
        win.grab().save(os.path.join(SHOT, "sweep.png"))


def run(shot_dir):
    tmp = tempfile.mkdtemp()
    dbfile = os.path.join(tmp, "devices.json")
    shutil.copy(DB, dbfile)
    before = db.load(dbfile)
    win = MainWindow(dbfile)
    win.show()
    check("opens the summary workbook", win.open_workbook(SUMMARY))
    m = win.model
    check("one row per MAC", len(m.rows) == 30, len(m.rows))
    check("table chooser lists the device sheets", win.table_box.count() == 4, win.table_box.count())

    # The rack case: B1's switch, processor and DSP named for the rooms they serve.
    rack = [row_of(win, "B1 r5"), row_of(win, "B1 r6"), row_of(win, "B1 r7")]
    check("rack rows start with the rack's room", [m.rows[r].name for r in rack] ==
          ["B1-101-Switch", "B1-101-Processor", "B1-101-DSP"], [m.rows[r].name for r in rack])
    m.set_part(rack, "room", "10X")
    check("set room renames every selected row", [m.rows[r].name for r in rack] ==
          ["B1-10X-Switch", "B1-10X-Processor", "B1-10X-DSP"], [m.rows[r].name for r in rack])

    # Typing over a name keeps it, a room edit doesn't touch it, Reset names gives it back.
    r = row_of(win, "B1 r6")
    m.setData(m.index(r, NAME), "B1-10X-RMC4")
    m.set_part([r], "room", "101")
    check("a typed name survives a room edit", m.rows[r].name == "B1-10X-RMC4", m.rows[r].name)
    check("a typed name is bold", m.data(m.index(r, NAME), Qt.FontRole).bold())
    m.reset_names([r])
    check("reset names builds it from the parts", m.rows[r].name == "B1-101-Processor", m.rows[r].name)

    # Ticking a differs row means replace; a database name on another MAC is flagged, not renumbered.
    m.setData(m.index(r, ROOM), "105")
    m.setData(m.index(r, TICK), Qt.Checked.value, Qt.CheckStateRole)
    adds, repl, _ = win.pending()
    check("ticking a row already in the database replaces it", m.rows[r] in repl)
    other = row_of(win, "B1 r8")
    m.setData(m.index(other, NAME), "B1-10X-Switch")
    m.set_ticked([other], True)
    check("a name the database has on another MAC is flagged", any("already has B1-10X-Switch" in c for c in m.checks(m.rows[other])),
          m.checks(m.rows[other]))
    check("the flag shows in the Checks column", "⚠" in m.data(m.index(other, CHECKS)))

    # A row already in the database highlights only the cells that differ from its entry.
    dsp, tp = row_of(win, "B1 r7"), row_of(win, "B1 r15")
    lit = lambda r, c: m.data(m.index(r, c), Qt.BackgroundRole) is not None   # noqa: E731
    check("differing name and model cells are highlighted, the matching jobsite is not",
          lit(dsp, NAME) and lit(dsp, MODEL) and not lit(dsp, JOBSITE), (lit(dsp, NAME), lit(dsp, MODEL), lit(dsp, JOBSITE)))
    check("the tooltip gives the database's value", m.data(m.index(dsp, MODEL), Qt.ToolTipRole).startswith("database: X400"))
    check("a formatting-only difference says so, in a paler shade", "only spacing" in m.data(m.index(tp, MODEL), Qt.ToolTipRole)
          and m.data(m.index(tp, MODEL), Qt.BackgroundRole) != m.data(m.index(dsp, MODEL), Qt.BackgroundRole))
    check("a new row highlights nothing", not any(lit(row_of(win, "B1 r12"), c) for c in (NAME, MODEL, JOBSITE)))
    m.setData(m.index(tp, MODEL), "TSW 770")
    check("editing a cell to match clears its highlight", not lit(tp, MODEL))

    # Use the database's value for one column of a differing row.
    from PySide6.QtWidgets import QMenu
    import iptable.gui as gui_mod
    shown = []

    class NoPopupMenu(QMenu):                   # records the menu's items instead of opening it
        def exec(self, *a):
            shown.extend(x.text() for x in self.actions() if x.text())
    real_menu, gui_mod.QMenu = gui_mod.QMenu, NoPopupMenu
    select(win, [dsp])
    win.view.scrollTo(win.proxy.mapFromSource(m.index(dsp, MODEL)))
    win.context_menu(win.view.visualRect(win.proxy.mapFromSource(m.index(dsp, MODEL))).center())
    check("right-click on a differing cell offers its database value", "Use database value for Model: X400" in shown, shown[:3])
    shown.clear()
    win.context_menu(win.view.visualRect(win.proxy.mapFromSource(m.index(dsp, JOBSITE))).center())
    check("no offer on a cell that matches", not any(s.startswith("Use database") for s in shown), shown[:3])
    gui_mod.QMenu = real_menu
    m.use_database([dsp], MODEL)
    check("only that column takes the database value", m.rows[dsp].model == "X400" and m.rows[dsp].name == "B1-10X-DSP")
    m.use_database([dsp], NAME)
    check("then the name, kept as if typed", m.rows[dsp].name == "B1-10X-Tesira" and m.rows[dsp].locked)
    check("a row that now matches its entry is In database", category(m.rows[dsp]) == "same" and not m.rows[dsp].ticked)
    enc = [row_of(win, "B1 r9"), row_of(win, "B1 r10"), tp]   # the touch panel's model already matches
    check("over a selection, only the rows that differ change", m.use_database(enc, MODEL) == enc[:2]
          and [m.rows[r].model for r in enc[:2]] == ["NVX E30", "NVX E30"])

    # IP for reference: as the table has it, not editable, not written, sorted as numbers.
    check("IP column shows the table's IP", m.data(m.index(row_of(win, "B1 r6"), IP)) == "172.22.0.5")
    check("IP is not editable", not (m.flags(m.index(0, IP)) & Qt.ItemIsEditable))
    check("IP is not part of the entry", "ip" not in m.rows[row_of(win, "B1 r6")].entry())
    win.proxy.set(table=win.found[[t.sheet for t in win.found].index("B1")])
    win.view.sortByColumn(IP, Qt.AscendingOrder)
    ips = [win.proxy.data(win.proxy.index(i, IP)) for i in range(win.proxy.rowCount())]
    real = [x for x in ips if x]
    check("IP sorts numerically, blanks last", real == sorted(real, key=lambda s: [int(n) for n in s.split(".")])
          and ips.index("172.22.0.5") < ips.index("172.22.0.21") and (not ips[-1] or "" not in ips), ips[:8])
    win.view.sortByColumn(-1, Qt.AscendingOrder)
    win.proxy.set(table=None)

    # Filters and search.
    win.proxy.set(mode="flagged")
    shown = {win.proxy.mapToSource(win.proxy.index(i, 0)).row() for i in range(win.proxy.rowCount())}
    check("Flagged shows the copied camera hostnames", row_of(win, "B1 r13") in shown and row_of(win, "B1 r14") in shown)
    win.proxy.set(mode="all", text="CV605")
    check("search finds by any cell", win.proxy.rowCount() == 2, win.proxy.rowCount())
    win.proxy.set(text="")

    select(win, [row_of(win, "B1 r9")])
    win.select_like()
    like = {win.proxy.mapToSource(i).row() for i in win.view.selectionModel().selectedRows()}
    check("rows like this: same model and description", like == {row_of(win, "B1 r9"), row_of(win, "B1 r10")}, like)

    # The database tab: everything on disk, filterable, linked to the open table.
    dp, dm = win.db_proxy, win.db_model
    check("database tab shows every entry", dp.rowCount() == len(before) and win.tabs.tabText(1) == "Database (%d)" % len(before),
          (dp.rowCount(), win.tabs.tabText(1)))
    win.db_jobsite.setCurrentIndex(win.db_jobsite.findData("Example Campus B1"))
    want = sum(1 for e in before if e.get("jobsite") == "Example Campus B1")
    check("jobsite filter", dp.rowCount() == want and win.db_count.text() == "%d of %d shown" % (want, len(before)),
          (dp.rowCount(), want, win.db_count.text()))
    win.db_jobsite.setCurrentIndex(0)
    win.db_search.setText("c44268002001")
    check("search finds a MAC typed without colons", dp.rowCount() == 1)
    cells = [dm.data(dm.index(dp.mapToSource(dp.index(0, 0)).row(), c)) for c in range(6)]
    check("row shows manufacturer and where it is in the open table",
          cells[4].upper().startswith("CRESTRON") and cells[5] == "B1 r6", cells)
    win.goto_import_row(dp.index(0, 0))
    picked = [win.proxy.mapToSource(i).row() for i in win.view.selectionModel().selectedRows()]
    check("double-click jumps to the row on the Import tab", win.tabs.currentIndex() == 0 and picked == [row_of(win, "B1 r6")],
          (win.tabs.currentIndex(), picked))
    win.db_search.setText("")
    win.db_view.sortByColumn(1, Qt.AscendingOrder)
    names_shown = [dm.rows[dp.mapToSource(dp.index(i, 0)).row()]["entry"].get("name", "") for i in range(dp.rowCount())]
    check("sorts by name", names_shown == sorted(names_shown, key=str.lower), names_shown[:5])
    if shot_dir:
        os.makedirs(shot_dir, exist_ok=True)
        win.tabs.setCurrentIndex(1)
        win.resize(1500, 860)
        win.grab().save(os.path.join(shot_dir, "database.png"))
        win.tabs.setCurrentIndex(0)

    # Review and write.
    m.set_ticked([other], False)
    text = win.review_text()
    check("review lists adds and replacements", "Add %d:" % len(win.pending()[0]) in text and "Replace 1:" in text, text[:200])
    if shot_dir:
        os.makedirs(shot_dir, exist_ok=True)
        win.resize(1500, 860)
        select(win, [row_of(win, "B1 r13")])
        win.show_details()
        win.grab().save(os.path.join(shot_dir, "window.png"))
        win.proxy.set(mode="db")
        win.view.clearSelection()
        win.view.horizontalScrollBar().setValue(0)
        win.grab().save(os.path.join(shot_dir, "differs.png"))
        win.proxy.set(mode="all")
        dlg = ReviewDialog(win, text, True)
        dlg.show()
        dlg.grab().save(os.path.join(shot_dir, "review.png"))
    adds, repl, _ = win.pending()
    backup = win.write()
    after = db.load(dbfile)
    check("write leaves a backup", backup and os.path.exists(backup))
    check("write replaces the ticked database row in place",
          after[[db.key(e) for e in before].index(m.rows[r].mac)]["name"] == "B1-105-Processor")
    check("write adds the ticked new rows at the end", len(after) == len(before) + len(adds), (len(after), len(before), len(adds)))
    check("written rows become In database", category(m.rows[r]) == "same" and not win.pending()[0])
    untouched = [e for e in before if db.key(e) != m.rows[r].mac]
    check("every other database entry unchanged", all(e in after for e in untouched))
    check("database tab follows the write", win.db_proxy.rowCount() == len(after) and win.tabs.tabText(1) == "Database (%d)" % len(after))

    database_edit_tests(win, dbfile)
    # A room change over a selection keeps equal names numbered, as the first suggestion did.
    check("opens the single-site workbook", win.open_workbook(SINGLE))
    m = win.model
    rows = [i for i, p in enumerate(m.rows) if p.parts["room"] == "260"]
    before_names = [m.rows[i].name for i in rows]
    m.setData(m.index(rows[0], NAME), "260-Boardroom-Front-Panel")          # typed: must stay as typed
    m.set_part(rows, "room", "260 Boardroom")
    after_names = [m.rows[i].name for i in rows]
    want = [after_names[0]] + [n.replace("260-", "260-Boardroom-", 1) for n in before_names[1:]]
    check("set room keeps -2, -3 on equal names", after_names == want, list(zip(before_names, after_names)))
    check("a typed name is untouched by the room change", after_names[0] == "260-Boardroom-Front-Panel")
    check("a row with no model or description is named for its maker", "260-Boardroom-Shure" in after_names
          and "260-Boardroom-Shure-2" in after_names, after_names)
    m.reset_names(rows[:1])
    check("reset names renumbers with the rest", [m.rows[i].name for i in rows][:2] == ["260-Boardroom-Touch-Panel", "260-Boardroom-Touch-Panel-2"],
          [m.rows[i].name for i in rows][:2])

    check("opens the network-sheet workbook", win.open_workbook(SHEET))
    sweep_tests(win, dbfile)
    win.close()
    shutil.rmtree(tmp)


if __name__ == "__main__":
    shot = SHOT = sys.argv[sys.argv.index("--shot") + 1] if "--shot" in sys.argv else None
    app = QApplication(sys.argv[:1])
    app.setStyle("Fusion")
    from PySide6.QtWidgets import QMessageBox
    QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.Yes)   # no dialogs in tests
    run(shot)
    shutil.rmtree(SETTINGS, True)
    print("%d/%d passed" % (sum(RESULTS), len(RESULTS)))
    sys.exit(0 if all(RESULTS) else 1)
