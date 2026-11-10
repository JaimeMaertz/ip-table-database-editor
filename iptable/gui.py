"""The IP table database editor's window: review an IP table row by row, fix names, decide what goes into
pinger's database, and edit the database itself.

    python -m iptable.gui [TABLE.xlsx | SWEEP.json | -] [--db devices.json]
    pinger -json 172.22.0.1/24 | python -m iptable.gui              a sweep, piped in (here, the prompt waits
                                                                    for the window; iptable.launch doesn't)

The tool suggests, you decide. Every row starts with a suggested name and a suggested tick; nothing is
written until the review dialog's Write. A row already in the database is replaced only when you tick it.
A name is built from its Room and Role cells, so editing either builds it again; typing over the name
itself keeps what you typed (shown in bold) until Reset names.
"""
import argparse
import os
import re
import sys

from PySide6.QtCore import (QAbstractTableModel, QItemSelectionModel, QModelIndex, QSettings, QSortFilterProxyModel,
                            Qt, Signal)
from PySide6.QtGui import QAction, QColor, QFont, QKeySequence, QShortcut
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QButtonGroup, QComboBox, QDialog, QDialogButtonBox,
                               QFileDialog, QHBoxLayout, QHeaderView, QInputDialog, QLabel, QLineEdit, QMainWindow,
                               QMenu, QMessageBox, QPlainTextEdit, QPushButton, QSplitter, QTableView, QTabWidget,
                               QTextBrowser, QVBoxLayout, QWidget)

from . import db, launch, macs, names, sweep
from .oui import Registry
from .plan import Planner

COLUMNS = ["", "Status", "Row", "MAC", "IP", "Name", "Room", "Role", "Model", "Jobsite", "Checks"]
TICK, STATUS, WHERE, MAC, IP, NAME, ROOM, ROLE, MODEL, JOBSITE, CHECKS = range(len(COLUMNS))
EDITABLE = {NAME, ROOM, ROLE, MODEL, JOBSITE}

CATEGORY = {"new": "New", "maybe": "Maybe", "same": "In database", "differs": "Differs"}
OPENS = (".xlsx", ".xlsm", ".json")         # what Open and a drop take: IP tables and pinger -json sweeps
FILTERS = [("all", "All"), ("write", "Ticked"), ("maybe", "Maybes"), ("flagged", "Flagged"), ("db", "In database")]


def category(p):
    if p.status == "new":
        return "maybe" if p.maybe else "new"
    return p.status


def html_escape(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


class ImportModel(QAbstractTableModel):
    """One row per proposal. Editing Room or Role rebuilds the name unless it was typed over (locked)."""
    changed = Signal()

    def __init__(self):
        super().__init__()
        self.rows = []
        self.db_names = {}          # lower-case name -> MAC, for the database entries
        self.conflicts = {}         # id(proposal) -> [text]
        self.colors = {}

    def load(self, proposals, entries):
        self.beginResetModel()
        self.rows = proposals
        for p in proposals:
            p.locked = False
        self.set_database(entries, reset=False)
        self.endResetModel()
        self.changed.emit()

    def set_database(self, entries, reset=True):
        self.db_names = {}
        for e in entries:
            if e.get("name"):
                self.db_names.setdefault(e["name"].lower(), db.key(e))
        self.refresh_conflicts()

    def set_colors(self, dark):
        self.colors = ({"new": "#1f3d24", "maybe": "#4a3d12", "differs": "#173a5c", "same": None,
                        "diff": "#5c3a12", "diff_format": "#3a2e22", "warn": "#ef5350", "hint": "#9a9a9a"} if dark else
                       {"new": "#e3f4e3", "maybe": "#fff4cc", "differs": "#e1eefb", "same": None,
                        "diff": "#ffd9a8", "diff_format": "#fdf0e0", "warn": "#c62828", "hint": "#757575"})

    # The duplicate-name check runs on every edit: a name already used by another ticked row, or by a
    # database entry for a different MAC, is flagged rather than renumbered.
    def refresh_conflicts(self):
        ticked = {}
        for p in self.rows:
            if p.ticked:
                ticked.setdefault(p.name.lower(), []).append(p)
        self.conflicts = {}
        for p in self.rows:
            out = []
            owner = self.db_names.get(p.name.lower()) if p.name else None
            if owner and owner != p.mac:
                out.append("the database already has %s on %s" % (p.name, owner))
            if p.ticked and not p.name.strip():
                out.append("no name: set Room and Role or type one before writing")
            others = [o for o in ticked.get(p.name.lower(), []) if o is not p and p.name.strip()]
            if others and p.ticked:
                out.append("%s is also the name of %s" % (p.name, ", ".join(o.where for o in others)))
            self.conflicts[id(p)] = out

    def checks(self, p):
        return list(p.flags) + self.conflicts.get(id(p), [])

    DIFF_FIELDS = {NAME: "name", MODEL: "model", JOBSITE: "jobsite"}

    @staticmethod
    def _loose(s):
        return re.sub(r"[\s\-_]", "", str(s)).lower()

    def differs(self, p, col):
        """(database value, only-formatting) when this cell of a row already in the database differs from
        the entry, else None."""
        field = self.DIFF_FIELDS.get(col)
        if not field or not p.existing:
            return None
        mine, theirs = getattr(p, field) or "", p.existing.get(field, "")
        if mine == theirs:
            return None
        return theirs, self._loose(mine) == self._loose(theirs)

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(COLUMNS)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return COLUMNS[section]
        if orientation == Qt.Horizontal and role == Qt.ToolTipRole:
            return {TICK: "Ticked rows are written: a new row is added, an existing one replaced",
                    IP: "The IP the table gives, for reference; it is not stored in the database",
                    NAME: "Built from Room and Role; type over it to keep your own (bold)",
                    ROOM: "The room the device serves; editing it rebuilds the name",
                    ROLE: "What the device is; editing it rebuilds the name"}.get(section)
        return None

    def flags(self, index):
        f = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        p = self.rows[index.row()]
        if index.column() == TICK and p.status != "same":
            f |= Qt.ItemIsUserCheckable
        if index.column() in EDITABLE:
            f |= Qt.ItemIsEditable
        return f

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        p, col = self.rows[index.row()], index.column()
        if role == Qt.CheckStateRole and col == TICK:
            if p.status == "same":
                return None
            return Qt.Checked if p.ticked else Qt.Unchecked
        if role in (Qt.DisplayRole, Qt.EditRole):
            checks = self.checks(p)
            return {
                STATUS: CATEGORY[category(p)], WHERE: p.where, MAC: p.mac, IP: p.row.values.get("ip", ""), NAME: p.name,
                ROOM: p.parts.get("room", ""), ROLE: p.parts.get("role", ""), MODEL: p.model, JOBSITE: p.jobsite,
                CHECKS: ("⚠ %d: %s" % (len(checks), checks[0])) if checks else ("maybe: " + p.maybe[0] if p.maybe else ""),
            }.get(col)
        diff = self.differs(p, col)
        if role == Qt.BackgroundRole and diff:
            return QColor(self.colors["diff_format" if diff[1] else "diff"])
        if role == Qt.ToolTipRole and diff:
            return "database: %s%s" % (diff[0] or "(empty)", "\n(only spacing, dashes or case differ)" if diff[1] else "")
        if role == Qt.ToolTipRole:
            lines = ["⚠ " + c for c in self.checks(p)] + ["maybe: " + m for m in p.maybe]
            if p.existing:
                lines.append("in the database as %s" % p.existing.get("name", ""))
            return "\n".join(lines) or None
        if role == Qt.FontRole and col == NAME and getattr(p, "locked", False):
            f = QFont()
            f.setBold(True)
            return f
        if role == Qt.ForegroundRole and col == IP:
            return QColor(self.colors["hint"])
        if role == Qt.ForegroundRole and col == CHECKS:
            return QColor(self.colors["warn"] if self.checks(p) else self.colors["hint"])
        if role == Qt.BackgroundRole and col == STATUS:
            c = self.colors.get(category(p))
            return QColor(c) if c else None
        return None

    def setData(self, index, value, role=Qt.EditRole):
        p, col = self.rows[index.row()], index.column()
        if role == Qt.CheckStateRole and col == TICK:
            self.set_ticked([index.row()], Qt.CheckState(value) == Qt.Checked)
            return True
        if role != Qt.EditRole or col not in EDITABLE:
            return False
        value = str(value).strip()
        if col == NAME:
            if not value or value == p.name:
                return False
            p.name, p.locked = value, True
        elif col == ROOM:
            self.set_part([index.row()], "room", value)
            return True
        elif col == ROLE:
            self.set_part([index.row()], "role", value)
            return True
        elif col == MODEL:
            p.model = value
        elif col == JOBSITE:
            p.jobsite = value
        self._rows_changed([index.row()])
        return True

    # Bulk edits, used by the toolbar for every selected row.
    def set_ticked(self, rows, on):
        for r in rows:
            if self.rows[r].status != "same":
                self.rows[r].ticked = on
        self._rows_changed(rows, all_conflicts=True)

    def set_part(self, rows, part, value):
        for r in rows:
            self.rows[r].parts[part] = value
        self.renumber({id(self.rows[r].table) for r in rows})
        self._rows_changed(rows, all_conflicts=True)

    def renumber(self, tables):
        """Build every name in these tables from its parts again, numbering equal names -2, -3 in table order,
        as the first suggestion does. A name typed over (locked) is left as it is but keeps its place in the
        count, so typing over the first of two touch panels doesn't turn the second's -2 into nothing."""
        for t in tables:
            seen = {}
            for p in self.rows:
                if id(p.table) != t:
                    continue
                base = names.compose(p.parts)
                if not base:            # nothing to build from (a sweep row before Room or Role): keep the name
                    continue
                n = seen[base.lower()] = seen.get(base.lower(), 0) + 1
                if not getattr(p, "locked", False):
                    p.name = base if n == 1 else "%s-%d" % (base, n)

    def set_field(self, rows, field, value):
        for r in rows:
            setattr(self.rows[r], field, value)
        self._rows_changed(rows)

    def use_database(self, rows, col):
        """Take the database's value for one column (Name, Model or Jobsite) on the rows where it differs. A
        name taken this way is kept like a typed one; a row that then matches its entry is In database."""
        field = self.DIFF_FIELDS.get(col)
        changed = []
        for r in rows:
            p = self.rows[r]
            if not field or not self.differs(p, col):
                continue
            setattr(p, field, p.existing.get(field, ""))
            if field == "name":
                p.locked = True
            if all(p.existing.get(k, "") == v for k, v in p.entry().items() if k != "mac"):
                p.status, p.ticked = "same", False
            changed.append(r)
        if changed:
            self._rows_changed(changed, all_conflicts=True)
        return changed

    def reset_names(self, rows):
        for r in rows:
            self.rows[r].locked = False
        self.renumber({id(self.rows[r].table) for r in rows})
        self._rows_changed(rows, all_conflicts=True)

    def _rows_changed(self, rows, all_conflicts=True):
        self.refresh_conflicts()
        if rows and not all_conflicts:
            for r in rows:
                self.dataChanged.emit(self.index(r, 0), self.index(r, len(COLUMNS) - 1))
        else:
            self.dataChanged.emit(self.index(0, 0), self.index(len(self.rows) - 1, len(COLUMNS) - 1))
        self.changed.emit()


class ImportFilter(QSortFilterProxyModel):
    def __init__(self):
        super().__init__()
        self.mode, self.table, self.text = "all", None, ""

    def set(self, **kw):
        self.__dict__.update(kw)
        self.invalidateFilter()

    @staticmethod
    def ip_key(text):
        """172.22.0.9 before 172.22.0.10; a cell that is not an address ("Self-Assigned", blank) sorts last."""
        parts = re.findall(r"\d+", str(text))
        return (0, [int(x) for x in parts]) if len(parts) == 4 else (1, [str(text).lower()])

    def lessThan(self, left, right):
        if left.column() == IP:
            m = self.sourceModel()
            return self.ip_key(m.data(left)) < self.ip_key(m.data(right))
        return super().lessThan(left, right)

    def filterAcceptsRow(self, row, parent):
        m = self.sourceModel()
        p = m.rows[row]
        if self.table is not None and p.table is not self.table:
            return False
        cat = category(p)
        if self.mode == "write" and not p.ticked:
            return False
        if self.mode == "maybe" and cat != "maybe":
            return False
        if self.mode == "flagged" and not m.checks(p):
            return False
        if self.mode == "db" and cat not in ("same", "differs"):
            return False
        if self.text:
            hay = " ".join([p.mac, p.name, p.model, p.jobsite, p.where] + list(p.row.values.values())).lower()
            if not all(w in hay for w in self.text.lower().split()):
                return False
        return True


DB_COLUMNS = ["MAC", "Name", "Model", "Jobsite", "Manufacturer", "In open table"]


class DatabaseModel(QAbstractTableModel):
    """A working copy of the database, with each MAC's registered manufacturer and where the MAC appears in the
    open workbook. MAC, Name, Model and Jobsite edit in place and Delete marks a row; nothing reaches the
    file until it is saved. An edited cell is bold, a deleted row struck through."""
    changed = Signal()
    refused = Signal(str)
    FIELDS = ["mac", "name", "model", "jobsite"]

    def __init__(self, registry):
        super().__init__()
        self.registry = registry
        self.rows, self.where, self.colors = [], {}, {}

    def load(self, entries, where):
        """Start over from the entries on disk; pending edits are dropped."""
        self.beginResetModel()
        self.rows = [{"entry": dict(e), "orig": e, "deleted": False} for e in entries]
        self.where = where
        self.endResetModel()
        self.changed.emit()

    def set_where(self, where):
        self.where = where
        self._all_changed()

    def entries(self):
        return [r["entry"] for r in self.rows]

    def working(self):
        """The database as it would be saved."""
        return [r["entry"] for r in self.rows if not r["deleted"]]

    def edits(self):
        """(edited rows, deleted rows) as (old entry, new entry) pairs."""
        edited = [(r["orig"], r["entry"]) for r in self.rows if not r["deleted"] and r["entry"] != r["orig"]]
        deleted = [(r["orig"], r["entry"]) for r in self.rows if r["deleted"]]
        return edited, deleted

    def dirty(self):
        edited, deleted = self.edits()
        return bool(edited or deleted)

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(DB_COLUMNS)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return DB_COLUMNS[section]
        return None

    def flags(self, index):
        f = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if index.column() < len(self.FIELDS) and not self.rows[index.row()]["deleted"]:
            f |= Qt.ItemIsEditable
        return f

    def _mac(self, e):
        return db.key(e) or e.get("mac", "")

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        r, col = self.rows[index.row()], index.column()
        e, mac = r["entry"], self._mac(r["entry"])
        if role in (Qt.DisplayRole, Qt.EditRole):
            if col < len(self.FIELDS):
                return mac if col == 0 else e.get(self.FIELDS[col], "")
            return self.registry.lookup(mac) if col == 4 and db.key(e) else ", ".join(p.where for p in self.where.get(mac, []))
        edited = col < len(self.FIELDS) and e.get(self.FIELDS[col], "") != r["orig"].get(self.FIELDS[col], "")
        if role == Qt.FontRole and (edited or r["deleted"]):
            f = QFont()
            f.setBold(edited)
            f.setStrikeOut(r["deleted"])
            return f
        if role == Qt.ToolTipRole and edited:
            return "was: %s" % (r["orig"].get(self.FIELDS[col], "") or "(empty)")
        if role == Qt.ToolTipRole and r["deleted"]:
            return "deleted when saved (Revert keeps it)"
        if role == Qt.BackgroundRole and edited and self.colors.get("differs"):
            return QColor(self.colors["differs"])
        if role == Qt.ForegroundRole and r["deleted"]:
            return QColor(self.colors.get("hint", "#757575"))
        return None

    def setData(self, index, value, role=Qt.EditRole):
        if role != Qt.EditRole or index.column() >= len(self.FIELDS):
            return False
        return self.set_field([index.row()], self.FIELDS[index.column()], str(value).strip())

    def set_field(self, rows, field, value):
        if field == "mac":
            if len(rows) != 1:
                return False
            found = macs.find(value)
            if not found:
                self.refused.emit("%r is not a MAC address" % value)
                return False
            value = found[0]
            others = [self._mac(r["entry"]) for i, r in enumerate(self.rows) if i != rows[0] and not r["deleted"]]
            if value in others:
                self.refused.emit("%s is already in the database" % value)
                return False
        for i in rows:
            e = self.rows[i]["entry"]
            if self.rows[i]["deleted"]:
                continue
            if value or field == "mac":
                e[field] = value
            else:
                e.pop(field, None)              # an emptied cell leaves the key out, as pinger's file does
        self._all_changed()
        return True

    def delete(self, rows, on=True):
        for i in rows:
            self.rows[i]["deleted"] = on
        self._all_changed()

    def revert(self, rows):
        for i in rows:
            self.rows[i]["entry"] = dict(self.rows[i]["orig"])
            self.rows[i]["deleted"] = False
        self._all_changed()

    def _all_changed(self):
        if self.rows:
            self.dataChanged.emit(self.index(0, 0), self.index(len(self.rows) - 1, len(DB_COLUMNS) - 1))
        self.changed.emit()


class DatabaseFilter(QSortFilterProxyModel):
    def __init__(self):
        super().__init__()
        self.text, self.jobsite = "", None
        self.setSortCaseSensitivity(Qt.CaseInsensitive)

    def set(self, **kw):
        self.__dict__.update(kw)
        self.invalidateFilter()

    def filterAcceptsRow(self, row, parent):
        m = self.sourceModel()
        cells = [str(m.data(m.index(row, c)) or "") for c in range(len(DB_COLUMNS))]
        if self.jobsite is not None and cells[3] != self.jobsite:
            return False
        hay = " ".join(cells + [cells[0].replace(":", "")]).lower()       # a MAC matches with or without colons
        return all(w.replace("-", ":") in hay or w in hay for w in self.text.lower().split())


class Table(QTableView):
    """Space ticks or unticks the selection."""
    toggle = Signal()

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Space and self.state() != QAbstractItemView.EditingState:
            self.toggle.emit()
            return
        super().keyPressEvent(e)


class ReviewDialog(QDialog):
    def __init__(self, parent, text, can_write, title="Write to the device database"):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(820, 520)
        lay = QVBoxLayout(self)
        view = QPlainTextEdit(text)
        view.setReadOnly(True)
        view.setFont(QFont("Consolas" if sys.platform == "win32" else "Monospace", 9))
        lay.addWidget(view)
        buttons = QDialogButtonBox()
        self.write = buttons.addButton("Write", QDialogButtonBox.AcceptRole)
        self.write.setEnabled(can_write)
        buttons.addButton(QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        lay.addWidget(buttons)


class MainWindow(QMainWindow):
    def __init__(self, db_path=None):
        super().__init__()
        # The format is the default one, which a test can point at a folder; QSettings(org, app) is always
        # the native format, on Windows the real registry, whatever a test sets.
        self.settings = QSettings(QSettings.defaultFormat(), QSettings.UserScope, "ip-table-database-editor", "iptable")
        self.db_path = db_path or os.environ.get("IPTABLE_DB") or self.settings.value("db", db.DEFAULT)
        self.entries = db.load(self.db_path)
        self.db_mtime = os.path.getmtime(self.db_path) if os.path.exists(self.db_path) else None
        self.registry = Registry()
        self.workbook = None
        self.sweep = None
        self.found = []
        self.setWindowTitle("IP table database editor")
        self.setAcceptDrops(True)

        self.model = ImportModel()
        self.model.set_colors(self.palette().window().color().lightness() < 128)
        self.proxy = ImportFilter()
        self.proxy.setSourceModel(self.model)
        self.proxy.setDynamicSortFilter(False)
        self.view = Table()
        self.view.setModel(self.proxy)
        self.view.setSortingEnabled(True)
        self.view.sortByColumn(-1, Qt.AscendingOrder)      # table order until a header is clicked
        self.view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.view.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed
                                  | QAbstractItemView.AnyKeyPressed)
        self.view.verticalHeader().setVisible(False)
        self.view.verticalHeader().setDefaultSectionSize(22)
        hh = self.view.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.Interactive)
        hh.setStretchLastSection(True)
        for col, w in ((TICK, 28), (STATUS, 76), (WHERE, 76), (MAC, 124), (IP, 96), (NAME, 220), (ROOM, 60), (ROLE, 120),
                       (MODEL, 112), (JOBSITE, 100)):
            self.view.setColumnWidth(col, w)
        self.view.toggle.connect(self.toggle_selected)
        self.view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.view.customContextMenuRequested.connect(self.context_menu)
        self.view.selectionModel().selectionChanged.connect(self.show_details)
        self.model.changed.connect(self.update_summary)
        self.model.changed.connect(self.show_details)

        self.details = QTextBrowser()
        self.details.setOpenLinks(False)

        # Top: the workbook and the database.
        self.file_label = QLabel("Open an IP table (.xlsx) or a pinger sweep (pinger -json, .json), or drop one on the window")
        self.db_label = QLabel()
        open_btn = QPushButton("Open table or sweep…")
        open_btn.clicked.connect(self.choose_workbook)
        db_btn = QPushButton("Database…")
        db_btn.clicked.connect(self.choose_database)
        top = QHBoxLayout()
        top.addWidget(open_btn)
        top.addWidget(self.file_label, 1)
        top.addWidget(self.db_label)
        top.addWidget(db_btn)

        # Filters.
        self.table_box = QComboBox()
        self.table_box.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.table_box.currentIndexChanged.connect(self.pick_table)
        self.filter_group = QButtonGroup(self)
        filters = QHBoxLayout()
        filters.addWidget(QLabel("Table:"))
        filters.addWidget(self.table_box)
        filters.addSpacing(12)
        for i, (mode, label) in enumerate(FILTERS):
            b = QPushButton(label)
            b.setCheckable(True)
            b.setChecked(i == 0)
            b.setProperty("mode", mode)
            self.filter_group.addButton(b)
            filters.addWidget(b)
        self.filter_group.buttonClicked.connect(lambda b: self.proxy.set(mode=b.property("mode")))
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search MAC, name, model, any cell…")
        self.search.setClearButtonEnabled(True)
        self.search.textChanged.connect(lambda t: self.proxy.set(text=t))
        filters.addSpacing(12)
        filters.addWidget(self.search, 1)

        # Actions on the selected rows.
        actions = QHBoxLayout()
        for label, slot, tip in (
                ("Tick", lambda: self.tick_selected(True), "Tick the selected rows (Space toggles)"),
                ("Untick", lambda: self.tick_selected(False), "Untick the selected rows"),
                ("Set room…", lambda: self.ask_part("room"), "Room for every selected row: the rooms a rack serves, e.g. 12X or 121-123"),
                ("Set role…", lambda: self.ask_part("role"), "Role for every selected row"),
                ("Set jobsite…", self.ask_jobsite, "Jobsite for every selected row"),
                ("Reset names", self.reset_selected, "Build the selected names from Room and Role again"),
                ("Rows like this", self.select_like, "Select every row with the same model and description")):
            b = QPushButton(label)
            b.setToolTip(tip)
            b.clicked.connect(slot)
            actions.addWidget(b)
        actions.addStretch(1)
        self.summary = QLabel()
        self.write_btn = QPushButton("Write to database…")
        self.write_btn.clicked.connect(self.review)
        actions.addWidget(self.summary)
        actions.addWidget(self.write_btn)

        split = QSplitter(Qt.Horizontal)
        split.addWidget(self.view)
        split.addWidget(self.details)
        split.setStretchFactor(0, 1)
        split.setStretchFactor(1, 0)
        split.setSizes([1150, 330])

        import_page = QWidget()
        lay = QVBoxLayout(import_page)
        lay.setContentsMargins(0, 6, 0, 0)
        lay.addLayout(filters)
        lay.addWidget(split, 1)
        lay.addLayout(actions)

        self.tabs = QTabWidget()
        self.tabs.addTab(import_page, "Import")
        self.tabs.addTab(self.database_page(), "Database")
        self.refresh_database_view()
        central = QWidget()
        lay = QVBoxLayout(central)
        lay.addLayout(top)
        lay.addWidget(self.tabs, 1)
        self.setCentralWidget(central)
        QShortcut(QKeySequence.Open, self, self.choose_workbook)
        QShortcut(QKeySequence("Ctrl+1"), self, lambda: self.tabs.setCurrentIndex(0))
        QShortcut(QKeySequence("Ctrl+2"), self, lambda: self.tabs.setCurrentIndex(1))
        self.resize(1400, 820)
        geom = self.settings.value("geometry")
        if geom:
            self.restoreGeometry(geom)
        self.update_db_label()
        self.update_summary()

    # --- the database view -----------------------------------------------------------------------------
    def database_page(self):
        self.db_model = DatabaseModel(self.registry)
        self.db_proxy = DatabaseFilter()
        self.db_proxy.setSourceModel(self.db_model)
        self.db_proxy.setDynamicSortFilter(False)          # an edited row stays where it is until the next sort
        self.db_view = QTableView()
        self.db_view.setModel(self.db_proxy)
        self.db_view.setSortingEnabled(True)
        self.db_view.sortByColumn(-1, Qt.AscendingOrder)
        self.db_view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.db_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.db_view.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed
                                     | QAbstractItemView.AnyKeyPressed)
        self.db_view.verticalHeader().setVisible(False)
        self.db_view.verticalHeader().setDefaultSectionSize(22)
        self.db_view.horizontalHeader().setStretchLastSection(True)
        for col, w in enumerate((130, 260, 170, 190, 260)):
            self.db_view.setColumnWidth(col, w)
        self.db_view.doubleClicked.connect(lambda i: self.goto_import_row(i) if i.column() == 5 else None)
        self.db_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.db_view.customContextMenuRequested.connect(self.db_context_menu)
        self.db_model.colors = self.model.colors
        self.db_model.changed.connect(self.database_changed)
        self.db_model.refused.connect(lambda text: self.statusBar().showMessage(text, 6000))
        QShortcut(QKeySequence.Delete, self.db_view, lambda: self.db_delete(True))

        self.db_jobsite = QComboBox()
        self.db_jobsite.currentIndexChanged.connect(
            lambda i: self.db_proxy.set(jobsite=self.db_jobsite.itemData(i)) or self.update_db_count())
        self.db_search = QLineEdit()
        self.db_search.setPlaceholderText("Search MAC (with or without colons), name, model, manufacturer…")
        self.db_search.setClearButtonEnabled(True)
        self.db_search.textChanged.connect(lambda t: self.db_proxy.set(text=t) or self.update_db_count())
        self.db_count = QLabel()
        bar = QHBoxLayout()
        bar.addWidget(QLabel("Jobsite:"))
        bar.addWidget(self.db_jobsite)
        bar.addSpacing(12)
        bar.addWidget(self.db_search, 1)
        bar.addWidget(self.db_count)

        buttons = QHBoxLayout()
        for label, slot, tip in (
                ("Set name…", lambda: self.db_ask("name"), "Name for the selected rows"),
                ("Set model…", lambda: self.db_ask("model"), "Model for the selected rows"),
                ("Set jobsite…", lambda: self.db_ask("jobsite"), "Jobsite for the selected rows"),
                ("Delete", lambda: self.db_delete(True), "Mark the selected rows for deletion (Del); saved with the rest"),
                ("Revert", self.db_revert, "Undo every edit and deletion on the selected rows"),
                ("Show in table", lambda: self.goto_import_row(self.db_view.currentIndex()),
                 "Select this device's row on the Import tab (or double-click its In open table cell)")):
            b = QPushButton(label)
            b.setToolTip(tip)
            b.clicked.connect(slot)
            buttons.addWidget(b)
        buttons.addStretch(1)
        self.db_summary = QLabel()
        self.db_save_btn = QPushButton("Save changes…")
        self.db_save_btn.clicked.connect(self.review_database)
        buttons.addWidget(self.db_summary)
        buttons.addWidget(self.db_save_btn)

        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(0, 6, 0, 0)
        lay.addLayout(bar)
        lay.addWidget(self.db_view, 1)
        lay.addLayout(buttons)
        return page

    def where_in_table(self):
        where = {}
        for p in self.model.rows:
            where.setdefault(p.mac, []).append(p)
        return where

    def refresh_database_view(self, reload=True):
        """reload: start the tab over from self.entries (after a write or another database); otherwise only
        the open table changed and the tab's pending edits stay."""
        if reload:
            self.db_model.load(self.entries, self.where_in_table())
        else:
            self.db_model.set_where(self.where_in_table())

    def database_changed(self):
        """The tab's edits moved: jobsite list, counts, the tab's title and its Save button follow."""
        working = self.db_model.working()
        current = self.db_jobsite.currentData() if self.db_jobsite.count() else None
        sites = sorted({e.get("jobsite", "") for e in working}, key=str.lower)
        self.db_jobsite.blockSignals(True)
        self.db_jobsite.clear()
        self.db_jobsite.addItem("All jobsites", None)
        for s in sites:
            n = sum(1 for e in working if e.get("jobsite", "") == s)
            self.db_jobsite.addItem("%s (%d)" % (s or "(no jobsite)", n), s)
        i = self.db_jobsite.findData(current)
        self.db_jobsite.setCurrentIndex(i if i >= 0 else 0)
        self.db_jobsite.blockSignals(False)
        self.db_proxy.set(jobsite=self.db_jobsite.currentData())
        edited, deleted = self.db_model.edits()
        if hasattr(self, "tabs"):
            pending = len(edited) + len(deleted)
            self.tabs.setTabText(1, "Database (%d)%s" % (len(self.entries), (" · %d unsaved" % pending) if pending else ""))
        self.db_summary.setText("%d edited · %d to delete" % (len(edited), len(deleted)) if edited or deleted else "no unsaved edits")
        self.db_save_btn.setEnabled(bool(edited or deleted))
        self.update_db_count()
        if hasattr(self, "summary"):
            self.update_summary()

    def update_db_count(self):
        self.db_count.setText("%d of %d shown" % (self.db_proxy.rowCount(), self.db_model.rowCount()))

    def db_selected(self):
        rows = {self.db_proxy.mapToSource(i).row() for i in self.db_view.selectionModel().selectedRows()}
        if not rows and self.db_view.currentIndex().isValid():
            rows = {self.db_proxy.mapToSource(self.db_view.currentIndex()).row()}
        return sorted(rows)

    def db_ask(self, field):
        rows = self.db_selected()
        if not rows:
            return
        values = {self.db_model.rows[r]["entry"].get(field, "") for r in rows}
        value, ok = QInputDialog.getText(self, "Set " + field, "%s for %d rows:" % (field.capitalize(), len(rows)),
                                         text=values.pop() if len(values) == 1 else "")
        if ok:
            self.db_model.set_field(rows, field, value.strip())

    def db_delete(self, on):
        rows = self.db_selected()
        if rows:
            self.db_model.delete(rows, on and not all(self.db_model.rows[r]["deleted"] for r in rows))

    def db_revert(self):
        self.db_model.revert(self.db_selected())

    def db_context_menu(self, pos):
        menu = QMenu(self)
        for label, slot in (("Set name…", lambda: self.db_ask("name")), ("Set model…", lambda: self.db_ask("model")),
                            ("Set jobsite…", lambda: self.db_ask("jobsite")), (None, None),
                            ("Delete / undelete", lambda: self.db_delete(True)), ("Revert", self.db_revert), (None, None),
                            ("Show in table", lambda: self.goto_import_row(self.db_view.currentIndex())),
                            ("Copy MAC", lambda: QApplication.clipboard().setText(
                                "\n".join(self.db_model.data(self.db_model.index(r, 0)) for r in self.db_selected())))):
            if label is None:
                menu.addSeparator()
            else:
                a = QAction(label, menu)
                a.triggered.connect(slot)
                menu.addAction(a)
        menu.exec(self.db_view.viewport().mapToGlobal(pos))

    def database_edits_text(self):
        edited, deleted = self.db_model.edits()
        out = []
        if edited:
            out.append("Edit %d database entries:" % len(edited))
            for old, new in edited:
                out.append("  %-17s  %-34s %-24s %s" % (db.key(old) or old.get("mac", ""), old.get("name", ""),
                                                        old.get("model", ""), old.get("jobsite", "")))
                out.append("  %-17s  %-34s %-24s %s" % ("  becomes " + (db.key(new) if db.key(new) != db.key(old) else ""),
                                                        new.get("name", ""), new.get("model", ""), new.get("jobsite", "")))
        if deleted:
            out += ([""] if out else []) + ["Delete %d database entries:" % len(deleted)]
            out += ["  %-17s  %-34s %-24s %s" % (db.key(old) or old.get("mac", ""), old.get("name", ""), old.get("model", ""),
                                                 old.get("jobsite", "")) for old, _ in deleted]
        return out

    def review_database(self):
        text = "\n".join(["Database: %s (%d devices now)" % (self.db_path, len(self.entries)), ""]
                         + self.database_edits_text() + ["", "The current file is copied to a .bak beside it first."])
        dlg = ReviewDialog(self, text, self.db_model.dirty(), "Save database changes")
        if dlg.exec() == QDialog.Accepted:
            self.save_database()

    def file_changed_since_load(self):
        try:
            return os.path.getmtime(self.db_path) != self.db_mtime
        except OSError:
            return False

    def save_database(self, confirm_overwrite=True):
        """Write the tab's working copy. If something else changed the file since it was loaded, ask first:
        saving would undo that change."""
        if self.file_changed_since_load() and confirm_overwrite and not self.confirm(
                "%s changed on disk since it was loaded. Save over it anyway? The other change is lost "
                "(it stays in the backup)." % os.path.basename(self.db_path)):
            return None
        backup = self.commit(self.db_model.working())
        self.statusBar().showMessage("Saved the database edits to %s%s" % (
            self.db_path, ("; backup %s" % os.path.basename(backup)) if backup else ""))
        return backup

    def commit(self, entries):
        """Write entries as the database and bring both tabs in line with it."""
        backup = db.write(self.db_path, entries)
        self.entries = entries
        self.db_mtime = os.path.getmtime(self.db_path)
        self.restatus()
        self.update_db_label()
        self.refresh_database_view()
        return backup

    def restatus(self, written=()):
        """Compare every import row with the database again: a row whose MAC left the database is new again,
        one whose entry changed is same or differs. Rows just written are unticked."""
        by_mac = {db.key(e): e for e in self.entries}
        for p in self.model.rows:
            e = by_mac.get(p.mac)
            if e is None:
                if p.status != "new":
                    p.status, p.existing, p.ticked = "new", None, False
                continue
            same = all(e.get(k, "") == v for k, v in p.entry().items() if k != "mac")
            p.status, p.existing = ("same" if same else "differs"), e
            if id(p) in written or p.status == "same":
                p.ticked, p.maybe = False, []
        self.model.set_database(self.entries)
        self.model._rows_changed([])

    def goto_import_row(self, index):
        if not index.isValid():
            return
        mac = self.db_model.data(self.db_model.index(self.db_proxy.mapToSource(index).row(), 0))
        rows = [r for r, p in enumerate(self.model.rows) if p.mac == mac]
        if not rows:
            self.statusBar().showMessage("%s is not in the open table" % mac, 4000)
            return
        self.filter_group.buttons()[0].click()          # All, so the row is not filtered out
        self.search.clear()
        self.table_box.setCurrentIndex(0)
        self.tabs.setCurrentIndex(0)
        i = self.proxy.mapFromSource(self.model.index(rows[0], NAME))
        self.view.selectionModel().select(i, QItemSelectionModel.ClearAndSelect | QItemSelectionModel.Rows)
        self.view.setCurrentIndex(i)
        self.view.scrollTo(i, QAbstractItemView.PositionAtCenter)

    # --- loading ---------------------------------------------------------------------------------------
    def choose_workbook(self):
        start = self.settings.value("dir", os.path.expanduser("~"))
        path, _ = QFileDialog.getOpenFileName(self, "Open an IP table or a pinger sweep", start,
                                              "IP tables and sweeps (*.xlsx *.xlsm *.json);;Excel workbooks (*.xlsx *.xlsm);;"
                                              "pinger -json sweeps (*.json)")
        if path:
            self.open_workbook(path)

    def open_workbook(self, path, data=None):
        """Open an IP table, or a pinger sweep: a .json file, or with data, the bytes piped in ("-")."""
        if self.unsaved() and not self.confirm("Open another table? The edits to this one are not written."):
            return False
        planner = Planner(self.entries, self.registry)
        try:
            if data is not None or sweep.is_sweep_path(path):
                self.sweep = sweep.parse(data) if data is not None else sweep.read(path)
                found, proposals, unreadable, no_mac = planner.plan_sweep(self.sweep, self.source_label(path))
            else:
                self.sweep = None
                found, proposals, unreadable, no_mac = planner.plan(path)
        except Exception as e:  # a corrupt or password-protected workbook, or something that isn't a sweep
            QMessageBox.warning(self, "Can't read the sweep" if sweep.is_sweep_path(path) else "Can't read the table",
                                "%s\n\n%s" % ("piped in" if path == "-" else path, e))
            return False
        self.workbook, self.found = path, [t for t in found if t.rows]
        self.unreadable, self.no_mac = unreadable, no_mac
        if path != "-":
            self.settings.setValue("dir", os.path.dirname(path))
        self.model.load(proposals, self.entries)
        self.table_box.blockSignals(True)
        self.table_box.clear()
        self.table_box.addItem("All tables (%d)" % len(self.found))
        for t in self.found:
            self.table_box.addItem("%s (%d MACs)" % (t.sheet, len(t.rows)) if t.sweep is not None else
                                   "%s, header row %d (%d rows)" % (t.sheet, t.header_row, len(t.rows)) if t.header_row
                                   else "%s, MAC column guessed (%d rows)" % (t.sheet, len(t.rows)))
        self.table_box.blockSignals(False)
        self.proxy.set(table=None)
        self.file_label.setText(self.source_label(path))
        self.file_label.setToolTip("" if path == "-" else path)
        self.refresh_database_view(reload=False)
        self.details.setHtml(self.workbook_notes())
        return True

    @staticmethod
    def source_label(path):
        return "Sweep, piped in" if path == "-" else os.path.basename(path)

    def sweep_notes(self):
        sw = self.sweep
        out = ["<b>%s</b>" % html_escape(self.source_label(self.workbook)),
               "%d of %d hosts answered in %.1f s; %d with a MAC, one row each. Rows start with no name and "
               "unticked: set Room and Role, or type a name, then tick." % (len(sw.hosts), sw.targets, sw.seconds,
                                                                          len(self.model.rows))]
        if not sw.arp_cleared:
            out.append("<span style='color:%s'>⚠ The ARP cache was not cleared (pinger needs an Administrator prompt "
                       "for that), so a MAC can belong to a device that has since left that IP.</span>" % self.model.colors["warn"])
        if self.no_mac:
            out.append("Left out, answered with no MAC (routed, not on this machine's subnet): %s"
                       % (", ".join(self.no_mac[:40]) + (" …" if len(self.no_mac) > 40 else "")))
        if sw.skipped:
            out.append("pinger left out the sweeping machine: %s" % ", ".join(sw.skipped))
        return "<br>".join(out)

    def workbook_notes(self):
        if getattr(self, "sweep", None) is not None:
            return self.sweep_notes()
        out = ["<b>%s</b>" % html_escape(os.path.basename(self.workbook)),
               "%d tables, %d MACs. Select a row to see where it came from." % (len(self.found), len(self.model.rows))]
        if self.unreadable:
            out.append("<br><b>MAC cells that hold no MAC:</b>")
            out += ["&nbsp;&nbsp;%s: %r, %s%s" % (u.where, html_escape(u.cell), html_escape(u.why),
                                                   "; meant %s?" % u.guess if u.guess else "") for u in self.unreadable]
        if self.no_mac:
            out.append("<br>%d rows have no MAC and are left out: %s" % (len(self.no_mac), ", ".join(self.no_mac[:40])
                                                                       + (" …" if len(self.no_mac) > 40 else "")))
        return "<br>".join(out)

    def choose_database(self):
        path, _ = QFileDialog.getOpenFileName(self, "pinger device database", os.path.dirname(self.db_path),
                                              "Device database (*.json)")
        if not path:
            return
        if self.unsaved() and not self.confirm("Use another database? The table is read again, and its edits and the "
                                               "Database tab's unsaved edits are lost."):
            return
        self.db_path = path
        self.settings.setValue("db", path)
        self.entries = db.load(path)
        self.db_mtime = os.path.getmtime(path)
        self.update_db_label()
        self.refresh_database_view()
        if self.workbook:
            wb, sw, self.workbook = self.workbook, self.sweep, None
            self.model.rows = []
            if sw is not None and wb == "-":     # piped in: plan the same sweep against the other database
                self.workbook, self.found, self.sweep = wb, [], sw
                found, proposals, self.unreadable, self.no_mac = Planner(self.entries, self.registry).plan_sweep(
                    sw, self.source_label(wb))
                self.found = found
                self.model.load(proposals, self.entries)
                self.refresh_database_view(reload=False)
                self.details.setHtml(self.workbook_notes())
            else:
                self.open_workbook(wb)

    def update_db_label(self):
        self.db_label.setText("Database: %s (%d devices)" % (os.path.basename(self.db_path), len(self.entries)))
        self.db_label.setToolTip(self.db_path)

    def pick_table(self, i):
        self.proxy.set(table=self.found[i - 1] if i > 0 else None)

    def dragEnterEvent(self, e):
        if any(u.toLocalFile().lower().endswith(OPENS) for u in e.mimeData().urls()):
            e.acceptProposedAction()

    def dropEvent(self, e):
        for u in e.mimeData().urls():
            if u.toLocalFile().lower().endswith(OPENS):
                self.open_workbook(u.toLocalFile())
                return

    # --- selection and bulk edits ----------------------------------------------------------------------
    def selected(self):
        rows = {self.proxy.mapToSource(i).row() for i in self.view.selectionModel().selectedRows()}
        if not rows and self.view.currentIndex().isValid():
            rows = {self.proxy.mapToSource(self.view.currentIndex()).row()}
        return sorted(rows)

    def tick_selected(self, on):
        self.model.set_ticked(self.selected(), on)

    def toggle_selected(self):
        rows = self.selected()
        if rows:
            self.model.set_ticked(rows, not all(self.model.rows[r].ticked for r in rows))

    def common(self, rows, get):
        values = {get(self.model.rows[r]) for r in rows}
        return values.pop() if len(values) == 1 else ""

    def ask_part(self, part):
        rows = self.selected()
        if not rows:
            return
        label = {"room": "Room for %d rows (the rooms a rack serves, e.g. 12X):",
                 "role": "Role for %d rows:"}[part] % len(rows)
        value, ok = QInputDialog.getText(self, "Set " + part, label, text=self.common(rows, lambda p: p.parts.get(part, "")))
        if ok:
            self.model.set_part(rows, part, value.strip())

    def ask_jobsite(self):
        rows = self.selected()
        if not rows:
            return
        value, ok = QInputDialog.getText(self, "Set jobsite", "Jobsite for %d rows:" % len(rows),
                                         text=self.common(rows, lambda p: p.jobsite))
        if ok:
            self.model.set_field(rows, "jobsite", value.strip())

    def reset_selected(self):
        self.model.reset_names(self.selected())

    def select_like(self):
        rows = self.selected()
        if not rows:
            return
        p = self.model.rows[rows[0]]
        like = (p.model.lower(), p.row.values.get("description", "").lower(), p.table)
        sel = self.view.selectionModel()
        sel.clearSelection()
        for r, o in enumerate(self.model.rows):
            if (o.model.lower(), o.row.values.get("description", "").lower(), o.table) == like:
                i = self.proxy.mapFromSource(self.model.index(r, 0))
                if i.isValid():
                    sel.select(i, QItemSelectionModel.Select | QItemSelectionModel.Rows)

    def context_menu(self, pos):
        menu = QMenu(self)
        cell = self.proxy.mapToSource(self.view.indexAt(pos))
        col = cell.column() if cell.isValid() else -1
        if col in ImportModel.DIFF_FIELDS:
            rows = self.selected()
            if cell.row() not in rows:
                rows = [cell.row()]
            hits = [r for r in rows if self.model.differs(self.model.rows[r], col)]
            if hits:
                p = self.model.rows[cell.row()]
                diff = self.model.differs(p, col)
                label = ("Use database value for %s: %s" % (COLUMNS[col], diff[0] or "(empty)") if len(hits) == 1 and diff
                         else "Use database value for %s (%d rows)" % (COLUMNS[col], len(hits)))
                a = QAction(label, menu)
                a.triggered.connect(lambda: self.model.use_database(hits, col))
                menu.addAction(a)
                menu.addSeparator()
        for label, slot in (("Tick", lambda: self.tick_selected(True)), ("Untick", lambda: self.tick_selected(False)),
                            (None, None), ("Set room…", lambda: self.ask_part("room")),
                            ("Set role…", lambda: self.ask_part("role")), ("Set jobsite…", self.ask_jobsite),
                            ("Reset names", self.reset_selected), (None, None), ("Select rows like this", self.select_like),
                            ("Copy MAC", self.copy_mac)):
            if label is None:
                menu.addSeparator()
            else:
                a = QAction(label, menu)
                a.triggered.connect(slot)
                menu.addAction(a)
        menu.exec(self.view.viewport().mapToGlobal(pos))

    def copy_mac(self):
        QApplication.clipboard().setText("\n".join(self.model.rows[r].mac for r in self.selected()))

    def show_details(self, *_):
        rows = self.selected()
        if not rows or not self.model.rows:
            return
        p = self.model.rows[rows[0]]
        v = p.row.values
        out = ["<b>%s</b> &nbsp; %s &nbsp; <span style='color:gray'>%s</span>" % (html_escape(p.name), p.mac, html_escape(p.where))]
        out.append("Registered to: %s" % html_escape(p.oui or "no registered block"))
        cells = [("MAC" if p.table.sweep is not None else "MAC cell", p.row.mac_cell)] + [(k.replace("_", " / "), v[k]) for k in
                 ("room", "location", "make", "model", "make_model", "description", "hostname", "ip", "reply", "serial", "notes")
                 if v.get(k)]
        out.append("<table cellspacing=0 cellpadding=2>" + "".join(
            "<tr><td style='color:gray'>%s</td><td>%s</td></tr>" % (html_escape(k), html_escape(x)) for k, x in cells) + "</table>")
        for c in self.model.checks(p):
            out.append("<span style='color:%s'>⚠ %s</span>" % (self.model.colors["warn"], html_escape(c)))
        for m in p.maybe:
            out.append("maybe: %s" % html_escape(m))
        if p.existing:
            e = p.existing
            out.append("In the database: %s" % html_escape(", ".join("%s %s" % (k, e[k]) for k in ("name", "model", "jobsite") if e.get(k))))
            if p.status == "differs":
                out.append("Tick this row to replace that entry with the one above.")
        if len(rows) > 1:
            out.append("<i>%d rows selected</i>" % len(rows))
        self.details.setHtml("<br>".join(out))

    # --- writing ---------------------------------------------------------------------------------------
    def pending(self):
        """(additions, replacements, skipped) for the ticked rows."""
        adds, repl, skipped, seen = [], [], [], set()
        for p in self.model.rows:
            if not p.ticked or p.status == "same":
                continue
            if p.mac in seen:
                skipped.append(p)
                continue
            seen.add(p.mac)
            (adds if p.status == "new" else repl).append(p)
        return adds, repl, skipped

    def unsaved(self):
        adds, repl, _ = self.pending()
        table = self.workbook and (adds or repl or any(getattr(p, "locked", False) for p in self.model.rows))
        return bool(table or self.db_model.dirty())

    def update_summary(self):
        adds, repl, _ = self.pending()
        flagged = sum(1 for p in adds + repl if self.model.checks(p))
        maybes = sum(1 for p in self.model.rows if category(p) == "maybe" and not p.ticked)
        dbe = self.db_model.dirty()
        self.summary.setText("%d to add · %d to replace · %d maybes left · %s%s" % (
            len(adds), len(repl), maybes, ("%d ticked rows flagged" % flagged) if flagged else "no ticked row flagged",
            " · Database tab edits are written too" if dbe else ""))
        self.write_btn.setEnabled(bool(adds or repl or dbe))

    def review_text(self):
        adds, repl, skipped = self.pending()
        width = max([34] + [len(p.name) + 1 for p in adds + repl] + [len(p.existing.get("name", "")) + 1 for p in repl])
        line = "  %-17s  %-" + str(width) + "s %-24s %s"
        out = ["Database: %s (%d devices now)" % (self.db_path, len(self.entries)), ""]
        flagged = [p for p in adds + repl if self.model.checks(p)]
        if flagged:
            out.append("Ticked rows with checks (look at these before writing):")
            for p in flagged:
                out.append("  %-17s  %-34s %s" % (p.mac, p.name, "; ".join(self.model.checks(p))))
            out.append("")
        out.append("Add %d:" % len(adds))
        out += [line % (p.mac, p.name, p.model, p.jobsite) for p in adds]
        if repl:
            out += ["", "Replace %d:" % len(repl)]
            for p in repl:
                e = p.existing
                out.append(line % (p.mac, e.get("name", ""), e.get("model", ""), e.get("jobsite", "")))
                fields = [f for c, f in ImportModel.DIFF_FIELDS.items() if self.model.differs(p, c)]
                out.append((line % ("  becomes", p.name, p.model, p.jobsite)).rstrip() + "   (%s)" % ", ".join(fields))
        if skipped:
            out += ["", "Left out, MAC already ticked on another row:"] + ["  %s  %s" % (p.mac, p.where) for p in skipped]
        dbe = self.database_edits_text()
        if dbe:
            out += ["", "Also written, from the Database tab:"] + dbe
        out += ["", "The current file is copied to a .bak beside it first."]
        return "\n".join(out)

    def unnamed(self):
        adds, repl, _ = self.pending()
        return [p for p in adds + repl if not p.name.strip()]

    def review(self):
        adds, repl, _ = self.pending()
        unnamed = self.unnamed()
        text = self.review_text()
        if unnamed:
            text = ("Nothing can be written while %d ticked rows have no name: %s\n\n"
                    % (len(unnamed), ", ".join(p.where for p in unnamed[:20]) + (" …" if len(unnamed) > 20 else ""))) + text
        dlg = ReviewDialog(self, text, bool(adds or repl or self.db_model.dirty()) and not unnamed)
        if dlg.exec() == QDialog.Accepted:
            self.write()

    def write(self):
        """Write the ticked rows. With Database tab edits pending, they are the base and are written too;
        without, the rows are merged onto the file as it is now, so a change made elsewhere is kept."""
        adds, repl, _ = self.pending()
        if self.unnamed():
            return None
        if self.db_model.dirty():
            if self.file_changed_since_load() and not self.confirm(
                    "%s changed on disk since it was loaded, and the Database tab has unsaved edits. Write anyway? "
                    "The other change is lost (it stays in the backup)." % os.path.basename(self.db_path)):
                return None
            base = self.db_model.working()
        else:
            base = db.load(self.db_path)
        new = db.merge(base, [p.entry() for p in adds], [p.entry() for p in repl])
        backup = db.write(self.db_path, new)
        self.entries = new
        self.db_mtime = os.path.getmtime(self.db_path)
        self.restatus(written={id(p) for p in adds + repl})
        self.update_db_label()
        self.refresh_database_view()
        self.statusBar().showMessage("Wrote %d new and %d replaced devices to %s%s" % (
            len(adds), len(repl), self.db_path, ("; backup %s" % os.path.basename(backup)) if backup else ""))
        return backup

    def confirm(self, text):
        return QMessageBox.question(self, "IP table database editor", text) == QMessageBox.Yes

    def closeEvent(self, e):
        if self.unsaved() and not self.confirm("Close without writing? The ticked rows, edited names and the Database "
                                               "tab's unsaved edits are lost."):
            e.ignore()
            return
        self.settings.setValue("geometry", self.saveGeometry())
        super().closeEvent(e)


def main(argv=None):
    ap = argparse.ArgumentParser(prog="iptable.gui", description="Review an IP table and import it into pinger's database")
    ap.add_argument("table", nargs="?", help="an IP table (.xlsx), a pinger -json sweep (.json), or - for one piped in "
                                             "(a pipe with no table named is read too)")
    ap.add_argument("--db", help="pinger device database (default: the last one used, else devices.json beside pinger)")
    ap.add_argument("--piped", action="store_true", help=argparse.SUPPRESS)    # started by iptable.launch with a sweep
    args = ap.parse_args(argv)
    data = None
    if args.table == "-" or args.piped or (args.table is None and launch.stdin_piped()):
        data = sys.stdin.buffer.read() if sys.stdin is not None else b""
        args.table = "-"
    app = QApplication(sys.argv[:1])
    app.setStyle("Fusion")
    win = MainWindow(args.db)
    win.show()
    if args.table:
        win.open_workbook(args.table, data)
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
