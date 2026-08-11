import csv
import io
import json
import sqlite3
import sys
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QTabWidget, QMessageBox, QDialog, QFormLayout, QComboBox,
    QDateEdit, QSpinBox, QHeaderView, QAbstractItemView, QGroupBox,
    QGridLayout, QTextEdit, QFileDialog
)
from PyQt6.QtCore import Qt, QDate, QSize, QThread, pyqtSignal
from PyQt6.QtGui import QColor, QFont, QPixmap, QIcon
from PyQt6.QtCharts import (
    QChart, QChartView, QPieSeries, QPieSlice, QBarSet, QBarSeries,
    QBarCategoryAxis, QValueAxis
)

COVERS_DIR = Path(__file__).resolve().parent / "copertine"
COVER_URL_CACHE = {}


def trova_copertina_url(titolo, autore):
    key = (titolo.lower().strip(), autore.lower().strip())
    if key in COVER_URL_CACHE:
        return COVER_URL_CACHE[key]
    q = urllib.parse.quote(f"{titolo} {autore}")
    url = (f"https://openlibrary.org/search.json?q={q}"
           "&limit=3&fields=title,author_name,cover_i")
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "RubricaBiblioteca/1.0"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.load(resp)
        for d in data.get("docs", []):
            ci = d.get("cover_i")
            if ci:
                COVER_URL_CACHE[key] = f"https://covers.openlibrary.org/b/id/{ci}-L.jpg"
                return COVER_URL_CACHE[key]
    except Exception:
        pass
    COVER_URL_CACHE[key] = None
    return None


def file_copertina(libro_id):
    return COVERS_DIR / f"{libro_id}.jpg"


class CopertinaWorker(QThread):
    fatta = pyqtSignal(int)

    def __init__(self, libro_id, url, percorso):
        super().__init__()
        self.libro_id = libro_id
        self.url = url
        self.percorso = percorso

    def run(self):
        try:
            req = urllib.request.Request(
                self.url, headers={"User-Agent": "RubricaBiblioteca/1.0"})
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = resp.read()
            if len(data) > 1500:
                self.percorso.parent.mkdir(parents=True, exist_ok=True)
                self.percorso.write_bytes(data)
                self.fatta.emit(self.libro_id)
        except Exception:
            pass


def scarica_copertina(titolo, autore):
    url = trova_copertina_url(titolo, autore)
    if not url:
        return None
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "RubricaBiblioteca/1.0"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = resp.read()
        pix = QPixmap()
        if pix.loadFromData(data):
            return pix
    except Exception:
        pass
    return None

DB_PATH = Path(__file__).resolve().parent / "biblioteca.db"


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_conn()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS libri (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            titolo TEXT NOT NULL,
            autore TEXT NOT NULL,
            anno INTEGER,
            genere TEXT,
            isbn TEXT,
            note TEXT,
            data_aggiunta TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS prestiti (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            libro_id INTEGER NOT NULL REFERENCES libri(id) ON DELETE CASCADE,
            persona TEXT NOT NULL,
            data_prestito TEXT NOT NULL,
            data_restituzione TEXT,
            scadenza TEXT NOT NULL
        );
    """)
    conn.commit()
    conn.close()


STATUS_COLORS = {
    "Disponibile": QColor("#2e7d32"),
    "In prestito": QColor("#e65100"),
    "In ritardo": QColor("#c62828"),
}


class LibroDialog(QDialog):
    def __init__(self, parent=None, libro=None):
        super().__init__(parent)
        self.setWindowTitle("Aggiungi libro" if libro is None else "Modifica libro")
        self.setMinimumWidth(380)

        form = QFormLayout(self)

        self.titolo = QLineEdit()
        self.autore = QLineEdit()
        self.anno = QSpinBox()
        self.anno.setRange(0, 2100)
        self.anno.setSpecialValueText("Sconosciuto")
        self.genere = QLineEdit()
        self.isbn = QLineEdit()
        self.note = QTextEdit()
        self.note.setMaximumHeight(70)

        form.addRow("Titolo *", self.titolo)
        form.addRow("Autore *", self.autore)
        form.addRow("Anno", self.anno)
        form.addRow("Genere", self.genere)
        form.addRow("ISBN", self.isbn)
        form.addRow("Note", self.note)

        if libro is not None:
            self.titolo.setText(libro["titolo"])
            self.autore.setText(libro["autore"])
            self.anno.setValue(libro["anno"] or 0)
            self.genere.setText(libro["genere"] or "")
            self.isbn.setText(libro["isbn"] or "")
            self.note.setPlainText(libro["note"] or "")

        btns = QHBoxLayout()
        save = QPushButton("Salva")
        cancel = QPushButton("Annulla")
        save.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        btns.addWidget(save)
        btns.addWidget(cancel)
        form.addRow(btns)

    def data(self):
        return {
            "titolo": self.titolo.text().strip(),
            "autore": self.autore.text().strip(),
            "anno": self.anno.value() or None,
            "genere": self.genere.text().strip() or None,
            "isbn": self.isbn.text().strip() or None,
            "note": self.note.toPlainText().strip() or None,
        }


class PrestitoDialog(QDialog):
    def __init__(self, parent=None, libro_titolo="", giorni=30):
        super().__init__(parent)
        self.setWindowTitle("Nuovo prestito")
        self.setMinimumWidth(360)

        form = QFormLayout(self)
        form.addRow("Libro", QLabel(libro_titolo))

        self.persona = QLineEdit()
        self.data_prestito = QDateEdit(QDate.currentDate())
        self.data_prestito.setCalendarPopup(True)
        self.data_prestito.setDisplayFormat("dd/MM/yyyy")
        self.scadenza = QDateEdit(QDate.currentDate().addDays(giorni))
        self.scadenza.setCalendarPopup(True)
        self.scadenza.setDisplayFormat("dd/MM/yyyy")

        form.addRow("Persona *", self.persona)
        form.addRow("Data prestito", self.data_prestito)
        form.addRow("Scadenza", self.scadenza)

        btns = QHBoxLayout()
        save = QPushButton("Registra")
        cancel = QPushButton("Annulla")
        save.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        btns.addWidget(save)
        btns.addWidget(cancel)
        form.addRow(btns)


class DettagliLibroDialog(QDialog):
    def __init__(self, parent=None, libro=None, stato=""):
        super().__init__(parent)
        self.setWindowTitle("Dettagli libro")
        self.setMinimumSize(480, 320)

        lay = QHBoxLayout(self)
        self.cover_label = QLabel("Nessuna copertina\n(senza ISBN)")
        self.cover_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cover_label.setFixedSize(150, 220)
        self.cover_label.setStyleSheet(
            "border: 1px dashed #888; border-radius: 6px; color: #777;")
        lay.addWidget(self.cover_label)

        info = QVBoxLayout()
        titolo = QLabel(libro["titolo"])
        titolo.setFont(QFont("", 16, QFont.Weight.Bold))
        titolo.setWordWrap(True)
        info.addWidget(titolo)
        autore = QLabel(libro["autore"])
        autore.setFont(QFont("", 12))
        info.addWidget(autore)

        det = [
            ("Anno", str(libro["anno"]) if libro["anno"] else "—"),
            ("Genere", libro["genere"] or "—"),
            ("ISBN", libro["isbn"] or "—"),
            ("Stato", stato),
        ]
        for k, v in det:
            row = QHBoxLayout()
            kk = QLabel(f"<b>{k}:</b>")
            vv = QLabel(v)
            row.addWidget(kk)
            row.addWidget(vv)
            row.addStretch()
            info.addLayout(row)

        if libro["note"]:
            info.addWidget(QLabel("<b>Note:</b>"))
            note = QLabel(libro["note"])
            note.setWordWrap(True)
            note.setStyleSheet("color: #666;")
            info.addWidget(note)

        info.addStretch()
        chiudi = QPushButton("Chiudi")
        chiudi.clicked.connect(self.accept)
        info.addWidget(chiudi, alignment=Qt.AlignmentFlag.AlignRight)
        lay.addLayout(info)

        pix = scarica_copertina(libro["titolo"], libro["autore"])
        if pix and not pix.isNull():
            self.cover_label.setPixmap(pix.scaled(
                150, 220, Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation))
            self.cover_label.setStyleSheet("border: none;")


def status_of(libro_id, prestiti_map):
    if libro_id in prestiti_map:
        p = prestiti_map[libro_id]
        if p["data_restituzione"] is None:
            scad = datetime.strptime(p["scadenza"], "%Y-%m-%d").date()
            return "In ritardo" if scad < date.today() else "In prestito"
    return "Disponibile"


def prestiti_per_libro(conn):
    rows = conn.execute("SELECT * FROM prestiti ORDER BY data_prestito DESC").fetchall()
    m = {}
    for r in rows:
        m.setdefault(r["libro_id"], r)
    return m


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Rubrica Biblioteca")
        self.resize(920, 620)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        self.build_libri_tab()
        self.build_prestiti_tab()
        self.build_statistiche_tab()

        self.statusBar().showMessage("Pronto")

    # ------------------------------------------------------- LIBRI
    def build_libri_tab(self):
        tab = QWidget()
        lay = QVBoxLayout(tab)

        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Cerca:"))
        self.search = QLineEdit()
        self.search.setPlaceholderText("Titolo, autore, genere o ISBN...")
        self.search.textChanged.connect(self.load_libri)
        search_row.addWidget(self.search)

        search_row.addWidget(QLabel("Genere:"))
        self.filter_genre = QComboBox()
        self.filter_genre.addItem("Tutti")
        self.filter_genre.currentTextChanged.connect(self.load_libri)
        search_row.addWidget(self.filter_genre)

        search_row.addWidget(QLabel("Stato:"))
        self.filter_status = QComboBox()
        self.filter_status.addItems(["Tutti", "Disponibile", "In prestito", "In ritardo"])
        self.filter_status.currentTextChanged.connect(self.load_libri)
        search_row.addWidget(self.filter_status)
        lay.addLayout(search_row)

        year_row = QHBoxLayout()
        year_row.addWidget(QLabel("Anno dal:"))
        self.year_from = QSpinBox()
        self.year_from.setRange(0, 2100)
        self.year_from.setSpecialValueText("Tutti")
        self.year_from.valueChanged.connect(self.load_libri)
        year_row.addWidget(self.year_from)
        year_row.addWidget(QLabel("al:"))
        self.year_to = QSpinBox()
        self.year_to.setRange(0, 2100)
        self.year_to.setValue(2100)
        self.year_to.valueChanged.connect(self.load_libri)
        year_row.addWidget(self.year_to)
        year_row.addWidget(QLabel("Ordinamento:"))
        self.sort_by = QComboBox()
        self.sort_by.addItems(["Titolo", "Autore", "Anno", "Genere"])
        self.sort_by.currentTextChanged.connect(self.load_libri)
        year_row.addWidget(self.sort_by)
        year_row.addStretch()
        lay.addLayout(year_row)

        self.table_libri = QTableWidget(0, 7)
        self.table_libri.setHorizontalHeaderLabels(["Copertina", "Titolo", "Autore", "Anno", "Genere", "Stato", "Prestito a"])
        self.table_libri.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_libri.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table_libri.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_libri.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table_libri.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table_libri.verticalHeader().setDefaultSectionSize(60)
        self.table_libri.setIconSize(QSize(40, 55))
        self.table_libri.doubleClicked.connect(lambda: self.show_details())
        lay.addWidget(self.table_libri)

        btns = QHBoxLayout()
        b_add = QPushButton("Aggiungi libro")
        b_edit = QPushButton("Modifica")
        b_del = QPushButton("Elimina")
        b_dett = QPushButton("Dettagli")
        b_prestito = QPushButton("Registra prestito")
        b_rest = QPushButton("Restituisci")
        b_csv = QPushButton("Esporta CSV")
        b_add.clicked.connect(lambda: self.add_libro())
        b_edit.clicked.connect(lambda: self.edit_libro())
        b_del.clicked.connect(lambda: self.delete_libro())
        b_dett.clicked.connect(lambda: self.show_details())
        b_prestito.clicked.connect(lambda: self.new_prestito())
        b_rest.clicked.connect(lambda: self.return_libro())
        b_csv.clicked.connect(lambda: self.export_csv())
        for b in (b_add, b_edit, b_dett, b_del, b_prestito, b_rest, b_csv):
            btns.addWidget(b)
        btns.addStretch()
        lay.addLayout(btns)
        self.tabs.addTab(tab, "Libri")

    def current_libro(self):
        row = self.table_libri.currentRow()
        if row < 0:
            return None
        return self._libri[row]

    def load_libri(self):
        q = self.search.text().strip()
        filtro = self.filter_status.currentText()
        genere = self.filter_genre.currentText()
        y_from = self.year_from.value() if self.year_from.value() else -10000
        y_to = self.year_to.value() if self.year_to.value() else 2100
        ordine = self.sort_by.currentText()
        conn = get_conn()
        p_map = prestiti_per_libro(conn)
        rows = conn.execute("SELECT * FROM libri").fetchall()
        self._libri = []
        for libro in rows:
            st = status_of(libro["id"], p_map)
            if filtro != "Tutti" and st != filtro:
                continue
            if genere != "Tutti" and (libro["genere"] or "") != genere:
                continue
            anno = libro["anno"] or 0
            if anno and not (y_from <= anno <= y_to):
                continue
            if q and q.lower() not in " ".join(filter(None, [
                    libro["titolo"], libro["autore"], libro["genere"], libro["isbn"]
            ])).lower():
                continue
            self._libri.append(libro)

        order_key = {"Titolo": "titolo", "Autore": "autore", "Anno": "anno", "Genere": "genere"}[ordine]
        self._libri.sort(key=lambda l: (l[order_key] is None, l[order_key] or ""))

        self.table_libri.setRowCount(len(self._libri))
        self._cover_workers = getattr(self, "_cover_workers", [])
        for r, libro in enumerate(self._libri):
            st = status_of(libro["id"], p_map)
            persona = ""
            if libro["id"] in p_map and p_map[libro["id"]]["data_restituzione"] is None:
                persona = p_map[libro["id"]]["persona"]

            item_cover = QTableWidgetItem()
            f = file_copertina(libro["id"])
            if f.exists():
                pix = QPixmap(str(f))
                if not pix.isNull():
                    item_cover.setIcon(QIcon(pix.scaled(
                        40, 55, Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation)))
            self.table_libri.setItem(r, 0, item_cover)

            self.table_libri.setItem(r, 1, QTableWidgetItem(libro["titolo"]))
            self.table_libri.setItem(r, 2, QTableWidgetItem(libro["autore"]))
            self.table_libri.setItem(r, 3, QTableWidgetItem(str(libro["anno"]) if libro["anno"] else "—"))
            self.table_libri.setItem(r, 4, QTableWidgetItem(libro["genere"] or "—"))
            item_st = QTableWidgetItem(st)
            item_st.setForeground(STATUS_COLORS[st])
            self.table_libri.setItem(r, 5, item_st)
            self.table_libri.setItem(r, 6, QTableWidgetItem(persona))
            self._cover_workers = [w for w in self._cover_workers if w.isRunning()]

            if not f.exists():
                url = trova_copertina_url(libro["titolo"], libro["autore"])
                if url:
                    w = CopertinaWorker(libro["id"], url, f)
                    w.fatta.connect(self.on_copertina_pronta)
                    self._cover_workers.append(w)
                    w.start()
        conn.close()

    def on_copertina_pronta(self, libro_id):
        for r in range(self.table_libri.rowCount()):
            if self._libri and r < len(self._libri) and self._libri[r]["id"] == libro_id:
                f = file_copertina(libro_id)
                pix = QPixmap(str(f))
                if not pix.isNull():
                    self.table_libri.item(r, 0).setIcon(QIcon(pix.scaled(
                        40, 55, Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation)))
                break

    def aggiorna_generi(self):
        conn = get_conn()
        rows = conn.execute(
            "SELECT DISTINCT genere FROM libri WHERE genere IS NOT NULL AND genere != '' "
            "ORDER BY genere COLLATE NOCASE").fetchall()
        conn.close()
        current = self.filter_genre.currentText()
        self.filter_genre.blockSignals(True)
        self.filter_genre.clear()
        self.filter_genre.addItem("Tutti")
        for r in rows:
            self.filter_genre.addItem(r["genere"])
        if current in [self.filter_genre.itemText(i) for i in range(self.filter_genre.count())]:
            self.filter_genre.setCurrentText(current)
        self.filter_genre.blockSignals(False)

    def show_details(self):
        libro = self.current_libro()
        if not libro:
            QMessageBox.information(self, "Nessuna selezione", "Seleziona un libro dalla lista.")
            return
        conn = get_conn()
        p_map = prestiti_per_libro(conn)
        conn.close()
        st = status_of(libro["id"], p_map)
        dlg = DettagliLibroDialog(self, libro, st)
        dlg.exec()

    def export_csv(self):
        if not self._libri:
            QMessageBox.information(self, "Nessun libro", "Non c'è nulla da esportare.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Salva esportazione", "biblioteca.csv", "File CSV (*.csv)")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(["Titolo", "Autore", "Anno", "Genere", "ISBN", "Stato", "Prestito a"])
                for libro in self._libri:
                    conn = get_conn()
                    p_map = prestiti_per_libro(conn)
                    conn.close()
                    st = status_of(libro["id"], p_map)
                    persona = ""
                    if libro["id"] in p_map and p_map[libro["id"]]["data_restituzione"] is None:
                        persona = p_map[libro["id"]]["persona"]
                    w.writerow([
                        libro["titolo"], libro["autore"], libro["anno"] or "",
                        libro["genere"] or "", libro["isbn"] or "", st, persona
                    ])
            self.statusBar().showMessage(f"Esportati {len(self._libri)} libri in {path}", 6000)
        except OSError as e:
            QMessageBox.critical(self, "Errore", f"Impossibile salvare il file:\n{e}")

    def add_libro(self):
        dlg = LibroDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            data = dlg.data()
            if not data["titolo"] or not data["autore"]:
                QMessageBox.warning(self, "Dati mancanti", "Titolo e autore sono obbligatori.")
                return
            conn = get_conn()
            conn.execute(
                "INSERT INTO libri (titolo, autore, anno, genere, isbn, note, data_aggiunta) "
                "VALUES (?,?,?,?,?,?,?)",
                (data["titolo"], data["autore"], data["anno"], data["genere"],
                 data["isbn"], data["note"], date.today().isoformat()))
            conn.commit()
            conn.close()
            self.load_libri()
            self.aggiorna_generi()
            self.load_statistiche()

    def edit_libro(self):
        libro = self.current_libro()
        if not libro:
            QMessageBox.information(self, "Nessuna selezione", "Seleziona un libro dalla lista.")
            return
        dlg = LibroDialog(self, libro)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            data = dlg.data()
            if not data["titolo"] or not data["autore"]:
                QMessageBox.warning(self, "Dati mancanti", "Titolo e autore sono obbligatori.")
                return
            conn = get_conn()
            conn.execute(
                "UPDATE libri SET titolo=?, autore=?, anno=?, genere=?, isbn=?, note=? WHERE id=?",
                (data["titolo"], data["autore"], data["anno"], data["genere"],
                 data["isbn"], data["note"], libro["id"]))
            conn.commit()
            conn.close()
            self.load_libri()
            self.aggiorna_generi()
            self.load_statistiche()

    def delete_libro(self):
        libro = self.current_libro()
        if not libro:
            QMessageBox.information(self, "Nessuna selezione", "Seleziona un libro dalla lista.")
            return
        ans = QMessageBox.question(
            self, "Conferma",
            f"Eliminare '{libro['titolo']}'?\nI prestiti collegati verranno rimossi.")
        if ans == QMessageBox.StandardButton.Yes:
            conn = get_conn()
            conn.execute("DELETE FROM libri WHERE id=?", (libro["id"],))
            conn.commit()
            conn.close()
            self.load_libri()
            self.load_prestiti()
            self.aggiorna_generi()
            self.load_statistiche()

    # ------------------------------------------------------- PRESTITI
    def build_prestiti_tab(self):
        tab = QWidget()
        lay = QVBoxLayout(tab)

        self.table_prestiti = QTableWidget(0, 6)
        self.table_prestiti.setHorizontalHeaderLabels(
            ["Libro", "Persona", "Data prestito", "Scadenza", "Stato", "Restituito il"])
        self.table_prestiti.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_prestiti.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_prestiti.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table_prestiti.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table_prestiti.doubleClicked.connect(lambda: self.return_libro())
        lay.addWidget(self.table_prestiti)

        btns = QHBoxLayout()
        b_act = QPushButton("Mostra solo prestiti attivi")
        b_act.setCheckable(True)
        b_act.toggled.connect(self.load_prestiti)
        self.b_act = b_act
        b_ret = QPushButton("Restituisci selezionato")
        b_ret.clicked.connect(lambda: self.return_libro())
        btns.addWidget(b_act)
        btns.addWidget(b_ret)
        btns.addStretch()
        lay.addLayout(btns)
        self.tabs.addTab(tab, "Prestiti")

    def load_prestiti(self):
        conn = get_conn()
        if self.b_act.isChecked():
            rows = conn.execute("""
                SELECT p.*, l.titolo FROM prestiti p
                JOIN libri l ON l.id = p.libro_id
                WHERE p.data_restituzione IS NULL
                ORDER BY p.scadenza""").fetchall()
        else:
            rows = conn.execute("""
                SELECT p.*, l.titolo FROM prestiti p
                JOIN libri l ON l.id = p.libro_id
                ORDER BY p.data_prestito DESC""").fetchall()
        self.table_prestiti.setRowCount(0)
        for p in rows:
            scad = datetime.strptime(p["scadenza"], "%Y-%m-%d").date()
            attivo = p["data_restituzione"] is None
            st = "In ritardo" if attivo and scad < date.today() else ("Attivo" if attivo else "Restituito")
            r = self.table_prestiti.rowCount()
            self.table_prestiti.insertRow(r)
            self.table_prestiti.setItem(r, 0, QTableWidgetItem(p["titolo"]))
            self.table_prestiti.setItem(r, 1, QTableWidgetItem(p["persona"]))
            self.table_prestiti.setItem(r, 2, QTableWidgetItem(p["data_prestito"]))
            self.table_prestiti.setItem(r, 3, QTableWidgetItem(p["scadenza"]))
            item = QTableWidgetItem(st)
            item.setForeground(STATUS_COLORS.get(st, QColor("#555")))
            self.table_prestiti.setItem(r, 4, item)
            self.table_prestiti.setItem(r, 5, QTableWidgetItem(p["data_restituzione"] or "—"))
            self._prestiti_rows = rows
        conn.close()

    def new_prestito(self):
        libro = self.current_libro()
        if not libro:
            QMessageBox.information(self, "Nessuna selezione", "Seleziona un libro dalla lista.")
            return
        conn = get_conn()
        p_map = prestiti_per_libro(conn)
        conn.close()
        if libro["id"] in p_map and p_map[libro["id"]]["data_restituzione"] is None:
            QMessageBox.warning(self, "Libro occupato", "Il libro è già in prestito.")
            return
        dlg = PrestitoDialog(self, libro["titolo"])
        if dlg.exec() == QDialog.DialogCode.Accepted:
            persona = dlg.persona.text().strip()
            if not persona:
                QMessageBox.warning(self, "Dati mancanti", "Indica chi prende in prestito il libro.")
                return
            conn = get_conn()
            conn.execute(
                "INSERT INTO prestiti (libro_id, persona, data_prestito, scadenza) VALUES (?,?,?,?)",
                (libro["id"], persona,
                 dlg.data_prestito.date().toString("yyyy-MM-dd"),
                 dlg.scadenza.date().toString("yyyy-MM-dd")))
            conn.commit()
            conn.close()
            self.load_libri()
            self.load_prestiti()
            self.load_statistiche()

    def return_libro(self):
        if self.tabs.currentIndex() == 0:
            libro = self.current_libro()
            if not libro:
                QMessageBox.information(self, "Nessuna selezione", "Seleziona un libro dalla lista.")
                return
            conn = get_conn()
            row = conn.execute(
                "SELECT * FROM prestiti WHERE libro_id=? AND data_restituzione IS NULL "
                "ORDER BY data_prestito DESC LIMIT 1", (libro["id"],)).fetchone()
            if row is None:
                conn.close()
                QMessageBox.information(self, "Nessun prestito", "Questo libro non è in prestito.")
                return
        else:
            row = self.table_prestiti.currentRow()
            if row < 0:
                QMessageBox.information(self, "Nessuna selezione", "Seleziona un prestito dalla lista.")
                return
            p = self._prestiti_rows[row]
            conn = get_conn()
            row = conn.execute("SELECT * FROM prestiti WHERE id=?", (p["id"],)).fetchone()
        ans = QMessageBox.question(
            self, "Restituzione",
            f"Registrare la restituzione di '{row['titolo'] if 'titolo' in row.keys() else ''}' "
            f"da parte di {row['persona']}?")
        if ans == QMessageBox.StandardButton.Yes:
            conn.execute("UPDATE prestiti SET data_restituzione=? WHERE id=?",
                         (date.today().isoformat(), row["id"]))
            conn.commit()
            conn.close()
            self.load_libri()
            self.load_prestiti()
            self.load_statistiche()

    # ------------------------------------------------------- STATISTICHE
    def build_statistiche_tab(self):
        tab = QWidget()
        lay = QVBoxLayout(tab)

        self.stats_grid = QGroupBox("Panoramica")
        grid = QGridLayout(self.stats_grid)
        self.stats_labels = {}
        stats_items = [
            ("totale", "Libri totali"),
            ("disponibili", "Disponibili"),
            ("prestito", "In prestito"),
            ("ritardo", "In ritardo"),
            ("persone", "Persone che hanno preso in prestito"),
            ("restituzioni", "Restituzioni registrate"),
        ]
        for i, (key, label) in enumerate(stats_items):
            grid.addWidget(QLabel(f"{label}:"), i % 2, (i // 2) * 2)
            val = QLabel("0")
            val.setFont(QFont("", 12, QFont.Weight.Bold))
            grid.addWidget(val, i % 2, (i // 2) * 2 + 1)
            self.stats_labels[key] = val
        lay.addWidget(self.stats_grid)

        self.stats_text = QLabel("")
        self.stats_text.setWordWrap(True)
        self.stats_text.setStyleSheet("padding: 8px;")
        lay.addWidget(self.stats_text)

        charts_row = QHBoxLayout()

        self.chart_stato = self._make_chart_view("Stato dei libri", self._make_stato_chart)
        charts_row.addWidget(self.chart_stato)

        self.chart_genere = self._make_chart_view("Distribuzione per genere", self._make_genere_chart)
        charts_row.addWidget(self.chart_genere)

        self.chart_decenni = self._make_chart_view("Libri per decennio", self._make_decenni_chart)
        charts_row.addWidget(self.chart_decenni)

        lay.addLayout(charts_row)

        refresh = QPushButton("Aggiorna")
        refresh.clicked.connect(self.load_statistiche)
        lay.addWidget(refresh)
        self.tabs.addTab(tab, "Statistiche")

    def _make_chart_view(self, title, builder):
        chart = QChart()
        chart.setTitle(title)
        chart.setAnimationOptions(QChart.AnimationOption.SeriesAnimations)
        chart.legend().setAlignment(Qt.AlignmentFlag.AlignBottom)
        chart.setBackgroundVisible(False)
        view = QChartView(chart)
        view.setRenderHint(view.renderHints())
        view.setMinimumHeight(220)
        builder(chart)
        return view

    def _make_stato_chart(self, chart):
        conn = get_conn()
        p_map = prestiti_per_libro(conn)
        libri = conn.execute("SELECT id FROM libri").fetchall()
        counts = {"Disponibile": 0, "In prestito": 0, "In ritardo": 0}
        for l in libri:
            counts[status_of(l["id"], p_map)] += 1
        conn.close()
        series = QPieSeries()
        colors = {"Disponibile": QColor("#2e7d32"),
                  "In prestito": QColor("#e65100"),
                  "In ritardo": QColor("#c62828")}
        for k, v in counts.items():
            if v > 0:
                sl = series.append(f"{k} ({v})", v)
                sl.setColor(colors[k])
        chart.addSeries(series)

    def _make_genere_chart(self, chart):
        conn = get_conn()
        rows = conn.execute(
            "SELECT genere, COUNT(*) AS n FROM libri "
            "WHERE genere IS NOT NULL AND genere != '' "
            "GROUP BY genere ORDER BY n DESC LIMIT 8").fetchall()
        conn.close()
        series = QPieSeries()
        palette = [QColor("#5c6bc0"), QColor("#26a69a"), QColor("#ff7043"),
                   QColor("#ab47bc"), QColor("#ffa726"), QColor("#66bb6a"),
                   QColor("#ef5350"), QColor("#42a5f5")]
        for i, r in enumerate(rows):
            sl = series.append(f"{r['genere']} ({r['n']})", r["n"])
            sl.setColor(palette[i % len(palette)])
        chart.addSeries(series)

    def _make_decenni_chart(self, chart):
        conn = get_conn()
        rows = conn.execute(
            "SELECT CAST(CAST(anno/10 AS INTEGER)*10 AS TEXT) AS decennio, COUNT(*) AS n "
            "FROM libri WHERE anno IS NOT NULL AND anno > 0 "
            "GROUP BY decennio ORDER BY decennio").fetchall()
        conn.close()
        bar = QBarSet("Libri")
        cats = []
        for r in rows:
            cats.append(r["decennio"])
            bar.append(r["n"])
        series = QBarSeries()
        series.append(bar)
        chart.addSeries(series)
        axis_x = QBarCategoryAxis()
        axis_x.append(cats)
        chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)
        series.attachAxis(axis_x)
        axis_y = QValueAxis()
        axis_y.setLabelFormat("%d")
        axis_y.setMin(0)
        chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)
        series.attachAxis(axis_y)

    def load_statistiche(self):
        conn = get_conn()
        p_map = prestiti_per_libro(conn)
        libri = conn.execute("SELECT id FROM libri").fetchall()
        totale = len(libri)
        disponibili = ritardo = prestito = 0
        for l in libri:
            st = status_of(l["id"], p_map)
            if st == "Disponibile":
                disponibili += 1
            elif st == "In ritardo":
                ritardo += 1
            else:
                prestito += 1
        persone = conn.execute(
            "SELECT COUNT(DISTINCT persona) FROM prestiti").fetchone()[0]
        restituzioni = conn.execute(
            "SELECT COUNT(*) FROM prestiti WHERE data_restituzione IS NOT NULL").fetchone()[0]

        self.stats_labels["totale"].setText(str(totale))
        self.stats_labels["disponibili"].setText(str(disponibili))
        self.stats_labels["prestito"].setText(str(prestito))
        self.stats_labels["ritardo"].setText(str(ritardo))
        self.stats_labels["persone"].setText(str(persone))
        self.stats_labels["restituzioni"].setText(str(restituzioni))

        prossime = conn.execute("""
            SELECT l.titolo, p.persona, p.scadenza FROM prestiti p
            JOIN libri l ON l.id = p.libro_id
            WHERE p.data_restituzione IS NULL
            ORDER BY p.scadenza LIMIT 5""").fetchall()
        if prossime:
            oggi = date.today()
            txt = "Prossime scadenze:\n"
            for p in prossime:
                scad = datetime.strptime(p["scadenza"], "%Y-%m-%d").date()
                diff = (scad - oggi).days
                suff = " (in ritardo!)" if diff < 0 else ("" if diff == 0 else f" (tra {diff} giorni)")
                txt += f"  • {p['titolo']} → {p['persona']}, scade il {p['scadenza']}{suff}\n"
            self.stats_text.setText(txt)
        else:
            self.stats_text.setText("Nessun prestito attivo.")
        conn.close()

        for name, builder in (("chart_stato", self._make_stato_chart),
                              ("chart_genere", self._make_genere_chart),
                              ("chart_decenni", self._make_decenni_chart)):
            view = getattr(self, name)
            old = view.chart()
            new = QChart()
            new.setTitle(old.title())
            new.setAnimationOptions(QChart.AnimationOption.SeriesAnimations)
            new.legend().setAlignment(Qt.AlignmentFlag.AlignBottom)
            new.setBackgroundVisible(False)
            builder(new)
            view.setChart(new)


def main():
    init_db()
    app = QApplication(sys.argv)
    win = MainWindow()
    win.aggiorna_generi()
    win.load_libri()
    win.load_prestiti()
    win.load_statistiche()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
