import sys
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QTabWidget, QMessageBox, QDialog, QFormLayout, QComboBox,
    QDateEdit, QSpinBox, QHeaderView, QAbstractItemView, QGroupBox,
    QGridLayout, QTextEdit
)
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtGui import QColor, QFont

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
        status_row = QHBoxLayout()
        status_row.addWidget(QLabel("Filtro stato:"))
        self.filter_status = QComboBox()
        self.filter_status.addItems(["Tutti", "Disponibile", "In prestito", "In ritardo"])
        self.filter_status.currentTextChanged.connect(self.load_libri)
        status_row.addWidget(self.filter_status)
        status_row.addStretch()
        lay.addLayout(search_row)
        lay.addLayout(status_row)

        self.table_libri = QTableWidget(0, 6)
        self.table_libri.setHorizontalHeaderLabels(["Titolo", "Autore", "Anno", "Genere", "Stato", "Prestito a"])
        self.table_libri.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_libri.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_libri.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table_libri.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table_libri.doubleClicked.connect(lambda: self.edit_libro())
        lay.addWidget(self.table_libri)

        btns = QHBoxLayout()
        b_add = QPushButton("Aggiungi libro")
        b_edit = QPushButton("Modifica")
        b_del = QPushButton("Elimina")
        b_prestito = QPushButton("Registra prestito")
        b_rest = QPushButton("Restituisci")
        b_add.clicked.connect(lambda: self.add_libro())
        b_edit.clicked.connect(lambda: self.edit_libro())
        b_del.clicked.connect(lambda: self.delete_libro())
        b_prestito.clicked.connect(lambda: self.new_prestito())
        b_rest.clicked.connect(lambda: self.return_libro())
        for b in (b_add, b_edit, b_del, b_prestito, b_rest):
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
        conn = get_conn()
        p_map = prestiti_per_libro(conn)
        rows = conn.execute("SELECT * FROM libri ORDER BY titolo COLLATE NOCASE").fetchall()
        self._libri = []
        self.table_libri.setRowCount(0)
        for libro in rows:
            st = status_of(libro["id"], p_map)
            if filtro != "Tutti" and st != filtro:
                continue
            if q and q.lower() not in " ".join(filter(None, [
                    libro["titolo"], libro["autore"], libro["genere"], libro["isbn"]
            ])).lower():
                continue
            persona = ""
            if libro["id"] in p_map and p_map[libro["id"]]["data_restituzione"] is None:
                persona = p_map[libro["id"]]["persona"]
            r = self.table_libri.rowCount()
            self.table_libri.insertRow(r)
            self.table_libri.setItem(r, 0, QTableWidgetItem(libro["titolo"]))
            self.table_libri.setItem(r, 1, QTableWidgetItem(libro["autore"]))
            self.table_libri.setItem(r, 2, QTableWidgetItem(str(libro["anno"]) if libro["anno"] else "—"))
            self.table_libri.setItem(r, 3, QTableWidgetItem(libro["genere"] or "—"))
            item_st = QTableWidgetItem(st)
            item_st.setForeground(STATUS_COLORS[st])
            self.table_libri.setItem(r, 4, item_st)
            self.table_libri.setItem(r, 5, QTableWidgetItem(persona))
            self._libri.append(libro)
        conn.close()

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
            grid.addWidget(QLabel(f"{label}:"), i, 0)
            val = QLabel("0")
            val.setFont(QFont("", 12, QFont.Weight.Bold))
            grid.addWidget(val, i, 1)
            self.stats_labels[key] = val
        lay.addWidget(self.stats_grid)

        self.stats_text = QLabel("")
        self.stats_text.setWordWrap(True)
        lay.addWidget(self.stats_text)
        lay.addStretch()

        refresh = QPushButton("Aggiorna")
        refresh.clicked.connect(self.load_statistiche)
        lay.addWidget(refresh)
        self.tabs.addTab(tab, "Statistiche")

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


def main():
    init_db()
    app = QApplication(sys.argv)
    win = MainWindow()
    win.load_libri()
    win.load_prestiti()
    win.load_statistiche()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
