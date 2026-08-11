# ============================================================================
# RUBRICA BIBLIOTECA - gestione di libri e prestiti (applicazione desktop)
# ----------------------------------------------------------------------------
# Questo programma è scritto in Python 3 e usa PyQt6 (interfaccia grafica) e
# SQLite (database locale). Il file biblioteca.db contiene tutti i dati e
# viene creato automaticamente nella stessa cartella di questo script.
#
# STRUTTURA DEL PROGRAMMA
# -----------------------
# 1. IMPORTAZIONI      : le librerie di cui abbiamo bisogno
# 2. FUNZIONI AIUTO    : download copertine, connessione al database
# 3. DIALOG (finestre) : finestre per inserire/vedere dati
# 4. CLASSE PRINCIPALE : la finestra principale dell'app con tutte le schede
# 5. AVVIO            : il punto di ingresso (main)
#
# Ogni blocco di codice è commentato riga per riga in italiano.
# ============================================================================

# ---------------------------------------------------------------------------
# 1. IMPORTAZIONI DELLE LIBRERIE
# ---------------------------------------------------------------------------
import csv          # Modulo standard: serve per esportare i libri in file CSV (Excel).
import io           # Modulo standard: flussi di input/output in memoria (non sempre usato, ma utile).
import json         # Modulo standard: per leggere le risposte JSON dell'API OpenLibrary.
import sqlite3      # Modulo standard: driver per il database SQLite (salvataggio dati).
import sys          # Modulo standard: per gli argomenti della riga di comando e uscita del programma.
import urllib.parse # Modulo standard: per costruire URL sicuri (codifica testo nell'indirizzo web).
import urllib.request  # Modulo standard: per scaricare dati da internet (copertine dei libri).
from datetime import date, datetime, timedelta  # Per lavorare con date e scadenze.
from pathlib import Path  # Per gestire i percorsi dei file in modo semplice e cross-platform.

# Widget grafici di PyQt6: ogni nome è una classe che disegna una parte dell'interfaccia.
from PyQt6.QtWidgets import (
    QApplication,       # Oggetto che gestisce l'intera applicazione Qt (crea l'event loop).
    QMainWindow,        # Classe base per la finestra principale (con barra del titolo, menu, statusbar).
    QWidget,            # Widget generico: contenitore base per qualsiasi elemento.
    QVBoxLayout,        # Layout verticale: impila i widget uno sotto l'altro.
    QHBoxLayout,        # Layout orizzontale: mette i widget uno accanto all'altro.
    QLabel,             # Etichetta di testo statico.
    QLineEdit,          # Casella di testo a riga singola (es. per digitare un nome).
    QPushButton,        # Pulsante cliccabile.
    QTableWidget,       # Tabella: mostra dati in righe e colonne (la lista dei libri).
    QTableWidgetItem,   # Cella di una QTableWidget (il singolo contenuto della cella).
    QTabWidget,         # Contenitore a schede (Libri / Prestiti / Statistiche).
    QMessageBox,        # Finestre di dialogo standard (avvisi, conferme, errori).
    QDialog,            # Finestra di dialogo personalizzata (modale).
    QFormLayout,        # Layout a modulo: righe composte da "etichetta + campo".
    QComboBox,          # Menu a tendina (per scegliere un'opzione da una lista).
    QDateEdit,          # Campo per inserire una data (con calendario).
    QSpinBox,           # Campo numerico con frecce su/giù.
    QHeaderView,        # Gestisce le intestazioni della tabella (larghezze colonne).
    QAbstractItemView,  # Classe base per la tabella: contiene impostazioni di selezione.
    QGroupBox,          # Riquadro con bordo e titolo, per raggruppare elementi.
    QGridLayout,        # Layout a griglia: posiziona i widget in righe e colonne.
    QTextEdit,          # Area di testo multilinea (per le note).
    QFileDialog         # Finestra per scegliere file da salvare/aprire.
)
from PyQt6.QtCore import Qt, QDate, QSize, QThread, pyqtSignal  # Funzioni base di Qt.
from PyQt6.QtGui import QColor, QFont, QPixmap, QIcon            # Elementi grafici (colori, font, immagini, icone).
from PyQt6.QtCharts import (    # Libreria per disegnare i grafici delle statistiche.
    QChart,                     # Il contenitore del grafico.
    QChartView,                 # Il widget che mostra il grafico a schermo.
    QPieSeries,                 # Serie dati per i grafici a torta.
    QPieSlice,                  # Una singola "fetta" del grafico a torta.
    QBarSet,                    # Un gruppo di barre nel grafico a barre.
    QBarSeries,                 # Serie dati per i grafici a barre.
    QBarCategoryAxis,           # Asse orizzontale delle categorie (etichette testuali).
    QValueAxis                  # Asse numerico (per i valori sull'asse Y).
)

# ---------------------------------------------------------------------------
# 2. FUNZIONI DI SUPPORTO
# ---------------------------------------------------------------------------

# COVERS_DIR: la cartella dove salviamo le immagini delle copertine scaricate.
# __file__ è il percorso di questo script; .resolve() elimina eventuali link simbolici;
# .parent prende la cartella che contiene lo script. Quindi i file delle copertine
# vengono messi in una sottocartella "copertine" accanto a biblioteca.py.
COVERS_DIR = Path(__file__).resolve().parent / "copertine"

# COVER_URL_CACHE: un dizionario che ricorda gli URL delle copertine già cercati.
# Serve per non fare la stessa richiesta a internet ogni volta: la chiave è la
# coppia (titolo, autore), il valore è l'URL della copertina (o None se non trovata).
COVER_URL_CACHE = {}


def trova_copertina_url(titolo, autore):
    """Cerca su OpenLibrary l'URL della copertina di un libro.

    - titolo: il titolo del libro (stringa)
    - autore: l'autore del libro (stringa)
    Restituisce una stringa URL se la copertina esiste, altrimenti None.
    I risultati vengono memorizzati in COVER_URL_CACHE per evitare ripetute
    richieste di rete (la rete è lenta e va usata il meno possibile).
    """
    # Creo una chiave normalizzata: minuscole e senza spazi in più, per confrontare
    # in modo uniforme anche se l'utente digita con maiuscole diverse.
    key = (titolo.lower().strip(), autore.lower().strip())

    # Se questa coppia è già stata cercata, restituisco subito il risultato salvato
    # senza contattare internet (velocizza molto quando la lista è lunga).
    if key in COVER_URL_CACHE:
        return COVER_URL_CACHE[key]

    # Costruisco la query di ricerca: l'API di OpenLibrary accetta "?q=testo".
    # urllib.parse.quote codifica il testo in modo che spazi e caratteri speciali
    # siano validi dentro un URL (es. lo spazio diventa %20).
    q = urllib.parse.quote(f"{titolo} {autore}")

    # URL completo dell'API di ricerca. Richiedo solo i campi che mi servono:
    # title, author_name e cover_i (l'identificativo della copertina).
    url = (f"https://openlibrary.org/search.json?q={q}"
           "&limit=3&fields=title,author_name,cover_i")

    # Inizio un blocco di gestione errori: se la rete non funziona o l'API cambia,
    # il programma non deve andare in crash ma semplicemente rinunciare alla copertina.
    try:
        # Preparo la richiesta HTTP. Il "User-Agent" identifica il nostro programma
        # al server remoto: alcuni server lo richiedono e permette di essere educati.
        req = urllib.request.Request(url, headers={"User-Agent": "RubricaBiblioteca/1.0"})

        # Eseguo la richiesta con un timeout di 8 secondi: se il server non risponde
        # entro quel tempo, viene sollevata un'eccezione ed esco dal try.
        # "with ... as resp" garantisce che la connessione venga chiusa da sola.
        with urllib.request.urlopen(req, timeout=8) as resp:
            # Leggo la risposta e la converto da formato JSON a struttura Python
            # (dizionari e liste). È il formato che OpenLibrary usa per i risultati.
            data = json.load(resp)

        # "docs" è la lista dei libri trovati. Scorro i primi risultati finché
        # non trovo uno che ha una copertina (campo cover_i valorizzato).
        for d in data.get("docs", []):
            ci = d.get("cover_i")   # Leggo l'identificativo della copertina (numero).
            if ci:                  # Se esiste un numero, la copertina c'è davvero.
                # Costruisco l'URL finale dell'immagine. La "-L" indica la versione
                # grande dell'immagine, adatta per essere mostrata nei dettagli.
                COVER_URL_CACHE[key] = f"https://covers.openlibrary.org/b/id/{ci}-L.jpg"
                return COVER_URL_CACHE[key]   # Salvo in cache e restituisco l'URL.
    except Exception:
        pass   # Qualunque errore di rete: semplicemente ignoro e vado avanti.

    # Se sono arrivato qui significa che nessun risultato aveva una copertina.
    # Salvo "None" in cache per non ricercarlo di nuovo, e restituisco None.
    COVER_URL_CACHE[key] = None
    return None


def file_copertina(libro_id):
    """Restituisce il percorso del file immagine locale di una copertina.

    - libro_id: l'id numerico del libro nel database
    Ogni copertina viene salvata come "copertine/<id>.jpg": usiamo l'id come nome
    file perché è unico e non cambia nemmeno se il titolo viene modificato.
    """
    # Costruisco un oggetto Path unendo la cartella e il nome file.
    return COVERS_DIR / f"{libro_id}.jpg"


class CopertinaWorker(QThread):
    """Thread per scaricare una copertina in background senza bloccare l'interfaccia.

    Le richieste a internet sono lente (alcuni secondi). Se le facessi sul filo
    principale, la finestra "si congelerebbe" finché il download non finisce.
    Per questo uso un QThread: il download avviene su un filo separato e la
    finestra resta sempre reattiva. Quando il download è finito, il thread
    emette il segnale "fatta" con l'id del libro.
    """
    # Segnale personalizzato: viene emesso a download completato.
    # Il parametro int sarà l'id del libro. I segnali sono il modo in cui
    # un thread comunica con l'interfaccia (in modo sicuro).
    fatta = pyqtSignal(int)

    def __init__(self, libro_id, url, percorso):
        """Costruttore: prepara i dati che serviranno durante il download.

        - libro_id: id del libro (per sapere a quale riga assegnare l'immagine)
        - url: indirizzo internet da cui scaricare l'immagine
        - percorso: dove salvare il file immagine sul disco
        """
        super().__init__()      # Chiamo il costruttore della classe madre QThread.
        self.libro_id = libro_id  # Salvo l'id del libro su questo oggetto.
        self.url = url            # Salvo l'URL da scaricare.
        self.percorso = percorso  # Salvo il percorso del file di destinazione.

    def run(self):
        """Metodo eseguito automaticamente quando il thread parte (.start()).

        Questo codice gira in background: se fallisce, l'interfaccia non se ne accorge.
        """
        try:
            # Preparo la richiesta HTTP con User-Agent.
            req = urllib.request.Request(
                self.url, headers={"User-Agent": "RubricaBiblioteca/1.0"})

            # Scarico il contenuto dell'immagine (i byte del file .jpg) con timeout.
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = resp.read()   # data contiene i byte grezzi dell'immagine.

            # Controllo che l'immagine non sia vuota: le risposte "non trovato"
            # di OpenLibrary sono piccolissime (meno di 1500 byte). Se è troppo
            # piccola, la scarto (probabilmente è un segnaposto vuoto).
            if len(data) > 1500:
                # Creo la cartella "copertine" se non esiste (parents=True crea
                # anche le sottocartelle mancanti, exist_ok=True evita errori
                # se esiste già).
                self.percorso.parent.mkdir(parents=True, exist_ok=True)
                # Salvo i byte dell'immagine nel file su disco.
                self.percorso.write_bytes(data)
                # Emetto il segnale "fatta" con l'id del libro: così l'interfaccia
                # sa che può mostrare la copertina appena scaricata.
                self.fatta.emit(self.libro_id)
        except Exception:
            pass   # Errore di rete o altro: rinuncio silenziosamente.


def scarica_copertina(titolo, autore):
    """Scarica e restituisce subito l'immagine della copertina come QPixmap.

    Usata dalla finestra "Dettagli libro" che ha bisogno dell'immagine pronta
    immediatamente. QPixmap è il tipo di Qt che rappresenta un'immagine disegnabile.
    """
    # Prima cerco l'URL della copertina (funzione precedente, usa la cache).
    url = trova_copertina_url(titolo, autore)

    # Se non c'è un URL valido, restituisco None (nessuna copertina disponibile).
    if not url:
        return None

    try:
        # Preparo e eseguo la richiesta HTTP come sopra.
        req = urllib.request.Request(url, headers={"User-Agent": "RubricaBiblioteca/1.0"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = resp.read()   # Leggo i byte dell'immagine.

        pix = QPixmap()          # Creo un contenitore immagine vuoto.
        if pix.loadFromData(data):  # Carico i byte come immagine (True se valida).
            return pix           # Restituisco l'immagine pronta per essere disegnata.
    except Exception:
        pass   # Errore di rete: restituisco None qui sotto.
    return None   # Nessuna immagine disponibile.


# DB_PATH: il percorso completo del file del database. Come per le copertine,
# viene creata accanto allo script. Qui SQLite salverà libri e prestiti.
DB_PATH = Path(__file__).resolve().parent / "biblioteca.db"

# PRESTITI_EXTRA_COLONNE: le colonne aggiuntive della tabella prestiti introdotte
# in una versione successiva dell'app. Se un utente ha già un database creato con
# la prima versione (che aveva solo: persona, data_prestito, data_restituzione,
# scadenza), queste colonne non esistono. Il dizionario mappa il nome della
# colonna con il suo tipo SQL, e serve alla migrazione automatica qui sotto.
PRESTITI_EXTRA_COLONNE = {
    "email": "TEXT",     # Indirizzo email della persona che prende il libro.
    "telefono": "TEXT",  # Numero di telefono della persona.
    "note": "TEXT",      # Note libere sul prestito (es. "chiamare prima della scadenza").
}


def get_conn():
    """Apre e restituisce una connessione al database SQLite.

    Ogni volta che dobbiamo leggere o scrivere dati chiamiamo questa funzione,
    che crea una nuova connessione già configurata correttamente.
    """
    # Apro la connessione al file del database (se il file non esiste lo crea).
    conn = sqlite3.connect(DB_PATH)

    # row_factory = sqlite3.Row rende ogni riga del database accessibile sia con
    # il nome della colonna (riga["titolo"]) sia con l'indice (riga[0]). Senza
    # questa riga, ogni risultato sarebbe solo una tupla numerica poco leggibile.
    conn.row_factory = sqlite3.Row

    # PRAGMA foreign_keys = ON attiva il controllo delle chiavi esterne: quando
    # eliminiamo un libro, SQLite eliminerà automaticamente i suoi prestiti
    # (grazie a ON DELETE CASCADE definito nella creazione della tabella).
    conn.execute("PRAGMA foreign_keys = ON")
    return conn   # Restituisco la connessione pronta all'uso.


def init_db():
    """Crea le tabelle del database se non esistono e fa la migrazione dati.

    Viene chiamata a ogni avvio dell'app. È sicuro chiamarla più volte perché
    CREATE TABLE IF NOT EXISTS non fa nulla se la tabella esiste già.
    """
    conn = get_conn()   # Apro la connessione.

    # executescript esegue più istruzioni SQL in un colpo solo, separate da ";".
    # Qui definisco lo schema delle due tabelle.
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS libri (          -- Tabella dei libri.
            id INTEGER PRIMARY KEY AUTOINCREMENT,   -- Numero progressivo unico (1,2,3...).
            titolo TEXT NOT NULL,                   -- Titolo del libro (obbligatorio).
            autore TEXT NOT NULL,                   -- Autore del libro (obbligatorio).
            anno INTEGER,                           -- Anno di pubblicazione (può mancare).
            genere TEXT,                            -- Genere letterario (romanzo, giallo...).
            isbn TEXT,                              -- Codice ISBN (può mancare).
            note TEXT,                              -- Note libere sul libro.
            data_aggiunta TEXT NOT NULL             -- Data di inserimento nel catalogo.
        );
        CREATE TABLE IF NOT EXISTS prestiti (       -- Tabella dei prestiti.
            id INTEGER PRIMARY KEY AUTOINCREMENT,   -- Numero progressivo unico.
            libro_id INTEGER NOT NULL REFERENCES libri(id) ON DELETE CASCADE,
                                                    -- A quale libro si riferisce il prestito.
                                                    -- REFERENCES + ON DELETE CASCADE: se il libro
                                                    -- viene eliminato, sparisce anche il prestito.
            persona TEXT NOT NULL,                  -- Nome di chi prende il libro (obbligatorio).
            email TEXT,                             -- Email della persona.
            telefono TEXT,                          -- Telefono della persona.
            note TEXT,                              -- Note sul prestito.
            data_prestito TEXT NOT NULL,            -- Data in cui è iniziato il prestito.
            data_restituzione TEXT,                 -- Data di restituzione (None = non restituito).
            scadenza TEXT NOT NULL                  -- Data entro cui va restituito.
        );
    """)

    # Migrazione: controllo quali colonne esistono davvero nella tabella prestiti.
    # PRAGMA table_info(prestiti) restituisce le definizioni delle colonne;
    # prendo solo i nomi (secondo elemento di ogni riga) e li metto in un set.
    colonne_esistenti = {r[1] for r in conn.execute("PRAGMA table_info(prestiti)")}

    # Per ogni colonna nuova prevista (email, telefono, note), se non esiste già
    # la aggiungo con ALTER TABLE. Così un database vecchio viene aggiornato
    # senza perdere i dati già salvati.
    for nome, tipo in PRESTITI_EXTRA_COLONNE.items():
        if nome not in colonne_esistenti:   # Se la colonna manca...
            conn.execute(f"ALTER TABLE prestiti ADD COLUMN {nome} {tipo}")  # ...la aggiungo.

    conn.commit()   # Salvo definitivamente le modifiche sul file.
    conn.close()    # Chiudo la connessione (importante per non bloccare il file).


# STATUS_COLORS: associa a ogni stato di un libro il colore con cui viene
# mostrato nelle tabelle. Verde = disponibile, arancione = in prestito,
# rosso = in ritardo. QColor è un colore usabile direttamente da Qt.
STATUS_COLORS = {
    "Disponibile": QColor("#2e7d32"),   # Verde scuro.
    "In prestito": QColor("#e65100"),   # Arancione scuro.
    "In ritardo": QColor("#c62828"),    # Rosso scuro.
}


# ---------------------------------------------------------------------------
# 3. FINESTRE DI DIALOGO
# ---------------------------------------------------------------------------

class LibroDialog(QDialog):
    """Finestra per aggiungere o modificare un libro.

    Ha un campo per ogni informazione del libro (titolo, autore, anno, genere,
    ISBN, note). Quando l'utente preme "Salva", il metodo data() restituisce un
    dizionario con tutti i valori inseriti. Se viene passato l'argomento libro
    (un record dal database), i campi vengono precompilati: stessa finestra
    usata sia per "aggiungi" sia per "modifica".
    """
    def __init__(self, parent=None, libro=None):
        """Costruttore.

        - parent: la finestra che apre questo dialogo (serve per la centratura).
        - libro: se fornito, i campi vengono riempiti con i suoi dati (modifica).
        """
        super().__init__(parent)   # Chiamo il costruttore di QDialog.

        # Imposto il titolo della finestra: se libro è None siamo in modalità
        # "aggiungi", altrimenti in modalità "modifica".
        self.setWindowTitle("Aggiungi libro" if libro is None else "Modifica libro")

        # Larghezza minima della finestra (in pixel): evita che venga schiacciata.
        self.setMinimumWidth(380)

        # Creo un layout a modulo: ogni riga è composta da "etichetta : campo".
        form = QFormLayout(self)

        # Creo i campi di inserimento (widget). Ogni riga crea il campo e poi lo
        # aggiunge al modulo con una etichetta. L'asterisco (*) indica obbligatorio.
        self.titolo = QLineEdit()      # Casella di testo per il titolo.
        self.autore = QLineEdit()      # Casella di testo per l'autore.
        self.anno = QSpinBox()         # Campo numerico per l'anno.
        self.anno.setRange(0, 2100)    # Intervallo accettato (dal 0 al 2100).
        self.anno.setSpecialValueText("Sconosciuto")  # Se il valore è 0 mostra la scritta "Sconosciuto".
        self.genere = QLineEdit()      # Casella di testo per il genere.
        self.isbn = QLineEdit()        # Casella di testo per l'ISBN.
        self.note = QTextEdit()        # Area di testo per le note.
        self.note.setMaximumHeight(70) # Limito l'altezza dell'area note.

        # Aggiungo ogni campo al modulo con la sua etichetta.
        form.addRow("Titolo *", self.titolo)   # Riga "Titolo *" con il campo titolo.
        form.addRow("Autore *", self.autore)   # Riga "Autore *" con il campo autore.
        form.addRow("Anno", self.anno)         # Riga "Anno".
        form.addRow("Genere", self.genere)     # Riga "Genere".
        form.addRow("ISBN", self.isbn)         # Riga "ISBN".
        form.addRow("Note", self.note)         # Riga "Note".

        # Se siamo in modalità "modifica" (libro passato), riempio i campi con i
        # valori attuali del libro. "or 0"/"or ''" gestisce i valori vuoti (None).
        if libro is not None:
            self.titolo.setText(libro["titolo"])     # Inserisco il titolo corrente.
            self.autore.setText(libro["autore"])     # Inserisco l'autore corrente.
            self.anno.setValue(libro["anno"] or 0)   # Inserisco l'anno (0 se manca).
            self.genere.setText(libro["genere"] or "")  # Inserisco il genere (vuoto se manca).
            self.isbn.setText(libro["isbn"] or "")   # Inserisco l'ISBN (vuoto se manca).
            self.note.setPlainText(libro["note"] or "")  # Inserisco le note.

        # Riga finale con i pulsanti "Salva" e "Annulla".
        btns = QHBoxLayout()            # Layout orizzontale per i due pulsanti.
        save = QPushButton("Salva")     # Pulsante Salva.
        cancel = QPushButton("Annulla") # Pulsante Annulla.
        save.clicked.connect(self.accept)    # Salva chiude il dialogo con risultato "accettato".
        cancel.clicked.connect(self.reject)  # Annulla chiude con risultato "rifiutato".
        btns.addWidget(save)            # Metto Salva nel layout.
        btns.addWidget(cancel)          # Metto Annulla nel layout.
        form.addRow(btns)               # Aggiungo la riga dei pulsanti al modulo.

    def data(self):
        """Raccoglie i valori inseriti nei campi e li restituisce come dizionario.

        .text() legge il testo della casella; .strip() toglie gli spazi all'inizio
        e alla fine (così "  Mario  " diventa "Mario"). I campi vuoti diventano
        None nel database (valore NULL) grazie a "or None".
        """
        return {
            "titolo": self.titolo.text().strip(),        # Titolo pulito da spazi.
            "autore": self.autore.text().strip(),        # Autore pulito.
            "anno": self.anno.value() or None,           # Anno; 0 diventa None (sconosciuto).
            "genere": self.genere.text().strip() or None,  # Genere; vuoto diventa None.
            "isbn": self.isbn.text().strip() or None,    # ISBN; vuoto diventa None.
            "note": self.note.toPlainText().strip() or None,  # Note; vuote diventano None.
        }


class PrestitoDialog(QDialog):
    """Finestra per registrare un nuovo prestito o vedere i dettagli di uno esistente.

    Oltre al nome della persona ora raccoglie anche i dati di contatto (email,
    telefono) e delle note, così la biblioteca trattiene più informazioni su chi
    prende in prestito i libri. Se viene passato un prestito già esistente e
    restituito, i campi vengono disabilitati (sola lettura).
    """
    def __init__(self, parent=None, libro_titolo="", giorni=30, prestito=None):
        """Costruttore.

        - parent: finestra che apre il dialogo.
        - libro_titolo: il titolo del libro da mostrare in alto (per chiarezza).
        - giorni: scadenza predefinita in giorni (default 30).
        - prestito: se fornito, mostra i dati di un prestito esistente (dettagli).
        """
        super().__init__(parent)   # Chiamo il costruttore di QDialog.

        # Titolo della finestra: "Nuovo prestito" o "Dettagli prestito".
        self.setWindowTitle("Nuovo prestito" if prestito is None else "Dettagli prestito")

        # Larghezza minima della finestra.
        self.setMinimumWidth(380)

        # Layout a modulo.
        form = QFormLayout(self)

        # Prima riga: il titolo del libro, in sola lettura (non modificabile).
        form.addRow("Libro", QLabel(libro_titolo))

        # Campo per il nome della persona che prende il libro (obbligatorio).
        self.persona = QLineEdit()

        # Nuovi campi di contatto (aggiunti per arricchire i dati della biblioteca).
        self.email = QLineEdit()          # Indirizzo email del lettore.
        self.telefono = QLineEdit()       # Numero di telefono del lettore.
        self.note = QTextEdit()           # Note sul prestito (es. condizioni, promemoria).
        self.note.setMaximumHeight(70)    # Altezza limitata dell'area note.

        # Data di inizio prestito: preimpostata a oggi, con calendario a comparsa.
        self.data_prestito = QDateEdit(QDate.currentDate())  # QDate.currentDate() = oggi.
        self.data_prestito.setCalendarPopup(True)            # Mostra il calendario.
        self.data_prestito.setDisplayFormat("dd/MM/yyyy")    # Formato italiano data.

        # Data di scadenza: oggi + "giorni" (default 30). Se l'utente la vuole
        # diversa può cambiarla dal calendario.
        self.scadenza = QDateEdit(QDate.currentDate().addDays(giorni))
        self.scadenza.setCalendarPopup(True)     # Mostra il calendario.
        self.scadenza.setDisplayFormat("dd/MM/yyyy")   # Formato italiano.

        # Aggiungo i campi al modulo con le loro etichette.
        form.addRow("Persona *", self.persona)     # Nome del lettore (obbligatorio).
        form.addRow("Email", self.email)           # Email del lettore.
        form.addRow("Telefono", self.telefono)     # Telefono del lettore.
        form.addRow("Note", self.note)             # Note sul prestito.
        form.addRow("Data prestito", self.data_prestito)  # Data di inizio.
        form.addRow("Scadenza", self.scadenza)     # Data di scadenza.

        # Se stiamo mostrando un prestito esistente (dettagli), riempio i campi.
        if prestito is not None:
            self.persona.setText(prestito["persona"])       # Nome salvato.
            self.email.setText(prestito["email"] or "")     # Email salvata (vuota se None).
            self.telefono.setText(prestito["telefono"] or "")  # Telefono salvato.
            self.note.setPlainText(prestito["note"] or "")  # Note salvate.
            # Converto la data dal formato di salvataggio (yyyy-MM-dd) al formato QDate.
            self.data_prestito.setDate(QDate.fromString(prestito["data_prestito"], "yyyy-MM-dd"))
            self.scadenza.setDate(QDate.fromString(prestito["scadenza"], "yyyy-MM-dd"))

            # Se il prestito è già stato restituito (data_restituzione non vuota),
            # disabilito tutti i campi: è una consultazione in sola lettura, così
            # l'utente non può modificare per errore una storia già chiusa.
            if prestito["data_restituzione"]:
                self.persona.setEnabled(False)       # Blocco il campo nome.
                self.email.setEnabled(False)         # Blocco il campo email.
                self.telefono.setEnabled(False)      # Blocco il campo telefono.
                self.note.setEnabled(False)          # Blocco l'area note.
                self.data_prestito.setEnabled(False) # Blocco la data prestito.
                self.scadenza.setEnabled(False)      # Blocco la scadenza.

        # Pulsanti Registra / Annulla come nella finestra libro.
        btns = QHBoxLayout()
        save = QPushButton("Registra")
        cancel = QPushButton("Annulla")
        save.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        btns.addWidget(save)
        btns.addWidget(cancel)
        form.addRow(btns)

    def data(self):
        """Restituisce tutti i valori del modulo come dizionario.

        Le date vengono convertite nel formato "yyyy-MM-dd" (es. 2026-08-10),
        che è quello standard usato nel database (e confrontabile come testo).
        """
        return {
            "persona": self.persona.text().strip(),  # Nome del lettore pulito.
            "email": self.email.text().strip() or None,  # Email; vuota diventa None.
            "telefono": self.telefono.text().strip() or None,  # Telefono; vuoto diventa None.
            "note": self.note.toPlainText().strip() or None,   # Note; vuote diventano None.
            "data_prestito": self.data_prestito.date().toString("yyyy-MM-dd"),  # Data inizio.
            "scadenza": self.scadenza.date().toString("yyyy-MM-dd"),  # Data scadenza.
        }


class DettagliLibroDialog(QDialog):
    """Finestra che mostra le informazioni complete di un libro con la copertina.

    A differenza della tabella (che mostra poche colonne), qui vediamo tutto:
    titolo in grande, autore, anno, genere, ISBN, stato, note e la copertina
    scaricata da internet.
    """
    def __init__(self, parent=None, libro=None, stato=""):
        """Costruttore.

        - parent: finestra che apre il dialogo.
        - libro: il record del libro da mostrare (dizionario stile sqlite.Row).
        - stato: lo stato del libro ("Disponibile", "In prestito", "In ritardo").
        """
        super().__init__(parent)   # Chiamo il costruttore di QDialog.
        self.setWindowTitle("Dettagli libro")   # Titolo della finestra.
        self.setMinimumSize(480, 320)           # Dimensione minima (larghezza, altezza).

        # Layout orizzontale: a sinistra la copertina, a destra le informazioni.
        lay = QHBoxLayout(self)

        # Etichetta per la copertina. Di default mostra un messaggio segnaposto.
        # (Il testo "senza ISBN" è rimasto dalla prima versione; ora la ricerca
        # è per titolo+autore, ma se non si trova copertina resta comunque.)
        self.cover_label = QLabel("Nessuna copertina\n(senza ISBN)")
        self.cover_label.setAlignment(Qt.AlignmentFlag.AlignCenter)  # Testo centrato.
        self.cover_label.setFixedSize(150, 220)   # Dimensioni fisse dell'area immagine.
        self.cover_label.setStyleSheet(           # Stile a cornice tratteggiata.
            "border: 1px dashed #888; border-radius: 6px; color: #777;")
        lay.addWidget(self.cover_label)           # Aggiungo l'etichetta a sinistra.

        # Layout verticale per le informazioni sulla destra.
        info = QVBoxLayout()

        # Titolo in grande e in grassetto (font size 16).
        titolo = QLabel(libro["titolo"])
        titolo.setFont(QFont("", 16, QFont.Weight.Bold))  # Grassetto.
        titolo.setWordWrap(True)   # Se il titolo è lungo, va a capo invece di allargare.
        info.addWidget(titolo)

        # Autore con font leggermente più piccolo.
        autore = QLabel(libro["autore"])
        autore.setFont(QFont("", 12))
        info.addWidget(autore)

        # Lista di dettagli base: coppie (etichetta, valore).
        det = [
            ("Anno", str(libro["anno"]) if libro["anno"] else "—"),  # Anno o trattino.
            ("Genere", libro["genere"] or "—"),                      # Genere o trattino.
            ("ISBN", libro["isbn"] or "—"),                          # ISBN o trattino.
            ("Stato", stato),                                        # Stato del libro.
        ]

        # Per ogni coppia creo una riga orizzontale "Etichetta: Valore".
        for k, v in det:
            row = QHBoxLayout()              # Layout orizzontale per la riga.
            kk = QLabel(f"<b>{k}:</b>")      # Etichetta in grassetto (HTML supportato).
            vv = QLabel(v)                   # Valore normale.
            row.addWidget(kk)                # Aggiungo l'etichetta.
            row.addWidget(vv)                # Aggiungo il valore.
            row.addStretch()                 # Spazio vuoto elastico per allineare a sinistra.
            info.addLayout(row)              # Aggiungo la riga al layout verticale.

        # Se il libro ha delle note, le mostro sotto i dettagli.
        if libro["note"]:
            info.addWidget(QLabel("<b>Note:</b>"))   # Etichetta "Note" in grassetto.
            note = QLabel(libro["note"])             # Il testo delle note.
            note.setWordWrap(True)                   # Va a capo se lungo.
            note.setStyleSheet("color: #666;")       # Testo grigio più discreto.
            info.addWidget(note)

        info.addStretch()   # Spazio elastico: spinge il pulsante in basso.

        # Pulsante Chiudi in basso a destra.
        chiudi = QPushButton("Chiudi")
        chiudi.clicked.connect(self.accept)   # Chiude la finestra.
        info.addWidget(chiudi, alignment=Qt.AlignmentFlag.AlignRight)  # Allineato a destra.
        lay.addLayout(info)   # Aggiungo la parte destra al layout principale.

        # Scarico la copertina reale (in modo sincrono, mostrando una piccola
        # attesa) e la mostro nell'etichetta. scarica_copertina restituisce una
        # QPixmap o None se non trovata.
        pix = scarica_copertina(libro["titolo"], libro["autore"])
        if pix and not pix.isNull():   # Se l'immagine esiste ed è valida...
            # La ridimensiono a massimo 150x220 mantenendo le proporzioni
            # (KeepAspectRatio) e con filtro morbido (SmoothTransformation).
            self.cover_label.setPixmap(pix.scaled(
                150, 220, Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation))
            self.cover_label.setStyleSheet("border: none;")  # Tolgo la cornice tratteggiata.


# ---------------------------------------------------------------------------
# 4. FUNZIONI UTILI SUI DATI (non legate all'interfaccia)
# ---------------------------------------------------------------------------

def status_of(libro_id, prestiti_map):
    """Calcola lo stato attuale di un libro.

    - libro_id: l'id del libro da valutare.
    - prestiti_map: un dizionario che associa ogni id di libro al suo prestito
      più recente (costruito da prestiti_per_libro). Passarlo già pronto evita
      di interrogare il database centinaia di volte.

    Restituisce una stringa: "Disponibile", "In prestito" o "In ritardo".
    """
    # Se il libro ha un prestito nel dizionario, c'è un prestito da esaminare.
    if libro_id in prestiti_map:
        p = prestiti_map[libro_id]   # Prendo il prestito del libro.

        # Se data_restituzione è None significa che il libro NON è stato ancora
        # restituito: quindi è fuori, bisogna capire se è in tempo o in ritardo.
        if p["data_restituzione"] is None:
            # Converto la stringa della scadenza in un oggetto data.
            scad = datetime.strptime(p["scadenza"], "%Y-%m-%d").date()
            # Se la scadenza è passata (minore di oggi) è in ritardo, altrimenti
            # è semplicemente in prestito. date.today() = data odierna.
            return "In ritardo" if scad < date.today() else "In prestito"

    # Se il libro non è nel dizionario (nessun prestito) o il prestito è stato
    # restituito, il libro è disponibile.
    return "Disponibile"


def prestiti_per_libro(conn):
    """Costruisce un dizionario: per ogni libro, il suo prestito più recente.

    - conn: la connessione al database già aperta.
    Serve per calcolare gli stati in modo efficiente: facciamo UNA sola query
    invece di una query per ogni libro. Il dizionario finale ha come chiave
    l'id del libro e come valore il record del prestito.
    """
    # Leggo tutti i prestiti, dal più recente al più vecchio (ORDER BY DESC).
    rows = conn.execute("SELECT * FROM prestiti ORDER BY data_prestito DESC").fetchall()

    m = {}   # Dizionario risultato: {id libro: prestito}.
    for r in rows:
        # setdefault aggiunge la riga SOLO se l'id non è ancora presente. Dato
        # che le righe arrivano dalla più recente, la prima che incontriamo per
        # ogni libro è proprio il prestito più recente (gli altri vengono ignorati).
        m.setdefault(r["libro_id"], r)
    return m   # Restituisco il dizionario.


# ---------------------------------------------------------------------------
# 5. FINESTRA PRINCIPALE
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    """La finestra principale dell'applicazione.

    Contiene tre schede (QTabWidget): "Libri", "Prestiti" e "Statistiche".
    Ogni scheda ha i propri widget e le proprie funzioni di caricamento dati.
    """
    def __init__(self):
        """Costruttore: prepara la finestra e crea le tre schede."""
        super().__init__()                       # Costruttore della classe madre.
        self.setWindowTitle("Rubrica Biblioteca")  # Titolo della finestra.
        self.resize(920, 620)                    # Dimensione iniziale (larghezza, altezza).

        central = QWidget()                      # Widget contenitore centrale.
        self.setCentralWidget(central)           # Lo imposto come centrale della finestra.
        layout = QVBoxLayout(central)            # Layout verticale sul contenitore.

        self.tabs = QTabWidget()                 # Contenitore a schede.
        layout.addWidget(self.tabs)              # Aggiungo le schede al layout.

        # Creo le tre schede (ogni metodo costruisce i widget della sua scheda).
        self.build_libri_tab()        # Scheda "Libri".
        self.build_prestiti_tab()     # Scheda "Prestiti".
        self.build_statistiche_tab()  # Scheda "Statistiche".

        # Messaggio nella barra di stato (in basso nella finestra).
        self.statusBar().showMessage("Pronto")

    # =====================================================================
    # SCHEDA LIBRI
    # =====================================================================

    def build_libri_tab(self):
        """Costruisce tutti i widget della scheda "Libri"."""
        tab = QWidget()                # Widget contenitore della scheda.
        lay = QVBoxLayout(tab)         # Layout verticale della scheda.

        # Riga 1: ricerca testo + filtri a tendina.
        search_row = QHBoxLayout()                     # Layout orizzontale.
        search_row.addWidget(QLabel("Cerca:"))         # Etichetta "Cerca:".
        self.search = QLineEdit()                      # Casella di testo per la ricerca.
        self.search.setPlaceholderText("Titolo, autore, genere o ISBN...")  # Testo guida.
        self.search.textChanged.connect(self.load_libri)  # A ogni tasto premuto ricarico la lista.
        search_row.addWidget(self.search)              # Aggiungo la casella.

        search_row.addWidget(QLabel("Genere:"))        # Etichetta "Genere:".
        self.filter_genre = QComboBox()                # Tendina dei generi.
        self.filter_genre.addItem("Tutti")             # Opzione predefinita "Tutti".
        self.filter_genre.currentTextChanged.connect(self.load_libri)  # Cambio filtro = ricarico.
        search_row.addWidget(self.filter_genre)        # Aggiungo la tendina.

        search_row.addWidget(QLabel("Stato:"))         # Etichetta "Stato:".
        self.filter_status = QComboBox()               # Tendina dello stato.
        self.filter_status.addItems(["Tutti", "Disponibile", "In prestito", "In ritardo"])
        self.filter_status.currentTextChanged.connect(self.load_libri)
        search_row.addWidget(self.filter_status)       # Aggiungo la tendina.
        lay.addLayout(search_row)                      # Aggiungo la riga al layout.

        # Riga 2: intervallo di anni + ordinamento.
        year_row = QHBoxLayout()                       # Layout orizzontale.
        year_row.addWidget(QLabel("Anno dal:"))        # Etichetta "Anno dal:".
        self.year_from = QSpinBox()                    # Campo numerico anno iniziale.
        self.year_from.setRange(0, 2100)               # Intervallo ammesso.
        self.year_from.setSpecialValueText("Tutti")    # 0 mostrato come "Tutti".
        self.year_from.valueChanged.connect(self.load_libri)  # Cambio = ricarico.
        year_row.addWidget(self.year_from)

        year_row.addWidget(QLabel("al:"))              # Etichetta "al:".
        self.year_to = QSpinBox()                      # Campo numerico anno finale.
        self.year_to.setRange(0, 2100)                 # Intervallo ammesso.
        self.year_to.setValue(2100)                    # Valore predefinito 2100 (nessun limite).
        self.year_to.valueChanged.connect(self.load_libri)
        year_row.addWidget(self.year_to)

        year_row.addWidget(QLabel("Ordinamento:"))     # Etichetta "Ordinamento:".
        self.sort_by = QComboBox()                     # Tendina ordinamento.
        self.sort_by.addItems(["Titolo", "Autore", "Anno", "Genere"])
        self.sort_by.currentTextChanged.connect(self.load_libri)
        year_row.addWidget(self.sort_by)
        year_row.addStretch()                          # Spazio elastico a destra.
        lay.addLayout(year_row)                        # Aggiungo la riga al layout.

        # La tabella principale dei libri: 7 colonne.
        self.table_libri = QTableWidget(0, 7)
        # Intestazioni delle colonne.
        self.table_libri.setHorizontalHeaderLabels(["Copertina", "Titolo", "Autore", "Anno", "Genere", "Stato", "Prestito a"])
        # Le colonne si allargano in modo uniforme per riempire lo spazio.
        self.table_libri.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        # La colonna 0 (copertina) si adatta al contenuto (immagine piccola).
        self.table_libri.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        # Selezione per righe intere (cliccando ovunque nella riga la seleziono tutta).
        self.table_libri.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        # Selezione singola (non posso selezionare più libri insieme).
        self.table_libri.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        # Le celle non sono modificabili direttamente (si modifica con la finestra).
        self.table_libri.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        # Altezza di ogni riga: 60 px, abbastanza per una copertina piccola.
        self.table_libri.verticalHeader().setDefaultSectionSize(60)
        # Dimensione delle icone delle copertine nella tabella.
        self.table_libri.setIconSize(QSize(40, 55))
        # Doppio clic su una riga = apri i dettagli del libro.
        self.table_libri.doubleClicked.connect(lambda: self.show_details())
        lay.addWidget(self.table_libri)    # Aggiungo la tabella al layout.

        # Riga dei pulsanti di azione.
        btns = QHBoxLayout()
        b_add = QPushButton("Aggiungi libro")     # Pulsante aggiungi.
        b_edit = QPushButton("Modifica")          # Pulsante modifica.
        b_del = QPushButton("Elimina")            # Pulsante elimina.
        b_dett = QPushButton("Dettagli")          # Pulsante dettagli.
        b_prestito = QPushButton("Registra prestito")  # Pulsante nuovo prestito.
        b_rest = QPushButton("Restituisci")       # Pulsante restituzione.
        b_csv = QPushButton("Esporta CSV")        # Pulsante esportazione.
        # Collego ogni pulsante alla sua funzione (lambda per passare zero argomenti).
        b_add.clicked.connect(lambda: self.add_libro())
        b_edit.clicked.connect(lambda: self.edit_libro())
        b_del.clicked.connect(lambda: self.delete_libro())
        b_dett.clicked.connect(lambda: self.show_details())
        b_prestito.clicked.connect(lambda: self.new_prestito())
        b_rest.clicked.connect(lambda: self.return_libro())
        b_csv.clicked.connect(lambda: self.export_csv())
        # Aggiungo tutti i pulsanti al layout orizzontale.
        for b in (b_add, b_edit, b_dett, b_del, b_prestito, b_rest, b_csv):
            btns.addWidget(b)
        btns.addStretch()       # Spazio elastico a destra dei pulsanti.
        lay.addLayout(btns)     # Aggiungo la riga pulsanti.

        self.tabs.addTab(tab, "Libri")   # Registro la scheda con nome "Libri".

    def current_libro(self):
        """Restituisce il record del libro selezionato nella tabella, o None.

        self._libri è la lista dei libri che stiamo mostrando (già filtrati e
        ordinati). La riga selezionata (currentRow) corrisponde a quella lista.
        """
        row = self.table_libri.currentRow()   # Numero della riga selezionata (-1 = nessuna).
        if row < 0:                           # Se non c'è selezione...
            return None                       # ...restituisco None.
        return self._libri[row]               # Altrimenti il libro alla riga scelta.

    def load_libri(self):
        """Carica i libri dal database, applica i filtri e li mostra in tabella.

        Viene chiamata a ogni modifica: all'avvio, quando si digita nella ricerca,
        quando si cambia un filtro, dopo aver aggiunto/modificato/eliminato un libro.
        """
        # Leggo i valori correnti dei filtri.
        q = self.search.text().strip()          # Testo cercato (spazi eliminati).
        filtro = self.filter_status.currentText()  # Filtro stato selezionato.
        genere = self.filter_genre.currentText()   # Filtro genere selezionato.

        # Intervallo anni: se il campo è 0 usiamo un valore "tutto". Per l'anno
        # iniziale -10000 è come dire "da sempre"; per quello finale 2100 "fino a sempre".
        y_from = self.year_from.value() if self.year_from.value() else -10000
        y_to = self.year_to.value() if self.year_to.value() else 2100

        ordine = self.sort_by.currentText()     # Criterio di ordinamento scelto.

        conn = get_conn()                       # Apro la connessione al database.
        p_map = prestiti_per_libro(conn)        # Prestiti recenti per calcolare gli stati.
        rows = conn.execute("SELECT * FROM libri").fetchall()  # Leggo tutti i libri.
        self._libri = []                        # Lista dei libri che passeranno i filtri.

        # Applico i filtri libro per libro.
        for libro in rows:
            st = status_of(libro["id"], p_map)   # Stato del libro.

            # Filtro per stato (se non è "Tutti").
            if filtro != "Tutti" and st != filtro:
                continue   # Saltiamo questo libro: non corrisponde.

            # Filtro per genere (se non è "Tutti").
            if genere != "Tutti" and (libro["genere"] or "") != genere:
                continue

            # Filtro per anno: l'anno deve stare dentro l'intervallo scelto.
            anno = libro["anno"] or 0   # Se manca l'anno usiamo 0.
            if anno and not (y_from <= anno <= y_to):
                continue

            # Filtro testo: cerco il testo in titolo, autore, genere e ISBN
            # (senza distinzione maiuscole/minuscole).
            if q and q.lower() not in " ".join(filter(None, [
                    libro["titolo"], libro["autore"], libro["genere"], libro["isbn"]
            ])).lower():
                continue

            self._libri.append(libro)   # Il libro supera tutti i filtri: lo tengo.

        # Ordinamento: mappa il testo scelto alla colonna del database corrispondente.
        order_key = {"Titolo": "titolo", "Autore": "autore", "Anno": "anno", "Genere": "genere"}[ordine]
        # Ordino con una chiave speciale: prima i libri con valore (False), poi quelli
        # senza valore (True). "or ''" fa sì che None e stringa siano confrontabili.
        self._libri.sort(key=lambda l: (l[order_key] is None, l[order_key] or ""))

        # Preparo la tabella: numero di righe = numero di libri filtrati.
        self.table_libri.setRowCount(len(self._libri))

        # Lista dei thread di download attivi (per non farli terminare da Python
        # prima che finiscano). getattr serve perché il primo avvio non ha l'attributo.
        self._cover_workers = getattr(self, "_cover_workers", [])

        # Riempio le righe della tabella con i dati dei libri.
        for r, libro in enumerate(self._libri):
            st = status_of(libro["id"], p_map)   # Stato del libro per la riga.
            persona = ""                         # Chi lo ha in prestito (vuoto se disponibile).
            if libro["id"] in p_map and p_map[libro["id"]]["data_restituzione"] is None:
                persona = p_map[libro["id"]]["persona"]   # Prendo il nome del lettore.

            # Cella 0: copertina. Inserisco l'icona se il file esiste già sul disco.
            item_cover = QTableWidgetItem()      # Cella vuota per la copertina.
            f = file_copertina(libro["id"])      # Percorso del file immagine.
            if f.exists():                       # Se il file esiste già...
                pix = QPixmap(str(f))            # ...carico l'immagine.
                if not pix.isNull():             # Se l'immagine è valida...
                    item_cover.setIcon(QIcon(pix.scaled(   # ...la metto come icona ridotta.
                        40, 55, Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation)))
            self.table_libri.setItem(r, 0, item_cover)   # Inserisco la cella copertina.

            # Celle da 1 a 6: dati di testo del libro.
            self.table_libri.setItem(r, 1, QTableWidgetItem(libro["titolo"]))  # Titolo.
            self.table_libri.setItem(r, 2, QTableWidgetItem(libro["autore"]))  # Autore.
            # Anno come testo; "—" se manca.
            self.table_libri.setItem(r, 3, QTableWidgetItem(str(libro["anno"]) if libro["anno"] else "—"))
            self.table_libri.setItem(r, 4, QTableWidgetItem(libro["genere"] or "—"))  # Genere.
            # Cella dello stato con il colore corrispondente.
            item_st = QTableWidgetItem(st)               # Testo dello stato.
            item_st.setForeground(STATUS_COLORS[st])     # Colore dello stato.
            self.table_libri.setItem(r, 5, item_st)      # Inserisco la cella stato.
            self.table_libri.setItem(r, 6, QTableWidgetItem(persona))  # Nome lettore.

            # Tengo solo i thread ancora in esecuzione (rimuovo quelli finiti
            # per non accumulare riferimenti inutili in memoria).
            self._cover_workers = [w for w in self._cover_workers if w.isRunning()]

            # Se la copertina non è ancora sul disco, la scarico in background.
            if not f.exists():
                url = trova_copertina_url(libro["titolo"], libro["autore"])  # Cerco l'URL.
                if url:                        # Se trovato...
                    w = CopertinaWorker(libro["id"], url, f)  # Creo il thread di download.
                    w.fatta.connect(self.on_copertina_pronta)  # Alla fine aggiorno la riga.
                    self._cover_workers.append(w)  # Lo tengo in memoria.
                    w.start()                  # Avvio il thread (scarica in background).

        conn.close()   # Chiudo la connessione al database.

    def on_copertina_pronta(self, libro_id):
        """Aggiorna la riga della tabella quando una copertina è stata scaricata.

        - libro_id: l'id del libro di cui è pronta la copertina.
        Viene chiamato dal segnale del thread CopertinaWorker.
        """
        # Cerco la riga della tabella corrispondente all'id del libro.
        for r in range(self.table_libri.rowCount()):
            if self._libri and r < len(self._libri) and self._libri[r]["id"] == libro_id:
                f = file_copertina(libro_id)   # Percorso del file appena scaricato.
                pix = QPixmap(str(f))          # Carico l'immagine.
                if not pix.isNull():           # Se valida...
                    # ...imposto l'icona nella cella copertina della riga trovata.
                    self.table_libri.item(r, 0).setIcon(QIcon(pix.scaled(
                        40, 55, Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation)))
                break   # Trovata la riga, esco dal ciclo.

    def aggiorna_generi(self):
        """Aggiorna la tendina dei generi con i generi presenti nel database.

        Da chiamare all'avvio e dopo ogni aggiunta/modifica/eliminazione di un libro,
        così la tendina riflette sempre i generi realmente esistenti.
        """
        conn = get_conn()   # Apro la connessione.
        # Leggo i generi distinti (DISTINCT), ordinati senza distinzione di maiuscole
        # (COLLATE NOCASE), escludendo vuoti.
        rows = conn.execute(
            "SELECT DISTINCT genere FROM libri WHERE genere IS NOT NULL AND genere != '' "
            "ORDER BY genere COLLATE NOCASE").fetchall()
        conn.close()   # Chiudo la connessione.

        current = self.filter_genre.currentText()   # Ricordo la selezione attuale.

        # Blocco i segnali della tendina: così durante il riempimento la tendina
        # non chiama load_libri decine di volte (solo alla fine con il valore scelto).
        self.filter_genre.blockSignals(True)
        self.filter_genre.clear()           # Svuoto la tendina.
        self.filter_genre.addItem("Tutti")  # Aggiungo l'opzione "Tutti".
        for r in rows:
            self.filter_genre.addItem(r["genere"])   # Aggiungo ogni genere trovato.

        # Se il genere che avevamo scelto esiste ancora, lo ripristino.
        if current in [self.filter_genre.itemText(i) for i in range(self.filter_genre.count())]:
            self.filter_genre.setCurrentText(current)

        self.filter_genre.blockSignals(False)   # Riattivo i segnali della tendina.

    def show_details(self):
        """Apre la finestra di dettaglio del libro selezionato."""
        libro = self.current_libro()   # Prendo il libro selezionato.
        if not libro:                  # Se nessun libro è selezionato...
            QMessageBox.information(self, "Nessuna selezione", "Seleziona un libro dalla lista.")
            return                     # ...avviso ed esco.

        conn = get_conn()              # Apro la connessione.
        p_map = prestiti_per_libro(conn)  # Calcolo i prestiti.
        conn.close()                   # Chiudo.
        st = status_of(libro["id"], p_map)  # Stato del libro.
        dlg = DettagliLibroDialog(self, libro, st)  # Creo la finestra dettagli.
        dlg.exec()                     # La apro in modo modale (blocca finché non si chiude).

    def export_csv(self):
        """Esporta i libri filtrati in un file CSV leggibile da Excel.

        Il separatore usato è ";" perché è quello che Excel italiano riconosce
        per impostazione predefinita. L'encoding utf-8-sig permette di vedere
        correttamente le lettere accentate anche in Excel.
        """
        if not self._libri:   # Se la lista filtrata è vuota...
            QMessageBox.information(self, "Nessun libro", "Non c'è nulla da esportare.")
            return

        # Apro la finestra di salvataggio; restituisce (percorso, filtro scelto).
        path, _ = QFileDialog.getSaveFileName(
            self, "Salva esportazione", "biblioteca.csv", "File CSV (*.csv)")
        if not path:   # Se l'utente ha annullato...
            return

        try:   # La scrittura su file può fallire (disco pieno, permessi...).
            # Apro il file in scrittura. newline="" evita righe vuote in più su Windows.
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f, delimiter=";")   # Scrittore CSV con separatore ";" .
                # Prima riga: le intestazioni delle colonne.
                w.writerow(["Titolo", "Autore", "Anno", "Genere", "ISBN", "Stato", "Prestito a"])
                # Poi una riga per ogni libro filtrato.
                for libro in self._libri:
                    conn = get_conn()              # Apro la connessione.
                    p_map = prestiti_per_libro(conn)  # Prestiti per lo stato.
                    conn.close()                   # Chiudo.
                    st = status_of(libro["id"], p_map)  # Stato del libro.
                    persona = ""                   # Lettore (vuoto se disponibile).
                    if libro["id"] in p_map and p_map[libro["id"]]["data_restituzione"] is None:
                        persona = p_map[libro["id"]]["persona"]
                    # Scrivo la riga con tutti i dati del libro.
                    w.writerow([
                        libro["titolo"], libro["autore"], libro["anno"] or "",
                        libro["genere"] or "", libro["isbn"] or "", st, persona
                    ])
            # Messaggio di conferma nella barra di stato (10 secondi = 6000 ms).
            self.statusBar().showMessage(f"Esportati {len(self._libri)} libri in {path}", 6000)
        except OSError as e:   # Se c'è un errore di scrittura...
            QMessageBox.critical(self, "Errore", f"Impossibile salvare il file:\n{e}")

    def add_libro(self):
        """Apre la finestra per aggiungere un nuovo libro e lo salva."""
        dlg = LibroDialog(self)   # Finestra in modalità "aggiungi" (nessun libro).
        if dlg.exec() == QDialog.DialogCode.Accepted:   # Se l'utente preme Salva...
            data = dlg.data()      # Raccolgo i dati inseriti.
            if not data["titolo"] or not data["autore"]:   # Controllo obbligatori.
                QMessageBox.warning(self, "Dati mancanti", "Titolo e autore sono obbligatori.")
                return
            conn = get_conn()      # Apro la connessione.
            # INSERT aggiunge una nuova riga nella tabella libri.
            conn.execute(
                "INSERT INTO libri (titolo, autore, anno, genere, isbn, note, data_aggiunta) "
                "VALUES (?,?,?,?,?,?,?)",
                # I "?" vengono sostituiti da questi valori in ordine. date.today()
                # .isoformat() produce la data odierna nel formato yyyy-mm-dd.
                (data["titolo"], data["autore"], data["anno"], data["genere"],
                 data["isbn"], data["note"], date.today().isoformat()))
            conn.commit()      # Salvo la modifica sul file.
            conn.close()       # Chiudo la connessione.
            self.load_libri()  # Ricarico la lista (mostro il nuovo libro).
            self.aggiorna_generi()   # Aggiorno la tendina dei generi.
            self.load_statistiche()  # Aggiorno i grafici.

    def edit_libro(self):
        """Apre la finestra per modificare il libro selezionato."""
        libro = self.current_libro()   # Libro selezionato.
        if not libro:                  # Nessuno selezionato.
            QMessageBox.information(self, "Nessuna selezione", "Seleziona un libro dalla lista.")
            return
        dlg = LibroDialog(self, libro)   # Finestra precompilata (modalità modifica).
        if dlg.exec() == QDialog.DialogCode.Accepted:
            data = dlg.data()
            if not data["titolo"] or not data["autore"]:   # Obbligatori.
                QMessageBox.warning(self, "Dati mancanti", "Titolo e autore sono obbligatori.")
                return
            conn = get_conn()
            # UPDATE modifica la riga del libro con i nuovi valori (WHERE id=?).
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
        """Elimina il libro selezionato dopo una conferma."""
        libro = self.current_libro()
        if not libro:
            QMessageBox.information(self, "Nessuna selezione", "Seleziona un libro dalla lista.")
            return
        # Chiedo conferma all'utente: la finestra mostra Si/No.
        ans = QMessageBox.question(
            self, "Conferma",
            f"Eliminare '{libro['titolo']}'?\nI prestiti collegati verranno rimossi.")
        if ans == QMessageBox.StandardButton.Yes:   # Se risponde Sì...
            conn = get_conn()
            # DELETE rimuove il libro. Grazie a ON DELETE CASCADE vengono
            # eliminati anche i suoi prestiti automaticamente.
            conn.execute("DELETE FROM libri WHERE id=?", (libro["id"],))
            conn.commit()
            conn.close()
            self.load_libri()
            self.load_prestiti()       # Aggiorno anche i prestiti.
            self.aggiorna_generi()
            self.load_statistiche()

    # =====================================================================
    # SCHEDA PRESTITI
    # =====================================================================

    def build_prestiti_tab(self):
        """Costruisce tutti i widget della scheda "Prestiti"."""
        tab = QWidget()
        lay = QVBoxLayout(tab)

        # Tabella dei prestiti: 7 colonne.
        self.table_prestiti = QTableWidget(0, 7)
        self.table_prestiti.setHorizontalHeaderLabels(
            ["Libro", "Persona", "Email", "Telefono", "Data", "Scadenza", "Stato"])
        self.table_prestiti.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_prestiti.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_prestiti.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table_prestiti.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table_prestiti.doubleClicked.connect(lambda: self.dettagli_prestito())  # Doppio clic = dettagli.
        lay.addWidget(self.table_prestiti)

        # Pulsanti della scheda prestiti.
        btns = QHBoxLayout()
        b_act = QPushButton("Mostra solo prestiti attivi")  # Filtro prestiti attivi.
        b_act.setCheckable(True)              # Pulsante che resta "premuto" (toggle).
        b_act.toggled.connect(self.load_prestiti)   # Allo switch ricarico la tabella.
        self.b_act = b_act                    # Lo salvo per leggerlo in load_prestiti.
        b_det = QPushButton("Dettagli prestito")
        b_det.clicked.connect(lambda: self.dettagli_prestito())
        b_ret = QPushButton("Restituisci selezionato")
        b_ret.clicked.connect(lambda: self.return_libro())
        btns.addWidget(b_act)
        btns.addWidget(b_det)
        btns.addWidget(b_ret)
        btns.addStretch()
        lay.addLayout(btns)
        self.tabs.addTab(tab, "Prestiti")   # Registro la scheda.

    def load_prestiti(self):
        """Carica i prestiti dal database e li mostra nella tabella."""
        conn = get_conn()   # Apro la connessione.
        if self.b_act.isChecked():   # Se il filtro "solo attivi" è premuto...
            # ...mostro solo i prestiti non ancora restituiti, ordinati per scadenza
            # (i più urgenti in cima). JOIN unisce la tabella prestiti con quella libri
            # per poter mostrare anche il titolo.
            rows = conn.execute("""
                SELECT p.*, l.titolo FROM prestiti p
                JOIN libri l ON l.id = p.libro_id
                WHERE p.data_restituzione IS NULL
                ORDER BY p.scadenza""").fetchall()
        else:
            # Altrimenti mostro tutti i prestiti, dal più recente.
            rows = conn.execute("""
                SELECT p.*, l.titolo FROM prestiti p
                JOIN libri l ON l.id = p.libro_id
                ORDER BY p.data_prestito DESC""").fetchall()

        self.table_prestiti.setRowCount(0)   # Svuoto la tabella.
        for p in rows:   # Per ogni prestito...
            # Converto la scadenza in oggetto data.
            scad = datetime.strptime(p["scadenza"], "%Y-%m-%d").date()
            attivo = p["data_restituzione"] is None   # Vero se non ancora restituito.
            # Stato del prestito: in ritardo / attivo / restituito.
            st = "In ritardo" if attivo and scad < date.today() else ("Attivo" if attivo else "Restituito")

            r = self.table_prestiti.rowCount()   # Numero di righe attuali (per la nuova).
            self.table_prestiti.insertRow(r)     # Aggiungo una riga vuota.

            # Riempio le celle della riga.
            self.table_prestiti.setItem(r, 0, QTableWidgetItem(p["titolo"]))    # Titolo libro.
            self.table_prestiti.setItem(r, 1, QTableWidgetItem(p["persona"]))   # Nome lettore.
            self.table_prestiti.setItem(r, 2, QTableWidgetItem(p["email"] or "—"))     # Email.
            self.table_prestiti.setItem(r, 3, QTableWidgetItem(p["telefono"] or "—"))  # Telefono.
            self.table_prestiti.setItem(r, 4, QTableWidgetItem(p["data_prestito"]))    # Data inizio.
            self.table_prestiti.setItem(r, 5, QTableWidgetItem(p["scadenza"]))         # Scadenza.
            # Cella dello stato con colore (STATUS_COLORS o grigio per "Restituito").
            item = QTableWidgetItem(st)
            item.setForeground(STATUS_COLORS.get(st, QColor("#555")))
            self.table_prestiti.setItem(r, 6, item)

            self._prestiti_rows = rows   # Salvo le righe grezze per i dettagli.

        conn.close()   # Chiudo la connessione.

    def dettagli_prestito(self):
        """Apre la finestra di dettaglio del prestito selezionato."""
        row = self.table_prestiti.currentRow()   # Riga selezionata.
        if row < 0:   # Nessuna selezione.
            QMessageBox.information(self, "Nessuna selezione", "Seleziona un prestito dalla lista.")
            return
        p = self._prestiti_rows[row]   # Record del prestito selezionato.
        dlg = PrestitoDialog(self, p["titolo"], prestito=p)  # Finestra precompilata (dettagli).
        dlg.setWindowTitle("Dettagli prestito")   # Titolo chiaro.
        dlg.exec()   # Apro in modo modale.

    def new_prestito(self):
        """Registra un nuovo prestito per il libro selezionato nella scheda Libri."""
        libro = self.current_libro()   # Libro selezionato (nella scheda libri).
        if not libro:
            QMessageBox.information(self, "Nessuna selezione", "Seleziona un libro dalla lista.")
            return

        conn = get_conn()              # Apro la connessione.
        p_map = prestiti_per_libro(conn)   # Prestiti correnti.
        conn.close()                   # Chiudo.
        # Se il libro ha già un prestito non restituito, non si può prestare di nuovo.
        if libro["id"] in p_map and p_map[libro["id"]]["data_restituzione"] is None:
            QMessageBox.warning(self, "Libro occupato", "Il libro è già in prestito.")
            return

        dlg = PrestitoDialog(self, libro["titolo"])   # Finestra nuovo prestito.
        if dlg.exec() == QDialog.DialogCode.Accepted:
            data = dlg.data()   # Raccolgo i dati (ora con email, telefono, note).
            if not data["persona"]:   # Il nome è obbligatorio.
                QMessageBox.warning(self, "Dati mancanti", "Indica chi prende in prestito il libro.")
                return
            conn = get_conn()
            # INSERT del nuovo prestito con tutti i campi, compresi i contatti.
            conn.execute(
                "INSERT INTO prestiti (libro_id, persona, email, telefono, note, "
                "data_prestito, scadenza) VALUES (?,?,?,?,?,?,?)",
                (libro["id"], data["persona"], data["email"], data["telefono"],
                 data["note"], data["data_prestito"], data["scadenza"]))
            conn.commit()
            conn.close()
            self.load_libri()        # Ricarico (lo stato del libro cambia).
            self.load_prestiti()     # Ricarico i prestiti.
            self.load_statistiche()  # Aggiorno i grafici.

    def return_libro(self):
        """Registra la restituzione di un libro.

        Funziona da entrambe le schede: dalla scheda Libri (con il libro
        selezionato) o dalla scheda Prestiti (con il prestito selezionato).
        """
        # Se siamo nella scheda Libri (indice 0)...
        if self.tabs.currentIndex() == 0:
            libro = self.current_libro()
            if not libro:
                QMessageBox.information(self, "Nessuna selezione", "Seleziona un libro dalla lista.")
                return
            conn = get_conn()
            # Cerco il prestito aperto (non restituito) più recente del libro.
            row = conn.execute(
                "SELECT * FROM prestiti WHERE libro_id=? AND data_restituzione IS NULL "
                "ORDER BY data_prestito DESC LIMIT 1", (libro["id"],)).fetchone()
            if row is None:   # Nessun prestito aperto.
                conn.close()
                QMessageBox.information(self, "Nessun prestito", "Questo libro non è in prestito.")
                return
        else:   # Siamo nella scheda Prestiti.
            row = self.table_prestiti.currentRow()
            if row < 0:
                QMessageBox.information(self, "Nessuna selezione", "Seleziona un prestito dalla lista.")
                return
            p = self._prestiti_rows[row]
            conn = get_conn()
            # Ricarico il prestito completo dal database (il record in memoria ha
            # solo alcune colonne, quello dal DB le ha tutte).
            row = conn.execute("SELECT * FROM prestiti WHERE id=?", (p["id"],)).fetchone()

        # Chiedo conferma all'utente.
        ans = QMessageBox.question(
            self, "Restituzione",
            f"Registrare la restituzione di '{row['titolo'] if 'titolo' in row.keys() else ''}' "
            f"da parte di {row['persona']}?")
        if ans == QMessageBox.StandardButton.Yes:   # Se conferma...
            # UPDATE imposta la data di restituzione a oggi: da quel momento lo
            # stato del libro torna "Disponibile".
            conn.execute("UPDATE prestiti SET data_restituzione=? WHERE id=?",
                         (date.today().isoformat(), row["id"]))
            conn.commit()
            conn.close()
            self.load_libri()        # Aggiorno stati libri.
            self.load_prestiti()     # Aggiorno tabella prestiti.
            self.load_statistiche()  # Aggiorno grafici.

    # =====================================================================
    # SCHEDA STATISTICHE
    # =====================================================================

    def build_statistiche_tab(self):
        """Costruisce la scheda "Statistiche" con numeri e grafici."""
        tab = QWidget()
        lay = QVBoxLayout(tab)

        # Riquadro "Panoramica" con i numeri principali.
        self.stats_grid = QGroupBox("Panoramica")
        grid = QGridLayout(self.stats_grid)   # Layout a griglia.
        self.stats_labels = {}                # Dizionario {chiave: etichetta valore}.

        # Elenco delle statistiche: chiave + etichetta descrittiva.
        stats_items = [
            ("totale", "Libri totali"),
            ("disponibili", "Disponibili"),
            ("prestito", "In prestito"),
            ("ritardo", "In ritardo"),
            ("persone", "Persone che hanno preso in prestito"),
            ("restituzioni", "Restituzioni registrate"),
        ]

        # Creo le etichette a coppie su due colonne (i % 2 dà riga alternata,
        # i // 2 dà la coppia di colonne). Così 6 valori stanno in 3 righe x 2 colonne.
        for i, (key, label) in enumerate(stats_items):
            grid.addWidget(QLabel(f"{label}:"), i % 2, (i // 2) * 2)   # Etichetta.
            val = QLabel("0")                      # Etichetta del valore (inizia a 0).
            val.setFont(QFont("", 12, QFont.Weight.Bold))   # In grassetto.
            grid.addWidget(val, i % 2, (i // 2) * 2 + 1)     # Casella del valore.
            self.stats_labels[key] = val           # Salvo il riferimento per aggiornarlo.

        lay.addWidget(self.stats_grid)

        # Etichetta con le prossime scadenze (testo, non grafico).
        self.stats_text = QLabel("")
        self.stats_text.setWordWrap(True)          # Va a capo se il testo è lungo.
        self.stats_text.setStyleSheet("padding: 8px;")   # Un po' di spazio interno.
        lay.addWidget(self.stats_text)

        # Prima riga di grafici: stato, genere, decenni.
        charts_row = QHBoxLayout()
        self.chart_stato = self._make_chart_view("Stato dei libri", self._make_stato_chart)
        charts_row.addWidget(self.chart_stato)
        self.chart_genere = self._make_chart_view("Distribuzione per genere", self._make_genere_chart)
        charts_row.addWidget(self.chart_genere)
        self.chart_decenni = self._make_chart_view("Libri per decennio", self._make_decenni_chart)
        charts_row.addWidget(self.chart_decenni)
        lay.addLayout(charts_row)

        # Seconda riga di grafici: prestiti per mese e autori più letti.
        charts_row2 = QHBoxLayout()
        self.chart_prestiti_mese = self._make_chart_view("Prestiti per mese", self._make_prestiti_mese_chart)
        charts_row2.addWidget(self.chart_prestiti_mese)
        self.chart_autori = self._make_chart_view("Autori più letti", self._make_autori_chart)
        charts_row2.addWidget(self.chart_autori)
        lay.addLayout(charts_row2)

        # Pulsante per aggiornare i grafici manualmente.
        refresh = QPushButton("Aggiorna")
        refresh.clicked.connect(self.load_statistiche)
        lay.addWidget(refresh)

        self.tabs.addTab(tab, "Statistiche")   # Registro la scheda.

    def _make_chart_view(self, title, builder):
        """Crea un contenitore per un grafico già configurato.

        - title: titolo mostrato sopra il grafico.
        - builder: la funzione che riempie il grafico con i dati (passata come riferimento).
        Restituisce un QChartView pronto da mettere nel layout.
        """
        chart = QChart()              # Oggetto grafico.
        chart.setTitle(title)         # Titolo del grafico.
        chart.setAnimationOptions(QChart.AnimationOption.SeriesAnimations)  # Animazioni morbide.
        chart.legend().setAlignment(Qt.AlignmentFlag.AlignBottom)   # Legenda sotto.
        chart.setBackgroundVisible(False)   # Sfondo trasparente (stile coerente).
        view = QChartView(chart)      # Widget che disegna il grafico.
        view.setRenderHint(view.renderHints())   # Antialias per bordi lisci.
        view.setMinimumHeight(220)    # Altezza minima.
        builder(chart)                # Chiamo la funzione che popola il grafico.
        return view                   # Restituisco il widget pronto.

    def _make_stato_chart(self, chart):
        """Grafico a torta dello stato dei libri (disponibile/in prestito/in ritardo)."""
        conn = get_conn()              # Apro la connessione.
        p_map = prestiti_per_libro(conn)   # Prestiti per gli stati.
        libri = conn.execute("SELECT id FROM libri").fetchall()   # Tutti gli id dei libri.
        counts = {"Disponibile": 0, "In prestito": 0, "In ritardo": 0}   # Contatori.
        for l in libri:   # Per ogni libro...
            counts[status_of(l["id"], p_map)] += 1   # ...incremento il suo stato.
        conn.close()   # Chiudo.

        series = QPieSeries()   # Serie a torta.
        colors = {"Disponibile": QColor("#2e7d32"),   # Verde.
                  "In prestito": QColor("#e65100"),   # Arancione.
                  "In ritardo": QColor("#c62828")}    # Rosso.
        for k, v in counts.items():   # Per ogni stato...
            if v > 0:   # ...se ha almeno un libro, aggiungo la fetta.
                sl = series.append(f"{k} ({v})", v)   # Fetta con etichetta e valore.
                sl.setColor(colors[k])                # Colore dedicato.
        chart.addSeries(series)   # Aggiungo la serie al grafico.

    def _make_genere_chart(self, chart):
        """Grafico a torta della distribuzione per genere (primi 8 generi)."""
        conn = get_conn()
        # Conto i libri per genere, escludendo i vuoti, ordinati dal più presente,
        # limitati ai primi 8.
        rows = conn.execute(
            "SELECT genere, COUNT(*) AS n FROM libri "
            "WHERE genere IS NOT NULL AND genere != '' "
            "GROUP BY genere ORDER BY n DESC LIMIT 8").fetchall()
        conn.close()

        series = QPieSeries()
        palette = [QColor("#5c6bc0"), QColor("#26a69a"), QColor("#ff7043"),
                   QColor("#ab47bc"), QColor("#ffa726"), QColor("#66bb6a"),
                   QColor("#ef5350"), QColor("#42a5f5")]   # Palette di 8 colori.
        for i, r in enumerate(rows):   # Per ogni genere...
            sl = series.append(f"{r['genere']} ({r['n']})", r["n"])   # Fetta con etichetta.
            sl.setColor(palette[i % len(palette)])   # Colore ciclico dalla palette.
        chart.addSeries(series)

    def _make_decenni_chart(self, chart):
        """Grafico a barre dei libri pubblicati per decennio."""
        conn = get_conn()
        # SQL che raggruppa per decennio: anno/10 tronca al decennio (es. 1987 -> 198),
        # *10 lo riporta a 1980. CAST ... AS TEXT lo converte in testo per l'asse.
        rows = conn.execute(
            "SELECT CAST(CAST(anno/10 AS INTEGER)*10 AS TEXT) AS decennio, COUNT(*) AS n "
            "FROM libri WHERE anno IS NOT NULL AND anno > 0 "
            "GROUP BY decennio ORDER BY decennio").fetchall()
        conn.close()

        bar = QBarSet("Libri")   # Gruppo di barre chiamato "Libri".
        cats = []                # Etichette dell'asse orizzontale (i decenni).
        for r in rows:
            cats.append(r["decennio"])   # Es. "1980".
            bar.append(r["n"])           # Numero di libri di quel decennio.
        series = QBarSeries()            # Serie a barre.
        series.append(bar)               # Aggiungo le barre alla serie.
        chart.addSeries(series)
        axis_x = QBarCategoryAxis()     # Asse orizzontale a categorie (testi).
        axis_x.append(cats)             # Aggiungo le etichette.
        chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)   # Asse in basso.
        series.attachAxis(axis_x)       # Collego la serie all'asse.
        axis_y = QValueAxis()           # Asse verticale numerico.
        axis_y.setLabelFormat("%d")     # Numeri interi.
        axis_y.setMin(0)                # Parte da zero.
        chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)   # Asse a sinistra.
        series.attachAxis(axis_y)       # Collego la serie all'asse.

    def _make_prestiti_mese_chart(self, chart):
        """Grafico a barre dei prestiti registrati per mese."""
        conn = get_conn()
        # substr(data_prestito, 1, 7) prende i primi 7 caratteri della data
        # (es. "2026-08-10" -> "2026-08"): raggruppo per mese.
        rows = conn.execute(
            "SELECT substr(data_prestito, 1, 7) AS mese, COUNT(*) AS n "
            "FROM prestiti GROUP BY mese ORDER BY mese").fetchall()
        conn.close()

        bar = QBarSet("Prestiti")   # Gruppo di barre "Prestiti".
        cats = []                   # Etichette mesi (es. "2026-08").
        for r in rows:
            cats.append(r["mese"])
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

    def _make_autori_chart(self, chart):
        """Grafico a torta degli autori più letti (con più prestiti)."""
        conn = get_conn()
        # Conto i prestiti per autore, prendendo i primi 8. JOIN unisce prestiti e libri.
        rows = conn.execute(
            "SELECT l.autore, COUNT(p.id) AS n FROM prestiti p "
            "JOIN libri l ON l.id = p.libro_id "
            "GROUP BY l.autore ORDER BY n DESC LIMIT 8").fetchall()
        conn.close()

        series = QPieSeries()
        palette = [QColor("#ec407a"), QColor("#7e57c2"), QColor("#29b6f6"),
                   QColor("#9ccc65"), QColor("#ffca28"), QColor("#ff7043"),
                   QColor("#8d6e63"), QColor("#26a69a")]   # Palette colori.
        for i, r in enumerate(rows):
            sl = series.append(f"{r['autore']} ({r['n']})", r["n"])
            sl.setColor(palette[i % len(palette)])
        chart.addSeries(series)

    def load_statistiche(self):
        """Aggiorna numeri, prossime scadenze e tutti i grafici.

        Viene chiamata all'avvio e dopo ogni operazione che cambia i dati.
        """
        conn = get_conn()              # Apro la connessione.
        p_map = prestiti_per_libro(conn)   # Prestiti per lo stato.
        libri = conn.execute("SELECT id FROM libri").fetchall()   # Tutti i libri.

        totale = len(libri)            # Numero totale di libri.
        disponibili = ritardo = prestito = 0   # Contatori inizializzati a zero.
        for l in libri:   # Per ogni libro conto il suo stato.
            st = status_of(l["id"], p_map)
            if st == "Disponibile":
                disponibili += 1       # Incremento i disponibili.
            elif st == "In ritardo":
                ritardo += 1           # Incremento i in ritardo.
            else:
                prestito += 1          # Incremento i in prestito.

        # Conto le persone distinte che hanno preso libri e le restituzioni.
        persone = conn.execute(
            "SELECT COUNT(DISTINCT persona) FROM prestiti").fetchone()[0]
        restituzioni = conn.execute(
            "SELECT COUNT(*) FROM prestiti WHERE data_restituzione IS NOT NULL").fetchone()[0]

        # Aggiorno le etichette numeriche della panoramica.
        self.stats_labels["totale"].setText(str(totale))
        self.stats_labels["disponibili"].setText(str(disponibili))
        self.stats_labels["prestito"].setText(str(prestito))
        self.stats_labels["ritardo"].setText(str(ritardo))
        self.stats_labels["persone"].setText(str(persone))
        self.stats_labels["restituzioni"].setText(str(restituzioni))

        # Cerco le prossime 5 scadenze di prestiti ancora attivi.
        prossime = conn.execute("""
            SELECT l.titolo, p.persona, p.scadenza FROM prestiti p
            JOIN libri l ON l.id = p.libro_id
            WHERE p.data_restituzione IS NULL
            ORDER BY p.scadenza LIMIT 5""").fetchall()

        # Preparo il testo da mostrare con le prossime scadenze.
        if prossime:
            oggi = date.today()
            txt = "Prossime scadenze:\n"
            for p in prossime:
                scad = datetime.strptime(p["scadenza"], "%Y-%m-%d").date()   # Data scadenza.
                diff = (scad - oggi).days   # Giorni mancanti (negativi = in ritardo).
                # Testo di aiuto: in ritardo, oggi, o fra N giorni.
                suff = " (in ritardo!)" if diff < 0 else ("" if diff == 0 else f" (tra {diff} giorni)")
                txt += f"  • {p['titolo']} → {p['persona']}, scade il {p['scadenza']}{suff}\n"
            self.stats_text.setText(txt)   # Mostro il testo.
        else:
            self.stats_text.setText("Nessun prestito attivo.")   # Nessuna scadenza.

        conn.close()   # Chiudo la connessione.

        # Ricreo tutti i grafici con i dati aggiornati.
        for name, builder in (("chart_stato", self._make_stato_chart),
                              ("chart_genere", self._make_genere_chart),
                              ("chart_decenni", self._make_decenni_chart),
                              ("chart_prestiti_mese", self._make_prestiti_mese_chart),
                              ("chart_autori", self._make_autori_chart)):
            view = getattr(self, name)   # Prendo il QChartView dal nome.
            old = view.chart()           # Il grafico attuale (per copiare il titolo).
            new = QChart()               # Creo un grafico nuovo.
            new.setTitle(old.title())    # Copio il titolo.
            # Riapplico le impostazioni estetiche.
            new.setAnimationOptions(QChart.AnimationOption.SeriesAnimations)
            new.legend().setAlignment(Qt.AlignmentFlag.AlignBottom)
            new.setBackgroundVisible(False)
            builder(new)                 # Popolo il nuovo grafico con i dati.
            view.setChart(new)           # Sostituisco il grafico nel widget.


# ---------------------------------------------------------------------------
# 6. AVVIO DEL PROGRAMMA
# ---------------------------------------------------------------------------

def main():
    """Punto di ingresso del programma: inizializza e avvia l'app."""
    init_db()                    # Creo/aggiorno il database (tabelle + migrazione).

    app = QApplication(sys.argv) # Creo l'applicazione Qt (serve SEMPRE prima di altri widget).
    win = MainWindow()           # Creo la finestra principale.
    win.aggiorna_generi()        # Popolo la tendina dei generi.
    win.load_libri()             # Carico i libri nella tabella.
    win.load_prestiti()          # Carico i prestiti nella tabella.
    win.load_statistiche()       # Carico numeri e grafici.
    win.show()                   # Mostro la finestra a schermo.

    # app.exec() avvia il ciclo degli eventi: il programma resta in esecuzione
    # finché l'utente non chiude la finestra. sys.exit passa il codice di uscita.
    sys.exit(app.exec())


# Se questo file viene eseguito direttamente (python biblioteca.py), esegue main().
# Se invece viene importato da un altro file (es. aggiungi_libri.py), non esegue
# l'interfaccia ma espone solo le funzioni e le classi.
if __name__ == "__main__":
    main()
