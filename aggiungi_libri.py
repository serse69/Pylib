#!/usr/bin/env python3
# ============================================================================
# AGGIUNGI LIBRI - script per caricare una serie di libri nel database
# ----------------------------------------------------------------------------
# Questo script NON apre l'interfaccia grafica: è un semplice strumento da
# terminale che riempie il database biblioteca.db con una lista di libri.
#
# Come si usa:
#   python3 aggiungi_libri.py            -> aggiunge tutti i libri della lista
#   python3 aggiungi_libri.py --skip     -> aggiunge solo i libri che non sono
#                                          già presenti (evita i duplicati)
#
# Funziona perché importa le funzioni dal file biblioteca.py (senza eseguirne
# l'interfaccia, grazie al blocco if __name__ == "__main__" presente lì).
# ============================================================================

import sys                       # Modulo standard: per leggere gli argomenti della riga di comando.
from datetime import date        # Modulo standard: per la data odierna (data di inserimento).
from pathlib import Path         # Modulo standard: per gestire i percorsi dei file.

# Aggiungo la cartella che contiene questo script al percorso di ricerca dei
# moduli Python. Serve per poter scrivere "import biblioteca" e farlo trovare:
# se lanciamo lo script da un'altra cartella, Python deve sapere dove cercare.
# __file__ è il percorso di questo script, quindi con .parent prendiamo la
# cartella "Pylib" dove sta anche biblioteca.py.
sys.path.insert(0, str(Path(__file__).resolve().parent))

import biblioteca as b   # Importo le funzioni utili di biblioteca.py (init_db, get_conn).

# LIBRI: la lista dei libri da inserire. È una lista di tuple, e ogni tupla ha
# questa struttura:
#   (titolo, autore, anno, genere, isbn, note)
# - anno: anno di pubblicazione (numero). Per i testi antichi (es. Omero) viene
#   usato un anno approssimativo e negativo per indicare "a.C.".
# - isbn: codice ISBN (stringa). Alcuni libri non lo hanno: usiamo una stringa
#   vuota "" che poi diventerà NULL nel database.
# - note: una breve descrizione del libro.
LIBRI = [
    # ---- Classici italiani e mondiali ----
    # (titolo, autore, anno, genere, isbn, note)
    ("Il nome della rosa", "Umberto Eco", 1980, "Romanzo storico", "9788804159450", "Il monastero del 1327, un delitto e il francescano Guglielmo da Baskerville."),
    ("1984", "George Orwell", 1949, "Distopia", "9788804668237", "Il Grande Fratello e la sorveglianza totale."),
    ("Il Piccolo Principe", "Antoine de Saint-Exupéry", 1943, "Fiaba", "9788852075498", "Un aviatore incontra un piccolo principe."),
    ("Cent'anni di solitudine", "Gabriel García Márquez", 1967, "Realismo magico", "9788804677482", "La saga della famiglia Buendía a Macondo."),
    ("Orgoglio e pregiudizio", "Jane Austen", 1813, "Romanzo", "9788804720987", "Elizabeth Bennet e Mr. Darcy nell'Inghilterra georgiana."),
    ("Delitto e castigo", "Fëdor Dostoevskij", 1866, "Romanzo", "9788804687849", "Raskol'nikov e la sua teoria del superuomo."),
    ("I promessi sposi", "Alessandro Manzoni", 1827, "Romanzo storico", "9788804662372", "Renzo e Lucia tra carestie, briganti e peste."),
    ("La divina commedia", "Dante Alighieri", 1321, "Poema epico", "9788804711435", "Inferno, Purgatorio e Paradiso."),
    ("Il vecchio e il mare", "Ernest Hemingway", 1952, "Romanzo", "9788804770326", "Il vecchio Santiago e il suo duello con un marlin."),
    ("Guerra e pace", "Lev Tolstoj", 1869, "Romanzo storico", "9788804736902", "Le famiglie aristocratiche russe durante le guerre napoleoniche."),
    ("Il ritratto di Dorian Gray", "Oscar Wilde", 1890, "Romanzo", "9788804665311", "Un patto fa invecchiare il ritratto al posto di Dorian."),
    ("Moby Dick", "Herman Melville", 1851, "Avventura", "9788804703805", "Il capitano Achab alla caccia della balena bianca."),
    ("Anna Karenina", "Lev Tolstoj", 1877, "Romanzo", "9788804693048", "La passione travolgente di Anna Karenina."),
    ("L'odissea", "Omero", -800, "Poema epico", "9788804661856", "Il ritorno di Ulisse a Itaca dopo la guerra di Troia."),
    ("Il giro del mondo in 80 giorni", "Jules Verne", 1873, "Avventura", "9788817058601", "Phileas Fogg e la sua scommessa attorno al globo."),

    # ---- Gialli, fantascienza e fantasy ----
    ("Assassinio sull'Orient Express", "Agatha Christie", 1934, "Giallo", "9788804705687", "Hercule Poirot investiga su un treno bloccato dalla neve."),
    ("Dune", "Frank Herbert", 1965, "Fantascienza", "9788804757258", "Il deserto di Arrakis, le spezie e la casa Atreides."),
    ("Fahrenheit 451", "Ray Bradbury", 1953, "Fantascienza", "9788804661771", "I pompieri bruciano i libri in una società senza lettura."),
    ("Il codice da Vinci", "Dan Brown", 2003, "Thriller", "9788804660773", "Robert Langdon a caccia di un segreto millenario."),
    ("Lo Hobbit", "J. R. R. Tolkien", 1937, "Fantasy", "9788804682295", "Bilbo Baggins parte verso la Montagna Solitaria."),
    ("Harry Potter e la pietra filosofale", "J. K. Rowling", 1997, "Fantasy", "9788804570158", "L'inizio delle avventure del giovane mago."),
    ("Il Signore degli Anelli - La Compagnia dell'Anello", "J. R. R. Tolkien", 1954, "Fantasy", "9788804737657", "Frodo lascia la Contea con la Compagnia."),

    # ---- Saggi, scienza e manuali ----
    ("Breve storia del tempo", "Stephen Hawking", 1988, "Scienza", "9788804539954", "Dal Big Bang ai buchi neri."),
    ("L'arte della guerra", "Sun Tzu", -500, "Saggio", "9788804620967", "Il trattato strategico più antico del mondo."),
    ("Il gene egoista", "Richard Dawkins", 1976, "Scienza", "9788804672608", "L'evoluzione spiegata dal punto di vista dei geni."),
    ("Pensieri", "Blaise Pascal", 1670, "Filosofia", "9788804710056", "Frammenti e aforismi del filosofo e matematico francese."),
    ("Cosmos", "Carl Sagan", 1980, "Scienza", "9788804674814", "Il viaggio tra le stelle e la storia dell'universo."),
    ("Manuale del giardinaggio", "AA.VV.", 2020, "Manuale", "", "Guida pratica per l'orto e il giardino di casa."),
    ("Il grande libro della cucina italiana", "AA.VV.", 2018, "Manuale", "", "Ricette tradizionali regionali."),
]

# Riscrivo LIBRI con una comprensione di lista che converte ogni tupla in lista
# [titolo, autore, anno, genere, isbn, note]. Non cambia il contenuto: serve
# solo per lavorare in modo più comodo (le liste sono modificabili, le tuple no).
LIBRI = [tuple(l) for l in LIBRI]


def main():
    """Funzione principale: inserisce i libri nel database e mostra un riepilogo."""
    # Leggo gli argomenti passati da riga di comando (es. "--skip").
    args = set(sys.argv[1:])

    # Se l'utente ha scritto "--skip", attiviamo la modalità anti-duplicati:
    # i libri già presenti nel database vengono saltati, non reinseriti.
    skip_existing = "--skip" in args

    # Inizializzo il database: crea le tabelle se non esistono (e fa la
    # migrazione se serve). Importato da biblioteca.py.
    b.init_db()

    # Apro la connessione al database. Tutte le INSERT verranno fatte su questa.
    conn = b.get_conn()

    # Contatori per il riepilogo finale.
    aggiunti = 0     # Quanti libri abbiamo inserito.
    saltati = 0      # Quanti libri abbiamo saltato perché già presenti.

    # Ciclo su ogni libro della lista.
    for titolo, autore, anno, genere, isbn, note in LIBRI:
        # In modalità --skip controllo se esiste già un libro con lo stesso
        # titolo E lo stesso autore (la coppia identifica il libro).
        if skip_existing:
            esistente = conn.execute(
                "SELECT id FROM libri WHERE titolo=? AND autore=?",
                (titolo, autore)).fetchone()
            if esistente:   # Se esiste già...
                saltati += 1   # ...incremento il contatore dei saltati...
                continue       # ...e passo al libro successivo senza inserirlo.

        # INSERT: aggiungo il libro al database. I valori "?" vengono sostituiti
        # con i dati del libro. "anno or None" e "isbn or None" trasformano i
        # valori vuoti/zero in NULL (significa "dato mancante" in SQLite).
        # date.today().isoformat() mette la data di oggi come data di inserimento.
        conn.execute(
            "INSERT INTO libri (titolo, autore, anno, genere, isbn, note, data_aggiunta) "
            "VALUES (?,?,?,?,?,?,?)",
            (titolo, autore, anno or None, genere, isbn or None,
             note or None, date.today().isoformat()))
        aggiunti += 1   # Incremento il contatore dei libri inseriti.

    # Commit: rende permanenti TUTTE le modifiche sul file del database.
    # Senza commit, chiudendo la connessione le insert andrebbero perse.
    conn.commit()

    # Conto quanti libri ci sono ora in totale nel database.
    totale = conn.execute("SELECT COUNT(*) FROM libri").fetchone()[0]

    # Chiudo la connessione (importante per liberare il file).
    conn.close()

    # Mostro il riepilogo a video (standard output del terminale).
    print(f"Aggiunti: {aggiunti}")
    if saltati:   # Se ci sono libri saltati, lo dico.
        print(f"Saltati (già presenti): {saltati}")
    print(f"Libri totali nel database: {totale}")


# Punto di ingresso: se eseguo questo file direttamente, parte main().
# Se venisse importato da un altro file, non eseguirebbe nulla da solo.
if __name__ == "__main__":
    main()
