#!/usr/bin/env python3
"""Carica una serie di libri di esempio nel database della Rubrica Biblioteca.

Uso:
    python3 aggiungi_libri.py              # aggiunge tutti i libri
    python3 aggiungi_libri.py --skip       # salta i libri già presenti (stesso titolo+autore)
"""

import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import biblioteca as b

LIBRI = [
    # ---- Classici italiani ----
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

    # ---- Gialli e fantascienza ----
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

LIBRI = [tuple(l) for l in LIBRI]


def main():
    args = set(sys.argv[1:])
    skip_existing = "--skip" in args

    b.init_db()
    conn = b.get_conn()

    aggiunti = 0
    saltati = 0
    for titolo, autore, anno, genere, isbn, note in LIBRI:
        if skip_existing:
            esistente = conn.execute(
                "SELECT id FROM libri WHERE titolo=? AND autore=?",
                (titolo, autore)).fetchone()
            if esistente:
                saltati += 1
                continue
        conn.execute(
            "INSERT INTO libri (titolo, autore, anno, genere, isbn, note, data_aggiunta) "
            "VALUES (?,?,?,?,?,?,?)",
            (titolo, autore, anno or None, genere, isbn or None,
             note or None, date.today().isoformat()))
        aggiunti += 1

    conn.commit()
    totale = conn.execute("SELECT COUNT(*) FROM libri").fetchone()[0]
    conn.close()

    print(f"Aggiunti: {aggiunti}")
    if saltati:
        print(f"Saltati (già presenti): {saltati}")
    print(f"Libri totali nel database: {totale}")


if __name__ == "__main__":
    main()
