#!/usr/bin/env python3
"""Generate a self-contained Minitab exec for per-test_id capability analysis.

Reads a tall-table test-data export (one row per measurement, columns include
test_id, test_value, lower_limit, upper_limit), drops test_id steps whose
test_value has zero standard deviation (pass/fail style flags with no real
variation), and emits ONE self-contained Minitab exec (.mtb) file.

The exec has no external data dependency: each remaining test_id's measured
values are written INLINE via a SET ... END block into its own column, the
column is renamed to the test_id, and a Capability Analysis (Capa) is run on
it. Because the data is inlined, there is no separate CSV/data file to import
and no Subset step — you just run the exec (File > Run an Exec) and everything
(data + analyses) executes.
"""

# argparse: parancssori kapcsolók (pl. --decimal-separator) feldolgozásához.
import argparse
# csv: a bemeneti CSV fájl soronkénti, oszloponkénti beolvasásához.
import csv
# os: az elérési utak összefűzéséhez (a script saját könyvtárának megtalálásához).
import os
# statistics: a szórás (stdev) kiszámításához, hogy kiszűrhessük a 0 szórású lépéseket.
import statistics
# sys: kilépési kód (sys.exit) és a program futásának vezérléséhez.
import sys
# defaultdict: olyan dict, ami hiányzó kulcsra automatikusan üres listát ad vissza,
# így nem kell külön ellenőrizni, hogy egy test_id-hez már van-e lista létrehozva.
from collections import defaultdict

# tkinter: a Python beépített GUI könyvtára. Csak akkor kell, ha a felhasználó
# nem ad meg parancssori argumentumot, és emiatt felugró ablakra van szükség.
# Egyszer importáljuk itt a fájl tetején, nem minden függvényben külön-külön.
# try/except: néhány minimál Python telepítésből (pl. bizonyos céges/Linux
# csomagokból) hiányzik a tkinter. Ilyenkor NEM akarunk azonnal elhasalni:
# ha a felhasználó parancssori argumentumokkal futtatja a scriptet, GUI-ra
# nincs is szükség. Ezért ha az import elbukik, tk = None marad, és csak akkor
# adunk hibát, ha ténylegesen ablakot próbálnánk nyitni (lásd a prompt_* fv.-eket).
try:
    import tkinter as tk
    from tkinter import filedialog
except ImportError:
    tk = None
    filedialog = None

# Hány mérési értéket írjunk egy sorba a SET blokk adatrészében. Csak
# olvashatóság kérdése — a Minitab a SET után az END-ig minden számot beolvas,
# akárhány sorban is vannak.
VALUES_PER_LINE = 10

# A script saját könyvtára. A __file__ maga a jelenlegi .py fájl elérési útja;
# abspath -> teljes (abszolút) útvonal, dirname -> ebből a mappa. Így az alapértelmezett
# kimeneti fájl mindig a script MELLÉ kerül, függetlenül attól, honnan (melyik
# munkakönyvtárból) indítjuk el a scriptet.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUTPUT = os.path.join(SCRIPT_DIR, "minitab_capability_analysis.mtb")


def prompt_csv_path():
    """Fájlválasztó ablakot nyit, és visszaadja a kiválasztott CSV fájl elérési útját.

    Ha a felhasználó nem választ fájlt (bezárja az ablakot), a program leáll.
    """
    if tk is None:
        sys.exit(
            "A tkinter nem érhető el ebben a Python telepítésben, ezért nem tudok "
            "fájlválasztó ablakot nyitni. Add meg a CSV elérési útját parancssori "
            "argumentumként, pl.: python generate_minitab_capa_script.py adat.csv"
        )
    root = tk.Tk()
    # A tk.Tk() létrehozza a fő ablakot, de nekünk csak a fájlválasztó dialógus
    # kell, magát a fő ablakot nem akarjuk megjeleníteni -> withdraw() elrejti.
    root.withdraw()
    path = filedialog.askopenfilename(
        title="Válaszd ki a bemeneti CSV fájlt",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
    )
    root.destroy()
    if not path:
        print("Nem lett CSV fájl kiválasztva, kilépés.")
        sys.exit(1)
    return path


def prompt_decimal_separator():
    """Egy kis ablakot nyit két gombbal (vessző / pont), és visszaadja a választást.

    Erre azért van szükség, mert Minitabban a tizedes elválasztó a Windows
    regionális beállításától függ (pl. magyar rendszeren vessző: 6,5), és ha
    a script a rossz karaktert írja a számokba (SET adat, Lspec/Uspec), a
    Minitab "Invalid name or invalid syntax" hibát dob, vagy szövegként olvassa
    be a számokat.
    """
    # Egy kis dict-ben tároljuk a kiválasztott értéket, mert a gombok
    # "command" callback-jei (lambda-k) csak úgy tudnak "kifelé" írni egy
    # értéket, ha az egy már létező, megosztott objektum (itt: choice dict).
    if tk is None:
        sys.exit(
            "A tkinter nem érhető el ebben a Python telepítésben, ezért nem tudok "
            "választó ablakot nyitni. Add meg a tizedes elválasztót parancssori "
            "argumentumként, pl.: --decimal-separator ,"
        )
    choice = {"value": ","}

    def pick(sep):
        choice["value"] = sep
        root.destroy()  # a gombnyomás után bezárjuk az ablakot, a mainloop() ki tud lépni

    root = tk.Tk()
    root.title("Tizedes elválasztó")
    tk.Label(
        root, text="Milyen tizedes elválasztót használjon a Minitab a számoknál?",
        padx=20, pady=10,
    ).pack()
    frame = tk.Frame(root, padx=20, pady=10)
    frame.pack()
    tk.Button(frame, text="Vessző (6,5)", width=15, command=lambda: pick(",")).pack(side="left", padx=5)
    tk.Button(frame, text="Pont (6.5)", width=15, command=lambda: pick(".")).pack(side="left", padx=5)
    # Ha a felhasználó az ablak bezáró (X) gombjával zárja be, ne fagyjon le a
    # program: essünk vissza az alapértelmezett (vessző) értékre.
    root.protocol("WM_DELETE_WINDOW", lambda: pick(choice["value"]))
    root.mainloop()  # itt vár a program, amíg a felhasználó nem választ egy gombot
    return choice["value"]


def load_test_steps(csv_path):
    """Beolvassa a tall-table CSV-t, és test_id szerint csoportosítja az adatokat.

    Visszaadott értékek:
      order  - a test_id-k listája, első előfordulásuk sorrendjében
               (ezzel a kimeneti fájlban is megmarad az eredeti sorrend)
      values - dict: test_id -> az adott lépéshez tartozó test_value-k listája,
               NYERS string formában (ebből számoljuk a szórást float()-tal,
               és ezt írjuk a SET blokkokba is)
      limits - dict: test_id -> (lower_limit, upper_limit) pár (LSL/USL)
    """
    values = defaultdict(list)
    limits = {}
    order = []

    # encoding="utf-8-sig": a CSV elején lévő BOM (byte order mark) karaktert
    # levágja, hogy az első oszlopnév ("timestamp") ne kapjon szemetet elé.
    # newline="": a csv modul kéri, hogy ne a Python, hanem ő kezelje a
    # sortöréseket (fontos, mert pl. a "description" mezőben többsoros szöveg is van).
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        # DictReader: minden sort dict-ként ad vissza, a fejléc oszlopnevei
        # lesznek a kulcsok (pl. row["test_id"], row["test_value"]).
        reader = csv.DictReader(f)
        for row in reader:
            test_id = row["test_id"]
            if test_id not in limits:
                # Csak az adott test_id ELSŐ sorában jegyezzük meg a sorrendet
                # és a limiteket (minden sorban ugyanaz a limit, nem kell újra tárolni).
                order.append(test_id)
                limits[test_id] = (row["lower_limit"], row["upper_limit"])
            # A NYERS (string) test_value-t tároljuk, nem a float() eredményét.
            # Így a SET blokkba pontosan az eredeti szám kerül vissza, nem torzul
            # el a float -> str oda-vissza konverzión (pl. hosszú tizedestörteknél).
            # A szóráshoz úgyis külön float()-oljuk majd.
            values[test_id].append(row["test_value"])

    return order, values, limits


def format_number(raw, decimal_separator):
    """Egy CSV-ből kiolvasott szám-szöveget ("6.5", "50.0", "15.529"...) Minitab-
    kompatibilis alakra hoz.

    - Ha a szám egész (pl. 50.0), simán "50"-et adunk vissza, tizedesjel nélkül.
    - Ha nem egész, a "." tizedespontot lecseréljük a kért elválasztóra
      (pl. "," -> "6,5"), mert a magyar Minitab/Windows a vesszőt várja.

    Ugyanez a formázás kell a limitekhez (Lspec/Uspec) ÉS a SET-be írt mérési
    értékekhez is, ezért egy közös függvény.
    """
    value = float(raw)
    if value == int(value):
        return str(int(value))
    return repr(value).replace(".", decimal_separator)


def build_capa_block(test_id, raw_values, lsl, usl, column, decimal_separator):
    """Összeállítja egy adott test_id-hez tartozó teljes, önálló Minitab blokkot.

    A blokk lépései:
      1. SET C<n> / <adatok> / END  -> az adott test_id mérési értékeit közvetlenül
                                        beírjuk egy üres oszlopba (nincs külön data
                                        fájl, nincs import). A számokat szóközzel
                                        választjuk el, a tizedesjel a kiválasztott
                                        elválasztó.
      2. NAME C<n> "<test_id>"      -> az oszlopot elnevezzük a test_id-re, így a
                                        Capability Analysis címe a valódi tesztlépés
                                        neve lesz (nem "C1" vagy "test_value").
      3. Capa '<test_id>' 1; ...    -> lefuttatjuk a capability analysist a
                                        megadott LSL/USL limitekkel.

    Minden test_id a SAJÁT oszlopába kerül (C1, C2, ...), így nem kell Subset,
    nem kell worksheet-váltogatás, és nincs esély a táblák "összekeveredésére".
    """
    # A mérési értékeket a helyes tizedesjellel, szóközzel elválasztva soronként
    # VALUES_PER_LINE darabonként tördeljük (csak olvashatóság miatt).
    formatted = [format_number(v, decimal_separator) for v in raw_values]
    data_lines = []
    for i in range(0, len(formatted), VALUES_PER_LINE):
        chunk = formatted[i:i + VALUES_PER_LINE]
        data_lines.append("  " + " ".join(chunk))
    data_block = "\n".join(data_lines)

    return (
        f"SET {column}\n"
        f"{data_block}\n"
        "END\n"
        f'NAME {column} "{test_id}".\n'
        f"Capa '{test_id}' 1;\n"
        f"  Lspec {lsl};\n"
        f"  Uspec {usl};\n"
        "  Pooled;\n"
        "  AMR;\n"
        "  UnBiased;\n"
        "  OBiased;\n"
        "  Toler 6;\n"
        "  Within;\n"
        "  Overall;\n"
        "  NoCI;\n"
        "  PPM;\n"
        "  CStat.\n"
    )


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "csv_path", nargs="?",
        help="Path to the tall-table test-data CSV (omit to pick it via a file dialog)",
    )
    parser.add_argument(
        "-o", "--output", default=DEFAULT_OUTPUT,
        help="Output Minitab exec (.mtb) path (default: a script mellé, %(default)s)",
    )
    parser.add_argument(
        "--decimal-separator", choices=[",", "."],
        help="Decimal separator for the numbers (omit to pick it interactively)",
    )
    args = parser.parse_args()

    # Ha a felhasználó nem adta meg parancssorból, kérdezzük meg felugró ablakban.
    csv_path = args.csv_path or prompt_csv_path()
    decimal_separator = args.decimal_separator or prompt_decimal_separator()

    order, values, limits = load_test_steps(csv_path)

    # Statisztika a végső összefoglalóhoz: hány test_id-t hagytunk ki és miért.
    skipped_zero_std = []
    skipped_no_variance = []
    blocks = []
    column_index = 0   # melyik oszlopba (C1, C2, ...) írjuk a következő test_id-t

    for test_id in order:
        # A nyers string értékeket float()-oljuk, hogy szórást tudjunk számolni.
        test_values = [float(v) for v in values[test_id]]
        if len(test_values) < 2:
            # Szórást csak legalább 2 adatpontból lehet számolni.
            skipped_no_variance.append(test_id)
            continue
        std = statistics.stdev(test_values)
        if std == 0:
            # Ez a lényegi szűrés: ha minden mérési érték ugyanaz (pl. egy
            # pass/fail flag, ami mindig 1.0), nincs értelme capability
            # analysisnek, ezért kihagyjuk ezt a test_id-t.
            skipped_zero_std.append(test_id)
            continue

        column_index += 1
        column = f"C{column_index}"
        lower_raw, upper_raw = limits[test_id]
        lsl = format_number(lower_raw, decimal_separator)
        usl = format_number(upper_raw, decimal_separator)
        blocks.append(
            build_capa_block(test_id, values[test_id], lsl, usl, column, decimal_separator)
        )

    # Az önálló exec: minden blokkot üres sorral elválasztva egyetlen fájlba.
    # Minitabban File > Run an Exec -> betölti az inline adatot ÉS lefuttat mindent.
    with open(args.output, "w", encoding="utf-8") as f:
        f.write("\n".join(blocks))

    total = len(order)
    print(f"Total test_id steps found:       {total}")
    print(f"Skipped (zero std dev):          {len(skipped_zero_std)}")
    print(f"Skipped (fewer than 2 points):   {len(skipped_no_variance)}")
    print(f"Capability blocks written:       {len(blocks)}")
    print(f"Exec written to:                 {args.output}")


if __name__ == "__main__":
    sys.exit(main())
