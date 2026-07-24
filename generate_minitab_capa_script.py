#!/usr/bin/env python3
"""Generate a Minitab session-command script for per-test_id capability analysis.

Reads a tall-table test-data export (one row per measurement, columns include
test_id, test_value, lower_limit, upper_limit), drops test_id steps whose
test_value has zero standard deviation (pass/fail style flags with no real
variation), and emits one Subset + Capability Analysis block per remaining
test_id into a single .txt file that can be pasted into / run from Minitab's
Session window (Edit > Command Line Editor).
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

# A Minitab worksheet neve, amire minden Subset parancs előtt vissza kell váltani,
# különben a következő Subset már csak az előző (leszűkített) táblát látná,
# nem a teljes eredeti adatsort.
SOURCE_WORKSHEET = "Worksheet 1"

# A script saját könyvtára. A __file__ maga a jelenlegi .py fájl elérési útja;
# abspath -> teljes (abszolút) útvonal, dirname -> ebből a mappa. Így az alapértelmezett
# kimeneti fájl mindig a script MELLÉ kerül, függetlenül attól, honnan (melyik
# munkakönyvtárból) indítjuk el a scriptet.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUTPUT = os.path.join(SCRIPT_DIR, "minitab_capability_commands.txt")
# A tiszta, Minitab-barát adatfájl alapértelmezett helye (szintén a script mellé).
# Ebben csak a test_id és test_value oszlop van, TAB-bal tagolva, sortörés-mentesen,
# hogy a Minitab importja ne akadjon meg a nyers CSV többsoros mezőin.
DEFAULT_DATA = os.path.join(SCRIPT_DIR, "minitab_data.txt")


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
    a script a rossz karaktert írja a Lspec/Uspec értékekbe, a Minitab
    "Invalid name or invalid syntax" hibát dob.
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
        root, text="Milyen tizedes elválasztót használjon a Minitab a Lspec/Uspec értékeknél?",
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
               és ezt írjuk a tiszta adatfájlba is)
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
            # Így a tiszta adatfájlba pontosan az eredeti szám kerül vissza,
            # nem torzul el a float -> str oda-vissza konverzión (pl. hosszú
            # tizedestörteknél). A szóráshoz úgyis külön float()-oljuk majd.
            values[test_id].append(row["test_value"])

    return order, values, limits


def format_limit(raw, decimal_separator):
    """Egy CSV-ből kiolvasott limit-szöveget ("6.5", "50.0", ...) Minitab-kompatibilis
    számmá alakít.

    - Ha a szám egész (pl. 50.0), simán "50"-et adunk vissza, tizedesjel nélkül.
    - Ha nem egész (pl. 6.5), a Python alapértelmezett "." tizedespontját
      lecseréljük a kért elválasztóra (pl. "," -> "6,5"), mert a magyar
      Minitab/Windows beállítás a vesszőt várja.
    """
    value = float(raw)
    if value == int(value):
        return str(int(value))
    return repr(value).replace(".", decimal_separator)


def write_minitab_data(path, analyzed_ids, values, decimal_separator):
    """Kiír egy tiszta, TAB-tagolt adatfájlt Minitab-importáláshoz.

    Csak két oszlop kerül bele: test_id és test_value — ez a kettő az egyetlen,
    amit a subset/capa parancsok használnak. A nyers CSV problémás mezőit
    (description, testrun_info: ezekben van sortörés) meg sem érintjük, így a
    Minitab importja nem darabolódik szét.

    - analyzed_ids: azoknak a test_id-knek a listája, amikre tényleg futtatunk
      elemzést (a 0 szórásúakat kihagyjuk, felesleges lenne az adatuk).
    - A test_value tizedespontját is a kért elválasztóra cseréljük, hogy a
      Minitab a számokat helyesen (számként, ne szövegként) olvassa be.
    """
    with open(path, "w", encoding="utf-8", newline="") as f:
        # Fejléc: a Minitab az oszlopokat a fejlécsor alapján nevezi el, ezért
        # itt PONTOSAN a 'test_id' és 'test_value' nevek kellenek, amikre a
        # parancsok később hivatkoznak.
        f.write("test_id\ttest_value\n")
        for test_id in analyzed_ids:
            for raw_value in values[test_id]:
                value = raw_value.replace(".", decimal_separator)
                f.write(f"{test_id}\t{value}\n")


def build_commands(test_id, lsl, usl):
    """Összeállítja egy adott test_id-hez tartozó teljes Minitab parancsblokkot.

    A blokk lépései:
      1. Worksheet "Worksheet 1".   -> visszaváltunk az eredeti, teljes adatsorra
      2. Subset; Include; GE/LE...  -> kiszűrjük azokat a sorokat, ahol
                                        test_id pontosan egyenlő a jelenlegi lépéssel
                                        (GE + LE együtt = "egyenlő", mert az "EQ"
                                        subcommand ezen a Minitab verzión hibát dobott)
      3. Name 'test_value' "..."    -> a mérési oszlopot átnevezzük az aktuális
                                        test_id-re, így a Capability Analysis
                                        címe nem "test_value", hanem a valódi
                                        tesztlépés neve lesz
      4. Capa '...' 1; ...          -> lefuttatjuk a capability analysist a
                                        megadott LSL/USL limitekkel
    """
    return (
        f'Worksheet "{SOURCE_WORKSHEET}".\n'
        "Subset;\n"
        "  Include;\n"
        f'  GE \'test_id\' "{test_id}";\n'
        f'  LE \'test_id\' "{test_id}";\n'
        f'  Name "{test_id}".\n'
        f'Name \'test_value\' "{test_id}".\n'
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
        help="Command .txt path (default: a script mellé, %(default)s)",
    )
    parser.add_argument(
        "-d", "--data-output", default=DEFAULT_DATA,
        help="Clean Minitab data .txt path (default: a script mellé, %(default)s)",
    )
    parser.add_argument(
        "--decimal-separator", choices=[",", "."],
        help="Decimal separator for Lspec/Uspec (omit to pick it interactively)",
    )
    args = parser.parse_args()

    # Ha a felhasználó nem adta meg parancssorból, kérdezzük meg felugró ablakban.
    csv_path = args.csv_path or prompt_csv_path()
    decimal_separator = args.decimal_separator or prompt_decimal_separator()

    order, values, limits = load_test_steps(csv_path)

    # Statisztika a végső összefoglalóhoz: hány test_id-t hagytunk ki és miért.
    skipped_zero_std = []
    skipped_no_variance = []
    analyzed_ids = []   # azok a test_id-k, amikre tényleg futtatunk elemzést
    blocks = []

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

        analyzed_ids.append(test_id)
        lower_raw, upper_raw = limits[test_id]
        lsl = format_limit(lower_raw, decimal_separator)
        usl = format_limit(upper_raw, decimal_separator)
        blocks.append(build_commands(test_id, lsl, usl))

    # 1) A tiszta, Minitab-barát adatfájl (csak test_id + test_value, TAB-tagolva).
    #    Ezt importálod Minitabba a nyers CSV helyett, így nincs sortörés-gond.
    write_minitab_data(args.data_output, analyzed_ids, values, decimal_separator)

    # 2) A parancsfájl: minden blokkot üres sorral elválasztva egyetlen txt-be,
    #    amit a Minitab Command Line / Session ablakába másolsz (vagy exec-ként futtatsz).
    with open(args.output, "w", encoding="utf-8") as f:
        f.write("\n".join(blocks))

    total = len(order)
    print(f"Total test_id steps found:       {total}")
    print(f"Skipped (zero std dev):          {len(skipped_zero_std)}")
    print(f"Skipped (fewer than 2 points):   {len(skipped_no_variance)}")
    print(f"Capability blocks written:       {len(blocks)}")
    print(f"Data (clean) written to:         {args.data_output}")
    print(f"Commands written to:             {args.output}")


if __name__ == "__main__":
    sys.exit(main())
