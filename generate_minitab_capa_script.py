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

# --- Summary tab oszlopnevei ---------------------------------------------------
# Minden test_id a SAJÁT adatoszlopába kerül (C1, C2, ... C<N>). Ez szándékos:
# így a kész Minitab projekt minden nyers adatot megőriz, elmenthető és átadható
# másoknak további elemzésre.
#
# A Capa "Storage" opciója mindig az AKTÍV worksheet 1. SORÁBA tárol, és minden
# futásnál FELÜLÍRJA azt. Ezért két lépcsőben dolgozunk:
#   1) a Capa a TEMP oszlopokba tárol (mindig az 1. sorba),
#   2) egy LET átmásolja a temp 1. sorát a SUMMARY oszlop i-edik sorába.
# Így minden test_id eredménye a saját (i-edik) sorába kerül, egymás alá gyűlve.
# A futás végén a summary oszlopokat átmásoljuk egy külön "summary" worksheetre.
COL_STEP_ID = "step_id"   # summary: a tesztlépés azonosítója (szöveg)
COL_CP = "Cp"             # summary: Cp (Within)
COL_CPK = "Cpk"           # summary: Cpk (Within)
COL_RESULT = "Result"     # summary: PASS/FAIL a Cp/Cpk határérték alapján
COL_ID_TMP = "id_tmp"     # temp: a Capa ide tárolja a változónevet (1. sor)
COL_CP_TMP = "cp_tmp"     # temp: a Capa ide tárolja a Cp-t (1. sor)
COL_CPK_TMP = "cpk_tmp"   # temp: a Capa ide tárolja a Cpk-t (1. sor)

# Kapabilitási határérték: FAIL, ha Cp VAGY Cpk ez alatt van.
CAPABILITY_THRESHOLD = 1.33
# A Minitab képletben ELKERÜLJÜK a tizedesjegyet: vesszős tizedes-beállításnál a
# "1,33" ütközne a számformázással ("Expecting operators" hiba). Ezért a
# Cp < 1,33 helyett Cp*100 < 133 alakot használunk — a 133 egész, nincs benne
# vessző. (2 tizedesjegyű határértéket feltételez; ezért szorzunk 100-zal.)
THRESHOLD_X100 = int(round(CAPABILITY_THRESHOLD * 100))

# A summary worksheet neve, amire a végén átmásoljuk a summary oszlopokat.
SUMMARY_WORKSHEET = "summary"

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


def build_summary_header(n_analyzed):
    """Létrehozza (elnevezi) a summary és temp oszlopokat, a DATA oszlopok UTÁN.

    A mérési adat a C1..C<N> oszlopokba kerül (N = elemzett test_id-k száma), ezért
    a summary/temp oszlopokat C<N+1>-től kezdve helyezzük, hogy ne ütközzenek az
    adattal. A Name paranccsal előre elnevezzük őket, hogy a Capa storage és a LET
    névre tudjon hivatkozni.
    """
    return (
        f'Name C{n_analyzed + 1} "{COL_STEP_ID}"'
        f' C{n_analyzed + 2} "{COL_CP}"'
        f' C{n_analyzed + 3} "{COL_CPK}"'
        f' C{n_analyzed + 4} "{COL_RESULT}"'
        f' C{n_analyzed + 5} "{COL_ID_TMP}"'
        f' C{n_analyzed + 6} "{COL_CP_TMP}"'
        f' C{n_analyzed + 7} "{COL_CPK_TMP}"\n'
    )


def build_summary_result():
    """Kiszámolja a PASS/FAIL verdict oszlopot a Cp/Cpk határérték alapján.

    A loop UTÁN fut, amikor a Cp/Cpk oszlopok már fel vannak töltve. Egyetlen
    oszlopszintű LET, ami egyszerre minden sorra kiszámolja az eredményt:
      FAIL, ha Cp < 1,33 VAGY Cpk < 1,33; egyébként PASS.

    Két Minitab-sajátosság miatt így néz ki a képlet:
      - Cp*100 < 133  a  Cp < 1,33  helyett -> így nincs vesszős tizedes a
        képletben (a "1,33" ütközne a számformázással).
      - a függvény-argumentumokat PONTOSVESSZŐ választja el (nem vessző), mert
        vesszős tizedes-beállításnál a Minitab a pontosvesszőt várja
        (ezt a GUI Calculator capture is így generálta).
    """
    return (
        f"Let '{COL_RESULT}' = IF("
        f"'{COL_CP}'*100<{THRESHOLD_X100} Or '{COL_CPK}'*100<{THRESHOLD_X100}; "
        '"FAIL"; "PASS")\n'
    )


def build_summary_copy():
    """A summary oszlopokat átmásolja egy külön, üres worksheetre.

    Ez a futás LEGVÉGÉN fut le. Így a summary a saját "summary" nevű lapján lesz,
    a nyers adat + az összes elemzés pedig a fő worksheeten marad (elmenthető,
    átadható). A parancs pontosan az, amit a GUI generált (Data > Copy):
      Copy 'step_id' 'Cp' 'Cpk' 'Result';  -> mit másolunk
        Newws "summary";                    -> hova: ÚJ worksheet, "summary" néven
        Varnames.                           -> az oszlopneveket is vigye át
    """
    return (
        f"Copy '{COL_STEP_ID}' '{COL_CP}' '{COL_CPK}' '{COL_RESULT}';\n"
        f'  Newws "{SUMMARY_WORKSHEET}";\n'
        "  Varnames.\n"
    )


def build_summary_print():
    """Kiírja a summary táblát (step_id/Cp/Cpk/Result) a Session ablakba.

    A Minitab grafikonjai nem menthetők session commanddal, viszont a report
    kézi összeállításakor (az összes output kijelölése -> Send to Report / Word)
    a Session szöveges kimenete is bekerül. Ha tehát a summary táblát Print-tel
    a Sessionbe írjuk, akkor a végső Word-reportban a chartok MELLETT ott lesz a
    step_id/Cp/Cpk/Result összefoglaló táblázat is — egyetlen exporttal.
    """
    return f"Print '{COL_STEP_ID}' '{COL_CP}' '{COL_CPK}' '{COL_RESULT}'.\n"


def build_capa_block(index, test_id, raw_values, lsl, usl, column, decimal_separator):
    """Összeállítja egy adott test_id-hez tartozó teljes Minitab blokkot, storage-dzsel.

    A blokk lépései:
      1. SET C<n> / <adatok> / END  -> az adott test_id mérési értékeit közvetlenül
                                        beírjuk egy üres oszlopba (nincs külön data
                                        fájl, nincs import). A számokat szóközzel
                                        választjuk el, a tizedesjel a kiválasztott
                                        elválasztó.
      2. NAME C<n> "<test_id>"      -> az oszlopot elnevezzük a test_id-re, így a
                                        Capability Analysis címe a valódi tesztlépés
                                        neve lesz (nem "C1" vagy "test_value").
      3. Capa '<test_id>' 1; ...    -> lefuttatjuk a capability analysist a megadott
                                        LSL/USL limitekkel. A végén a Name/CP/CPK
                                        storage alparancsok a TEMP oszlopokba tárolják
                                        a változónevet, a Cp-t és a Cpk-t (1. sorba).
      4. LET <summary>(index) = <temp>(1) -> a temp 1. sorát átmásoljuk a summary
                                        oszlop i-edik (index) sorába, hogy minden
                                        eredmény a saját sorába kerüljön.

    Minden test_id a SAJÁT oszlopába kerül (C1, C2, ...), így nem kell Subset,
    nem kell worksheet-váltogatás, és nincs esély a táblák "összekeveredésére".

    index: 1-alapú sorszám (egyben a data oszlop C<index> ÉS a summary sor is).
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
        "  CStat;\n"
        f"  Name '{COL_ID_TMP}';\n"
        f"  CP '{COL_CP_TMP}';\n"
        f"  CPK '{COL_CPK_TMP}'.\n"
        f"LET '{COL_STEP_ID}'({index}) = '{COL_ID_TMP}'(1)\n"
        f"LET '{COL_CP}'({index}) = '{COL_CP_TMP}'(1)\n"
        f"LET '{COL_CPK}'({index}) = '{COL_CPK_TMP}'(1)\n"
    )


def launch_minitab(mtb_path):
    """Elindítja a Minitabot a legenerált .mtb exec-kel (Windowson).

    A .mtb a Minitab "exec" fájlkiterjesztése. Az os.startfile() gyakorlatilag
    ugyanaz, mintha duplán kattintanál a fájlra a Fájlkezelőben: a Windows a
    fájltársítás alapján megnyitja a Minitabbal, ami (ha a társítás "run exec")
    le is futtatja a benne lévő parancsokat.

    Csak Windowson működik (az os.startfile máshol nem létezik). Ha a társítás
    hiányzik vagy hiba történik, csak jelezünk, és a felhasználó kézzel is meg
    tudja nyitni (File > Run an Exec).
    """
    if sys.platform != "win32" or not hasattr(os, "startfile"):
        print("Auto-indítás csak Windowson érhető el — kihagyva.")
        print(f"Nyisd meg kézzel: Minitab > File > Run an Exec -> {mtb_path}")
        return
    try:
        os.startfile(mtb_path)  # noqa: S606  (a .mtb társítás indítja a Minitabot)
        print(f"Minitab indítása a következővel: {mtb_path}")
    except OSError as exc:
        print(f"Nem sikerült automatikusan elindítani a Minitabot: {exc}")
        print(f"Nyisd meg kézzel: Minitab > File > Run an Exec -> {mtb_path}")


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
    parser.add_argument(
        "--no-launch", action="store_true",
        help="Ne indítsa el automatikusan a Minitabot a kész .mtb-vel (Windows)",
    )
    args = parser.parse_args()

    # Ha a felhasználó nem adta meg parancssorból, kérdezzük meg felugró ablakban.
    csv_path = args.csv_path or prompt_csv_path()
    decimal_separator = args.decimal_separator or prompt_decimal_separator()

    order, values, limits = load_test_steps(csv_path)

    # 1. MENET: eldöntjük, mely test_id-k kerülnek elemzésre (a 0 szórásúakat
    # kihagyjuk), és összegyűjtjük a szükséges adatokat. Azért külön menetben,
    # mert a summary/temp oszlopok helyéhez előre tudni kell az elemzett
    # test_id-k SZÁMÁT (N), hogy a data oszlopok (C1..C<N>) után helyezzük őket.
    skipped_zero_std = []
    skipped_no_variance = []
    analyzed = []   # (test_id, raw_values, lsl, usl) az elemzendő lépésekhez

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

        lower_raw, upper_raw = limits[test_id]
        lsl = format_number(lower_raw, decimal_separator)
        usl = format_number(upper_raw, decimal_separator)
        analyzed.append((test_id, values[test_id], lsl, usl))

    n_analyzed = len(analyzed)

    # 2. MENET: a parancsblokkok összeállítása. A summary/temp oszlopok neveit
    # a header hozza létre (a data oszlopok után), majd minden blokk a saját
    # C<index> oszlopába teszi az adatot és a summary <index>. sorába az eredményt.
    blocks = [build_summary_header(n_analyzed)]
    for index, (test_id, raw_values, lsl, usl) in enumerate(analyzed, start=1):
        column = f"C{index}"
        blocks.append(
            build_capa_block(index, test_id, raw_values, lsl, usl, column, decimal_separator)
        )
    # A loop után: kiszámoljuk a PASS/FAIL verdict oszlopot, majd a summary
    # oszlopokat átmásoljuk egy külön "summary" worksheetre, és Print-tel a
    # Session ablakba is kiírjuk (hogy a kézi Word-exportba bekerüljön).
    if n_analyzed:
        blocks.append(build_summary_result())
        blocks.append(build_summary_copy())
        blocks.append(build_summary_print())

    # Az önálló exec: minden blokkot üres sorral elválasztva egyetlen fájlba.
    # Minitabban File > Run an Exec -> betölti az inline adatot, lefuttat minden
    # elemzést, ÉS feltölti a step_id/Cp/Cpk summary oszlopokat.
    with open(args.output, "w", encoding="utf-8") as f:
        f.write("\n".join(blocks))

    total = len(order)
    print(f"Total test_id steps found:       {total}")
    print(f"Skipped (zero std dev):          {len(skipped_zero_std)}")
    print(f"Skipped (fewer than 2 points):   {len(skipped_no_variance)}")
    print(f"Capability blocks written:       {n_analyzed}")
    print(f"Summary columns:                 {COL_STEP_ID}, {COL_CP}, {COL_CPK}, {COL_RESULT} "
          f"(C{n_analyzed + 1}-C{n_analyzed + 4})")
    print(f"Exec written to:                 {args.output}")

    # Ha nem tiltottuk le, indítsuk el a Minitabot a friss .mtb-vel.
    if not args.no_launch:
        launch_minitab(os.path.abspath(args.output))


if __name__ == "__main__":
    sys.exit(main())
