#!/usr/bin/env python3
"""Generate a Minitab exec for capability analysis grouped by date / station / parameter.

Second tool, sibling of generate_minitab_capa_script.py. The input here is a
different (already Minitab-ready) tall table with columns:

    Date, Station, W1, W2, Row, Filename, month

Instead of grouping by test_id, this tool groups the measurements by
    year_month  x  station  x  parameter (W1 / W2)
and runs a Capability Analysis on each group, e.g. a group named
"2025_5_DEB_W1" holds every W1 value measured at station DEB in 2025 May.

Everything downstream (inline SET data, Capa, per-group column, the
step_id/Cp/Cpk/Result summary, the separate "summary" worksheet, the Session
Print, the Windows auto-launch) is IDENTICAL to the first tool, so we import
and reuse its engine (build_exec_text + helpers) — the capability/summary logic
lives in ONE place.

The two differences from the first tool are handled here:
  1. Input parsing + grouping (Date/Station/W1/W2 -> composite group key).
  2. The spec limits are NOT in the data, so LSL/USL/Target are parameters
     (defaults below), and a Target subcommand is added to the Capa command.
"""

import argparse
import csv
import os
import re
import statistics
import sys
from collections import defaultdict

# A KÖZÖS motor: az első tool importálható modulként (a main()-je csak __main__
# alatt fut, tehát az import nem indítja el). Innen vesszük a Capa+summary
# generálást és a segédfüggvényeket, hogy a logika egy helyen maradjon.
import generate_minitab_capa_script as core

# Az elemzett paraméter-oszlopok (mindkettőre külön Capa + külön csoport készül).
PARAMETERS = ["W1", "W2"]

# A spec limitek NINCSENEK az adatban, ezért itt adjuk meg (a felhasználó által
# megadott értékek). Kanonikus, PONTOS tizedessel tároljuk (float()-olható), a
# tényleges kimeneti tizedesjelre a core.format_number formázza majd át.
DEFAULT_LSL = "54.75"
DEFAULT_USL = "55.25"
DEFAULT_TARGET = "55"

# Alapértelmezett kimeneti fájl a script mellé (a core SCRIPT_DIR-jét használjuk).
DEFAULT_OUTPUT = os.path.join(core.SCRIPT_DIR, "minitab_dated_capability_analysis.mtb")


def detect_delimiter(sample_line):
    """Kitalálja a mezőelválasztót a fejlécsorból.

    A tizedesjel VESSZŐ az értékekben (55,02...), ezért a mezőelválasztó NEM
    lehet vessző. A tipikus export TAB-bal vagy pontosvesszővel tagolt.
    """
    if "\t" in sample_line:
        return "\t"
    if ";" in sample_line:
        return ";"
    return "\t"  # ésszerű alapértelmezés


def parse_year_month(date_str):
    """"2025-05-26" / "2025.05.26" / "2025/05/26" -> (2025, 5).

    A dátum-elválasztó lehet kötőjel, pont vagy perjel (a különböző exportok
    máshogy formázzák), ezért bármelyikre bontunk. A hónap vezető nulla NÉLKÜL
    (int), tehát 05 -> 5.
    """
    parts = re.split(r"[.\-/]", date_str.strip())
    return int(parts[0]), int(parts[1])


def load_grouped(path):
    """Beolvassa a dátum/állomás/W1/W2 táblát, és csoportosít.

    Visszaad egy dict-et: group_name -> mérési értékek listája (kanonikus,
    PONT-tizedesű string formában, hogy a core.format_number float()-olni tudja).

    A group_name formátuma: "<év>_<hónap>_<állomás>_<paraméter>", pl. 2025_5_DEB_W1.
    Az év_hónap a Date oszlopból jön (a külön 'month' oszlop néhol hibás).
    """
    groups = defaultdict(list)

    with open(path, encoding="utf-8-sig", newline="") as f:
        first_line = f.readline()
        delimiter = detect_delimiter(first_line)
        f.seek(0)
        reader = csv.DictReader(f, delimiter=delimiter)
        for row in reader:
            date_str = (row.get("Date") or "").strip()
            station = (row.get("Station") or "").strip()
            if not date_str or not station:
                continue
            year, month = parse_year_month(date_str)
            for param in PARAMETERS:
                raw = (row.get(param) or "").strip()
                if not raw:
                    continue
                # A bemenet tizedesjele VESSZŐ -> kanonikus PONT-ra normáljuk,
                # így a float() és a core.format_number is helyesen kezeli.
                canonical = raw.replace(",", ".")
                try:
                    float(canonical)
                except ValueError:
                    continue  # nem szám (pl. üres/szemét) -> kihagyjuk
                group_name = f"{year}_{month}_{station}_{param}"
                groups[group_name].append(canonical)

    return groups


def sort_key(group_name):
    """Rendezési kulcs: (év, hónap, állomás, paraméter), hogy a kimenet
    kronologikus és determinisztikus legyen."""
    year, month, station, param = group_name.split("_")
    return (int(year), int(month), station, param)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "data_path", nargs="?",
        help="A dátum/állomás/W1/W2 adatfájl (kihagyva fájlválasztó ablak nyílik)",
    )
    parser.add_argument(
        "-o", "--output", default=DEFAULT_OUTPUT,
        help="Kimeneti Minitab exec (.mtb) útvonala (alap: a script mellé, %(default)s)",
    )
    parser.add_argument(
        "--decimal-separator", choices=[",", "."],
        help="Tizedes elválasztó a kimenethez (kihagyva interaktívan kérdez)",
    )
    parser.add_argument("--lsl", default=DEFAULT_LSL, help="Alsó spec limit (LSL)")
    parser.add_argument("--usl", default=DEFAULT_USL, help="Felső spec limit (USL)")
    parser.add_argument("--target", default=DEFAULT_TARGET, help="Célérték (Target)")
    parser.add_argument(
        "--no-launch", action="store_true",
        help="Ne indítsa el automatikusan a Minitabot a kész .mtb-vel (Windows)",
    )
    args = parser.parse_args()

    # A segédfüggvényeket (fájlválasztó, tizedes-választó) a közös motorból vesszük.
    data_path = args.data_path or core.prompt_csv_path()
    decimal_separator = args.decimal_separator or core.prompt_decimal_separator()

    # A limiteket/célértéket a kiválasztott tizedesjelre formázzuk (54.75 -> 54,75).
    lsl = core.format_number(args.lsl, decimal_separator)
    usl = core.format_number(args.usl, decimal_separator)
    target = core.format_number(args.target, decimal_separator)

    groups = load_grouped(data_path)

    # Szűrés: ugyanaz a logika, mint az első toolban — a 0 szórású és a <2 pontos
    # csoportokat kihagyjuk (nincs értelmes capability analysisük).
    skipped_zero_std = []
    skipped_no_variance = []
    analyzed = []   # (group_name, raw_values, lsl, usl)

    for group_name in sorted(groups, key=sort_key):
        raw_values = groups[group_name]
        numeric = [float(v) for v in raw_values]
        if len(numeric) < 2:
            skipped_no_variance.append(group_name)
            continue
        if statistics.stdev(numeric) == 0:
            skipped_zero_std.append(group_name)
            continue
        analyzed.append((group_name, raw_values, lsl, usl))

    # A KÖZÖS motor összeállítja a teljes exec-et (a Target-et most átadjuk).
    exec_text = core.build_exec_text(analyzed, decimal_separator, target=target)
    with open(args.output, "w", encoding="utf-8") as f:
        f.write(exec_text)

    print(f"Groups found:                    {len(groups)}")
    print(f"Skipped (zero std dev):          {len(skipped_zero_std)}")
    print(f"Skipped (fewer than 2 points):   {len(skipped_no_variance)}")
    print(f"Capability blocks written:       {len(analyzed)}")
    print(f"Spec: LSL={lsl}  USL={usl}  Target={target}")
    print(f"Exec written to:                 {args.output}")

    # Ha nem tiltottuk le, indítsuk el a Minitabot a friss .mtb-vel.
    if not args.no_launch:
        core.launch_minitab(os.path.abspath(args.output))


if __name__ == "__main__":
    sys.exit(main())
