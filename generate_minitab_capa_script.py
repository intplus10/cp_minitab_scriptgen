#!/usr/bin/env python3
"""Generate a Minitab session-command script for per-test_id capability analysis.

Reads a tall-table test-data export (one row per measurement, columns include
test_id, test_value, lower_limit, upper_limit), drops test_id steps whose
test_value has zero standard deviation (pass/fail style flags with no real
variation), and emits one Subset + Capability Analysis block per remaining
test_id into a single .txt file that can be pasted into / run from Minitab's
Session window (Edit > Command Line Editor).
"""

import argparse
import csv
import statistics
import sys
from collections import defaultdict


def load_test_steps(csv_path):
    values = defaultdict(list)
    limits = {}
    order = []

    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            test_id = row["test_id"]
            if test_id not in limits:
                order.append(test_id)
                limits[test_id] = (row["lower_limit"], row["upper_limit"])
            values[test_id].append(float(row["test_value"]))

    return order, values, limits


def format_limit(raw):
    value = float(raw)
    if value == int(value):
        return str(int(value))
    return repr(value)


def build_commands(test_id, lsl, uspec):
    return (
        "Subset;\n"
        "  Include;\n"
        f'  LT \'test_id\' "{test_id}";\n'
        f'  Name "{test_id}".\n'
        "Capa 'test_value' 1;\n"
        f"  Lspec {lsl};\n"
        f"  Uspec {uspec};\n"
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
    parser.add_argument("csv_path", help="Path to the tall-table test-data CSV")
    parser.add_argument(
        "-o", "--output", default="minitab_capability_commands.txt",
        help="Output .txt path (default: %(default)s)",
    )
    args = parser.parse_args()

    order, values, limits = load_test_steps(args.csv_path)

    skipped_zero_std = []
    skipped_no_variance = []
    blocks = []

    for test_id in order:
        test_values = values[test_id]
        if len(test_values) < 2:
            skipped_no_variance.append(test_id)
            continue
        std = statistics.stdev(test_values)
        if std == 0:
            skipped_zero_std.append(test_id)
            continue

        lower_raw, upper_raw = limits[test_id]
        lsl = format_limit(lower_raw)
        usl = format_limit(upper_raw)
        blocks.append(build_commands(test_id, lsl, usl))

    with open(args.output, "w", encoding="utf-8") as f:
        f.write("\n".join(blocks))

    total = len(order)
    print(f"Total test_id steps found:       {total}")
    print(f"Skipped (zero std dev):          {len(skipped_zero_std)}")
    print(f"Skipped (fewer than 2 points):   {len(skipped_no_variance)}")
    print(f"Capability blocks written:       {len(blocks)}")
    print(f"Output written to:               {args.output}")


if __name__ == "__main__":
    sys.exit(main())
