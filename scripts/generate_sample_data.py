from __future__ import annotations

import csv
import math
import random
from pathlib import Path

from openpyxl import Workbook


def build_rows(row_count: int = 5000, seed: int = 42) -> list[list[object]]:
    rng = random.Random(seed)
    headers = ["cluster", "x", "y", "z", "intensity", "temperature", "label"]
    rows: list[list[object]] = [headers]
    centers = [
        ("Orion", -42.0, 18.0, 12.0),
        ("Lyra", 21.0, -31.0, 44.0),
        ("Cygnus", 53.0, 29.0, -25.0),
        ("Hydra", -18.0, -46.0, -34.0),
    ]

    for index in range(row_count):
        cluster, cx, cy, cz = centers[index % len(centers)]
        wobble = 1 + (index % 7) * 0.12
        x = rng.gauss(cx, 8.5 * wobble)
        y = rng.gauss(cy, 6.8 * wobble)
        z = rng.gauss(cz, 7.4 * wobble)
        intensity = math.sqrt(x * x + y * y + z * z) + rng.random() * 9
        temperature = 280 + intensity * 1.8 + rng.gauss(0, 6)
        rows.append(
            [
                cluster,
                round(x, 4),
                round(y, 4),
                round(z, 4),
                round(intensity, 4),
                round(temperature, 4),
                f"{cluster}-{index:05d}",
            ]
        )
    return rows


def write_csv(path: Path, rows: list[list[object]], delimiter: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle, delimiter=delimiter)
        writer.writerows(rows)


def write_xlsx(path: Path, rows: list[list[object]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "clusters"
    for row in rows:
        sheet.append(row)
    workbook.save(path)


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    output_dir = root / "examples"
    rows = build_rows()
    write_csv(output_dir / "cluster_demo.csv", rows, delimiter=",")
    write_csv(output_dir / "cluster_demo.tsv", rows, delimiter="\t")
    write_xlsx(output_dir / "cluster_demo.xlsx", rows)
    print(f"Sample files written to: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
