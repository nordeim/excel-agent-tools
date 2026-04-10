"""Benchmark large dataset write performance."""

from __future__ import annotations

import tempfile
import time
from pathlib import Path

from openpyxl import Workbook


def bench_write_large() -> dict[str, float]:
    """Benchmark writing 100k rows."""
    rows, cols = 100_000, 5

    path = Path(tempfile.mktemp(suffix=".xlsx"))

    try:
        start = time.perf_counter()

        # Write using openpyxl write-only mode
        wb = Workbook(write_only=True)
        ws = wb.create_sheet("Data")

        # Header
        ws.append([f"Col{i}" for i in range(1, cols + 1)])

        # Data rows
        for i in range(1, rows + 1):
            ws.append([i, f"Item {i}", i * 1.5, i * 2, i % 2 == 0])

        wb.save(str(path))

        elapsed = time.perf_counter() - start

        return {
            "rows": rows,
            "cols": cols,
            "elapsed_seconds": round(elapsed, 2),
            "rows_per_second": round(rows / elapsed, 0),
            "target_met": elapsed < 5.0,
        }
    finally:
        path.unlink(missing_ok=True)


if __name__ == "__main__":
    results = bench_write_large()
    print(f"Wrote {results['rows']} rows in {results['elapsed_seconds']}s")
    print(f"Rows/sec: {results['rows_per_second']}")
    print(f"Target met (<5s): {results['target_met']}")
