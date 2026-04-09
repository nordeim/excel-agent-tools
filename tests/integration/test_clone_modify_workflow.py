"""
End-to-End Integration Test: Standard Data Pipeline Workflow

Simulates an AI agent executing a complete financial reporting workflow:
Clone → Metadata → Read → Write → Insert Rows → Recalculate → Validate → Export PDF.

All tools are invoked via subprocess to guarantee CLI contract compliance.
JSON outputs are chained between steps exactly as an orchestration framework would.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path

import pytest


# ---------------------------------------------------------------------------
# Subprocess Helper
# ---------------------------------------------------------------------------


def _run_tool(tool_module: str, *args: str, cwd: Path | None = None) -> tuple[dict, int]:
    """
    Execute a CLI tool via subprocess, mimicking AI agent invocation.

    Returns:
        (parsed_json_dict, return_code)
    """
    env = os.environ.copy()
    # Governance secret required for token generation/validation
    env["EXCEL_AGENT_SECRET"] = "e2e-test-secret-key-2026"

    cmd = [sys.executable, "-m", f"excel_agent.tools.{tool_module}", *args]

    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=45,
        env=env,
        cwd=cwd,
    )

    out = result.stdout.strip()
    if not out:
        raise AssertionError(
            f"Tool {tool_module} produced no stdout output.\nStderr: {result.stderr}"
        )

    try:
        data = json.loads(out)
    except json.JSONDecodeError as exc:
        raise AssertionError(f"Invalid JSON from {tool_module}: {exc}") from exc

    return data, result.returncode


# ---------------------------------------------------------------------------
# Test Suite
# ---------------------------------------------------------------------------


class TestCloneModifyWorkflow:
    """
    Validates the full 8-step data pipeline used by AI agents for
    safe, traceable workbook manipulation.
    """

    @pytest.fixture(scope="class")
    def workflow_env(self, tmp_path_factory: pytest.TempPathFactory) -> Path:
        """Prepare isolated workspace for the entire workflow."""
        tmp = tmp_path_factory.mktemp("workflow_e2e")
        return tmp

    def test_full_data_pipeline(self, workflow_env: Path, sample_workbook: Path) -> None:
        """Execute and validate the complete clone → modify → recalc → export chain."""
        start_time = time.monotonic()
        work_dir = workflow_env / "work"
        work_dir.mkdir()

        # Step 1: Clone source to safe working copy
        clone_data, clone_code = _run_tool(
            "governance.xls_clone_workbook",
            "--input",
            str(sample_workbook),
            "--output-dir",
            str(work_dir),
            cwd=workflow_env,
        )
        assert clone_code == 0
        assert clone_data["status"] == "success"
        clone_path = Path(clone_data["data"]["clone_path"])
        assert clone_path.exists(), "Clone file was not created"

        # Step 2: Get workbook metadata
        meta_data, meta_code = _run_tool(
            "read.xls_get_workbook_metadata",
            "--input",
            str(clone_path),
            cwd=workflow_env,
        )
        assert meta_code == 0
        assert meta_data["status"] == "success"
        assert meta_data["data"]["sheet_count"] >= 1
        assert meta_data["data"]["total_formulas"] > 0, "Test fixture must contain formulas"

        # Step 3: Read initial range
        read_data, read_code = _run_tool(
            "read.xls_read_range",
            "--input",
            str(clone_path),
            "--range",
            "A1:B2",
            "--sheet",
            "Sheet1",
            cwd=workflow_env,
        )
        assert read_code == 0
        assert read_data["status"] == "success"
        assert len(read_data["data"]["values"]) == 2, "Expected 2 rows from A1:B2"

        # Step 4: Write new data
        write_data, write_code = _run_tool(
            "write.xls_write_range",
            "--input",
            str(clone_path),
            "--output",
            str(clone_path),
            "--range",
            "F1",
            "--sheet",
            "Sheet1",
            "--data",
            '[["Agent", "E2E"], [42, true]]',
            cwd=workflow_env,
        )
        assert write_code == 0
        assert write_data["status"] == "success"
        assert write_data["impact"]["cells_modified"] == 4

        # Step 5: Insert rows (structural mutation)
        insert_data, insert_code = _run_tool(
            "structure.xls_insert_rows",
            "--input",
            str(clone_path),
            "--output",
            str(clone_path),
            "--sheet",
            "Sheet1",
            "--before-row",
            "3",
            "--count",
            "2",
            cwd=workflow_env,
        )
        assert insert_code == 0
        assert insert_data["status"] == "success"

        # Step 6: Recalculate (Tier 1 → Tier 2 fallback)
        recalc_data, recalc_code = _run_tool(
            "formulas.xls_recalculate",
            "--input",
            str(clone_path),
            "--output",
            str(clone_path),
            cwd=workflow_env,
        )
        assert recalc_code == 0
        assert recalc_data["status"] == "success"
        engine = recalc_data["data"].get("engine", "unknown")
        assert engine in ("tier1_formulas", "tier2_libreoffice"), (
            f"Unexpected calc engine: {engine}"
        )

        # Step 7: Validate integrity
        valid_data, valid_code = _run_tool(
            "governance.xls_validate_workbook",
            "--input",
            str(clone_path),
            cwd=workflow_env,
        )
        # Valid or warning (circular refs are non-fatal)
        assert valid_code in (0, 1)
        assert valid_data["status"] in ("success", "warning")
        assert "errors" in valid_data["data"]

        # Step 8: Export PDF (skip gracefully if LibreOffice unavailable)
        pdf_path = workflow_env / "report.pdf"
        lo_available = (
            subprocess.run(
                ["soffice", "--headless", "--version"], capture_output=True, text=True, timeout=5
            ).returncode
            == 0
            or subprocess.run(
                ["libreoffice", "--headless", "--version"],
                capture_output=True,
                text=True,
                timeout=5,
            ).returncode
            == 0
        )

        if lo_available:
            pdf_data, pdf_code = _run_tool(
                "export.xls_export_pdf",
                "--input",
                str(clone_path),
                "--outfile",
                str(pdf_path),
                "--recalc",
                cwd=workflow_env,
            )
            assert pdf_code == 0, f"PDF export failed: {pdf_data.get('error', 'unknown')}"
            assert pdf_data["status"] == "success"
            assert pdf_path.exists(), "PDF file was not created"
            assert pdf_path.stat().st_size > 100, "PDF file appears empty or corrupted"
        else:
            pytest.skip("LibreOffice not installed; skipping PDF export validation")

        # Final Timing Assertion
        elapsed = time.monotonic() - start_time
        assert elapsed < 60, f"Full pipeline took {elapsed:.1f}s — exceeds 60s SLA"

    def test_clone_isolation(self, workflow_env: Path, sample_workbook: Path) -> None:
        """Verify cloned workbooks are independent of source."""
        work_dir = workflow_env / "work_isolated"
        work_dir.mkdir()

        # Clone
        clone_data, _ = _run_tool(
            "governance.xls_clone_workbook",
            "--input",
            str(sample_workbook),
            "--output-dir",
            str(work_dir),
            cwd=workflow_env,
        )
        clone_path = Path(clone_data["data"]["clone_path"])

        # Modify clone
        _run_tool(
            "write.xls_write_cell",
            "--input",
            str(clone_path),
            "--output",
            str(clone_path),
            "--cell",
            "A1",
            "--value",
            "MODIFIED",
            cwd=workflow_env,
        )

        # Verify source unchanged
        source_data, _ = _run_tool(
            "read.xls_read_range",
            "--input",
            str(sample_workbook),
            "--range",
            "A1",
            cwd=workflow_env,
        )
        assert source_data["data"]["values"][0][0] == "Name"  # Original value preserved

    def test_chunked_read_large_dataset(self, workflow_env: Path, large_workbook: Path) -> None:
        """Verify chunked reading of large datasets completes efficiently."""
        # Chunked mode returns JSONL, so we need special handling
        env = os.environ.copy()
        env["EXCEL_AGENT_SECRET"] = "e2e-test-secret-key-2026"

        cmd = [
            sys.executable,
            "-m",
            "excel_agent.tools.read.xls_read_range",
            "--input",
            str(large_workbook),
            "--range",
            "A1:E100",
            "--chunked",
        ]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=45,
            env=env,
            cwd=workflow_env,
        )

        assert result.returncode == 0
        # Parse JSONL - each line is a JSON object
        lines = result.stdout.strip().split("\n")
        assert len(lines) > 0, "Expected at least one chunk"

        # Verify each chunk is valid JSON
        for line in lines:
            if line.strip():
                chunk = json.loads(line)
                assert chunk.get("status") == "success"

    def test_formula_preservation_through_pipeline(
        self, workflow_env: Path, formula_workbook: Path
    ) -> None:
        """Ensure formulas are preserved through clone → modify → save cycles."""
        work_dir = workflow_env / "work_formulas"
        work_dir.mkdir()

        # Clone
        clone_data, _ = _run_tool(
            "governance.xls_clone_workbook",
            "--input",
            str(formula_workbook),
            "--output-dir",
            str(work_dir),
            cwd=workflow_env,
        )
        clone_path = Path(clone_data["data"]["clone_path"])

        # Get original formula
        orig_formula, _ = _run_tool(
            "read.xls_get_formula",
            "--input",
            str(clone_path),
            "--cell",
            "C1",
            "--sheet",
            "Sheet1",
            cwd=workflow_env,
        )
        assert orig_formula["data"]["formula"] == "=B1+5"

        # Modify unrelated cell
        _run_tool(
            "write.xls_write_cell",
            "--input",
            str(clone_path),
            "--output",
            str(clone_path),
            "--cell",
            "D1",
            "--value",
            "999",
            cwd=workflow_env,
        )

        # Verify formula preserved
        preserved_formula, _ = _run_tool(
            "read.xls_get_formula",
            "--input",
            str(clone_path),
            "--cell",
            "C1",
            "--sheet",
            "Sheet1",
            cwd=workflow_env,
        )
        assert preserved_formula["data"]["formula"] == "=B1+5"

    def test_cross_sheet_references_after_insert(
        self, workflow_env: Path, sample_workbook: Path
    ) -> None:
        """Verify cross-sheet references are maintained after structural changes."""
        work_dir = workflow_env / "work_cross_sheet"
        work_dir.mkdir()

        # Clone
        clone_data, _ = _run_tool(
            "governance.xls_clone_workbook",
            "--input",
            str(sample_workbook),
            "--output-dir",
            str(work_dir),
            cwd=workflow_env,
        )
        clone_path = Path(clone_data["data"]["clone_path"])

        # Get original cross-sheet formula
        orig_data, _ = _run_tool(
            "read.xls_get_formula",
            "--input",
            str(clone_path),
            "--cell",
            "B1",
            "--sheet",
            "Sheet2",
            cwd=workflow_env,
        )
        assert "Sheet1" in orig_data["data"]["formula"]

        # Insert rows in Sheet1
        _run_tool(
            "structure.xls_insert_rows",
            "--input",
            str(clone_path),
            "--output",
            str(clone_path),
            "--sheet",
            "Sheet1",
            "--before-row",
            "5",
            "--count",
            "1",
            cwd=workflow_env,
        )

        # Recalculate
        _run_tool(
            "formulas.xls_recalculate",
            "--input",
            str(clone_path),
            "--output",
            str(clone_path),
            cwd=workflow_env,
        )

        # Verify cross-sheet reference still works
        recalc_data, _ = _run_tool(
            "read.xls_read_range",
            "--input",
            str(clone_path),
            "--range",
            "B1",
            "--sheet",
            "Sheet2",
            cwd=workflow_env,
        )
        # Should have a calculated value (not #REF!)
        value = recalc_data["data"]["values"][0][0]
        assert value is not None
        assert not isinstance(value, str) or not value.startswith("#")

    def test_export_csv_alternative(self, workflow_env: Path, sample_workbook: Path) -> None:
        """Verify CSV export as alternative to PDF."""
        work_dir = workflow_env / "work_csv"
        work_dir.mkdir()

        # Clone
        clone_data, _ = _run_tool(
            "governance.xls_clone_workbook",
            "--input",
            str(sample_workbook),
            "--output-dir",
            str(work_dir),
            cwd=workflow_env,
        )
        clone_path = Path(clone_data["data"]["clone_path"])

        # Export CSV
        csv_path = workflow_env / "export.csv"
        csv_data, csv_code = _run_tool(
            "export.xls_export_csv",
            "--input",
            str(clone_path),
            "--outfile",
            str(csv_path),
            "--sheet",
            "Sheet1",
            cwd=workflow_env,
        )
        assert csv_code == 0
        assert csv_data["status"] == "success"
        assert csv_path.exists()
        assert csv_path.stat().st_size > 0

        # Verify content
        content = csv_path.read_text()
        assert "Name" in content
        assert "Value" in content

    def test_export_json_alternative(self, workflow_env: Path, sample_workbook: Path) -> None:
        """Verify JSON export as alternative to PDF."""
        work_dir = workflow_env / "work_json"
        work_dir.mkdir()

        # Clone
        clone_data, _ = _run_tool(
            "governance.xls_clone_workbook",
            "--input",
            str(sample_workbook),
            "--output-dir",
            str(work_dir),
            cwd=workflow_env,
        )
        clone_path = Path(clone_data["data"]["clone_path"])

        # Export JSON
        json_path = workflow_env / "export.json"
        json_data, json_code = _run_tool(
            "export.xls_export_json",
            "--input",
            str(clone_path),
            "--outfile",
            str(json_path),
            "--sheet",
            "Sheet1",
            "--orient",
            "records",
            cwd=workflow_env,
        )
        assert json_code == 0
        assert json_data["status"] == "success"
        assert json_path.exists()

        # Verify valid JSON
        exported = json.loads(json_path.read_text())
        assert isinstance(exported, list)
        assert len(exported) > 0
