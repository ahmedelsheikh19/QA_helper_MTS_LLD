"""
Tests for generate_ut_traceability.py
"""
import os
import pytest
from bs4 import BeautifulSoup

from generate_ut_traceability import (
    extract_reqs_from_req_file,
    extract_reqs_from_ut_files,
    generate_html_report,
)


# ── extract_reqs_from_req_file ─────────────────────────────────────────────
# (same function as in LLD script — run a quick smoke-check set)

class TestExtractReqsFromReqFile:

    def test_reads_id_column(self, req_file_id_col):
        result = extract_reqs_from_req_file(str(req_file_id_col))
        assert result == {"REQ-001", "REQ-002", "REQ-003"}

    def test_reads_object_identifier_column(self, req_file_object_identifier_col):
        result = extract_reqs_from_req_file(str(req_file_object_identifier_col))
        assert result == {"REQ-001", "REQ-002"}

    def test_returns_empty_when_no_id_column(self, req_file_no_id_col):
        assert extract_reqs_from_req_file(str(req_file_no_id_col)) == set()

    def test_returns_empty_when_no_data(self, req_file_empty_rows):
        assert extract_reqs_from_req_file(str(req_file_empty_rows)) == set()

    def test_deduplicates_ids(self, req_file_duplicates):
        assert extract_reqs_from_req_file(str(req_file_duplicates)) == {"REQ-001", "REQ-002"}


# ── extract_reqs_from_ut_files ─────────────────────────────────────────────

class TestExtractReqsFromUtFiles:

    # Happy path ──────────────────────────────────────────────────────────

    def test_extracts_valid_requirements(self, ut_file_valid):
        reqs, errors = extract_reqs_from_ut_files([str(ut_file_valid)])
        assert "REQ-001" in reqs
        assert "REQ-002" in reqs
        assert "REQ-003" in reqs

    def test_no_format_errors_on_valid_file(self, ut_file_valid):
        _, errors = extract_reqs_from_ut_files([str(ut_file_valid)])
        assert errors == []

    def test_returns_set_and_list(self, ut_file_valid):
        result = extract_reqs_from_ut_files([str(ut_file_valid)])
        assert isinstance(result[0], set)
        assert isinstance(result[1], list)

    def test_empty_file_list(self):
        reqs, errors = extract_reqs_from_ut_files([])
        assert reqs == set()
        assert errors == []

    def test_no_labels_in_file(self, ut_file_no_labels):
        reqs, errors = extract_reqs_from_ut_files([str(ut_file_no_labels)])
        assert reqs == set()
        assert errors == []

    # Multi-file & multi-sheet ─────────────────────────────────────────────

    def test_processes_multiple_files(self, ut_file_valid, ut_file_multi_sheet):
        reqs, _ = extract_reqs_from_ut_files(
            [str(ut_file_valid), str(ut_file_multi_sheet)])
        # ut_file_valid  → REQ-001, REQ-002, REQ-003
        # ut_file_multi_sheet → REQ-001, REQ-002  (different cells)
        assert {"REQ-001", "REQ-002", "REQ-003"}.issubset(reqs)

    def test_processes_multiple_sheets(self, ut_file_multi_sheet):
        reqs, _ = extract_reqs_from_ut_files([str(ut_file_multi_sheet)])
        assert "REQ-001" in reqs   # from Suite1
        assert "REQ-002" in reqs   # from Suite2

    def test_deduplicates_across_files(self, ut_file_valid, ut_file_multi_sheet):
        reqs, _ = extract_reqs_from_ut_files(
            [str(ut_file_valid), str(ut_file_multi_sheet)])
        # REQ-001 appears in both — should appear once in the set
        assert len([r for r in reqs if r == "REQ-001"]) == 1

    # Format validation ────────────────────────────────────────────────────

    def test_detects_spaces_in_value(self, ut_file_format_errors):
        _, errors = extract_reqs_from_ut_files([str(ut_file_format_errors)])
        reasons = " ".join(e["reason"] for e in errors)
        assert "Contains spaces" in reasons

    def test_detects_newlines_in_value(self, ut_file_format_errors):
        _, errors = extract_reqs_from_ut_files([str(ut_file_format_errors)])
        reasons = " ".join(e["reason"] for e in errors)
        assert "Contains new lines" in reasons

    def test_detects_missing_trailing_comma(self, ut_file_format_errors):
        _, errors = extract_reqs_from_ut_files([str(ut_file_format_errors)])
        reasons = " ".join(e["reason"] for e in errors)
        assert "Does not end with a comma" in reasons

    def test_format_error_contains_file_info(self, ut_file_format_errors):
        _, errors = extract_reqs_from_ut_files([str(ut_file_format_errors)])
        assert len(errors) > 0
        err = errors[0]
        assert "file" in err
        assert "sheet" in err
        assert "cell" in err
        assert "value" in err
        assert "reason" in err

    def test_no_error_for_valid_format(self, ut_file_valid):
        _, errors = extract_reqs_from_ut_files([str(ut_file_valid)])
        assert len(errors) == 0

    def test_still_extracts_reqs_despite_format_errors(self, ut_file_format_errors):
        """Format errors must not block extraction."""
        reqs, errors = extract_reqs_from_ut_files([str(ut_file_format_errors)])
        assert len(reqs) > 0
        assert len(errors) > 0

    def test_multiple_violations_in_one_cell(self, ut_file_format_errors):
        """A cell with both spaces and no comma should have both reasons."""
        _, errors = extract_reqs_from_ut_files([str(ut_file_format_errors)])
        multi = [e for e in errors if "Contains spaces" in e["reason"]
                 and "Does not end with a comma" in e["reason"]]
        assert len(multi) >= 1

    # Invalid / missing files ──────────────────────────────────────────────

    def test_skips_nonexistent_file_gracefully(self, ut_file_valid):
        """A missing file should be skipped; valid files still processed."""
        reqs, errors = extract_reqs_from_ut_files(
            ["/nonexistent/path/ut.xlsx", str(ut_file_valid)])
        assert "REQ-001" in reqs   # valid file was still processed

    def test_invalid_file_does_not_raise(self, tmp_path):
        """A corrupt/non-Excel file must not crash the function."""
        bad = tmp_path / "bad.xlsx"
        bad.write_text("this is not an xlsx file")
        reqs, errors = extract_reqs_from_ut_files([str(bad)])
        assert isinstance(reqs, set)
        assert isinstance(errors, list)


# ── generate_html_report ───────────────────────────────────────────────────

class TestGenerateHtmlReport:

    def test_creates_output_file(self, tmp_path):
        out = str(tmp_path / "ut_report.html")
        generate_html_report({"REQ-001"}, {"REQ-001"}, [], out)
        assert os.path.exists(out)

    def test_covered_count_in_html(self, tmp_path):
        out = str(tmp_path / "ut_report.html")
        generate_html_report({"REQ-001", "REQ-002"}, {"REQ-001", "REQ-002"}, [], out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        counts = [int(c.text) for c in soup.select(".card h3")]
        assert counts[0] == 2   # covered

    def test_missing_in_ut_count(self, tmp_path):
        out = str(tmp_path / "ut_report.html")
        generate_html_report({"REQ-001", "REQ-002"}, {"REQ-001"}, [], out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        counts = [int(c.text) for c in soup.select(".card h3")]
        assert counts[1] == 1   # missing in ut

    def test_unmapped_count(self, tmp_path):
        out = str(tmp_path / "ut_report.html")
        generate_html_report({"REQ-001"}, {"REQ-001", "REQ-EXTRA"}, [], out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        counts = [int(c.text) for c in soup.select(".card h3")]
        assert counts[2] == 1   # unmapped

    def test_format_violations_count(self, tmp_path):
        out = str(tmp_path / "ut_report.html")
        errors = [{"file": "f.xlsx", "sheet": "S1", "cell": "B1",
                   "value": "REQ-001", "reason": "Does not end with a comma"}]
        generate_html_report({"REQ-001"}, {"REQ-001"}, errors, out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        counts = [int(c.text) for c in soup.select(".card h3")]
        assert counts[3] == 1   # format violations

    def test_zero_format_violations_shows_success_message(self, tmp_path):
        out = str(tmp_path / "ut_report.html")
        generate_html_report({"REQ-001"}, {"REQ-001"}, [], out)
        content = open(out, encoding="utf-8").read()
        assert "No formatting errors found" in content

    def test_all_zero_counts_when_both_empty(self, tmp_path):
        out = str(tmp_path / "ut_report.html")
        generate_html_report(set(), set(), [], out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        counts = [int(c.text) for c in soup.select(".card h3")]
        assert counts == [0, 0, 0, 0]

    def test_uses_default_filename(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        generate_html_report({"REQ-001"}, {"REQ-001"}, [])
        assert os.path.exists(tmp_path / "ut_traceability_report.html")

    def test_format_error_details_appear_in_html(self, tmp_path):
        out = str(tmp_path / "ut_report.html")
        errors = [{"file": "myfile.xlsx", "sheet": "MySuite", "cell": "B3",
                   "value": "REQ BAD", "reason": "Contains spaces"}]
        generate_html_report(set(), set(), errors, out)
        content = open(out, encoding="utf-8").read()
        assert "myfile.xlsx" in content
        assert "MySuite" in content
        assert "B3" in content


# ── Integration ────────────────────────────────────────────────────────────

class TestIntegration:

    def test_full_valid_pipeline(self, req_file_id_col, ut_file_valid, tmp_path):
        """
        req_file has REQ-001..003.
        ut_file has REQ-001, REQ-002, REQ-003 → covered=3, missing=0, format errors=0.
        """
        out = str(tmp_path / "report.html")
        req_ids = extract_reqs_from_req_file(str(req_file_id_col))
        ut_reqs, fmt_errors = extract_reqs_from_ut_files([str(ut_file_valid)])
        generate_html_report(req_ids, ut_reqs, fmt_errors, out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        counts = [int(c.text) for c in soup.select(".card h3")]
        assert counts[0] == 3   # covered
        assert counts[1] == 0   # missing in ut
        assert counts[3] == 0   # no format errors

    def test_pipeline_with_format_errors(self, req_file_id_col, ut_file_format_errors, tmp_path):
        out = str(tmp_path / "report.html")
        req_ids = extract_reqs_from_req_file(str(req_file_id_col))
        ut_reqs, fmt_errors = extract_reqs_from_ut_files([str(ut_file_format_errors)])
        generate_html_report(req_ids, ut_reqs, fmt_errors, out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        counts = [int(c.text) for c in soup.select(".card h3")]
        assert counts[3] > 0    # format errors present

    def test_pipeline_with_multiple_ut_files(
            self, req_file_id_col, ut_file_valid, ut_file_multi_sheet, tmp_path):
        out = str(tmp_path / "report.html")
        req_ids = extract_reqs_from_req_file(str(req_file_id_col))
        ut_reqs, fmt_errors = extract_reqs_from_ut_files(
            [str(ut_file_valid), str(ut_file_multi_sheet)])
        generate_html_report(req_ids, ut_reqs, fmt_errors, out)
        assert os.path.exists(out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        counts = [int(c.text) for c in soup.select(".card h3")]
        assert counts[0] >= 1   # at least something covered
