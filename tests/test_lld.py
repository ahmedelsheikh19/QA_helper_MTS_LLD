"""
Tests for generate_traceabilityLLD.py
"""
import os
import pytest
from bs4 import BeautifulSoup

from generate_traceabilityLLD import (
    extract_reqs_from_req_file,
    extract_reqs_from_lld_file,
    generate_html_report,
)


# ── extract_reqs_from_req_file ─────────────────────────────────────────────

class TestExtractReqsFromReqFile:

    def test_reads_id_column(self, req_file_id_col):
        result = extract_reqs_from_req_file(str(req_file_id_col))
        assert result == {"REQ-001", "REQ-002", "REQ-003"}

    def test_reads_object_identifier_column(self, req_file_object_identifier_col):
        result = extract_reqs_from_req_file(str(req_file_object_identifier_col))
        assert result == {"REQ-001", "REQ-002"}

    def test_returns_empty_set_when_no_id_column(self, req_file_no_id_col):
        result = extract_reqs_from_req_file(str(req_file_no_id_col))
        assert result == set()

    def test_returns_empty_set_when_no_data_rows(self, req_file_empty_rows):
        result = extract_reqs_from_req_file(str(req_file_empty_rows))
        assert result == set()

    def test_deduplicates_ids(self, req_file_duplicates):
        result = extract_reqs_from_req_file(str(req_file_duplicates))
        assert result == {"REQ-001", "REQ-002"}

    def test_strips_whitespace_from_ids(self, req_file_whitespace):
        result = extract_reqs_from_req_file(str(req_file_whitespace))
        assert result == {"REQ-001", "REQ-002"}

    def test_returns_set_type(self, req_file_id_col):
        result = extract_reqs_from_req_file(str(req_file_id_col))
        assert isinstance(result, set)

    def test_column_detection_is_case_insensitive(self, tmp_path):
        """Column header in various cases should all be recognised."""
        import openpyxl
        for header in ["id", "Id", "iD", "ID"]:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append([header])
            ws.append(["REQ-001"])
            p = tmp_path / f"reqs_{header}.xlsx"
            wb.save(p)
            result = extract_reqs_from_req_file(str(p))
            assert "REQ-001" in result, f"Failed for header '{header}'"


# ── extract_reqs_from_lld_file ─────────────────────────────────────────────

class TestExtractReqsFromLldFile:

    def test_reads_req_column(self, lld_file_single_reqs):
        result = extract_reqs_from_lld_file(str(lld_file_single_reqs))
        assert result == {"REQ-001", "REQ-002"}

    def test_splits_comma_separated_reqs(self, lld_file_multi_reqs):
        result = extract_reqs_from_lld_file(str(lld_file_multi_reqs))
        assert {"REQ-001", "REQ-002"}.issubset(result)

    def test_splits_newline_separated_reqs(self, lld_file_multi_reqs):
        result = extract_reqs_from_lld_file(str(lld_file_multi_reqs))
        assert {"REQ-003", "REQ-004"}.issubset(result)

    def test_splits_space_separated_reqs(self, lld_file_multi_reqs):
        result = extract_reqs_from_lld_file(str(lld_file_multi_reqs))
        assert {"REQ-005", "REQ-006"}.issubset(result)

    def test_returns_empty_set_when_no_req_column(self, lld_file_no_req_col):
        result = extract_reqs_from_lld_file(str(lld_file_no_req_col))
        assert result == set()

    def test_returns_empty_set_when_no_data_rows(self, lld_file_empty_rows):
        result = extract_reqs_from_lld_file(str(lld_file_empty_rows))
        assert result == set()

    def test_deduplicates_reqs(self, tmp_path):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["REQ"])
        ws.append(["REQ-001"])
        ws.append(["REQ-001"])
        p = tmp_path / "lld_dup.xlsx"
        wb.save(p)
        result = extract_reqs_from_lld_file(str(p))
        assert result == {"REQ-001"}

    def test_returns_set_type(self, lld_file_single_reqs):
        result = extract_reqs_from_lld_file(str(lld_file_single_reqs))
        assert isinstance(result, set)


# ── generate_html_report ───────────────────────────────────────────────────

class TestGenerateHtmlReport:

    def test_creates_output_file(self, tmp_path):
        out = str(tmp_path / "report.html")
        generate_html_report({"REQ-001"}, {"REQ-001"}, out)
        assert os.path.exists(out)

    def test_covered_count_in_html(self, tmp_path):
        out = str(tmp_path / "report.html")
        generate_html_report({"REQ-001", "REQ-002"}, {"REQ-001", "REQ-002"}, out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        cards = soup.select(".card h3")
        counts = [int(c.text) for c in cards]
        assert counts[0] == 2   # covered

    def test_missing_in_lld_count(self, tmp_path):
        out = str(tmp_path / "report.html")
        generate_html_report({"REQ-001", "REQ-002"}, {"REQ-001"}, out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        counts = [int(c.text) for c in soup.select(".card h3")]
        assert counts[1] == 1   # missing in lld

    def test_missing_in_reqs_count(self, tmp_path):
        out = str(tmp_path / "report.html")
        generate_html_report({"REQ-001"}, {"REQ-001", "REQ-EXTRA"}, out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        counts = [int(c.text) for c in soup.select(".card h3")]
        assert counts[2] == 1   # missing in reqs

    def test_all_zero_counts_when_both_empty(self, tmp_path):
        out = str(tmp_path / "report.html")
        generate_html_report(set(), set(), out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        counts = [int(c.text) for c in soup.select(".card h3")]
        assert counts == [0, 0, 0]

    def test_req_ids_appear_in_covered_section(self, tmp_path):
        out = str(tmp_path / "report.html")
        generate_html_report({"REQ-001"}, {"REQ-001"}, out)
        content = open(out, encoding="utf-8").read()
        assert "REQ-001" in content

    def test_missing_req_appears_in_missing_section(self, tmp_path):
        out = str(tmp_path / "report.html")
        generate_html_report({"REQ-001", "REQ-002"}, {"REQ-001"}, out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        missing_section = soup.find("div", {"id": "missing-lld"})
        assert "REQ-002" in missing_section.text

    def test_uses_default_filename(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        generate_html_report({"REQ-001"}, {"REQ-001"})
        assert os.path.exists(tmp_path / "traceability_report.html")

    def test_overwrites_existing_file(self, tmp_path):
        out = str(tmp_path / "report.html")
        generate_html_report({"REQ-001"}, {"REQ-001"}, out)
        generate_html_report({"REQ-002"}, {"REQ-002"}, out)
        content = open(out, encoding="utf-8").read()
        assert "REQ-002" in content


# ── Integration ────────────────────────────────────────────────────────────

class TestIntegration:

    def test_full_coverage(self, req_file_id_col, lld_file_single_reqs, tmp_path):
        """Requirements match exactly what's in the LLD — covered=2, missing=1 (REQ-003 not in LLD)."""
        out = str(tmp_path / "report.html")
        req_ids  = extract_reqs_from_req_file(str(req_file_id_col))
        lld_reqs = extract_reqs_from_lld_file(str(lld_file_single_reqs))
        generate_html_report(req_ids, lld_reqs, out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        counts = [int(c.text) for c in soup.select(".card h3")]
        assert counts[0] == 2   # covered (REQ-001, REQ-002)
        assert counts[1] == 1   # missing in lld (REQ-003)
        assert counts[2] == 0   # nothing extra in lld

    def test_extra_lld_items(self, req_file_id_col, tmp_path):
        """LLD has reqs that don't exist in requirements file."""
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["REQ"])
        ws.append(["REQ-001"])
        ws.append(["REQ-EXTRA"])
        lld_path = tmp_path / "lld_extra.xlsx"
        wb.save(lld_path)

        out = str(tmp_path / "report.html")
        req_ids  = extract_reqs_from_req_file(str(req_file_id_col))
        lld_reqs = extract_reqs_from_lld_file(str(lld_path))
        generate_html_report(req_ids, lld_reqs, out)
        soup = BeautifulSoup(open(out, encoding="utf-8").read(), "html.parser")
        counts = [int(c.text) for c in soup.select(".card h3")]
        assert counts[2] == 1   # REQ-EXTRA missing in reqs
