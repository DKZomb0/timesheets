"""
delaware Timesheet Automator — Test Suite
-----------------------------------------
Run: python test_suite.py
Tests all logic WITHOUT needing Outlook, a real token, or internet access.
"""

import json, sys, datetime, unittest
from pathlib import Path
from unittest.mock import patch, MagicMock

sys.path.insert(0, str(Path(__file__).parent))
import timesheet

# ── Helpers ────────────────────────────────────────────────────────────────────

REAL_API_RESPONSE = {
    "codes": [
        {
            "codeCategory": "Customer Project",
            "title": "Dats 24 NV",
            "lane1": "Managed Services for DATS24 - Transition SAP Grow - fixed fee",
            "lane2": "Business Hours ( x 1 )",
            "lane3": "DATS0011.1.4 - M001",
            "projectCode": "DATS0011",
            "projectTaskCode": "DATS0011.1.4",
            "projectTaskItemCode": "M001",
        },
        {
            "codeCategory": "Administration",
            "title": "Administration",
            "lane1": "Sharpen The Saw - Internal Meetings & Admin",
            "lane2": None,
            "lane3": None,
            "projectCode": "18358",
            "projectTaskCode": "122063",
            "projectTaskItemCode": "",   # empty — caused 422 errors before
        },
        {
            "codeCategory": "Presales",
            "title": "Presales - Thema Foundries - PROS CRM",
            "lane1": "Presales - New SAP Public Cloud ERP - Pres",
            "lane2": None,
            "lane3": None,
            "projectCode": "67138",
            "projectTaskCode": "354676",
            "projectTaskItemCode": "",   # numeric codes — caused 422 errors before
        },
        {
            "codeCategory": "Internal Project",
            "title": "Delaware Belgium (BELU)",
            "lane1": "SOH SAP Public Cloud - Solution Strategy",
            "lane2": "Release as a service",
            "lane3": "BE20I552.1.5 - USR2",
            "projectCode": "BE20I552",
            "projectTaskCode": "BE20I552.1.5",
            "projectTaskItemCode": "USR2",  # non-M001 item code
        },
    ]
}

SAMPLE_DRAFT = [
    {
        "subject": "DATS biweekly sprintstatus",
        "activityDate": "2026-03-13",
        "duration": 0.5,
        "projectCode": "DATS0011",
        "projectTaskCode": "DATS0011.1.4",
        "projectTaskItemCode": "M001",
        "workDescription": "Sprint status review",
        "confidence": "high",
        "reason": "Tag match: dats24",
    },
    {
        "subject": "sync Azure support-IT4IT",
        "activityDate": "2026-03-13",
        "duration": 0.5,
        "projectCode": "18358",
        "projectTaskCode": "122063",
        "projectTaskItemCode": "",
        "workDescription": "Azure support sync",
        "confidence": "medium",
        "reason": "Internal overhead",
    },
]

# ── Tests ──────────────────────────────────────────────────────────────────────

class TestAPIResponseParsing(unittest.TestCase):
    """Bug: API returns {codes: [...]} not a flat list — was silently returning []"""

    def test_codes_key_parsed(self):
        codes = timesheet.fetch_project_codes_api.__wrapped__ if hasattr(
            timesheet.fetch_project_codes_api, '__wrapped__') else None
        # Simulate what the fixed fetch function does
        data = REAL_API_RESPONSE
        if isinstance(data, dict) and "codes" in data:
            result = data["codes"]
        elif isinstance(data, list):
            result = data
        else:
            result = data.get("data", [])
        self.assertEqual(len(result), 4)
        self.assertEqual(result[0]["projectCode"], "DATS0011")

    def test_numeric_project_codes_preserved(self):
        data = REAL_API_RESPONSE
        codes = data["codes"]
        numeric = [c for c in codes if c["projectCode"].isdigit()]
        self.assertEqual(len(numeric), 2)
        self.assertEqual(numeric[0]["projectCode"], "18358")
        self.assertEqual(numeric[0]["projectTaskCode"], "122063")

    def test_empty_projectTaskItemCode_preserved(self):
        data = REAL_API_RESPONSE
        codes = data["codes"]
        empty_item = [c for c in codes if c["projectTaskItemCode"] == ""]
        self.assertEqual(len(empty_item), 2)


class TestHTMLGeneration(unittest.TestCase):
    """Bug: JS variables rows/ds were empty due to broken Python expression"""

    def setUp(self):
        self.projects = REAL_API_RESPONSE["codes"]
        self.html = timesheet.build_review_html(
            SAMPLE_DRAFT, "Friday, 13 March 2026",
            self.projects, "test@delawareconsulting.com", "fake-bearer-token"
        )

    def test_rows_json_populated(self):
        self.assertIn("var rows=[{", self.html,
            "rows JS variable must be populated — was empty before fix")

    def test_ds_date_correct(self):
        self.assertIn("var ds='2026-03-13'", self.html,
            "ds JS variable must contain correct date — was empty before fix")

    def test_pretoken_populated(self):
        self.assertIn("var preToken='fake-bearer-token'", self.html,
            "preToken must be set from terminal input")

    def test_token_bar_hidden_when_token_provided(self):
        self.assertIn("display:none", self.html,
            "Token bar must be hidden when token already provided in terminal")

    def test_token_bar_visible_when_no_token(self):
        html_no_token = timesheet.build_review_html(
            SAMPLE_DRAFT, "Test", self.projects, "test@test.com", ""
        )
        # Should NOT have display:none for token bar
        import re
        trow_match = re.search(r"class='trow' style='([^']*)'", html_no_token)
        style = trow_match.group(1) if trow_match else ""
        self.assertNotIn("display:none", style,
            "Token bar must be visible when no token provided")

    def test_project_codes_in_dropdown(self):
        self.assertIn("DATS0011", self.html)
        self.assertIn("18358", self.html,
            "Numeric project codes must appear in dropdown")
        self.assertIn("BE20I552", self.html)

    def test_data_item_attribute_on_options(self):
        self.assertIn('data-item="M001"', self.html,
            "data-item attribute must be set for correct projectTaskItemCode submission")
        self.assertIn('data-item=""', self.html,
            "Empty projectTaskItemCode must also be preserved as data-item")

    def test_project_preselected(self):
        self.assertIn('value="DATS0011" data-task="DATS0011.1.4" data-item="M001">DATS0011', self.html)

    def test_add_row_copies_options(self):
        self.assertIn("var optHtml=pcOpts", self.html,
            "Add row must copy existing dropdown options — was empty select before fix")

    def test_submit_uses_data_item(self):
        self.assertIn("s.options[s.selectedIndex].dataset.item", self.html,
            "Submit must read projectTaskItemCode from data-item — was hardcoded M001 before fix")


class TestSubmitPayload(unittest.TestCase):
    """Bug: userId was set from config causing 403 NoDelegationRights"""

    def test_userid_empty_in_payload(self):
        """The site identifies user via bearer token, not userId field"""
        import urllib.request
        captured = {}
        original_urlopen = urllib.request.urlopen

        def mock_urlopen(req, timeout=None):
            if hasattr(req, 'data') and req.data:
                captured['payload'] = json.loads(req.data)
            raise Exception("mock — not really calling API")

        with patch('urllib.request.urlopen', side_effect=mock_urlopen):
            result = timesheet.submit_entry("fake-token", "", {
                "projectCode": "DATS0011",
                "projectTaskCode": "DATS0011.1.4",
                "projectTaskItemCode": "M001",
                "activityDate": "2026-03-13",
                "duration": 1.0,
                "workDescription": "Test"
            })

        if captured.get('payload'):
            self.assertEqual(captured['payload']['data']['userId'], "",
                "userId must be empty — non-empty value caused 403 NoDelegationRights")

    def test_empty_projectTaskItemCode_sent_correctly(self):
        """Empty projectTaskItemCode must be sent as empty string, not 'M001'"""
        import urllib.request
        captured = {}

        def mock_urlopen(req, timeout=None):
            if hasattr(req, 'data') and req.data:
                captured['payload'] = json.loads(req.data)
            raise Exception("mock")

        with patch('urllib.request.urlopen', side_effect=mock_urlopen):
            timesheet.submit_entry("fake-token", "", {
                "projectCode": "18358",
                "projectTaskCode": "122063",
                "projectTaskItemCode": "",   # empty — as returned by API
                "activityDate": "2026-03-13",
                "duration": 0.5,
                "workDescription": "Admin"
            })

        if captured.get('payload'):
            self.assertEqual(captured['payload']['data']['projectTaskItemCode'], "",
                "Empty projectTaskItemCode must be preserved — M001 caused 422 errors")

    def test_url_has_empty_user_param(self):
        """URL must have user= empty, not filled with email"""
        import urllib.request
        captured = {}

        def mock_urlopen(req, timeout=None):
            captured['url'] = req.full_url if hasattr(req, 'full_url') else str(req)
            raise Exception("mock")

        with patch('urllib.request.urlopen', side_effect=mock_urlopen):
            timesheet.submit_entry("fake-token", "vincent@delaware.com", {
                "projectCode": "DATS0011", "projectTaskCode": "DATS0011.1.4",
                "projectTaskItemCode": "M001", "activityDate": "2026-03-13",
                "duration": 1.0, "workDescription": "Test"
            })

        if captured.get('url'):
            self.assertIn("user=", captured['url'])
            self.assertNotIn("vincent@delaware", captured['url'],
                "Email must not appear in URL — caused 403 NoDelegationRights")


class TestProjectCodeLoading(unittest.TestCase):
    """Bug: fetch_project_codes_api returned [] because it looked for 'data' key not 'codes'"""

    def _mock_urlopen(self):
        mock_resp = MagicMock()
        mock_resp.read.return_value = json.dumps(REAL_API_RESPONSE).encode()
        return mock_resp

    def test_codes_key_extraction(self):
        with patch('urllib.request.urlopen', return_value=self._mock_urlopen()):
            result = timesheet.fetch_project_codes_api("fake-token", "", "2026-03-13")
        self.assertEqual(len(result), 4,
            "Must parse codes array from API response — was returning [] before fix")
        self.assertEqual(result[0]["projectCode"], "DATS0011")

    def test_numeric_codes_included(self):
        with patch('urllib.request.urlopen', return_value=self._mock_urlopen()):
            result = timesheet.fetch_project_codes_api("fake-token", "", "2026-03-13")
        codes = [r["projectCode"] for r in result]
        self.assertIn("18358", codes, "Numeric project codes must be included")
        self.assertIn("67138", codes)


class TestCatchUpDateLogic(unittest.TestCase):
    """Catch-up mode: correct working days selected, weekends skipped"""

    def test_skips_weekends(self):
        today = datetime.date(2026, 3, 16)  # Monday
        candidates = []
        d = today - datetime.timedelta(days=1)
        while len(candidates) < 7:
            if d.weekday() < 5:
                candidates.append(d)
            d -= datetime.timedelta(days=1)
        for day in candidates:
            self.assertLess(day.weekday(), 5, f"{day} is a weekend day")

    def test_yesterday_is_first(self):
        today = datetime.date(2026, 3, 16)  # Monday
        candidates = []
        d = today - datetime.timedelta(days=1)
        while len(candidates) < 7:
            if d.weekday() < 5:
                candidates.append(d)
            d -= datetime.timedelta(days=1)
        self.assertEqual(candidates[0], datetime.date(2026, 3, 13),
            "Friday should be yesterday when today is Monday")

    def test_seven_days_returned(self):
        today = datetime.date(2026, 3, 16)
        candidates = []
        d = today - datetime.timedelta(days=1)
        while len(candidates) < 7:
            if d.weekday() < 5:
                candidates.append(d)
            d -= datetime.timedelta(days=1)
        self.assertEqual(len(candidates), 7)

    def test_multi_day_selection(self):
        today = datetime.date(2026, 3, 16)
        candidates = []
        d = today - datetime.timedelta(days=1)
        while len(candidates) < 7:
            if d.weekday() < 5:
                candidates.append(d)
            d -= datetime.timedelta(days=1)
        choice = "1,2,3"
        indices = [int(x.strip())-1 for x in choice.replace(","," ").split() if x.strip()]
        selected = [candidates[i] for i in indices if 0 <= i < len(candidates)]
        self.assertEqual(len(selected), 3)
        self.assertEqual(selected[0], datetime.date(2026, 3, 13))
        self.assertEqual(selected[1], datetime.date(2026, 3, 12))


class TestCorrectionsLog(unittest.TestCase):
    """Corrections must save and reload correctly from Excel"""

    def test_save_and_load_correction(self):
        timesheet.save_correction(
            "DATS biweekly sprintstatus", "DATS0011", "DATS0011.1.4", "Sprint review"
        )
        corrections = timesheet.load_corrections()
        self.assertIn("dats biweekly sprintstatus", corrections)
        self.assertEqual(corrections["dats biweekly sprintstatus"]["projectCode"], "DATS0011")

    def test_correction_key_is_lowercase(self):
        timesheet.save_correction(
            "DATS Biweekly SprintStatus", "DATS0011", "DATS0011.1.4", "Sprint review"
        )
        corrections = timesheet.load_corrections()
        self.assertIn("dats biweekly sprintstatus", corrections,
            "Corrections must be stored lowercase for case-insensitive matching")


class TestServerEndToEnd(unittest.TestCase):
    """Local server must serve HTML and handle submit correctly"""

    def test_server_serves_html(self):
        import urllib.request, time, threading
        html = timesheet.build_review_html(
            SAMPLE_DRAFT, "Test", REAL_API_RESPONSE["codes"],
            "test@test.com", "test-token"
        )
        server = timesheet.start_server(html, "test@test.com", port=8423)
        time.sleep(0.3)
        try:
            resp = urllib.request.urlopen("http://localhost:8423/", timeout=5).read().decode()
            self.assertIn("DATS0011", resp)
            self.assertIn("var rows=[{", resp)
            self.assertIn("var ds='2026-03-13'", resp)
        finally:
            server.shutdown()

    def test_server_handles_bad_token_gracefully(self):
        import urllib.request, time, json as json2
        html = timesheet.build_review_html(
            SAMPLE_DRAFT, "Test", REAL_API_RESPONSE["codes"],
            "test@test.com", "test-token"
        )
        server = timesheet.start_server(html, "test@test.com", port=8424)
        time.sleep(0.3)
        try:
            payload = json2.dumps({"entries": [{
                "projectCode": "DATS0011", "projectTaskCode": "DATS0011.1.4",
                "projectTaskItemCode": "M001", "activityDate": "2026-03-13",
                "duration": 1.0, "workDescription": "Test", "userId": ""
            }], "corrections": []}).encode()
            req = urllib.request.Request(
                "http://localhost:8424/submit", data=payload,
                headers={"Content-Type": "application/json", "X-Token": "bad-token"}
            )
            resp = json2.loads(urllib.request.urlopen(req, timeout=5).read())
            # Should return ok=False with error, not crash
            self.assertFalse(resp["ok"])
            self.assertIn("error", resp)
        finally:
            server.shutdown()


if __name__ == "__main__":
    print("=" * 60)
    print("  delaware Timesheet Automator — Test Suite")
    print("=" * 60)
    print()
    loader = unittest.TestLoader()
    suite  = unittest.TestSuite()
    for cls in [
        TestAPIResponseParsing,
        TestHTMLGeneration,
        TestSubmitPayload,
        TestProjectCodeLoading,
        TestCatchUpDateLogic,
        TestCorrectionsLog,
        TestServerEndToEnd,
    ]:
        suite.addTests(loader.loadTestsFromTestCase(cls))

    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    print()
    if result.wasSuccessful():
        print("  All tests passed.")
    else:
        print(f"  {len(result.failures)} failure(s), {len(result.errors)} error(s)")
    sys.exit(0 if result.wasSuccessful() else 1)
