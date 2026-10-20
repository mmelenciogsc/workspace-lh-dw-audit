import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parents[1] / "src"))

from audit_cleanup import analyze_workbook, annotate_column_g


class AuditCleanupTests(unittest.TestCase):
    def setUp(self):
        self.source = Path(__file__).parents[1] / "data" / "Workspace_LH_and_DW_Audit.xlsx"

    def test_workbook_selects_only_nonblank_no_rows(self):
        workbook, decisions, metadata = analyze_workbook(self.source)
        self.assertEqual(metadata["action_sheet"], "Sheet1")
        self.assertTrue(decisions)
        self.assertTrue(all(decision.keep.upper() == "NO" for decision in decisions))
        self.assertEqual(len(decisions), 43)


    def test_dependency_rows_are_not_eligible_and_annotations_go_to_g(self):
        workbook, decisions, metadata = analyze_workbook(self.source)
        annotate_column_g(workbook, decisions, metadata)
        row = next(row for row in decisions if row.item == "Datamart.Warehouse")
        self.assertTrue(row.dependency_known)
        self.assertFalse(row.eligible_for_deletion)
        sheet = workbook[metadata["action_sheet"]]
        self.assertIn("ACTION=SKIP_DEPENDENCY", sheet.cell(row=11, column=7).value)


    def test_no_dependency_rows_are_eligible_but_not_deleted_by_default(self):
        _, decisions, _ = analyze_workbook(self.source)
        eligible = [decision for decision in decisions if decision.eligible_for_deletion]
        self.assertTrue(eligible)
        self.assertTrue(all(decision.action == "ELIGIBLE_AFTER_LIVE_RESOLUTION" for decision in eligible))


if __name__ == "__main__":
    unittest.main()
