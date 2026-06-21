import json
import unittest

import generate_docs


EMP_ID = "\u5de5\u53f7"
PASSPORT = "\u62a4\u7167\u53f7"
DEPART_DATE = "\u51fa\u53d1\u65f6\u95f4"
RETURN_DATE = "\u8fd4\u7a0b\u65f6\u95f4"
TEMPLATE_TYPE = "\u6a21\u677f\u7c7b\u578b"
STATUS = "\u751f\u6210\u72b6\u6001"
PROJECT = "\u9879\u76ee\u7ec4"
OPERATION = "\u8fd0\u8425\u7ec4"
HR = "\u4eba\u4e8b\u7ec4"
PROJECT_SHORT = "\u9879\u76ee"


class GenerateDocsTests(unittest.TestCase):
    def test_parse_feishu_json_keeps_template_type(self):
        payload = {
            "data": {
                "fields": [EMP_ID, PASSPORT, DEPART_DATE, RETURN_DATE, TEMPLATE_TYPE, STATUS],
                "record_id_list": ["rec_1"],
                "data": [["1", "e 123", "2026-07-01", "2026-07-05", PROJECT, ""]],
            }
        }

        records = generate_docs.parse_feishu_json(json.dumps(payload, ensure_ascii=False))

        self.assertEqual(records[0]["record_id"], "rec_1")
        self.assertEqual(records[0][EMP_ID], "001")
        self.assertEqual(records[0][TEMPLATE_TYPE], PROJECT)

    def test_parse_feishu_record_items_accepts_english_template_type(self):
        payload = {
            "items": [
                {
                    "record_id": "rec_2",
                    "fields": {
                        "emp_id": "2",
                        "passport": "E00000002",
                        "depart_date": "2026-07-03",
                        "return_date": "2026-07-08",
                        "template_type": OPERATION,
                    },
                }
            ]
        }

        records = generate_docs.parse_feishu_json(json.dumps(payload, ensure_ascii=False))

        self.assertEqual(records[0][EMP_ID], "002")
        self.assertEqual(records[0][TEMPLATE_TYPE], OPERATION)

    def test_template_type_takes_priority_over_department(self):
        warnings = []
        row = {EMP_ID: "001", TEMPLATE_TYPE: HR, "department": PROJECT_SHORT}

        resolved = generate_docs.resolve_template_for_row(row, warnings)

        self.assertEqual(resolved, HR)
        self.assertEqual(warnings, [])

    def test_empty_template_type_falls_back_to_department(self):
        warnings = []
        row = {EMP_ID: "001", TEMPLATE_TYPE: "", "department": PROJECT_SHORT}

        resolved = generate_docs.resolve_template_for_row(row, warnings)

        self.assertEqual(resolved, PROJECT)
        self.assertEqual(warnings, [])

    def test_fingerprint_changes_when_template_type_changes(self):
        base_row = {
            EMP_ID: "001",
            "name": "\u5f20\u4e09",
            "department": PROJECT_SHORT,
            "position": "Engineer",
            "gender": "\u7537",
            "passport": "E00000001",
            DEPART_DATE: "2026-07-01",
            RETURN_DATE: "2026-07-05",
            TEMPLATE_TYPE: PROJECT,
        }
        changed_row = {**base_row, TEMPLATE_TYPE: OPERATION}

        self.assertNotEqual(
            generate_docs.fingerprint_for_row(base_row),
            generate_docs.fingerprint_for_row(changed_row),
        )

    def test_normalize_passport(self):
        self.assertEqual(generate_docs.normalize_passport(" e 0001 "), "E0001")
        self.assertEqual(generate_docs.normalize_passport("N/A"), "")
        self.assertEqual(generate_docs.normalize_passport(None), "")


if __name__ == "__main__":
    unittest.main()
