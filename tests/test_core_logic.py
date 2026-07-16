import os
import tempfile
import unittest

from core.logic import (
    _format_rule_decision_log,
    filter_vorlagen_nach_auswahl,
    sammle_vorlagen_pfade,
)


class TemplateDiscoveryTests(unittest.TestCase):
    def test_sammle_vorlagen_pfade_filters_expected_template_folders(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            paths = [
                os.path.join(temp_dir, "anlagen", "anlage.docx"),
                os.path.join(temp_dir, "allgemein", "deckblatt.docx"),
                os.path.join(temp_dir, "sonstiges", "ignore.docx"),
                os.path.join(temp_dir, "anlagen", "~$lock.docx"),
                os.path.join(temp_dir, "anlagen", "ignore.txt"),
            ]
            for path in paths:
                os.makedirs(os.path.dirname(path), exist_ok=True)
                with open(path, "w", encoding="utf-8"):
                    pass

            anlagen, allgemein = sammle_vorlagen_pfade(temp_dir)

            self.assertEqual([os.path.basename(p) for p in anlagen], ["anlage.docx"])
            self.assertEqual([os.path.basename(p) for p in allgemein], ["deckblatt.docx"])

    def test_filter_vorlagen_nach_auswahl_preserves_all_when_selection_is_none(self):
        anlagen = ["C:/templates/anlagen/a.docx"]
        allgemein = ["C:/templates/allgemein/b.docx"]

        filtered_anlagen, filtered_allgemein = filter_vorlagen_nach_auswahl(anlagen, allgemein, None)

        self.assertEqual(filtered_anlagen, anlagen)
        self.assertEqual(filtered_allgemein, allgemein)

    def test_filter_vorlagen_nach_auswahl_uses_normalized_paths(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            anlage_path = os.path.join(temp_dir, "anlagen", "..", "anlagen", "a.docx")
            allgemein_path = os.path.join(temp_dir, "allgemein", "b.docx")
            selected = [os.path.normpath(anlage_path)]

            filtered_anlagen, filtered_allgemein = filter_vorlagen_nach_auswahl(
                [anlage_path],
                [allgemein_path],
                selected,
            )

            self.assertEqual(filtered_anlagen, [anlage_path])
            self.assertEqual(filtered_allgemein, [])


class RuleDecisionLogTests(unittest.TestCase):
    def test_format_rule_decision_log_handles_missing_rule(self):
        message = _format_rule_decision_log(
            {"has_rule": False, "reason": "keine Bedingung"},
            {"anlage_seriennummer": "SN-1"},
            "C:/templates/anlagen/a.docx",
        )

        self.assertIn("SN-1", message)
        self.assertIn("a.docx", message)
        self.assertIn("keine aktive Regel", message)


if __name__ == "__main__":
    unittest.main()
