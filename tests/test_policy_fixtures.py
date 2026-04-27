import json
import os
import unittest

from main import BaseClassifier, Classification, EmailRecord, PolicyContext, decide_message_action


class StubClassifier(BaseClassifier):
    backend_name = "stub"
    model_name = "stub-model"

    def __init__(self, action: str, confidence: float, reason: str) -> None:
        self.action = action
        self.confidence = confidence
        self.reason = reason

    def classify(self, record: EmailRecord) -> Classification:
        return Classification(
            action=self.action,
            confidence=self.confidence,
            reason=self.reason,
            raw_response=(
                '{"action":"%s","confidence":%s,"reason":"%s"}'
                % (self.action, self.confidence, self.reason)
            ),
        )


class PolicyFixturesTest(unittest.TestCase):
    def test_policy_fixtures(self) -> None:
        fixtures_path = os.path.join(
            os.path.dirname(__file__), "..", "fixtures", "policy_fixtures.json"
        )
        with open(fixtures_path, "r", encoding="utf-8") as handle:
            cases = json.load(handle)

        for case in cases:
            with self.subTest(case=case["name"]):
                record = EmailRecord(
                    uid="1",
                    from_header=case["record"]["from_header"],
                    subject=case["record"]["subject"],
                    date=case["record"]["date"],
                    message_id=case["record"]["message_id"],
                    list_unsubscribe=case["record"]["list_unsubscribe"],
                    snippet=case["record"]["snippet"],
                )
                classifier = StubClassifier(
                    action=case["llm"]["action"],
                    confidence=float(case["llm"]["confidence"]),
                    reason=case["llm"]["reason"],
                )
                context = PolicyContext(
                    mode="aggressive",
                    folder="[Gmail]/All Mail",
                    gmail_query="category:promotions" if case["is_promotions_scan"] else None,
                    is_promotions_scan=bool(case["is_promotions_scan"]),
                    folder_default_policy="promotions-delete",
                    allowlist_emails=set(),
                    allowlist_domains=set(),
                )
                decision = decide_message_action(
                    record=record,
                    classifier=classifier,
                    context=context,
                )

                self.assertEqual(case["expected_action"], decision.final_action)
                self.assertIn(case["expected_rule_contains"], decision.rule_match)


if __name__ == "__main__":
    unittest.main()
