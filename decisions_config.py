
import os
import json
import asyncio
import time
from openai import AsyncOpenAI


class TransactionDecisions:

    def __init__(self):
        self.client = AsyncOpenAI(
            api_key=os.environ["OPENAI_API_KEY"],
            timeout=60.0,
            max_retries=2
        )

    async def classify(self, transactions, rules):
        if not transactions:
            return json.dumps({"results": []})

        # IDs are used as choices, so ensure they are
        # unique, valid, and cannot collide with NONE.
        rule_map = {}
        for rule in rules:
            rid = str(rule["rule_id"])
            if rid == "NONE" or rid in rule_map:
                raise ValueError(
                    f"Invalid or duplicate rule ID: {rid}"
                )
            rule_map[rid] = rule

        choices = list(rule_map) + ["NONE"]

        instructions = (
            "Choose the single best matching rule. "
            "Apply only the provided classification rules. "
            "Use NONE when no rule is established or "
            "required evidence is missing. "
            "Do not infer missing supplier, tax, location "
            "or currency information."
        )

        evidence = json.dumps(
            {"rules": rules, "transactions": transactions},
            ensure_ascii=False,
            default=str
        )

        questions = []
        for index, txn in enumerate(transactions):
            questions.append({
                "type": "choice",
                "name": f"txn_{index}",
                "instructions": (
                    f"For transaction index {index}, "
                    f"{instructions}"
                ),
                "choices": choices
            })

        start = time.perf_counter()

        decision = await self.client.decisions.create(
            model="gpt-6-luna",
            input=evidence,
            questions=questions
        )

        if len(decision.answers) != len(transactions):
            raise ValueError(
                "Decision answer count mismatch"
            )

        results = []

        for index, (txn, answer) in enumerate(
            zip(transactions, decision.answers)
        ):
            if answer.name != f"txn_{index}":
                raise ValueError(
                    "Decision answer order mismatch"
                )

            if answer.type == "refusal":
                raise ValueError(
                    f"Decision refused for transaction {index}"
                )

            if answer.type != "choice":
                raise ValueError(
                    "Unexpected decision answer type"
                )

            rid = answer.choice

            if rid not in choices:
                raise ValueError(
                    f"Unknown rule returned: {rid}"
                )

            rule = rule_map.get(rid)

            results.append({
                "id": str(txn["transaction_id"]),
                "target_classification": (
                    rule["target_classification"]
                    if rule else None
                ),
                "reasoning": (
                    f"Selected {rid}: "
                    f"{rule.get('rule_definition', '')}"
                    if rule
                    else "No rule established from evidence"
                ),
                "applied_rule_id": (
                    rid if rule else None
                ),
                "confidence_score": (
                    round(answer.confidence, 4)
                    if rule else None
                )
            })

        elapsed = time.perf_counter() - start

        # Decisions API currently bills input tokens.
        # Check regional/long-context premiums separately.
        estimated_cost = (
            decision.usage.input_tokens
            / 1_000_000 * 0.10
        )

        return json.dumps({
            "results": results
        }, ensure_ascii=False)

    async def close(self):
        await self.client.close()
