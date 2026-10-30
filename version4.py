
def check_duplicate_transaction_ids(transactions):
    transaction_ids = [
        str(txn.get("transaction_id", "")).strip()
        for txn in transactions
    ]

    # Reject missing IDs.
    if any(not txn_id for txn_id in transaction_ids):
        return {
            "error": "INVALID_TRANSACTION_ID",
            "message": "Every transaction must have a transaction_id."
        }

    # Check duplicates within the submitted payload.
    counts = Counter(transaction_ids)

    payload_duplicates = {
        txn_id
        for txn_id, count in counts.items()
        if count > 1
    }

    # Check transactions already submitted in any batch.
    existing_ids = set(
        ResponseMapping.objects.filter(
            transaction_id__in=transaction_ids
        ).values_list("transaction_id", flat=True)
    )

    if payload_duplicates or existing_ids:
        return {
            "error": "DUPLICATE_TRANSACTION",
            "message": "Transaction already processed or submitted.",
            "duplicates_in_payload": sorted(payload_duplicates),
            "already_submitted": sorted(existing_ids),
        }

    return None
