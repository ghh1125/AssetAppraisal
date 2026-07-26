from __future__ import annotations

from collections import Counter
from typing import Any


def aggregate_reviews(reviews: dict[str, dict[str, Any]]) -> dict[str, Any]:
    findings = [
        finding
        for review in reviews.values()
        for finding in review.get("findings", [])
    ]
    severity_counts = Counter(str(item.get("severity", "low")) for item in findings)
    failed_reviews = [
        name for name, review in reviews.items() if review.get("status") == "failed"
    ]
    return {
        "status": "completed_with_issues" if findings or failed_reviews else "completed",
        "review_count": len(reviews),
        "finding_count": len(findings),
        "severity_counts": {
            "high": severity_counts.get("high", 0),
            "medium": severity_counts.get("medium", 0),
            "low": severity_counts.get("low", 0),
        },
        "failed_reviews": failed_reviews,
        "findings": findings,
    }
