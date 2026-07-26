from demo.domain.review import aggregate_reviews


def test_aggregate_reviews_preserves_findings_and_counts_severity():
    result = aggregate_reviews(
        {
            "format": {"status": "completed", "findings": [{"severity": "medium"}]},
            "data": {"status": "completed", "findings": [{"severity": "high"}, {"severity": "high"}]},
            "semantic": {"status": "failed", "findings": []},
        }
    )

    assert result["status"] == "completed_with_issues"
    assert result["review_count"] == 3
    assert result["finding_count"] == 3
    assert result["severity_counts"] == {"high": 2, "medium": 1, "low": 0}
    assert result["failed_reviews"] == ["semantic"]
