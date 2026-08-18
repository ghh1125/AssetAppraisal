from demo.domain.mapping_agent import mapping_agent_request, validate_mapping_agent_response


def test_mapping_agent_can_only_map_an_unresolved_location_to_an_existing_field():
    request = mapping_agent_request(
        [{"location_id": "P1-X1", "context": "被评估单位 XXX", "marker": "XXX"}],
        ["target_company_name"],
    )

    mapped, issues = validate_mapping_agent_response(
        {"mappings": {"P1-X1": "target_company_name"}},
        unresolved_location_ids=[item["location_id"] for item in request["locations"]],
        allowed_fields=request["allowed_field_keys"],
    )

    assert mapped == {"P1-X1": "target_company_name"}
    assert issues == []


def test_mapping_agent_cannot_invent_fields_or_override_confirmed_location():
    mapped, issues = validate_mapping_agent_response(
        {"mappings": {"P1-X1": "invented", "P2-X1": "target_company_name"}},
        unresolved_location_ids=["P1-X1"],
        allowed_fields=["target_company_name"],
    )

    assert mapped == {}
    assert len(issues) == 2
