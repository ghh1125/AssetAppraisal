from __future__ import annotations

from enum import StrEnum
from typing import Any, Iterable

from pydantic import BaseModel, ConfigDict, Field


class RouteKind(StrEnum):
    """黄色位置唯一允许使用的数据来源。"""

    PDF_OCR_XLSX = "pdf_ocr_xlsx"
    QICHACHA_API = "qichacha_api"
    BAILIAN_GLM = "bailian_glm"
    NODE_INPUT = "node_input"


class YellowRoute(BaseModel):
    """一个黄色位置到业务字段及唯一来源的不可变契约。"""

    model_config = ConfigDict(extra="forbid", frozen=True)

    location_id: str = Field(description="Word 黄色位置编号")
    field_key: str = Field(description="标准业务字段键")
    route_kind: RouteKind = Field(description="该字段唯一允许的数据来源")
    replacement_mode: str = Field(
        default="replace_paragraph",
        description="黄色内容替换方式；用于保留模板原有段落开头、字体和括号结构",
    )


def load_yellow_routes(items: Iterable[dict[str, Any]]) -> list[YellowRoute]:
    routes = [YellowRoute.model_validate(item) for item in items]
    location_ids = [route.location_id for route in routes]
    field_keys = [route.field_key for route in routes]
    if len(location_ids) != len(set(location_ids)):
        raise ValueError("黄色位置编号重复")
    if len(field_keys) != len(set(field_keys)):
        raise ValueError("黄色业务字段重复")
    return routes


def validate_yellow_routes(
    routes: Iterable[YellowRoute], *, expected_location_ids: set[str]
) -> None:
    actual = {route.location_id for route in routes}
    if actual != expected_location_ids:
        missing = sorted(expected_location_ids - actual)
        extra = sorted(actual - expected_location_ids)
        raise ValueError(f"黄色路由与模板位置不一致：缺少={missing}，多余={extra}")


def fields_for_route(routes: Iterable[YellowRoute], route_kind: RouteKind) -> set[str]:
    return {route.field_key for route in routes if route.route_kind == route_kind}
