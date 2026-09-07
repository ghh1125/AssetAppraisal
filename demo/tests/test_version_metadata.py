import re
import tomllib
from pathlib import Path

from demo import __version__


ROOT = Path(__file__).resolve().parents[2]
VERSION_HEADING = re.compile(r"^## (\d+\.\d+\.\d+) — \d{4}-\d{2}-\d{2}(?: .*)?$", re.MULTILINE)


def test_package_and_changelog_versions_are_aligned() -> None:
    project = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    changelog = (ROOT / "demo" / "CHANGELOG.md").read_text(encoding="utf-8")
    releases = VERSION_HEADING.findall(changelog)

    assert releases, "CHANGELOG 必须至少包含一个带日期的发布版本"
    assert len(releases) == len(set(releases)), "CHANGELOG 不能包含重复版本标题"
    assert project["project"]["version"] == __version__ == releases[0]


def test_changelog_keeps_an_unreleased_section_before_releases() -> None:
    changelog = (ROOT / "demo" / "CHANGELOG.md").read_text(encoding="utf-8")
    first_release = VERSION_HEADING.search(changelog)

    assert first_release is not None
    assert changelog.index("## Unreleased") < first_release.start()
