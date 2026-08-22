"""
QGIS API Documentation Search — Keyword search across the QGIS PyQGIS API.

Builds an in-memory index from a curated API reference directory (API/*.md,
same convention as SAP2000). Each top-level heading (# ClassName.method)
becomes one searchable section with signature, parameters, and examples.

If the API/ directory is absent, the index is empty and search tools
return guidance pointing at the live-discovery tools instead
(list_processing_algorithms, get_algorithm_help) — graceful degradation.
"""

import re
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# API/ lives at the connector root, one level up from officemcp/
API_DIR = Path(__file__).resolve().parent.parent / "API"

# Pre-defined category mapping (filename stem → category label)
CATEGORIES = {
    "Word_Application": "Word Application Object",
    "Word_Document": "Word Document Object",
    "Word_Selection_Range": "Word Selection and Range",
    "Excel_Application": "Excel Application Object",
    "Excel_Workbook": "Excel Workbook Object",
    "Excel_Worksheet": "Excel Worksheet Object",
    "Excel_Range": "Excel Range Object",
    "PowerPoint_Application": "PowerPoint Application Object",
    "PowerPoint_Presentation": "PowerPoint Presentation Object",
    "PowerPoint_Slide": "PowerPoint Slide Object",
    "OOXML_Python": "OOXML via Python (python-docx/openpyxl/python-pptx)",
}


class DocIndex:
    """
    In-memory index of all PyQGIS documentation sections.

    Each section corresponds to a top-level heading (# ClassName.method) and
    contains the full text under that heading until the next one.
    """

    def __init__(self):
        self._sections: list[dict] = []
        self._loaded = False
        self._api_dir_mtime: float | None = None

    def _reload_if_changed(self) -> None:
        try:
            current_mtime = max(
                (f.stat().st_mtime for f in API_DIR.glob("*.md")),
                default=0,
            )
            if self._api_dir_mtime is None or current_mtime != self._api_dir_mtime:
                self._loaded = False
                self._api_dir_mtime = current_mtime
        except Exception:
            pass

    def _load(self) -> None:
        """Parse all API/*.md files into sections."""
        self._reload_if_changed()
        if self._loaded:
            return

        if not API_DIR.is_dir():
            logger.warning("API documentation directory not found: %s", API_DIR)
            self._loaded = True
            return

        for md_file in sorted(API_DIR.glob("*.md")):
            category = CATEGORIES.get(md_file.stem, md_file.stem)
            self._parse_file(md_file, category)

        logger.info("Indexed %d PyQGIS doc sections.", len(self._sections))
        self._loaded = True

    def _parse_file(self, file_path: Path, category: str) -> None:
        """Split a markdown file into sections by top-level heading."""
        try:
            content = file_path.read_text(encoding="utf-8")
        except Exception:
            logger.warning("Could not read %s", file_path)
            return

        parts = re.split(r"^(?=# )", content, flags=re.MULTILINE)

        for part in parts:
            part = part.strip()
            if not part:
                continue

            first_line = part.split("\n", 1)[0].strip()
            function_name = first_line[2:].strip() if first_line.startswith("# ") else first_line

            syntax = ""
            syntax_match = re.search(
                r"##\s*Syntax\s*\n+(.+?)(?:\n\n|\n##)", part, re.DOTALL
            )
            if syntax_match:
                syntax = syntax_match.group(1).strip()

            description = ""
            desc_match = re.search(
                r"##\s*(?:Description|Summary)\s*\n+(.+?)(?:\n\n|\n##)", part, re.DOTALL
            )
            if desc_match:
                description = re.sub(r"\s+", " ", desc_match.group(1)).strip()[:500]

            example = ""
            example_match = re.search(
                r"```\w*\n(.+?)```", part, re.DOTALL
            )
            if example_match:
                example = example_match.group(1).strip()

            self._sections.append(
                {
                    "file": file_path.stem,
                    "category": category,
                    "function_name": function_name,
                    "syntax": syntax,
                    "description": description,
                    "example_snippet": example,
                    "text": re.sub(r"\s+", " ", part).lower(),
                }
            )

    def search(self, query: str, category: str | None = None, max_results: int = 10) -> list[dict]:
        """Keyword-scored search across indexed sections.

        Scoring: term frequency of each query word in the section text.
        All query words must appear (AND); results ranked by total hits.
        """
        self._load()
        terms = [t.lower() for t in re.findall(r"\w+", query) if len(t) > 2]
        if not terms:
            return []

        scored: list[tuple[int, dict]] = []
        for section in self._sections:
            if category and section["category"].lower() != category.lower():
                continue
            text = section["text"]
            total = 0
            missing = False
            for term in terms:
                count = text.count(term)
                if count == 0:
                    missing = True
                    break
                total += count
            if missing or total == 0:
                continue
            scored.append((total, section))

        scored.sort(key=lambda x: x[0], reverse=True)

        results: list[dict] = []
        for _, item in scored[:max_results]:
            results.append({k: v for k, v in item.items() if k != "text"})
        return results

    def list_categories(self) -> list[dict]:
        """Return available categories with section counts."""
        self._load()
        cat_counts: dict[str, int] = {}
        for section in self._sections:
            cat = section["category"]
            cat_counts[cat] = cat_counts.get(cat, 0) + 1
        return [{"category": k, "sections": v} for k, v in sorted(cat_counts.items())]


# Module-level singleton
doc_index = DocIndex()
