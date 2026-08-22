"""
QGIS Function Registry — Catalog of verified PyQGIS API functions.

Tracks which PyQGIS API calls have been verified against a real QGIS
instance, in which scripts they were used, and provides search/filter
capabilities for the agent. Same JSON-file convention as SAP2000's
scripts/registry.json.
"""

import json
import logging
from datetime import datetime, timezone
from pathlib import Path

logger = logging.getLogger(__name__)

# registry.json lives in scripts/ at the connector root
REGISTRY_PATH = Path(__file__).resolve().parent.parent / "scripts" / "registry.json"

_EMPTY_REGISTRY = {
    "version": "1.0",
    "functions": {},
    "summary": {
        "total_registered": 0,
        "total_verified": 0,
        "last_updated": None,
    },
}


class FunctionRegistry:
    """Manages a JSON file cataloging verified PyQGIS API functions."""

    def __init__(self, registry_path: Path | None = None):
        self._path = registry_path or REGISTRY_PATH
        self._data: dict | None = None
        self._last_mtime: float = 0

    def _reload_if_changed(self) -> None:
        if self._data is None:
            return
        try:
            disk_mtime = self._path.stat().st_mtime
            if disk_mtime > self._last_mtime:
                self._data = None
        except OSError:
            pass

    def _load(self) -> dict:
        self._reload_if_changed()
        if self._data is not None:
            return self._data
        try:
            self._data = json.loads(self._path.read_text(encoding="utf-8"))
        except (FileNotFoundError, json.JSONDecodeError):
            self._data = json.loads(json.dumps(_EMPTY_REGISTRY))
        self._touch_mtime()
        return self._data

    def _touch_mtime(self) -> None:
        try:
            self._last_mtime = self._path.stat().st_mtime
        except OSError:
            self._last_mtime = 0

    def _save(self) -> None:
        data = self._load()
        data["summary"]["total_registered"] = len(data["functions"])
        data["summary"]["total_verified"] = sum(
            1 for f in data["functions"].values() if f.get("verified")
        )
        data["summary"]["last_updated"] = datetime.now(timezone.utc).isoformat()
        self._path.parent.mkdir(parents=True, exist_ok=True)
        self._path.write_text(json.dumps(data, indent=2), encoding="utf-8")
        self._touch_mtime()

    # ── Registration ──────────────────────────────────────────────────

    def register_function(
        self,
        function_path: str,
        category: str,
        description: str = "",
        signature: str = "",
        parameter_notes: str = "",
        notes: str = "",
    ) -> dict:
        data = self._load()
        is_new = function_path not in data["functions"]
        entry = data["functions"].get(function_path, {})
        entry.update(
            {
                "function_path": function_path,
                "category": category,
                "description": description or entry.get("description", ""),
                "signature": signature or entry.get("signature", ""),
                "parameter_notes": parameter_notes or entry.get("parameter_notes", ""),
                "notes": notes or entry.get("notes", ""),
            }
        )
        data["functions"][function_path] = entry
        self._save()
        return {"registered": True, "function_path": function_path, "is_new": is_new}

    def mark_verified(self, function_path: str) -> dict:
        data = self._load()
        if function_path not in data["functions"]:
            return {"verified": False, "error": f"Unknown function: {function_path}"}
        data["functions"][function_path]["verified"] = True
        data["functions"][function_path]["verified_at"] = datetime.now(timezone.utc).isoformat()
        self._save()
        return {"verified": True, "function_path": function_path}

    # ── Query ─────────────────────────────────────────────────────────

    def get_summary(self) -> dict:
        return dict(self._load()["summary"])

    def get_function(self, function_path: str) -> dict | None:
        return self._load()["functions"].get(function_path)

    def list_functions(
        self,
        category: str | None = None,
        verified_only: bool = False,
        query: str | None = None,
    ) -> list[dict]:
        data = self._load()
        out: list[dict] = []
        q = query.lower() if query else None
        for fp, entry in sorted(data["functions"].items()):
            if category and entry.get("category", "").lower() != category.lower():
                continue
            if verified_only and not entry.get("verified"):
                continue
            if q:
                haystack = " ".join(
                    [fp, entry.get("description", ""), entry.get("signature", "")]
                ).lower()
                if q not in haystack:
                    continue
            out.append({k: v for k, v in entry.items()})
        return out

    def get_categories(self) -> dict[str, dict[str, int]]:
        data = self._load()
        cats: dict[str, dict[str, int]] = {}
        for entry in data["functions"].values():
            cat = entry.get("category", "Uncategorized")
            counts = cats.setdefault(cat, {"registered": 0, "verified": 0})
            counts["registered"] += 1
            if entry.get("verified"):
                counts["verified"] += 1
        return cats


# Module-level singleton
registry = FunctionRegistry()
