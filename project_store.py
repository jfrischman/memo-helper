from __future__ import annotations

import json
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional


BASE_DIR = Path(__file__).resolve().parent / "memo_projects"


def ensure_base_dir() -> Path:
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    return BASE_DIR


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _safe_slug(value: str, fallback: str = "project") -> str:
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", str(value or "").strip())
    text = text.strip("._-")
    return text or fallback


def project_dir(project_id: str) -> Path:
    return ensure_base_dir() / project_id


def project_json_path(project_id: str) -> Path:
    return project_dir(project_id) / "project.json"


def list_project_summaries() -> List[Dict[str, Any]]:
    ensure_base_dir()
    summaries: List[Dict[str, Any]] = []
    for path in sorted(BASE_DIR.glob("*/project.json")):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        summaries.append(
            {
                "project_id": data.get("project_id") or path.parent.name,
                "project_name": data.get("project_name") or "Untitled project",
                "memo_name": data.get("memo_name") or "",
                "updated_at": data.get("updated_at") or "",
                "fund_count": len(data.get("funds") or []),
            }
        )
    summaries.sort(key=lambda item: (item.get("updated_at") or "", item.get("project_name") or ""), reverse=True)
    return summaries


def create_project(project_name: str = "", memo_name: str = "") -> Dict[str, Any]:
    project_id = str(uuid.uuid4())
    data = {
        "project_id": project_id,
        "project_name": project_name or "Untitled project",
        "memo_name": memo_name or "",
        "rules": "",
        "funds": [],
        "created_at": _now_iso(),
        "updated_at": _now_iso(),
    }
    save_project(data)
    return data


def load_project(project_id: str) -> Dict[str, Any]:
    path = project_json_path(project_id)
    if not path.exists():
        raise FileNotFoundError(f"Project {project_id} not found")
    data = json.loads(path.read_text(encoding="utf-8"))
    data.setdefault("project_id", project_id)
    data.setdefault("project_name", "Untitled project")
    data.setdefault("memo_name", "")
    data.setdefault("rules", "")
    data.setdefault("funds", [])
    return data


def save_project(data: Dict[str, Any]) -> Dict[str, Any]:
    ensure_base_dir()
    project_id = data.get("project_id") or str(uuid.uuid4())
    data = dict(data)
    data["project_id"] = project_id
    data.setdefault("project_name", "Untitled project")
    data.setdefault("memo_name", "")
    data.setdefault("rules", "")
    data.setdefault("funds", [])
    data["updated_at"] = _now_iso()
    if "created_at" not in data:
        data["created_at"] = data["updated_at"]
    project_dir(project_id).mkdir(parents=True, exist_ok=True)
    project_json_path(project_id).write_text(json.dumps(data, indent=2, sort_keys=True), encoding="utf-8")
    return data


def store_upload_bytes(project_id: str, upload_id: str, filename: str, data: bytes) -> str:
    uploads_dir = project_dir(project_id) / "uploads"
    uploads_dir.mkdir(parents=True, exist_ok=True)
    safe_name = _safe_slug(Path(filename).stem) + Path(filename).suffix
    path = uploads_dir / f"{upload_id}__{safe_name}"
    path.write_bytes(data)
    return str(path)


def read_upload_bytes(source_path: str, project_id: Optional[str] = None) -> bytes:
    path = Path(source_path)
    if not path.is_absolute() and project_id:
        path = project_dir(project_id) / path
    return path.read_bytes()
