from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, Optional
import uuid


@dataclass
class ChangeSet:
    id: str
    kind: str  # text|table|document
    prompt: str
    before: Any
    after: Any
    diff: Dict[str, Any] = field(default_factory=dict)
    status: str = "draft"  # draft|previewed|approved|rejected|applied|failed
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())
    updated_at: str = field(default_factory=lambda: datetime.now().isoformat())


class SessionStore:
    def __init__(self):
        self._changesets: Dict[str, ChangeSet] = {}

    def create(self, kind: str, prompt: str, before: Any, after: Any, diff: Optional[Dict[str, Any]] = None) -> ChangeSet:
        cid = str(uuid.uuid4())
        cs = ChangeSet(
            id=cid,
            kind=kind,
            prompt=prompt,
            before=before,
            after=after,
            diff=diff or {},
        )
        self._changesets[cid] = cs
        return cs

    def get(self, changeset_id: str) -> Optional[ChangeSet]:
        return self._changesets.get(changeset_id)

    def update_status(self, changeset_id: str, status: str) -> Optional[ChangeSet]:
        cs = self._changesets.get(changeset_id)
        if not cs:
            return None
        cs.status = status
        cs.updated_at = datetime.now().isoformat()
        return cs
