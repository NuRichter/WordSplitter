"""Language neutral errors.

Every user facing failure carries a translation key and named parameters
instead of a finished sentence. The worker process sends the key across the
process boundary and the GUI renders it in the language the user picked, so a
message never appears half in one language and half in another.
"""

from __future__ import annotations

from typing import Any, Dict, Optional


class AppError(Exception):
    """A failure that is explained to the user through a locale key."""

    def __init__(self, key: str, detail: Optional[str] = None, **params: Any) -> None:
        super().__init__(key)
        self.key = key
        self.params: Dict[str, Any] = params
        self.detail = detail  # technical text, shown only under "Details"

    def to_dict(self) -> Dict[str, Any]:
        return {"key": self.key, "params": self.params, "detail": self.detail}

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "AppError":
        return cls(str(data.get("key") or "err.unexpected"), data.get("detail"), **(data.get("params") or {}))

    def __str__(self) -> str:  # useful in logs and tests
        return f"{self.key} {self.params}" + (f" | {self.detail}" if self.detail else "")


class ValidationError(AppError):
    """Raised when user supplied data violates a documented rule."""


class OperationCancelled(BaseException):
    """The user cancelled the running task.

    Derived from BaseException on purpose: the many defensive
    `except Exception` blocks around individual COM calls must not swallow a
    cancellation and carry on with the next step.
    """
