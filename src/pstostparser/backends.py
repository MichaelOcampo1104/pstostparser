"""Backend abstraction + auto-detection for PST/OST parsing.

Three backends:
  - aspose  (easy pip install, handles PST and OST, no Outlook needed)
  - outlook (reads through Outlook/MAPI exactly as Outlook sees it;
             Windows + Outlook required; best for files Aspose under-reads)
  - pypff   (free/open-source libyal bindings; needs a working build)

Both .pst and .ost share the same underlying format; Aspose's
PersonalStorage.from_file() opens either one transparently.
"""

from __future__ import annotations

import os
import sys


def detect_backends() -> dict[str, bool]:
    out = {"aspose": False, "outlook": False, "pypff": False}
    try:
        import aspose.email  # noqa: F401
        out["aspose"] = True
    except Exception:
        pass
    if sys.platform == "win32":
        try:
            import win32com.client  # noqa: F401
            out["outlook"] = True
        except Exception:
            pass
    try:
        import pypff  # noqa: F401
        # NOTE: PyPI's `pypff` 1.x is an unrelated astronomy package, not the
        # libyal bindings — only count it if it exposes the expected API.
        if not hasattr(pypff, "file"):
            out["pypff"] = False
        else:
            out["pypff"] = True
    except Exception:
        pass
    return out


BACKEND_CHOICES = ("auto", "aspose", "outlook", "pypff")


def resolve_backend(requested: str = "auto") -> str:
    """Return 'aspose', 'outlook' or 'pypff' or raise RuntimeError with install hints."""
    avail = detect_backends()
    if requested == "auto":
        if avail["aspose"]:
            return "aspose"
        if avail["outlook"]:
            return "outlook"
        if avail["pypff"]:
            return "pypff"
        raise RuntimeError(
            "No PST backend installed. Install one of:\n"
            "  uv pip install 'Aspose.Email-for-Python-via-NET'   # easiest on Windows (evaluation limits may apply)\n"
            "  uv pip install pywin32   # Outlook COM backend (needs Outlook)\n"
            "  uv pip install libpff-python  # free/open-source libyal bindings"
        )
    if requested not in ("aspose", "outlook", "pypff"):
        raise ValueError(f"Unknown backend {requested!r}; choose auto/aspose/outlook/pypff")
    if not avail[requested]:
        hint = {"aspose": "uv pip install 'Aspose.Email-for-Python-via-NET'",
                "outlook": "uv pip install pywin32 (Windows + Outlook required)",
                "pypff": "uv pip install libpff-python"}[requested]
        raise RuntimeError(f"Backend {requested!r} is not available. Run: {hint}")
    return requested


def guess_format(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".ost":
        return "ost"
    if ext == ".pst":
        return "pst"
    return "unknown"
