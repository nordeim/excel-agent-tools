"""
excel-agent-tools: 53 governance-first CLI tools for AI agents to safely
read, write, and manipulate Excel workbooks.

Headless, JSON-native, zero Excel dependency.

Phase 14 Complete: Agent SDK, Distributed State Protocols, Pre-commit Hooks
"""

__version__ = "1.0.0"

# Convenience imports — available after Phase 1+ implementations.
# These are lazy-imported to avoid ImportError during scaffolding.
__all__ = [
    "ApprovalTokenManager",
    "AuditTrail",
    "DependencyTracker",
    "ExcelAgent",
    "__version__",
]


def __getattr__(name: str) -> object:
    """Lazy imports for core classes — avoids import errors during scaffolding."""
    if name == "ExcelAgent":
        from excel_agent.core.agent import ExcelAgent  # type: ignore[import-untyped]

        return ExcelAgent
    if name == "DependencyTracker":
        from excel_agent.core.dependency import DependencyTracker  # type: ignore[import-untyped]

        return DependencyTracker
    if name == "ApprovalTokenManager":
        from excel_agent.governance.token_manager import (
            ApprovalTokenManager,  # type: ignore[import-untyped]
        )

        return ApprovalTokenManager
    if name == "AuditTrail":
        from excel_agent.governance.audit_trail import AuditTrail  # type: ignore[import-untyped]

        return AuditTrail
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
