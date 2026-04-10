"""
Pluggable storage protocols for governance state.

Provides Protocols for distributed token/nonce storage, enabling:
- In-memory storage (default, single-process)
- Redis storage (multi-process/orchestrator deployments)
- Custom backends for enterprise SIEM integration

Phase 14 Addition: Distributed State Management
"""

from __future__ import annotations

from typing import Protocol, runtime_checkable


@runtime_checkable
class TokenStore(Protocol):
    """Protocol for pluggable token storage backends.

    Implementations may use Redis, PostgreSQL, or other distributed stores
    for multi-agent orchestration scenarios.

    Example:
        class RedisTokenStore:
            def __init__(self, redis_url: str):
                self._redis = redis.from_url(redis_url)

            def add(self, nonce: str) -> None:
                self._redis.sadd("excel_agent:nonces", nonce)

            def __contains__(self, nonce: str) -> bool:
                return self._redis.sismember("excel_agent:nonces", nonce)

        # Usage
        manager = ApprovalTokenManager(nonce_store=RedisTokenStore("redis://localhost"))
    """

    def add(self, nonce: str) -> None:
        """Add a nonce to the store."""
        ...

    def __contains__(self, nonce: str) -> bool:
        """Check if nonce exists in store."""
        ...

    def clear(self) -> None:
        """Clear all nonces (mainly for testing)."""
        ...


@runtime_checkable
class AuditBackend(Protocol):
    """Protocol for pluggable audit logging backends.

    Default implementation writes to JSONL file.
    Alternative implementations may use Redis Streams, PostgreSQL,
    or webhook endpoints for SIEM integration.

    Phase 14: Enhanced Protocol with async support hints.
    """

    def log_event(self, event: dict[str, object]) -> None:
        """Log an audit event.

        Args:
            event: Dictionary containing audit fields:
                - timestamp: ISO 8601 UTC
                - tool: Tool name (e.g., "xls_delete_sheet")
                - scope: Token scope used
                - resource: Affected resource
                - action: Action performed
                - outcome: "success" | "denied" | "error"
                - token_used: Whether token was required
                - file_hash: Workbook geometry hash
                - pid: Process ID
                - details: Additional context
        """
        ...

    def query_events(
        self,
        *,
        tool: str | None = None,
        scope: str | None = None,
        start_time: str | None = None,
        end_time: str | None = None,
        limit: int = 100,
    ) -> list[dict[str, object]]:
        """Query audit events with filters.

        Args:
            tool: Filter by tool name
            scope: Filter by token scope
            start_time: ISO 8601 timestamp (inclusive)
            end_time: ISO 8601 timestamp (inclusive)
            limit: Maximum events to return

        Returns:
            List of event dictionaries
        """
        ...


class InMemoryTokenStore:
    """Default in-memory token store (single-process).

    Used when no distributed store is configured.
    Nonces are lost on process restart.
    """

    def __init__(self) -> None:
        self._nonces: set[str] = set()

    def add(self, nonce: str) -> None:
        self._nonces.add(nonce)

    def __contains__(self, nonce: str) -> bool:
        return nonce in self._nonces

    def clear(self) -> None:
        self._nonces.clear()


# Optional: Redis implementation stub for Phase 14
# Requires: pip install redis
# Usage:
#   from excel_agent.governance.backends.redis import RedisTokenStore
#   manager = ApprovalTokenManager(nonce_store=RedisTokenStore("redis://localhost:6379"))
