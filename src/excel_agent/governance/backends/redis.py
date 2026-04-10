"""
Redis-based backends for distributed governance state.

Requires: pip install redis

Usage:
    from excel_agent.governance.backends.redis import (
        RedisTokenStore,
        RedisAuditBackend,
    )

    # Distributed nonce tracking across agent clusters
    token_store = RedisTokenStore("redis://localhost:6379")
    manager = ApprovalTokenManager(
        secret="my-secret",
        nonce_store=token_store
    )

    # Centralized audit logging
    audit = AuditTrail(backend=RedisAuditBackend("redis://localhost:6379"))

Phase 14: Optional distributed state backend
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from redis import Redis

from excel_agent.governance.stores import TokenStore


class RedisTokenStore(TokenStore):
    """Redis-backed token nonce store for multi-agent deployments.

    Prevents replay attacks across distributed agent orchestrators.
    Nonces are stored in a Redis Set with configurable TTL.

    Args:
        redis_url: Redis connection URL (e.g., "redis://localhost:6379/0")
        key_prefix: Redis key prefix (default: "excel_agent:nonce")
        ttl_seconds: Nonce expiration time (default: 3600 = 1 hour)
    """

    def __init__(
        self,
        redis_url: str = "redis://localhost:6379/0",
        *,
        key_prefix: str = "excel_agent:nonce",
        ttl_seconds: int = 3600,
    ):
        try:
            import redis
        except ImportError as exc:
            raise ImportError(
                "Redis backend requires 'redis' package. Install with: pip install redis"
            ) from exc

        self._redis: Redis = redis.from_url(redis_url)
        self._key_prefix = key_prefix
        self._ttl_seconds = ttl_seconds
        self._key = f"{key_prefix}:used"

    def add(self, nonce: str) -> None:
        """Add nonce to Redis Set with TTL."""
        self._redis.sadd(self._key, nonce)
        self._redis.expire(self._key, self._ttl_seconds)

    def __contains__(self, nonce: str) -> bool:
        """Check if nonce exists in Redis Set."""
        return bool(self._redis.sismember(self._key, nonce))

    def clear(self) -> None:
        """Clear all nonces (use with caution)."""
        self._redis.delete(self._key)


class RedisAuditBackend:
    """Redis Streams-based audit logging backend.

    Publishes audit events to a Redis Stream for centralized logging.
    Enables real-time audit trail consumption by SIEM systems.

    Args:
        redis_url: Redis connection URL
        stream_key: Redis stream key (default: "excel_agent:audit")
        maxlen: Maximum stream length (default: 10000)
    """

    def __init__(
        self,
        redis_url: str = "redis://localhost:6379/0",
        *,
        stream_key: str = "excel_agent:audit",
        maxlen: int = 10000,
    ):
        try:
            import redis
        except ImportError as exc:
            raise ImportError(
                "Redis backend requires 'redis' package. Install with: pip install redis"
            ) from exc

        self._redis: Redis = redis.from_url(redis_url)
        self._stream_key = stream_key
        self._maxlen = maxlen

    def log_event(self, event: dict[str, object]) -> None:
        """Log event to Redis Stream.

        Event is JSON-serialized and added to stream with maxlen limit.
        """
        event_json = json.dumps(event, default=str)
        self._redis.xadd(
            self._stream_key,
            {"event": event_json},
            maxlen=self._maxlen,
        )

    def query_events(
        self,
        *,
        tool: str | None = None,
        scope: str | None = None,
        start_time: str | None = None,
        end_time: str | None = None,
        limit: int = 100,
    ) -> list[dict[str, object]]:
        """Query audit events from Redis Stream.

        Note: Full stream scan required; use sparingly.
        """
        # Read from beginning
        entries = self._redis.xrange(self._stream_key, count=limit)

        events: list[dict[str, object]] = []
        for _entry_id, fields in entries:
            event_json = fields.get("event", "{}")
            try:
                event = json.loads(event_json)
            except json.JSONDecodeError:
                continue

            # Apply filters
            if tool and event.get("tool") != tool:
                continue
            if scope and event.get("scope") != scope:
                continue
            if start_time and event.get("timestamp", "") < start_time:
                continue
            if end_time and event.get("timestamp", "") > end_time:
                continue

            events.append(event)

            if len(events) >= limit:
                break

        return events


# Convenience factory
def create_redis_backends(
    redis_url: str = "redis://localhost:6379/0",
) -> tuple[RedisTokenStore, RedisAuditBackend]:
    """Create both Redis backends with shared connection.

    Returns:
        Tuple of (token_store, audit_backend)
    """
    return (
        RedisTokenStore(redis_url),
        RedisAuditBackend(redis_url),
    )
