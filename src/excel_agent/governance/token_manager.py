"""
HMAC-SHA256 scoped approval token manager for excel-agent-tools.

This module implements the governance token system that protects destructive
operations. Tokens are:
- Scoped to specific operations (sheet:delete, range:delete, etc.)
- Bound to file hashes (preventing replay across different workbooks)
- Time-limited with TTL (default 300s, max 3600s)
- Single-use via nonce tracking
- Validated with constant-time comparison (hmac.compare_digest)

The token format is: scope|file_hash|nonce|issued_at|ttl|signature
where signature = HMAC-SHA256(secret, scope|file_hash|nonce|issued_at|ttl)
"""

from __future__ import annotations

import hashlib
import hmac
import logging
import secrets
import time
from dataclasses import dataclass
from pathlib import Path

from excel_agent.utils.exceptions import PermissionDeniedError

logger = logging.getLogger(__name__)


# Valid token scopes for destructive operations
VALID_SCOPES = frozenset(
    {
        "sheet:delete",
        "sheet:rename",
        "range:delete",
        "formula:convert",
        "macro:remove",
        "macro:inject",
        "structure:modify",
    }
)

DEFAULT_TTL = 300  # 5 minutes
MAX_TTL = 3600  # 1 hour


@dataclass(frozen=True)
class ApprovalToken:
    """A validated approval token.

    Attributes:
        scope: The operation scope (e.g., "sheet:delete")
        target_file_hash: SHA-256 hash of the target workbook
        nonce: UUID4, single-use identifier
        issued_at: Unix timestamp when token was issued
        ttl_seconds: Token lifetime in seconds
        signature: HMAC-SHA256 signature
    """

    scope: str
    target_file_hash: str
    nonce: str
    issued_at: float
    ttl_seconds: int
    signature: str

    def to_string(self) -> str:
        """Serialize token to string format."""
        return (
            f"{self.scope}|"
            f"{self.target_file_hash}|"
            f"{self.nonce}|"
            f"{self.issued_at:.6f}|"
            f"{self.ttl_seconds}|"
            f"{self.signature}"
        )

    @classmethod
    def from_string(cls, token_str: str) -> ApprovalToken:
        """Parse token from string format."""
        parts = token_str.split("|")
        if len(parts) != 6:
            raise ValueError(f"Invalid token format: expected 6 parts, got {len(parts)}")

        return cls(
            scope=parts[0],
            target_file_hash=parts[1],
            nonce=parts[2],
            issued_at=float(parts[3]),
            ttl_seconds=int(parts[4]),
            signature=parts[5],
        )


class ApprovalTokenManager:
    """Manages HMAC-SHA256 scoped approval tokens for destructive operations.

    Usage:
        manager = ApprovalTokenManager(secret="your-secret-key")

        # Generate token
        token = manager.generate_token(
            scope="sheet:delete",
            target_file=Path("workbook.xlsx"),
            ttl_seconds=300,
        )

        # Validate token
        try:
            manager.validate_token(token_string, scope="sheet:delete", file_path=path)
        except PermissionDeniedError:
            print("Token invalid or expired")
    """

    def __init__(
        self,
        secret: str | None = None,
        *,
        nonce_store: object | None = None,
    ):
        """Initialize the token manager.

        Args:
            secret: HMAC secret key. If None, generates a random key (for testing).
            nonce_store: Optional store for nonce tracking (default: in-memory set).
                         Use RedisTokenStore for distributed deployments.
        """
        if secret is None:
            secret = secrets.token_hex(32)
        self._secret = secret.encode("utf-8")

        # Phase 14: Support pluggable nonce stores
        if nonce_store is None:
            self._used_nonces: set[str] = set()
            self._nonce_store = None
        else:
            self._nonce_store = nonce_store
            self._used_nonces = set()  # Unused when external store provided

    def generate_token(
        self,
        scope: str,
        target_file: Path,
        ttl_seconds: int = DEFAULT_TTL,
    ) -> str:
        """Generate a new approval token.

        Args:
            scope: Operation scope (e.g., "sheet:delete")
            target_file: Path to the target workbook
            ttl_seconds: Token lifetime (default 300, max 3600)

        Returns:
            Serialized token string.

        Raises:
            ValueError: If scope is invalid or TTL is out of range.
        """
        if scope not in VALID_SCOPES:
            raise ValueError(f"Invalid scope: {scope!r}. Valid scopes: {sorted(VALID_SCOPES)}")

        if ttl_seconds < 1 or ttl_seconds > MAX_TTL:
            raise ValueError(f"TTL must be between 1 and {MAX_TTL} seconds")

        # Compute file hash
        from excel_agent.core.version_hash import compute_file_hash

        file_hash = compute_file_hash(target_file)

        # Generate nonce
        nonce = secrets.token_hex(16)

        # Current time
        issued_at = time.time()

        # Build message to sign
        message = f"{scope}|{file_hash}|{nonce}|{issued_at:.6f}|{ttl_seconds}"

        # Sign with HMAC-SHA256
        signature = hmac.new(
            self._secret,
            message.encode("utf-8"),
            hashlib.sha256,
        ).hexdigest()

        token = ApprovalToken(
            scope=scope,
            target_file_hash=file_hash,
            nonce=nonce,
            issued_at=issued_at,
            ttl_seconds=ttl_seconds,
            signature=signature,
        )

        logger.debug(
            "Generated token: scope=%s, file=%s, nonce=%s...",
            scope,
            target_file.name,
            nonce[:8],
        )

        return token.to_string()

    def validate_token(
        self,
        token_string: str | None,
        scope: str,
        file_path: Path,
    ) -> None:
        """Validate a token for a specific operation.

        Args:
            token_string: The token to validate (from generate_token).
            scope: Expected operation scope.
            file_path: Path to the target workbook.

        Raises:
            PermissionDeniedError: If token is invalid, expired, wrong scope,
                or already used.
        """
        if token_string is None:
            raise PermissionDeniedError(
                "No approval token provided",
                reason="missing_token",
            )

        try:
            token = ApprovalToken.from_string(token_string)
        except (ValueError, IndexError) as e:
            raise PermissionDeniedError(
                f"Invalid token format: {e}",
                reason="malformed_token",
            ) from e

        # Check scope
        if token.scope != scope:
            raise PermissionDeniedError(
                f"Token scope mismatch: expected {scope!r}, got {token.scope!r}",
                reason="scope_mismatch",
            )

        # Check expiration
        now = time.time()
        if now > token.issued_at + token.ttl_seconds:
            raise PermissionDeniedError(
                f"Token expired at {token.issued_at + token.ttl_seconds}",
                reason="expired",
            )

        # Check file hash
        from excel_agent.core.version_hash import compute_file_hash

        current_hash = compute_file_hash(file_path)
        if token.target_file_hash != current_hash:
            raise PermissionDeniedError(
                "Token bound to different file hash (file was modified)",
                reason="file_hash_mismatch",
            )

        # Check nonce (single-use) - Phase 14: Support external store
        if self._nonce_store is not None:
            nonce_used = token.nonce in self._nonce_store
        else:
            nonce_used = token.nonce in self._used_nonces

        if nonce_used:
            raise PermissionDeniedError(
                "Token already used (replay detected)",
                reason="replay",
            )

        # Verify signature (constant-time comparison)
        message = (
            f"{token.scope}|{token.target_file_hash}|{token.nonce}|"
            f"{token.issued_at:.6f}|{token.ttl_seconds}"
        )
        expected_sig = hmac.new(
            self._secret,
            message.encode("utf-8"),
            hashlib.sha256,
        ).hexdigest()

        if not hmac.compare_digest(token.signature, expected_sig):
            raise PermissionDeniedError(
                "Invalid token signature",
                reason="invalid_signature",
            )

        # Mark nonce as used - Phase 14: Support external store
        if self._nonce_store is not None:
            self._nonce_store.add(token.nonce)
        else:
            self._used_nonces.add(token.nonce)

        logger.debug(
            "Token validated: scope=%s, file=%s, nonce=%s...",
            scope,
            file_path.name,
            token.nonce[:8],
        )

    def clear_used_nonces(self) -> None:
        """Clear the used nonces set (for testing).

        Phase 14: Also clears external store if configured.
        """
        self._used_nonces.clear()
        if self._nonce_store is not None:
            self._nonce_store.clear()
