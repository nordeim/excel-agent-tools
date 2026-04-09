#!/usr/bin/env python3
"""Generate HMAC approval tokens for destructive PowerPoint operations."""

import argparse
import hashlib
import hmac
import os
import sys


def generate_token(scope: str, secret: str) -> str:
    """Generate HMAC-SHA256 approval token."""
    return hmac.new(secret.encode(), scope.encode(), hashlib.sha256).hexdigest()


def main():
    parser = argparse.ArgumentParser(
        description="Generate approval tokens for destructive PowerPoint operations"
    )
    parser.add_argument(
        "--scope",
        required=True,
        help="Token scope (e.g., 'slide:delete:2', 'shape:remove:0:3', 'merge:presentations:2')",
    )
    parser.add_argument(
        "--secret",
        default=None,
        help="Secret key (default: PPT_APPROVAL_SECRET env var, then 'dev_secret')",
    )
    parser.add_argument(
        "--quiet", action="store_true", help="Output only the token (no description)"
    )
    args = parser.parse_args()

    secret = args.secret or os.getenv("PPT_APPROVAL_SECRET", "dev_secret")
    token = generate_token(args.scope, secret)

    if args.quiet:
        print(token)
    else:
        print(f"Scope: {args.scope}")
        print(f"Token: {token}")
        print(f"\nUsage:")
        print(
            f'  uv run tools/ppt_*.py --file work.pptx ... --approval-token "{token}" --json'
        )


if __name__ == "__main__":
    main()
