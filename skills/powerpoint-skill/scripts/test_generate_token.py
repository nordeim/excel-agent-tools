#!/usr/bin/env python3
"""Tests for generate_token.py"""

import os
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))
from generate_token import generate_token


class TestGenerateToken(unittest.TestCase):
    def test_deterministic_output(self):
        """Same scope + secret produces same token."""
        t1 = generate_token("slide:delete:2", "secret123")
        t2 = generate_token("slide:delete:2", "secret123")
        self.assertEqual(t1, t2)

    def test_different_scopes_different_tokens(self):
        """Different scopes produce different tokens."""
        t1 = generate_token("slide:delete:0", "secret")
        t2 = generate_token("slide:delete:1", "secret")
        self.assertNotEqual(t1, t2)

    def test_different_secrets_different_tokens(self):
        """Different secrets produce different tokens."""
        t1 = generate_token("slide:delete:0", "secret1")
        t2 = generate_token("slide:delete:0", "secret2")
        self.assertNotEqual(t1, t2)

    def test_token_is_hex_sha256(self):
        """Token should be 64-char hex string (SHA-256)."""
        token = generate_token("test:scope", "key")
        self.assertEqual(len(token), 64)
        self.assertTrue(all(c in "0123456789abcdef" for c in token))

    def test_known_scope_values(self):
        """Test all three documented scope patterns."""
        for scope in ["slide:delete:0", "shape:remove:0:3", "merge:presentations:2"]:
            token = generate_token(scope, "dev_secret")
            self.assertEqual(len(token), 64)


if __name__ == "__main__":
    unittest.main()
