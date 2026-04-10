"""
Agent Orchestration SDK for excel-agent-tools.

Provides a Python client that wraps subprocess calls with retry logic,
JSON parsing, and token management to simplify AI agent integrations.
"""

from __future__ import annotations

import json
import logging
import subprocess
import time
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


class AgentClientError(Exception):
    """Base exception for AgentClient errors."""

    def __init__(self, message: str, exit_code: int | None = None, response: dict | None = None):
        super().__init__(message)
        self.exit_code = exit_code
        self.response = response


class ToolExecutionError(AgentClientError):
    """Tool execution failed."""

    pass


class TokenRequiredError(AgentClientError):
    """Approval token required but not provided."""

    pass


class ImpactDeniedError(AgentClientError):
    """Operation denied due to impact on formulas."""

    def __init__(self, message: str, guidance: str, impact: dict, **kwargs):
        super().__init__(message, **kwargs)
        self.guidance = guidance
        self.impact = impact


class AgentClient:
    """Python client for excel-agent-tools CLI.

    Wraps subprocess.run() with retry logic, JSON parsing, and token management.
    Simplifies integration with LangChain, AutoGen, and other agent frameworks.

    Usage:
        client = AgentClient(secret_key="your-secret")

        # Simple read operation
        result = client.run("read.xls_read_range",
                           input="data.xlsx",
                           range="A1:C10")

        # Token-gated operation
        token = client.generate_token("sheet:delete", "data.xlsx")
        result = client.run("structure.xls_delete_sheet",
                           input="data.xlsx",
                           name="OldSheet",
                           token=token)

        # Automatic retry on lock contention
        result = client.run("write.xls_write_range",
                           input="data.xlsx",
                           range="A1",
                           data=[["Value"]],
                           max_retries=3)
    """

    def __init__(
        self,
        secret_key: str | None = None,
        python_executable: str | None = None,
        timeout: int = 45,
        base_retry_delay: float = 0.5,
    ):
        """Initialize the AgentClient.

        Args:
            secret_key: EXCEL_AGENT_SECRET for token generation.
            python_executable: Path to Python executable (default: sys.executable).
            timeout: Default timeout for tool execution in seconds.
            base_retry_delay: Base delay for exponential backoff on retries.
        """
        self.secret_key = secret_key
        self.python_executable = python_executable or __import__("sys").executable
        self.timeout = timeout
        self.base_retry_delay = base_retry_delay
        self._used_nonces: set[str] = set()

    def run(
        self,
        tool: str,
        *,
        max_retries: int = 3,
        retry_on: list[int] | None = None,
        **kwargs: Any,
    ) -> dict[str, Any]:
        """Execute a CLI tool with automatic retry logic.

        Args:
            tool: Tool module path (e.g., "read.xls_read_range").
            max_retries: Maximum number of retry attempts.
            retry_on: List of exit codes to retry on (default: [3] for lock contention).
            **kwargs: Tool arguments as keyword args (e.g., input="file.xlsx").

        Returns:
            Parsed JSON response from the tool.

        Raises:
            ToolExecutionError: If tool fails after all retries.
            ImpactDeniedError: If operation denied due to formula impact.
            TokenRequiredError: If token required but missing.
        """
        if retry_on is None:
            retry_on = [3]  # Default: retry on lock contention

        cmd = self._build_command(tool, **kwargs)
        last_error = None

        for attempt in range(max_retries):
            try:
                return self._execute(cmd, tool)
            except AgentClientError as e:
                last_error = e

                # Don't retry on certain errors
                if e.exit_code not in retry_on:
                    raise

                if attempt < max_retries - 1:
                    delay = self.base_retry_delay * (2**attempt)
                    logger.warning(
                        "Tool %s failed (exit %s), retrying in %.1fs (attempt %d/%d)",
                        tool,
                        e.exit_code,
                        delay,
                        attempt + 1,
                        max_retries,
                    )
                    time.sleep(delay)

        raise ToolExecutionError(
            f"Tool {tool} failed after {max_retries} attempts: {last_error}",
            exit_code=last_error.exit_code if last_error else None,
        )

    def _build_command(self, tool: str, **kwargs: Any) -> list[str]:
        """Build subprocess command from tool and arguments."""
        cmd = [self.python_executable, "-m", f"excel_agent.tools.{tool}"]

        for key, value in kwargs.items():
            if value is None:
                continue

            # Convert snake_case to kebab-case for CLI
            cli_key = key.replace("_", "-")

            if isinstance(value, bool):
                if value:
                    cmd.append(f"--{cli_key}")
            elif isinstance(value, (list, dict)):
                # JSON encode complex types
                cmd.extend([f"--{cli_key}", json.dumps(value)])
            else:
                cmd.extend([f"--{cli_key}", str(value)])

        return cmd

    def _execute(self, cmd: list[str], tool: str) -> dict[str, Any]:
        """Execute command and parse response."""
        env = __import__("os").environ.copy()
        if self.secret_key:
            env["EXCEL_AGENT_SECRET"] = self.secret_key

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=self.timeout,
            env=env,
        )

        # Parse JSON output
        stdout = result.stdout.strip()
        if not stdout:
            raise ToolExecutionError(
                f"Tool {tool} produced no output",
                exit_code=result.returncode,
            )

        try:
            response = json.loads(stdout)
        except json.JSONDecodeError as e:
            raise ToolExecutionError(
                f"Invalid JSON from {tool}: {e}",
                exit_code=result.returncode,
            )

        # Handle different response statuses
        status = response.get("status")
        exit_code = response.get("exit_code", result.returncode)

        if status == "denied":
            guidance = response.get("guidance", "")
            impact = response.get("impact", {})
            raise ImpactDeniedError(
                response.get("denial_reason", "Operation denied"),
                guidance=guidance,
                impact=impact,
                exit_code=exit_code,
                response=response,
            )

        if status == "error":
            error_msg = response.get("error", "Unknown error")
            if "token" in error_msg.lower() or exit_code == 4:
                raise TokenRequiredError(error_msg, exit_code=exit_code, response=response)
            raise ToolExecutionError(error_msg, exit_code=exit_code, response=response)

        if exit_code != 0:
            raise ToolExecutionError(
                f"Tool {tool} exited with code {exit_code}",
                exit_code=exit_code,
                response=response,
            )

        return response

    def generate_token(
        self,
        scope: str,
        file_path: str | Path,
        ttl_seconds: int = 300,
    ) -> str:
        """Generate an approval token for destructive operations.

        Args:
            scope: Token scope (e.g., "sheet:delete", "range:delete").
            file_path: Path to the target workbook.
            ttl_seconds: Token lifetime in seconds (default: 300).

        Returns:
            Token string for use with token-gated tools.

        Raises:
            TokenRequiredError: If secret_key not configured.
        """
        if not self.secret_key:
            raise TokenRequiredError("EXCEL_AGENT_SECRET required for token generation")

        result = self.run(
            "governance.xls_approve_token",
            scope=scope,
            file=str(file_path),
            ttl=ttl_seconds,
            max_retries=1,  # Don't retry token generation
        )

        return result["data"]["token"]

    def clone(
        self,
        input_path: str | Path,
        output_dir: str | Path | None = None,
    ) -> Path:
        """Clone a workbook to a safe working copy.

        Args:
            input_path: Source workbook path.
            output_dir: Destination directory (default: current directory).

        Returns:
            Path to the cloned workbook.
        """
        kwargs: dict[str, Any] = {"input": str(input_path)}
        if output_dir:
            kwargs["output_dir"] = str(output_dir)

        result = self.run("governance.xls_clone_workbook", **kwargs)
        return Path(result["data"]["clone_path"])

    def read_range(
        self,
        input_path: str | Path,
        range_str: str,
        sheet: str | None = None,
    ) -> list[list[Any]]:
        """Read data from a cell range.

        Args:
            input_path: Workbook path.
            range_str: Cell range (e.g., "A1:C10").
            sheet: Sheet name (default: active sheet).

        Returns:
            2D array of cell values.
        """
        kwargs: dict[str, Any] = {"input": str(input_path), "range": range_str}
        if sheet:
            kwargs["sheet"] = sheet

        result = self.run("read.xls_read_range", **kwargs)
        return result["data"]["values"]

    def write_range(
        self,
        input_path: str | Path,
        output_path: str | Path,
        range_str: str,
        data: list[list[Any]],
        sheet: str | None = None,
    ) -> dict[str, Any]:
        """Write data to a cell range.

        Args:
            input_path: Source workbook path.
            output_path: Destination workbook path.
            range_str: Target cell range (e.g., "A1").
            data: 2D array of values to write.
            sheet: Sheet name (default: active sheet).

        Returns:
            Tool response with impact metrics.
        """
        kwargs: dict[str, Any] = {
            "input": str(input_path),
            "output": str(output_path),
            "range": range_str,
            "data": data,
        }
        if sheet:
            kwargs["sheet"] = sheet

        return self.run("write.xls_write_range", **kwargs)

    def recalculate(
        self,
        input_path: str | Path,
        output_path: str | Path,
        *,
        tier: int | None = None,
        circular: bool = False,
    ) -> dict[str, Any]:
        """Recalculate all formulas in a workbook.

        Args:
            input_path: Source workbook path.
            output_path: Destination workbook path.
            tier: Force specific tier (1 or 2, default: auto).
            circular: Enable circular reference support.

        Returns:
            Calculation result with stats.
        """
        kwargs: dict[str, Any] = {
            "input": str(input_path),
            "output": str(output_path),
            "circular": circular,
        }
        if tier:
            kwargs["tier"] = tier

        return self.run("formulas.xls_recalculate", **kwargs)


# Convenience function for quick usage
def run_tool(tool: str, **kwargs: Any) -> dict[str, Any]:
    """Execute a tool with default settings (no retries, no secret).

    Args:
        tool: Tool module path.
        **kwargs: Tool arguments.

    Returns:
        Parsed JSON response.
    """
    client = AgentClient()
    return client.run(tool, max_retries=1, **kwargs)
