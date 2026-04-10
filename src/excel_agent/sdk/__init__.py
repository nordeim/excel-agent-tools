"""
Agent Orchestration SDK for excel-agent-tools.

Simplifies integration with AI agent frameworks by providing a Pythonic wrapper
around the CLI tools with retry logic, JSON parsing, and token management.

Example:
    >>> from excel_agent.sdk import AgentClient
    >>>
    >>> client = AgentClient(secret_key="your-secret")
    >>>
    >>> # Clone and modify
    >>> clone_path = client.clone("data.xlsx", output_dir="./work")
    >>>
    >>> # Read data
    >>> data = client.read_range(clone_path, "A1:C10")
    >>>
    >>> # Write data
    >>> client.write_range(clone_path, clone_path, "A1", [["New", "Data"]])
    >>>
    >>> # Recalculate
    >>> client.recalculate(clone_path, clone_path)
    >>>
    >>> # Safe structural edit with token
    >>> token = client.generate_token("sheet:delete", clone_path)
    >>> client.run("structure.xls_delete_sheet",
    ...            input=clone_path, name="OldSheet", token=token)
"""

from excel_agent.sdk.client import (
    AgentClient,
    AgentClientError,
    ImpactDeniedError,
    TokenRequiredError,
    ToolExecutionError,
    run_tool,
)

__all__ = [
    "AgentClient",
    "AgentClientError",
    "ImpactDeniedError",
    "TokenRequiredError",
    "ToolExecutionError",
    "run_tool",
]
