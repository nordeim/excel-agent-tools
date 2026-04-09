# Development Guide: excel-agent-tools

**Version:** 1.0.0

This guide covers local setup, CI/CD, and adding new tools.

---

## Quick Start

```bash
# Clone and setup
git clone <repo>
cd excel-agent-tools
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -e ".[dev]"
pre-commit install

# Run tests
pytest

# Run specific test
pytest tests/integration/test_clone_modify_workflow.py -v
```

---

## Project Structure

```
src/excel_agent/
├── core/           # ExcelAgent, FileLock, RangeSerializer
├── governance/     # Tokens, AuditTrail
├── calculation/    # Tier 1/2 engines
└── tools/          # 53 CLI tools
    ├── governance/
    ├── read/
    ├── write/
    ├── structure/
    ├── cells/
    ├── formulas/
    ├── objects/
    ├── formatting/
    ├── macros/
    └── export/
```

---

## Code Standards

### Formatting

```bash
black src/ tests/ --line-length 99
ruff check src/ tests/
mypy src/ --strict
```

### Type Hints

- Enable strict mode
- Use `interface` for structures
- Use `type` for unions
- Never use `any`

---

## Adding a Tool

1. Create tool file: `src/excel_agent/tools/<category>/xls_<name>.py`
2. Register entry point in `pyproject.toml`
3. Add stub `_stub_main()` for testability
4. Write unit tests
5. Write integration tests
6. Update `docs/API.md`

---

## Testing

```bash
# Unit tests
pytest tests/unit/ -v

# Integration tests
pytest tests/integration/ -v --tb=short

# With coverage
pytest --cov=excel_agent --cov-report=html

# Specific marker
pytest -m "libreoffice"
```

---

## CI/CD

GitHub Actions runs on Python 3.12/3.13 with:
- Linting (black, ruff, mypy)
- Unit tests
- Integration tests
- Coverage gate (≥90%)

---

## Dependencies

See `requirements.txt` and `requirements-dev.txt`. Key packages:
- `openpyxl>=3.1.5`
- `formulas[excel]>=1.3.4`
- `defusedxml>=0.7.1`
- `oletools>=0.60.2`

---

**Maintained by:** Development Team
**License:** MIT
