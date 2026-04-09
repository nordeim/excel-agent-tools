# 🚀 Phase 14 Execution Plan: Hardening, Security & Release Preparation
**Status:** Planned | **Dependency:** Phases 0–13 Complete | **Estimated Duration:** 3 Days

## 🎯 1. Strategic Objectives
Phase 14 transitions `excel-agent-tools` from a "validated prototype" to an "enterprise-grade release." This phase focuses on fortifying the security perimeter, optimizing critical paths, validating cross-platform compatibility, and delivering the tooling necessary for seamless AI agent integration.

| Objective | Key Deliverables | Impact |
| :--- | :--- | :--- |
| **Supply Chain Security** | CycloneDX SBOM, Sigstore keyless signing, Secret scanning | Prevents supply chain attacks; ensures artifact integrity |
| **Automated Resilience** | Hypothesis fuzzing tests for parsers & tokens | Catches edge-case vulnerabilities in reference extraction & crypto |
| **Performance Scaling** | Dependency Graph Cache, Distributed State Protocols | Reduces re-parse overhead; enables multi-agent deployments |
| **Cross-Platform CI** | macOS/Windows runners, Lock validation | Guarantees portability across enterprise environments |
| **Agent Orchestration SDK** | `excel_agent.sdk` Python client | Lowers integration barrier for LangChain/AutoGen/Custom agents |

---

## 📂 2. Detailed Task Breakdown

### 🔒 Track A: Security & Supply Chain Hardening (Day 1)
| ID | Task | Description | Deliverables |
|:--|:---|:---|:---|
| **14.1** | **SBOM Generation** | Implement `scripts/generate_sbom.py` using `cyclonedx-python-lib`. Generates CycloneDX BOM from `requirements.txt`. | `scripts/generate_sbom.py` |
| **14.2** | **Sigstore Signing** | Integrate `sigstore-python` in CI release workflow. Sign wheels/sdist with OIDC-based keyless signatures. | Update `.github/workflows/ci.yml` |
| **14.3** | **Secret Detection** | Add `detect-secrets` pre-commit hook. Scan repo for accidental secrets. | `.pre-commit-config.yaml` |
| **14.4** | **Fuzz Testing** | Add `hypothesis` property-based tests for `_extract_references_from_formula`, `ApprovalTokenManager.validate_token`, and `DependencyTracker.build_graph`. | `tests/property/test_fuzzing.py` |

### ⚡ Track B: Performance & Scalability (Day 1–2)
| ID | Task | Description | Deliverables |
|:--|:---|:---|:---|
| **14.5** | **Graph Persistence** | Add `--cache-graph` to `xls-dependency-report.py`. Cache serialized graphs in `./work/.graph_cache/` keyed by file hash. | `src/excel_agent/core/graph_cache.py` |
| **14.6** | **Distributed State Protocol** | Define `TokenStore` and `NonceStore` Protocols in `governance/stores.py`. | `src/excel_agent/governance/stores.py` |
| **14.7** | **Redis Backend Stub** | Implement `RedisNonceStore` as an optional backend (requires `redis` package). | `src/excel_agent/governance/backends/redis.py` |
| **14.8** | **Tier 2 Pooling** | Add `--pool-size` logic to `Tier2Calculator` to serialize LibreOffice calls if needed. | Update `tier2_libreoffice.py` |

### 🌍 Track C: Cross-Platform CI & Validation (Day 2)
| ID | Task | Description | Deliverables |
|:--|:---|:---|:---|
| **14.9** | **CI Matrix Expansion** | Add `macos-13` and `windows-latest` runners to `ci.yml`. | `.github/workflows/ci.yml` |
| **14.10** | **Locking Validation** | Verify `msvcrt.locking()` on Windows and `fcntl` on macOS. Add explicit `@pytest.mark.os` decorators. | Update `tests/unit/test_locking.py` |
| **14.11** | **LibreOffice Path Discovery** | Validate `Tier2Calculator` auto-detection on macOS (`/Applications/.../soffice`) and Windows. | Verify `tier2_libreoffice.py` |

### 🛠 Track D: Developer Experience & Agent SDK (Day 2–3)
| ID | Task | Description | Deliverables |
|:--|:---|:---|:---|
| **14.12** | **Agent Orchestration SDK** | Create `excel_agent.sdk.AgentClient`. Wraps `subprocess.run` with retry logic, JSON parsing, and token management. | `src/excel_agent/sdk/client.py` |
| **14.13** | **Pre-commit Config** | Generate robust `.pre-commit-config.yaml` (black, ruff, mypy, detect-secrets, markdownlint). | `.pre-commit-config.yaml` |
| **14.14** | **API Doc Scraper** | Create `scripts/generate_api_docs.py` to scrape tool docstrings and validate `docs/API.md` coverage. | `scripts/generate_api_docs.py` |
| **14.15** | **Tier 1 Workflow Docs** | Update `docs/DEVELOPMENT.md` with explicit `save → calculate → reload` workflow for Tier 1. | Update `docs/DEVELOPMENT.md` |

### 📦 Track E: Final Release Polish (Day 3)
| ID | Task | Description | Deliverables |
|:--|:---|:---|:---|
| **14.16** | **Sanity Workflow Run** | Manual execution of "Clone → Modify → Recalc → Export" on a 50k-row dataset. | Validation Report |
| **14.17** | **PyPI Metadata** | Finalize `README.md` badges, classifiers, and `CHANGELOG.md`. | `CHANGELOG.md` |
| **14.18** | **Version Bump** | Bump version to `1.0.0-rc1` or `1.0.0` depending on test results. | `src/excel_agent/__init__.py` |

---

## 📁 3. File Generation & Modification List

### New Files
1.  **`src/excel_agent/governance/stores.py`**: Protocols for pluggable token/nonce storage.
2.  **`src/excel_agent/governance/backends/__init__.py`**: Init for optional backends.
3.  **`src/excel_agent/governance/backends/redis.py`**: Redis implementation for nonces/tokens.
4.  **`src/excel_agent/core/graph_cache.py`**: File-based cache for dependency graphs.
5.  **`src/excel_agent/sdk/__init__.py`**: Init for Agent SDK.
6.  **`src/excel_agent/sdk/client.py`**: `AgentClient` wrapper class with retry/backoff.
7.  **`tests/property/test_fuzzing.py`**: Hypothesis fuzzing tests for critical parsers.
8.  **`scripts/generate_sbom.py`**: CycloneDX SBOM generator.
9.  **`.pre-commit-config.yaml`**: Pre-commit hooks configuration.

### Modified Files
1.  **`.github/workflows/ci.yml`**: Add macOS/Windows matrix, Sigstore signing, SBOM artifact.
2.  **`src/excel_agent/governance/token_manager.py`**: Inject `NonceStore` Protocol support.
3.  **`src/excel_agent/governance/audit_trail.py`**: Inject `AuditBackend` Protocol support (if not already fully abstract).
4.  **`src/excel_agent/tools/governance/xls_dependency_report.py`**: Add `--cache-graph` flag.
5.  **`docs/DEVELOPMENT.md`**: Add Tier 1 workflow warning and SDK usage examples.
6.  **`pyproject.toml`**: Add `[project.optional-dependencies]` for `redis`, `sdk`, `dev-tools`.

---

## ✅ 4. Phase 14 Exit Criteria Checklist

| # | Criterion | Validation Method |
|:--|:---|:---|
| **1** | **SBOM Generated** | `python scripts/generate_sbom.py` produces valid `bom.json` |
| **2** | **Fuzz Tests Pass** | `pytest tests/property/test_fuzzing.py --hypothesis-seed=0` runs 1000+ iterations |
| **3** | **Distributed Protocols Defined** | `TokenStore` and `NonceStore` exist in `governance/stores.py` |
| **4** | **Graph Caching Works** | `xls-dependency-report` second run is >50% faster via cache hit |
| **5** | **CI Matrix Green** | Ubuntu, macOS, and Windows runners all pass `pytest` and `mypy` |
| **6** | **Agent SDK Usable** | `AgentClient` can execute `xls-read-range` and parse result in <5 lines of code |
| **7** | **Secret Scanning Active** | Pre-commit blocks commit with hardcoded API key |
| **8** | **Documentation Updated** | `DEVELOPMENT.md` explicitly warns about Tier 1 `save → calc → reload` |
| **9** | **PyPI Ready** | `twine check dist/*` passes, metadata complete, classifiers accurate |
| **10** | **Zero Critical Vulnerabilities** | `safety check -r requirements.txt` returns 0 vulnerabilities |

---

## 📅 5. Timeline & Effort Estimation

| Day | Focus | Tasks | Effort |
|:---|:---|:---|:---|
| **Day 1** | **Security & Performance** | SBOM, Fuzzing, Graph Caching, Store Protocols | 8 hours |
| **Day 2** | **CI, SDK & Tooling** | CI Matrix, Agent SDK, Pre-commit, Redis Stub | 8 hours |
| **Day 3** | **Release Polish** | Docs, Sanity Run, PyPI Prep, Version Bump | 6 hours |
| **Total** | | | **~22 Hours (3 Days)** |

---

## 🚨 6. Risk Mitigation

| Risk | Probability | Impact | Mitigation |
|:---|:---|:---|:---|
| **Fuzzing finds critical bug** | Medium | High | Fix immediately; do not block release unless it breaks the master plan. |
| **macOS/Windows CI fails** | Medium | Medium | Use `pytest.mark.skipif` for platform-specific features (e.g., `msvcrt`) if unfixable in 2 hours. |
| **Redis dependency bloat** | Low | Low | Keep `redis` as an *optional* dependency (`pip install excel-agent-tools[redis]`). |
| **Agent SDK scope creep** | Low | Low | Limit SDK to `subprocess` wrapper only. No direct library calls in v1.0.0. |

**Next Step:** Proceed to implementation of **Track A (Security & Supply Chain)** or **Track B (Performance)** based on your priority preference.

