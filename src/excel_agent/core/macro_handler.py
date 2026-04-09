"""VBA macro analysis using oletools library.

This module provides the MacroAnalyzer class that uses oletools
to detect, inspect, and analyze VBA macros in Excel workbooks.

Key capabilities:
- Detect VBA macro presence
- List VBA modules with source code
- Check digital signature status
- Scan for suspicious patterns (auto-exec, Shell, IOCs)
- Extract VBA project for forensics
"""

from __future__ import annotations

import logging
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Protocol

logger = logging.getLogger(__name__)

# Risk patterns for macro safety scanning
SUSPICIOUS_PATTERNS = {
    "auto_exec": [
        r"AutoOpen",
        r"AutoClose",
        r"AutoExec",
        r"AutoExit",
        r"AutoNew",
        r"Document_Open",
        r"Document_Close",
        r"Workbook_Open",
        r"Workbook_Close",
        r"Workbook_Activate",
    ],
    "shell": [
        r"Shell",
        r"CreateObject",
        r"WScript\.Shell",
        r"WScript\.Exec",
        r"Process\.Start",
        r"Application\.Run",
    ],
    "network": [
        r"WinHttp",
        r"XmlHttp",
        r"InternetOpen",
        r"URLDownloadToFile",
        r"CreateObject\([\"']MSXML2\.XMLHTTP[\"']",
        r"CreateObject\([\"']WinHttp\.WinHttpRequest[\"']",
    ],
    "obfuscation": [
        r"Chr\(",
        r"ChrW\(",
        r"Asc\(",
        r"StrReverse",
        r"Replace\(",
        r"Split\(",
        r"Join\(",
        r"&H[0-9A-Fa-f]+",
    ],
}


@dataclass
class MacroModule:
    """Represents a VBA module."""

    name: str
    code: str
    is_stream: bool = False
    risk_indicators: list[str] = field(default_factory=list)


@dataclass
class MacroAnalysisResult:
    """Result of macro analysis."""

    has_macros: bool = False
    is_signed: bool = False
    signature_valid: bool = False
    module_count: int = 0
    modules: list[MacroModule] = field(default_factory=list)
    risk_score: int = 0
    risk_level: str = "none"  # none, low, medium, high
    auto_exec_functions: list[str] = field(default_factory=list)
    iocs: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary."""
        return {
            "has_macros": self.has_macros,
            "is_signed": self.is_signed,
            "signature_valid": self.signature_valid,
            "module_count": self.module_count,
            "modules": [
                {
                    "name": m.name,
                    "code_preview": m.code[:500] if m.code else "",
                    "is_stream": m.is_stream,
                    "risk_indicators": m.risk_indicators,
                }
                for m in self.modules
            ],
            "risk_score": self.risk_score,
            "risk_level": self.risk_level,
            "auto_exec_functions": self.auto_exec_functions,
            "iocs": self.iocs,
            "errors": self.errors,
        }


class MacroAnalyzer(Protocol):
    """Protocol for VBA macro analysis."""

    def analyze(self, path: Path) -> MacroAnalysisResult:
        """Analyze workbook for macros."""
        ...


class OleToolsMacroAnalyzer:
    """Macro analyzer using oletools library."""

    def __init__(self) -> None:
        self._vba_parser = None
        try:
            from oletools import olevba

            self._olevba = olevba
        except ImportError:
            logger.warning("oletools not installed, macro analysis disabled")
            self._olevba = None

    def _check_vba_presence(self, path: Path) -> bool:
        """Check if file contains VBA macros."""
        try:
            with zipfile.ZipFile(path, "r") as zf:
                for name in zf.namelist():
                    if "vba" in name.lower():
                        return True
            return False
        except zipfile.BadZipFile:
            return False

    def analyze(self, path: Path) -> MacroAnalysisResult:
        """Analyze workbook for VBA macros."""
        result = MacroAnalysisResult()

        if not path.exists():
            result.errors.append(f"File not found: {path}")
            return result

        # Check for VBA presence
        if not self._check_vba_presence(path):
            result.has_macros = False
            return result

        result.has_macros = True

        if self._olevba is None:
            result.errors.append("oletools not available for detailed analysis")
            return result

        try:
            vba = self._olevba.VBA_Parser(str(path))

            if vba.detect_vba_macros():
                # Check signature
                try:
                    signed = vba.is_signed()
                    result.is_signed = signed
                    if signed:
                        result.signature_valid = vba.verify_signature()
                except Exception as e:
                    logger.debug("Signature check failed: %s", e)

                # Extract modules
                for module in vba.extract_all_macros():
                    if isinstance(module, tuple) and len(module) >= 3:
                        module_name = module[0] if module[0] else "Unknown"
                        code = module[2] if len(module) > 2 else ""

                        macro_module = MacroModule(
                            name=module_name,
                            code=code,
                            is_stream=False,
                        )

                        # Check for risk indicators
                        self._analyze_risk(macro_module)
                        result.modules.append(macro_module)

                result.module_count = len(result.modules)

                # Calculate risk score
                result.risk_score = self._calculate_risk_score(result)
                result.risk_level = self._get_risk_level(result.risk_score)

                # Extract auto-exec functions
                result.auto_exec_functions = self._find_auto_exec(result)

            vba.close()

        except Exception as e:
            result.errors.append(f"Analysis error: {e}")
            logger.exception("Macro analysis failed")

        return result

    def _analyze_risk(self, module: MacroModule) -> None:
        """Analyze module for risk indicators."""
        code = module.code.upper()

        for category, patterns in SUSPICIOUS_PATTERNS.items():
            for pattern in patterns:
                if re.search(pattern.upper(), code):
                    module.risk_indicators.append(category)
                    break

        # Remove duplicates
        module.risk_indicators = list(set(module.risk_indicators))

    def _calculate_risk_score(self, result: MacroAnalysisResult) -> int:
        """Calculate overall risk score (0-100)."""
        score = 0

        # If no macros, no risk
        if not result.has_macros or not result.modules:
            return 0

        # Base score for having macros
        score += 10

        for module in result.modules:
            # Points per risk indicator
            score += len(module.risk_indicators) * 15

            # Auto-exec is high risk
            if "auto_exec" in module.risk_indicators:
                score += 25

            # Shell execution is high risk
            if "shell" in module.risk_indicators:
                score += 20

            # Network activity is medium risk
            if "network" in module.risk_indicators:
                score += 15

            # Obfuscation is suspicious
            if "obfuscation" in module.risk_indicators:
                score += 10

        return min(score, 100)

    def _get_risk_level(self, score: int) -> str:
        """Convert score to risk level."""
        if score == 0:
            return "none"
        elif score < 25:
            return "low"
        elif score < 50:
            return "medium"
        else:
            return "high"

    def _find_auto_exec(self, result: MacroAnalysisResult) -> list[str]:
        """Find auto-execute functions."""
        auto_exec = []
        auto_patterns = SUSPICIOUS_PATTERNS["auto_exec"]

        for module in result.modules:
            for pattern in auto_patterns:
                if re.search(pattern, module.code, re.IGNORECASE):
                    auto_exec.append(f"{module.name}: {pattern}")

        return auto_exec


def get_analyzer() -> MacroAnalyzer:
    """Get the default macro analyzer."""
    return OleToolsMacroAnalyzer()


def has_macros(path: Path) -> bool:
    """Quick check if file has VBA macros."""
    try:
        with zipfile.ZipFile(path, "r") as zf:
            for name in zf.namelist():
                if "vba" in name.lower():
                    return True
        return False
    except (zipfile.BadZipFile, IOError):
        return False
