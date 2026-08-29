"""Domain errors used to fail the workbook transaction closed."""

from __future__ import annotations

import re
from typing import Any


def _snake_case(name: str) -> str:
    return re.sub(r"(?<!^)(?=[A-Z])", "_", name).lower()


class PLAutomationError(RuntimeError):
    code = "pl_automation_error"

    def __init__(self, message: str, *, evidence: Any = None, phase: str | None = None):
        super().__init__(message)
        self.message = message
        self.evidence = evidence
        self.phase = phase


class ConfigurationError(PLAutomationError):
    code = "configuration_error"


class UnsupportedScopeError(PLAutomationError):
    code = "unsupported_scope"


class ExcelNotInstalledError(PLAutomationError):
    code = "excel_not_installed"


class ExcelConnectionError(PLAutomationError):
    code = "excel_connection_error"


class NoRunningExcelError(ExcelConnectionError):
    code = "no_running_excel"


class WorkbookNotFoundError(PLAutomationError):
    code = "workbook_not_found"


class WorkbookIdentityError(PLAutomationError):
    code = "workbook_identity_error"


class UnsavedWorkbookError(PLAutomationError):
    code = "unsaved_workbook"


class WorkbookReadOnlyError(PLAutomationError):
    code = "workbook_read_only"


class DRMOpenBlockedError(PLAutomationError):
    code = "drm_open_blocked"


class MissingSheetError(PLAutomationError):
    code = "missing_sheet"


class WorkbookFormatError(PLAutomationError):
    code = "workbook_format_error"


class AmbiguousMonthBlockError(PLAutomationError):
    code = "ambiguous_month_block"


class MissingMonthBlockError(PLAutomationError):
    code = "missing_month_block"


class ExistingActualColumnError(PLAutomationError):
    code = "existing_actual_column"


class CopyCreationError(PLAutomationError):
    code = "copy_creation_error"


class FormulaCloneError(PLAutomationError):
    code = "formula_clone_error"


class FormulaRewriteError(PLAutomationError):
    code = "formula_rewrite_error"


class UnsupportedFormulaError(PLAutomationError):
    code = "unsupported_formula"


class MergeRepairError(PLAutomationError):
    code = "merge_repair_error"


class TotalPLMappingError(PLAutomationError):
    code = "total_pl_mapping_error"


class CalculationTimeoutError(PLAutomationError):
    code = "calculation_timeout"


class ValidationError(PLAutomationError):
    code = "validation_error"


class SaveError(PLAutomationError):
    code = "save_error"


class ReopenValidationError(PLAutomationError):
    code = "reopen_validation_error"


class PublicationError(PLAutomationError):
    code = "publication_error"


class SourceChangedError(PLAutomationError):
    code = "source_changed"
