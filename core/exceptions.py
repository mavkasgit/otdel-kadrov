"""
HRMS Custom Exceptions
Пользовательские исключения системы
"""


class HRMSException(Exception):
    """Base exception for HRMS system"""
    pass


# Data Access Errors
class DatabaseConnectionError(HRMSException):
    """Raised when cannot connect to Excel database"""
    pass


class SheetNotFoundError(HRMSException):
    """Raised when required sheet is not found in workbook"""
    pass


class DataIntegrityError(HRMSException):
    """Raised when data structure doesn't match expected format"""
    pass


# Validation Errors
class ValidationError(HRMSException):
    """Raised when business rule validation fails"""
    pass


class VacationOverlapError(ValidationError):
    """Raised when vacation dates overlap with existing vacation"""
    pass


# Document Generation Errors
class TemplateNotFoundError(HRMSException):
    """Raised when Word template file is not found"""
    pass


class TemplateMissingVariableError(HRMSException):
    """Raised when template contains variable not present in data"""
    pass


class DocumentSaveError(HRMSException):
    """Raised when cannot save generated document"""
    pass


# System Errors
class ConfigurationError(HRMSException):
    """Raised when configuration is invalid or missing"""
    pass
