"""
Protocols for dependency injection.
"""

from typing import Protocol


class LoggerProtocol(Protocol):
    """Protocol for logger objects to enable dependency injection."""
    
    def debug(self, msg: str, *args, **kwargs) -> None:
        """Log debug message."""
        ...
    
    def info(self, msg: str, *args, **kwargs) -> None:
        """Log info message."""
        ...
    
    def warning(self, msg: str, *args, **kwargs) -> None:
        """Log warning message."""
        ...
    
    def error(self, msg: str, *args, **kwargs) -> None:
        """Log error message."""
        ...
    
    def exception(self, msg: str, *args, **kwargs) -> None:
        """Log exception with traceback."""
        ...
