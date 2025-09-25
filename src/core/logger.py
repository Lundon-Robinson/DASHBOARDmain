"""
Advanced Logging System
=======================

Centralized logging with file rotation, structured logging,
and integration with external monitoring systems.
"""

import os
import sys
import logging
import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional, Any, Dict
from loguru import logger as loguru_logger

class DashboardLogger:
    """Custom logger with enhanced features"""
    
    def __init__(self, name: str = "dashboard", log_dir: Optional[Path] = None):
        self.name = name
        self.log_dir = log_dir or Path("logs")
        self.log_dir.mkdir(parents=True, exist_ok=True)
        
        # Configure loguru
        self._setup_loguru()
        
        # Setup standard logging integration
        self._setup_standard_logging()
    
    def _setup_loguru(self):
        """Configure loguru logger"""
        # Remove default handler
        loguru_logger.remove()
        
        # Console handler with colors
        loguru_logger.add(
            sys.stderr,
            format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>",
            level="INFO"
        )
        
        # File handler for all logs
        loguru_logger.add(
            self.log_dir / "dashboard.log",
            format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
            level="DEBUG",
            rotation="50 MB",
            retention="30 days",
            compression="zip"
        )
        
        # Error file handler
        loguru_logger.add(
            self.log_dir / "errors.log", 
            format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
            level="ERROR",
            rotation="10 MB",
            retention="60 days"
        )
        
        # JSON structured logs
        loguru_logger.add(
            self.log_dir / "dashboard.json",
            format="{message}",
            level="INFO",
            rotation="100 MB",
            retention="90 days",
            serialize=True
        )
    
    def _setup_standard_logging(self):
        """Setup integration with standard logging module"""
        class InterceptHandler(logging.Handler):
            def emit(self, record):
                # Get corresponding Loguru level if it exists
                try:
                    level = loguru_logger.level(record.levelname).name
                except ValueError:
                    level = record.levelno
                
                # Find caller from where originated the logged message
                frame, depth = logging.currentframe(), 2
                while frame.f_code.co_filename == logging.__file__:
                    frame = frame.f_back
                    depth += 1
                
                loguru_logger.opt(depth=depth, exception=record.exc_info).log(level, record.getMessage())
        
        # Intercept standard logging
        logging.basicConfig(handlers=[InterceptHandler()], level=0)
    
    def info(self, message: str, **kwargs):
        """Log info message"""
        loguru_logger.info(message, **kwargs)
    
    def debug(self, message: str, **kwargs):
        """Log debug message"""
        loguru_logger.debug(message, **kwargs)
    
    def warning(self, message: str, **kwargs):
        """Log warning message"""
        loguru_logger.warning(message, **kwargs)
    
    def error(self, message: str, exception: Optional[Exception] = None, **kwargs):
        """Log error message"""
        if exception:
            loguru_logger.error(f"{message}: {exception}", **kwargs)
            loguru_logger.error(f"Traceback: {traceback.format_exc()}")
        else:
            loguru_logger.error(message, **kwargs)
    
    def critical(self, message: str, **kwargs):
        """Log critical message"""
        loguru_logger.critical(message, **kwargs)
    
    def log_function_call(self, func_name: str, args: tuple = (), kwargs: Dict[str, Any] = None, 
                         result: Any = None, duration: Optional[float] = None):
        """Log function call details"""
        kwargs = kwargs or {}
        msg = f"Function: {func_name}"
        
        if args:
            msg += f" | Args: {args}"
        if kwargs:
            msg += f" | Kwargs: {kwargs}"
        if result is not None:
            msg += f" | Result: {result}"
        if duration is not None:
            msg += f" | Duration: {duration:.4f}s"
        
        self.debug(msg)
    
    def log_api_call(self, method: str, url: str, status_code: int, 
                     duration: float, request_data: Optional[Dict] = None,
                     response_data: Optional[Dict] = None):
        """Log API call details"""
        msg = f"API: {method} {url} | Status: {status_code} | Duration: {duration:.4f}s"
        
        if request_data:
            msg += f" | Request: {request_data}"
        if response_data:
            msg += f" | Response: {response_data}"
        
        if status_code >= 400:
            self.error(msg)
        else:
            self.info(msg)

def setup_logging(name: str = "dashboard", log_dir: Optional[Path] = None) -> DashboardLogger:
    """Setup and return dashboard logger"""
    return DashboardLogger(name, log_dir)

def log_uncaught_exceptions(exc_type, exc_value, exc_tb):
    """Global uncaught exception handler"""
    if issubclass(exc_type, KeyboardInterrupt):
        # Allow keyboard interrupts to exit normally
        sys.__excepthook__(exc_type, exc_value, exc_tb)
        return
    
    logger = setup_logging()
    logger.critical(
        "Uncaught exception",
        exception=exc_value
    )

# Set global exception handler
sys.excepthook = log_uncaught_exceptions

# Global logger instance
logger = setup_logging()

__all__ = ['DashboardLogger', 'setup_logging', 'logger']