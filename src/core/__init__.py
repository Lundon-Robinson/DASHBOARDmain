# Core application components
from .config import DashboardConfig
from .logger import setup_logging
from .database import DatabaseManager

# Application import is optional (requires GUI dependencies)
try:
    from .application import DashboardApplication
    __all__ = [
        'DashboardApplication',
        'DashboardConfig', 
        'setup_logging',
        'DatabaseManager'
    ]
except ImportError:
    __all__ = [
        'DashboardConfig', 
        'setup_logging',
        'DatabaseManager'
    ]