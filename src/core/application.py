"""
Main Dashboard Application
==========================

The core application class that coordinates all components
and provides the main GUI interface.
"""

import sys
from typing import Dict, Any

try:
    from PyQt6.QtWidgets import (
        QApplication, QMainWindow, QWidget,
        QTabWidget, QStatusBar, QMenuBar, QPushButton
    )
    from PyQt6.QtGui import QAction
    PYQT_AVAILABLE = True
except ImportError:
    PYQT_AVAILABLE = False
    print("PyQt6 not available, falling back to Tkinter")

if not PYQT_AVAILABLE:
    import tkinter as tk
    from tkinter import ttk

from .config import DashboardConfig
from .logger import logger
from .database import get_db_manager
from ..ui.main_window import MainWindow
from ..modules.excel_handler import ExcelHandler
from ..modules.email_handler import EmailHandler
from ..modules.script_runner import ScriptRunner
from ..modules.ai_assistant import AIAssistant

class DashboardApplication:
    """Main dashboard application"""
    
    def __init__(self, config: DashboardConfig):
        self.config = config
        self.db_manager = get_db_manager(config.database.url)
        
        # Initialize components
        self.excel_handler = ExcelHandler(config)
        self.email_handler = EmailHandler(config)
        self.script_runner = ScriptRunner(config)
        self.ai_assistant = AIAssistant(config) if config.ai.openai_api_key else None
        
        # GUI components
        self.app = None
        self.main_window = None
        self.tray_icon = None
        
        # Runtime state
        self.is_running = False
        
        logger.info("Dashboard application initialized")
    
    async def run(self):
        """Run the dashboard application"""
        try:
            if PYQT_AVAILABLE:
                await self._run_pyqt()
            else:
                await self._run_tkinter()
        except Exception as e:
            logger.error("Failed to run application", exception=e)
            raise
    
    async def _run_pyqt(self):
        """Run PyQt6 version of the application"""
        self.app = QApplication(sys.argv)
        self.app.setApplicationName("Advanced Finance Dashboard")
        self.app.setApplicationVersion(self.config.version)
        
        # Apply theme
        self._apply_pyqt_theme()
        
        # Create main window
        self.main_window = MainWindowPyQt(self)
        
        # Create system tray
        self._create_system_tray()
        
        # Show window
        self.main_window.show()
        
        # Center window
        self._center_window()
        
        self.is_running = True
        logger.info("PyQt6 application started")
        
        # Run event loop
        sys.exit(self.app.exec())
    
    async def _run_tkinter(self):
        """Run Tkinter version of the application"""
        self.app = tk.Tk()
        self.app.title("Advanced Finance Dashboard")
        self.app.geometry(f"{self.config.ui.window_width}x{self.config.ui.window_height}")
        
        # Apply theme
        self._apply_tkinter_theme()
        
        # Create main window
        self.main_window = MainWindowTkinter(self.app, self)
        
        self.is_running = True
        logger.info("Tkinter application started")
        
        # Run event loop
        self.app.mainloop()
    
    def _apply_pyqt_theme(self):
        """Apply theme to PyQt application"""
        if self.config.ui.theme == "dark":
            dark_style = """
            QMainWindow {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QWidget {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QPushButton {
                background-color: #3c3c3c;
                border: 1px solid #555555;
                padding: 8px;
                border-radius: 4px;
                color: #ffffff;
            }
            QPushButton:hover {
                background-color: #4c4c4c;
            }
            QPushButton:pressed {
                background-color: #1c1c1c;
            }
            QTabWidget::pane {
                border: 1px solid #555555;
                background-color: #2b2b2b;
            }
            QTabBar::tab {
                background-color: #3c3c3c;
                color: #ffffff;
                padding: 8px 16px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: #0078d4;
            }
            QTextEdit, QLineEdit {
                background-color: #1e1e1e;
                border: 1px solid #555555;
                color: #ffffff;
            }
            QMenuBar {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QMenuBar::item:selected {
                background-color: #0078d4;
            }
            QStatusBar {
                background-color: #1e1e1e;
                color: #ffffff;
            }
            """
            self.app.setStyleSheet(dark_style)
    
    def _apply_tkinter_theme(self):
        """Apply theme to Tkinter application"""
        # Configure ttk styles for dark theme
        style = ttk.Style()
        
        if self.config.ui.theme == "dark":
            style.theme_use("clam")
            style.configure(".", background="#2b2b2b", foreground="#ffffff")
            style.configure("TFrame", background="#2b2b2b")
            style.configure("TLabel", background="#2b2b2b", foreground="#ffffff")
            style.configure("TButton", background="#3c3c3c", foreground="#ffffff")
            style.map("TButton", background=[("active", "#4c4c4c")])
            style.configure("TEntry", background="#1e1e1e", foreground="#ffffff")
            style.configure("TText", background="#1e1e1e", foreground="#ffffff")
            
            self.app.configure(bg="#2b2b2b")
    
    def _create_system_tray(self):
        """Create system tray icon"""
        if not PYQT_AVAILABLE or not QSystemTrayIcon.isSystemTrayAvailable():
            return
        
        self.tray_icon = QSystemTrayIcon()
        
        # Create tray menu
        tray_menu = QMenu()
        
        show_action = QAction("Show", self.app)
        show_action.triggered.connect(self.main_window.show)
        tray_menu.addAction(show_action)
        
        hide_action = QAction("Hide", self.app)
        hide_action.triggered.connect(self.main_window.hide)
        tray_menu.addAction(hide_action)
        
        tray_menu.addSeparator()
        
        quit_action = QAction("Quit", self.app)
        quit_action.triggered.connect(self._quit_application)
        tray_menu.addAction(quit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
    
    def _center_window(self):
        """Center the main window on screen"""
        if PYQT_AVAILABLE and self.main_window:
            screen = self.app.primaryScreen().availableGeometry()
            window = self.main_window.frameGeometry()
            center = screen.center()
            window.moveCenter(center)
            self.main_window.move(window.topLeft())
    
    def _quit_application(self):
        """Quit the application"""
        logger.info("Shutting down dashboard application")
        
        # Clean up components
        if self.script_runner:
            self.script_runner.cleanup()
        
        if self.db_manager:
            self.db_manager.close()
        
        if self.app:
            self.app.quit()
        
        self.is_running = False
    
    def get_status(self) -> Dict[str, Any]:
        """Get application status"""
        return {
            "running": self.is_running,
            "version": self.config.version,
            "database_connected": self.db_manager is not None,
            "ai_enabled": self.ai_assistant is not None,
            "components": {
                "excel_handler": self.excel_handler is not None,
                "email_handler": self.email_handler is not None,
                "script_runner": self.script_runner is not None,
                "ai_assistant": self.ai_assistant is not None
            }
        }

# Simplified run function for non-async use
def run_dashboard(config: DashboardConfig = None):
    """Simple function to run the dashboard"""
    if config is None:
        config = DashboardConfig()
    
    app = DashboardApplication(config)
    
    # Run synchronously
    try:
        main_window = MainWindow(config)
        main_window.run()
    except Exception as e:
        logger.error("Dashboard run failed", exception=e)
        raise
    finally:
        if hasattr(app, 'script_runner') and app.script_runner:
            app.script_runner.cleanup()

__all__ = ['DashboardApplication', 'run_dashboard']