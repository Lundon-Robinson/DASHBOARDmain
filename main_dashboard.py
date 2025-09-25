"""
Advanced Finance/Admin Dashboard
=================================

A comprehensive, feature-packed finance and administration dashboard
with Excel integration, Outlook automation, AI assistance, and advanced analytics.

Main Entry Point
"""

import sys
from pathlib import Path

# Add src directory to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.core.config import DashboardConfig
from src.core.logger import setup_logging
from src.core.application import run_dashboard

def main():
    """Main application entry point"""
    try:
        # Setup logging
        logger = setup_logging()
        logger.info("Starting Advanced Finance Dashboard...")
        
        # Load configuration
        config = DashboardConfig()
        logger.info(f"Configuration loaded - Theme: {config.ui.theme}, Database: {config.database.url}")
        
        # Run dashboard
        run_dashboard(config)
        
    except KeyboardInterrupt:
        print("\nDashboard stopped by user")
        sys.exit(0)
    except Exception as e:
        print(f"Failed to start dashboard: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()