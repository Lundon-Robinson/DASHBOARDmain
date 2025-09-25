"""
Advanced Finance/Admin Dashboard
=================================

A comprehensive, feature-packed finance and administration dashboard
with Excel integration, Outlook automation, AI assistance, and advanced analytics.

Main Entry Point
"""

import sys
import os
import asyncio
from pathlib import Path

# Add src directory to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.core.application import DashboardApplication
from src.core.config import DashboardConfig
from src.core.logger import setup_logging

async def main():
    """Main application entry point"""
    try:
        # Setup logging
        logger = setup_logging()
        logger.info("Starting Advanced Finance Dashboard...")
        
        # Load configuration
        config = DashboardConfig()
        
        # Create and run application
        app = DashboardApplication(config)
        await app.run()
        
    except Exception as e:
        print(f"Failed to start dashboard: {e}")
        sys.exit(1)

if __name__ == "__main__":
    asyncio.run(main())