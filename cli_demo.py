"""
Command Line Interface for Dashboard
====================================

Non-GUI interface for testing and automation
"""

import sys
import os
from pathlib import Path

# Add src directory to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.core.config import DashboardConfig
from src.core.logger import setup_logging
from src.modules.excel_handler import ExcelHandler
from src.modules.email_handler import EmailHandler
from src.modules.script_runner import ScriptRunner
from src.modules.ai_assistant import AIAssistant

def main():
    """CLI demonstration"""
    print("🚀 Advanced Finance Dashboard - CLI Demo")
    print("=" * 60)
    
    try:
        # Initialize
        logger = setup_logging()
        config = DashboardConfig()
        
        print(f"✓ Configuration loaded (Theme: {config.ui.theme})")
        
        # Initialize modules
        excel_handler = ExcelHandler(config)
        email_handler = EmailHandler(config)
        script_runner = ScriptRunner(config)
        ai_assistant = AIAssistant(config)
        
        print("✓ All modules initialized successfully")
        
        # Demo functionality
        print("\n📊 Module Status:")
        print(f"  Excel Handler: Ready ({len(excel_handler.db_manager.get_cardholders())} cardholders)")
        print(f"  Email Handler: Ready ({len(email_handler.list_templates())} templates)")
        print(f"  Script Runner: Ready ({len(script_runner.list_scripts())} scripts)")
        print(f"  AI Assistant: {'Ready' if ai_assistant else 'Not configured'}")
        
        # Demo AI commands
        print("\n🤖 AI Assistant Demo:")
        test_commands = [
            "system status",
            "analyze logs today", 
            "generate statements for September 2025"
        ]
        
        for command in test_commands:
            result = ai_assistant.process_command(command)
            status = "✓" if result['success'] else "✗"
            print(f"  {status} '{command}': {result.get('result', result.get('error', 'Unknown'))[:60]}...")
        
        # Demo Excel operations
        print("\n📈 Excel Handler Demo:")
        import pandas as pd
        from datetime import datetime
        
        sample_data = pd.DataFrame({
            'date': [datetime.now()],
            'amount': [123.45],
            'merchant': ['Test Merchant']
        })
        
        html_export = excel_handler.export_to_html(sample_data)
        print(f"  ✓ Sample data exported to HTML: {Path(html_export).name}")
        
        duplicates = excel_handler.detect_duplicates(sample_data)
        print(f"  ✓ Duplicate detection: {len(duplicates)} found")
        
        # Demo Email operations
        print("\n📧 Email Handler Demo:")
        templates = email_handler.list_templates()
        print(f"  ✓ Available templates: {', '.join(templates)}")
        
        recipients = email_handler.parse_recipient_list("test1@example.com, test2@example.com")
        print(f"  ✓ Parsed recipients: {len(recipients)} valid emails")
        
        # Demo Script Runner
        print("\n🖥️ Script Runner Demo:")
        scripts = script_runner.list_scripts()
        print(f"  ✓ Discovered scripts: {len(scripts)}")
        for script in scripts[:5]:  # Show first 5
            print(f"    - {script.name} ({script.category}): {script.description[:40]}...")
        
        # System statistics
        print("\n📊 System Statistics:")
        
        # Email stats
        email_stats = email_handler.get_email_statistics()
        print(f"  Email Statistics: {email_stats['total_sent']} sent, {email_stats['total_failed']} failed")
        
        # Execution history
        execution_history = script_runner.get_execution_history(5)
        print(f"  Recent Executions: {len(execution_history)} in history")
        
        # System status
        system_status = ai_assistant.get_system_status()
        print(f"  System Health: {system_status['overall_health']}")
        
        print("\n🎉 Dashboard CLI demo completed successfully!")
        print("\nTo run the full GUI version, ensure you have a display environment and run:")
        print("  python main_dashboard.py")
        
    except Exception as e:
        print(f"\n❌ Demo failed: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        # Cleanup
        if 'script_runner' in locals():
            script_runner.cleanup()

if __name__ == "__main__":
    main()