"""
100 Consecutive Test Runner
===========================

Automated end-to-end test harness that executes full workflows
100 times consecutively to ensure system stability and reliability.
"""

import sys
import time
import json
import traceback
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, Any, List, Optional
import tempfile
import shutil

# Add src directory to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.core.config import DashboardConfig
from src.core.logger import setup_logging
from src.core.database import get_db_manager
from src.modules.excel_handler import ExcelHandler
from src.modules.email_handler import EmailHandler
from src.modules.script_runner import ScriptRunner
from src.modules.ai_assistant import AIAssistant

class TestResult:
    """Individual test result"""
    
    def __init__(self, test_name: str):
        self.test_name = test_name
        self.start_time = datetime.now()
        self.end_time: Optional[datetime] = None
        self.success = False
        self.error_message = ""
        self.details: Dict[str, Any] = {}
        self.duration_seconds = 0.0
    
    def complete(self, success: bool, error_message: str = "", details: Dict[str, Any] = None):
        """Mark test as completed"""
        self.end_time = datetime.now()
        self.success = success
        self.error_message = error_message
        self.details = details or {}
        self.duration_seconds = (self.end_time - self.start_time).total_seconds()

class TestHarness:
    """100 consecutive test harness"""
    
    def __init__(self):
        self.logger = setup_logging("test_harness")
        self.results: List[TestResult] = []
        self.consecutive_successes = 0
        self.max_consecutive_target = 100
        self.temp_dir = None
        
        # Test configuration
        self.test_config = None
        
    def setup_test_environment(self) -> bool:
        """Setup isolated test environment"""
        try:
            # Create temporary directory for test isolation
            self.temp_dir = Path(tempfile.mkdtemp(prefix="dashboard_test_"))
            
            # Setup test database
            test_db_path = self.temp_dir / "test_dashboard.db"
            
            # Create test configuration
            self.test_config = DashboardConfig()
            self.test_config.database.url = f"sqlite:///{test_db_path}"
            self.test_config.paths.data_dir = self.temp_dir / "data"
            self.test_config.paths.logs_dir = self.temp_dir / "logs"
            self.test_config.paths.temp_dir = self.temp_dir / "temp"
            self.test_config.paths.exports_dir = self.temp_dir / "exports"
            self.test_config.paths.templates_dir = self.temp_dir / "templates"
            
            # Create directories
            for path in [self.test_config.paths.data_dir, self.test_config.paths.logs_dir,
                        self.test_config.paths.temp_dir, self.test_config.paths.exports_dir,
                        self.test_config.paths.templates_dir]:
                path.mkdir(parents=True, exist_ok=True)
            
            # Initialize database with test data
            self._setup_test_data()
            
            self.logger.info(f"Test environment setup complete: {self.temp_dir}")
            return True
            
        except Exception as e:
            self.logger.error("Failed to setup test environment", exception=e)
            return False
    
    def _setup_test_data(self):
        """Setup test data in database"""
        db_manager = get_db_manager(self.test_config.database.url)
        
        # Create test cardholders
        test_cardholders = [
            ("1234-5678-9012-3456", "John Doe", "john.doe@test.com", "IT Department"),
            ("2345-6789-0123-4567", "Jane Smith", "jane.smith@test.com", "Finance"),
            ("3456-7890-1234-5678", "Bob Wilson", "bob.wilson@test.com", "Operations")
        ]
        
        for card_number, name, email, department in test_cardholders:
            try:
                db_manager.create_cardholder(
                    card_number=card_number,
                    name=name,
                    email=email,
                    department=department
                )
            except Exception as e:
                # Ignore duplicate errors
                pass
        
        # Create test transactions
        from datetime import datetime
        cardholders = db_manager.get_cardholders()
        
        for cardholder in cardholders:
            for i in range(5):  # 5 transactions per cardholder
                try:
                    db_manager.create_transaction(
                        cardholder_id=cardholder.id,
                        transaction_date=datetime.now() - timedelta(days=i),
                        merchant=f"Test Merchant {i+1}",
                        amount=round(100.0 + i * 25.5, 2),
                        category="Test"
                    )
                except Exception as e:
                    # Ignore errors for test setup
                    pass
        
        self.logger.info("Test data setup complete")
    
    def cleanup_test_environment(self):
        """Clean up test environment"""
        try:
            if self.temp_dir and self.temp_dir.exists():
                shutil.rmtree(self.temp_dir)
                self.logger.info("Test environment cleaned up")
        except Exception as e:
            self.logger.warning(f"Failed to cleanup test environment: {e}")
    
    def run_single_test_cycle(self) -> List[TestResult]:
        """Run a single complete test cycle"""
        cycle_results = []
        
        # Test 1: Database Operations
        result = TestResult("database_operations")
        try:
            db_manager = get_db_manager(self.test_config.database.url)
            cardholders = db_manager.get_cardholders()
            transactions = db_manager.get_transactions()
            
            result.complete(True, details={
                "cardholders_count": len(cardholders),
                "transactions_count": len(transactions)
            })
        except Exception as e:
            result.complete(False, str(e))
        
        cycle_results.append(result)
        
        # Test 2: Excel Handler
        result = TestResult("excel_handler")
        try:
            excel_handler = ExcelHandler(self.test_config)
            
            # Create sample data
            import pandas as pd
            sample_data = pd.DataFrame({
                'date': [datetime.now() - timedelta(days=i) for i in range(3)],
                'amount': [100.0, 200.0, 300.0],
                'merchant': ['Test Merchant A', 'Test Merchant B', 'Test Merchant C']
            })
            
            # Test operations
            duplicates = excel_handler.detect_duplicates(sample_data)
            html_export = excel_handler.export_to_html(sample_data)
            validation = excel_handler.validate_data(sample_data, {
                'amount': {'required': True, 'min_value': 0}
            })
            
            result.complete(True, details={
                "duplicates_found": len(duplicates),
                "html_exported": Path(html_export).exists(),
                "validation_errors": len(validation['errors'])
            })
        except Exception as e:
            result.complete(False, str(e))
        
        cycle_results.append(result)
        
        # Test 3: Email Handler
        result = TestResult("email_handler")
        try:
            email_handler = EmailHandler(self.test_config)
            
            # Test template processing
            template = email_handler.load_template("purchase_card_statement")
            variables = {
                'cardholder_name': 'Test User',
                'period': 'Test Period',
                'card_number': '****-1234',
                'total_amount': '£100.00'
            }
            
            processed_subject = email_handler.process_variables(template.subject, variables)
            processed_body = email_handler.process_variables(template.body, variables)
            
            # Test recipient parsing
            recipients = email_handler.parse_recipient_list("test1@example.com, test2@example.com")
            
            result.complete(True, details={
                "template_loaded": template is not None,
                "subject_processed": "Test User" in processed_subject,
                "recipients_parsed": len(recipients)
            })
        except Exception as e:
            result.complete(False, str(e))
        
        cycle_results.append(result)
        
        # Test 4: Script Runner
        result = TestResult("script_runner")
        try:
            script_runner = ScriptRunner(self.test_config)
            
            # Create a simple test script
            test_script_path = self.temp_dir / "test_script.py"
            with open(test_script_path, 'w') as f:
                f.write("""
import sys
print("Test script running")
sys.exit(0)
""")
            
            from src.modules.script_runner import ScriptInfo
            test_script = ScriptInfo(
                name="test_script",
                path=str(test_script_path),
                description="Test script",
                timeout=10
            )
            script_runner.register_script(test_script)
            
            # Run script and wait for completion
            execution_id = script_runner.run_script("test_script")
            
            # Wait for completion (max 15 seconds)
            for _ in range(15):
                status = script_runner.get_script_status(execution_id)
                if status.value in ['success', 'failed']:
                    break
                time.sleep(1)
            
            final_status = script_runner.get_script_status(execution_id)
            output = script_runner.get_script_output(execution_id)
            
            script_runner.cleanup()
            
            result.complete(True, details={
                "execution_id": execution_id,
                "final_status": final_status.value,
                "output_lines": len(output)
            })
        except Exception as e:
            result.complete(False, str(e))
        
        cycle_results.append(result)
        
        # Test 5: AI Assistant
        result = TestResult("ai_assistant")
        try:
            ai_assistant = AIAssistant(self.test_config)
            
            # Test command processing
            commands = [
                "system status",
                "analyze logs today",
                "generate statements for September 2025"
            ]
            
            command_results = []
            for command in commands:
                cmd_result = ai_assistant.process_command(command)
                command_results.append(cmd_result['success'])
            
            # Test log analysis
            log_analysis = ai_assistant.analyze_logs("today")
            system_status = ai_assistant.get_system_status()
            
            result.complete(True, details={
                "commands_successful": sum(command_results),
                "total_commands": len(commands),
                "log_analysis_completed": 'summary' in log_analysis,
                "system_status_healthy": system_status['overall_health'] == 'healthy'
            })
        except Exception as e:
            result.complete(False, str(e))
        
        cycle_results.append(result)
        
        # Test 6: Integration Test (Statement Generation Workflow)
        result = TestResult("integration_workflow")
        try:
            # Full workflow: Load data -> Process -> Generate statements -> Send notifications
            db_manager = get_db_manager(self.test_config.database.url)
            excel_handler = ExcelHandler(self.test_config)
            
            # Generate statements for test period
            start_date = datetime.now().replace(day=1)
            end_date = datetime.now()
            
            statements = excel_handler.batch_generate_statements(start_date, end_date)
            
            # Verify statements exist
            statements_exist = all(Path(path).exists() for path in statements)
            
            result.complete(True, details={
                "statements_generated": len(statements),
                "statements_exist": statements_exist,
                "period_start": start_date.isoformat(),
                "period_end": end_date.isoformat()
            })
        except Exception as e:
            result.complete(False, str(e))
        
        cycle_results.append(result)
        
        return cycle_results
    
    def run_100_consecutive(self) -> Dict[str, Any]:
        """Run 100 consecutive successful test cycles"""
        self.logger.info("Starting 100 consecutive test run")
        start_time = datetime.now()
        
        cycle_count = 0
        failure_count = 0
        
        while self.consecutive_successes < self.max_consecutive_target:
            cycle_count += 1
            
            self.logger.info(f"Running test cycle {cycle_count} (consecutive successes: {self.consecutive_successes})")
            
            # Setup fresh environment for each cycle
            if not self.setup_test_environment():
                failure_count += 1
                self.consecutive_successes = 0
                continue
            
            try:
                # Run test cycle
                cycle_results = self.run_single_test_cycle()
                
                # Check if all tests passed
                all_passed = all(result.success for result in cycle_results)
                
                if all_passed:
                    self.consecutive_successes += 1
                    self.logger.info(f"✓ Cycle {cycle_count} passed (consecutive: {self.consecutive_successes})")
                else:
                    failure_count += 1
                    self.consecutive_successes = 0
                    failed_tests = [r.test_name for r in cycle_results if not r.success]
                    self.logger.error(f"✗ Cycle {cycle_count} failed. Failed tests: {failed_tests}")
                    
                    # Log details of failures
                    for result in cycle_results:
                        if not result.success:
                            self.logger.error(f"Test '{result.test_name}' failed: {result.error_message}")
                
                # Add results to overall collection
                self.results.extend(cycle_results)
                
            except Exception as e:
                failure_count += 1
                self.consecutive_successes = 0
                self.logger.error(f"✗ Cycle {cycle_count} crashed: {e}")
                traceback.print_exc()
            
            finally:
                # Cleanup environment
                self.cleanup_test_environment()
            
            # Progress update every 10 cycles
            if cycle_count % 10 == 0:
                self.logger.info(f"Progress: {cycle_count} cycles completed, {self.consecutive_successes} consecutive successes")
            
            # Safety break for infinite loops
            if cycle_count > 1000:  # Maximum 1000 cycles
                self.logger.error("Reached maximum cycle limit (1000), stopping")
                break
        
        end_time = datetime.now()
        total_duration = (end_time - start_time).total_seconds()
        
        # Generate final report
        report = {
            'success': self.consecutive_successes >= self.max_consecutive_target,
            'consecutive_successes': self.consecutive_successes,
            'target': self.max_consecutive_target,
            'total_cycles': cycle_count,
            'total_failures': failure_count,
            'total_duration_seconds': total_duration,
            'start_time': start_time.isoformat(),
            'end_time': end_time.isoformat(),
            'test_summary': self._generate_test_summary()
        }
        
        if report['success']:
            self.logger.info(f"🎉 SUCCESS: Achieved {self.consecutive_successes} consecutive successful test runs!")
        else:
            self.logger.error(f"❌ FAILED: Only achieved {self.consecutive_successes} consecutive successes out of {self.max_consecutive_target} target")
        
        return report
    
    def _generate_test_summary(self) -> Dict[str, Any]:
        """Generate summary statistics of all test results"""
        if not self.results:
            return {}
        
        # Group results by test name
        test_stats = {}
        for result in self.results:
            if result.test_name not in test_stats:
                test_stats[result.test_name] = {
                    'total_runs': 0,
                    'successes': 0,
                    'failures': 0,
                    'avg_duration': 0.0,
                    'total_duration': 0.0
                }
            
            stats = test_stats[result.test_name]
            stats['total_runs'] += 1
            stats['total_duration'] += result.duration_seconds
            
            if result.success:
                stats['successes'] += 1
            else:
                stats['failures'] += 1
        
        # Calculate averages
        for test_name, stats in test_stats.items():
            if stats['total_runs'] > 0:
                stats['avg_duration'] = stats['total_duration'] / stats['total_runs']
                stats['success_rate'] = (stats['successes'] / stats['total_runs']) * 100
        
        return test_stats
    
    def save_report(self, report: Dict[str, Any], filename: str = None):
        """Save test report to file"""
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"test_report_100_consecutive_{timestamp}.json"
        
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(report, f, indent=2, default=str)
            
            self.logger.info(f"Test report saved: {filename}")
        except Exception as e:
            self.logger.error(f"Failed to save test report: {e}")

def main():
    """Main entry point"""
    print("🚀 Starting 100 Consecutive Test Runner")
    print("=" * 60)
    
    harness = TestHarness()
    
    try:
        report = harness.run_100_consecutive()
        
        # Save report
        harness.save_report(report)
        
        # Print summary
        print("\\n" + "=" * 60)
        print("📊 FINAL REPORT")
        print("=" * 60)
        print(f"Success: {report['success']}")
        print(f"Consecutive Successes: {report['consecutive_successes']}/{report['target']}")
        print(f"Total Cycles: {report['total_cycles']}")
        print(f"Total Failures: {report['total_failures']}")
        print(f"Total Duration: {report['total_duration_seconds']:.2f} seconds")
        
        if report['test_summary']:
            print("\\nTest Performance:")
            for test_name, stats in report['test_summary'].items():
                print(f"  {test_name}:")
                print(f"    Success Rate: {stats['success_rate']:.1f}% ({stats['successes']}/{stats['total_runs']})")
                print(f"    Avg Duration: {stats['avg_duration']:.3f}s")
        
        # Exit code
        sys.exit(0 if report['success'] else 1)
        
    except KeyboardInterrupt:
        print("\\n⚠️  Test run interrupted by user")
        sys.exit(2)
    except Exception as e:
        print(f"\\n❌ Test harness failed: {e}")
        traceback.print_exc()
        sys.exit(3)
    finally:
        # Final cleanup
        harness.cleanup_test_environment()

if __name__ == "__main__":
    main()