"""
AI Assistant Module
===================

AI-powered assistant for log analysis, natural language
commands, and intelligent automation.
"""

import re
import json
from typing import List, Dict, Any, Optional, Callable
from datetime import datetime, timedelta
from pathlib import Path

try:
    import openai
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

from ..core.logger import logger
from ..core.database import get_db_manager

class AIAssistant:
    """AI assistant handler"""
    
    def __init__(self, config):
        self.config = config
        self.db_manager = get_db_manager(config.database.url)
        
        # Initialize OpenAI if available
        if OPENAI_AVAILABLE and config.ai.openai_api_key:
            openai.api_key = config.ai.openai_api_key
            self.openai_enabled = True
            logger.info("OpenAI integration enabled")
        else:
            self.openai_enabled = False
            logger.warning("OpenAI not available or API key not configured")
        
        # Command patterns for natural language processing
        self.command_patterns = {
            'generate_statements': [
                r'generate.*statements?.*for.*(?P<period>\w+\s+\d{4})',
                r'create.*statements?.*(?P<period>\w+\s+\d{4})',
                r'statements?.*for.*(?P<period>\w+\s+\d{4})'
            ],
            'send_emails': [
                r'send.*emails?.*to.*(?P<recipients>.+)',
                r'email.*(?P<recipients>.+)',
                r'bulk.*email.*(?P<recipients>.+)'
            ],
            'analyze_logs': [
                r'analyze.*logs?',
                r'check.*logs?.*(?P<period>today|yesterday|week|month)?',
                r'log.*analysis.*(?P<period>today|yesterday|week|month)?'
            ],
            'run_script': [
                r'run.*script.*(?P<script_name>\w+)',
                r'execute.*(?P<script_name>\w+)',
                r'start.*(?P<script_name>\w+)'
            ],
            'get_status': [
                r'status',
                r'dashboard.*status',
                r'system.*status'
            ]
        }
        
        # Action handlers
        self.action_handlers = {
            'generate_statements': self._handle_generate_statements,
            'send_emails': self._handle_send_emails,
            'analyze_logs': self._handle_analyze_logs,
            'run_script': self._handle_run_script,
            'get_status': self._handle_get_status
        }
        
        logger.info("AI assistant initialized")
    
    def process_command(self, user_input: str) -> Dict[str, Any]:
        """Process natural language command"""
        try:
            user_input = user_input.lower().strip()
            
            # Try to match command patterns
            for command_type, patterns in self.command_patterns.items():
                for pattern in patterns:
                    match = re.search(pattern, user_input, re.IGNORECASE)
                    if match:
                        # Extract parameters
                        params = match.groupdict()
                        
                        # Execute command
                        handler = self.action_handlers.get(command_type)
                        if handler:
                            result = handler(user_input, params)
                            return {
                                'success': True,
                                'command_type': command_type,
                                'parameters': params,
                                'result': result
                            }
            
            # If no pattern matches, try AI interpretation
            if self.openai_enabled:
                return self._ai_interpret_command(user_input)
            
            # Fallback response
            return {
                'success': False,
                'error': 'Command not recognized. Try commands like "generate statements for September 2025" or "analyze logs"',
                'suggestions': [
                    'generate statements for [month year]',
                    'send emails to [recipients]',
                    'analyze logs',
                    'run script [name]',
                    'system status'
                ]
            }
            
        except Exception as e:
            logger.error("Failed to process AI command", exception=e)
            return {
                'success': False,
                'error': f'Error processing command: {str(e)}'
            }
    
    def _ai_interpret_command(self, user_input: str) -> Dict[str, Any]:
        """Use OpenAI to interpret complex commands"""
        try:
            prompt = f"""
You are an AI assistant for a finance dashboard system. The user said: "{user_input}"

Available commands:
1. generate_statements - Generate purchase card statements for a period
2. send_emails - Send bulk emails to recipients  
3. analyze_logs - Analyze system logs for issues
4. run_script - Execute a specific script
5. get_status - Get system status

Based on the user input, determine:
1. What command they want to execute
2. What parameters they provided
3. If the request is unclear, what clarification is needed

Respond with JSON format:
{{
    "command": "command_name or null",
    "parameters": {{}},
    "confidence": 0-100,
    "clarification_needed": "question if unclear"
}}
"""
            
            response = openai.ChatCompletion.create(
                model=self.config.ai.model,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=self.config.ai.max_tokens,
                temperature=self.config.ai.temperature
            )
            
            ai_response = response.choices[0].message.content
            parsed_response = json.loads(ai_response)
            
            if parsed_response.get('command') and parsed_response.get('confidence', 0) > 70:
                # Execute the interpreted command
                command = parsed_response['command']
                params = parsed_response.get('parameters', {})
                
                handler = self.action_handlers.get(command)
                if handler:
                    result = handler(user_input, params)
                    return {
                        'success': True,
                        'command_type': command,
                        'parameters': params,
                        'result': result,
                        'ai_interpretation': True
                    }
            
            # Low confidence or unclear command
            return {
                'success': False,
                'ai_interpretation': True,
                'clarification': parsed_response.get('clarification_needed', 'Could you please clarify your request?'),
                'confidence': parsed_response.get('confidence', 0)
            }
            
        except Exception as e:
            logger.error("AI interpretation failed", exception=e)
            return {
                'success': False,
                'error': 'AI interpretation failed',
                'fallback': True
            }
    
    def _handle_generate_statements(self, user_input: str, params: Dict[str, Any]) -> str:
        """Handle statement generation command"""
        period = params.get('period', 'current month')
        
        # Parse period
        try:
            from dateutil import parser as date_parser
            # Try to parse the period
            if 'current' in period.lower():
                start_date = datetime.now().replace(day=1)
                end_date = datetime.now()
            else:
                # Try to parse month/year
                parsed_date = date_parser.parse(period, fuzzy=True)
                start_date = parsed_date.replace(day=1)
                # Get last day of month
                if parsed_date.month == 12:
                    end_date = parsed_date.replace(year=parsed_date.year + 1, month=1, day=1) - timedelta(days=1)
                else:
                    end_date = parsed_date.replace(month=parsed_date.month + 1, day=1) - timedelta(days=1)
        except:
            start_date = datetime.now().replace(day=1)
            end_date = datetime.now()
        
        try:
            # Get Excel handler and generate statements
            from ..modules.excel_handler import ExcelHandler
            excel_handler = ExcelHandler(self.config)
            
            statements = excel_handler.batch_generate_statements(start_date, end_date)
            return f"Generated {len(statements)} statements for {period}"
            
        except Exception as e:
            return f"Failed to generate statements: {str(e)}"
    
    def _handle_send_emails(self, user_input: str, params: Dict[str, Any]) -> str:
        """Handle email sending command"""
        recipients_text = params.get('recipients', '')
        
        try:
            from ..modules.email_handler import EmailHandler
            email_handler = EmailHandler(self.config)
            
            recipients = email_handler.parse_recipient_list(recipients_text)
            
            if not recipients:
                return "No valid email recipients found"
            
            # For now, just count recipients
            return f"Would send emails to {len(recipients)} recipients: {', '.join(r.email for r in recipients[:3])}"
            
        except Exception as e:
            return f"Failed to process email command: {str(e)}"
    
    def _handle_analyze_logs(self, user_input: str, params: Dict[str, Any]) -> str:
        """Handle log analysis command"""
        period = params.get('period', 'today')
        
        try:
            # Get log analysis
            analysis = self.analyze_logs(period)
            return f"Log analysis complete: {analysis['summary']}"
            
        except Exception as e:
            return f"Failed to analyze logs: {str(e)}"
    
    def _handle_run_script(self, user_input: str, params: Dict[str, Any]) -> str:
        """Handle script execution command"""
        script_name = params.get('script_name', '')
        
        try:
            from ..modules.script_runner import ScriptRunner
            script_runner = ScriptRunner(self.config)
            
            # Check if script exists
            scripts = script_runner.list_scripts()
            script_names = [s.name.lower() for s in scripts]
            
            if script_name.lower() in script_names:
                execution_id = script_runner.run_script(script_name)
                return f"Started script '{script_name}' with ID: {execution_id}"
            else:
                available = ', '.join([s.name for s in scripts[:5]])
                return f"Script '{script_name}' not found. Available scripts: {available}"
                
        except Exception as e:
            return f"Failed to run script: {str(e)}"
    
    def _handle_get_status(self, user_input: str, params: Dict[str, Any]) -> str:
        """Handle status request"""
        try:
            # Get system status
            status = self.get_system_status()
            return f"System Status: {status['summary']}"
            
        except Exception as e:
            return f"Failed to get status: {str(e)}"
    
    def analyze_logs(self, period: str = 'today') -> Dict[str, Any]:
        """Analyze system logs for issues and patterns"""
        try:
            log_file = self.config.paths.logs_dir / "dashboard.log"
            
            if not log_file.exists():
                return {
                    'summary': 'No log file found',
                    'errors': 0,
                    'warnings': 0,
                    'info': 0
                }
            
            # Read recent log entries
            with open(log_file, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            # Filter by period
            now = datetime.now()
            if period == 'today':
                cutoff = now.replace(hour=0, minute=0, second=0)
            elif period == 'yesterday':
                cutoff = now.replace(hour=0, minute=0, second=0) - timedelta(days=1)
            elif period == 'week':
                cutoff = now - timedelta(days=7)
            elif period == 'month':
                cutoff = now - timedelta(days=30)
            else:
                cutoff = now - timedelta(days=1)
            
            # Count log levels
            errors = []
            warnings = []
            info_count = 0
            
            for line in lines:
                try:
                    # Try to parse timestamp
                    if '|' in line:
                        timestamp_part = line.split('|')[0].strip()
                        log_time = datetime.strptime(timestamp_part, '%Y-%m-%d %H:%M:%S')
                        
                        if log_time >= cutoff:
                            if '| ERROR' in line:
                                errors.append(line.strip())
                            elif '| WARNING' in line:
                                warnings.append(line.strip())
                            elif '| INFO' in line:
                                info_count += 1
                except:
                    continue
            
            # Generate summary
            issues = []
            if len(errors) > 0:
                issues.append(f"{len(errors)} errors")
            if len(warnings) > 0:
                issues.append(f"{len(warnings)} warnings")
            
            if issues:
                summary = f"Found {', '.join(issues)} in {period}"
            else:
                summary = f"No issues found in {period} ({info_count} info messages)"
            
            return {
                'summary': summary,
                'errors': len(errors),
                'warnings': len(warnings),
                'info': info_count,
                'period': period,
                'recent_errors': errors[-5:] if errors else [],
                'recent_warnings': warnings[-5:] if warnings else []
            }
            
        except Exception as e:
            logger.error("Log analysis failed", exception=e)
            return {
                'summary': f'Log analysis failed: {str(e)}',
                'errors': 0,
                'warnings': 0,
                'info': 0
            }
    
    def get_system_status(self) -> Dict[str, Any]:
        """Get overall system status"""
        try:
            status_info = {
                'timestamp': datetime.now().isoformat(),
                'database_connected': True,
                'components': {}
            }
            
            # Check database
            try:
                cardholders = self.db_manager.get_cardholders()
                status_info['components']['database'] = {
                    'status': 'healthy',
                    'cardholders_count': len(cardholders)
                }
            except Exception as e:
                status_info['components']['database'] = {
                    'status': 'error',
                    'error': str(e)
                }
                status_info['database_connected'] = False
            
            # Check logs
            log_analysis = self.analyze_logs('today')
            status_info['components']['logs'] = {
                'status': 'healthy' if log_analysis['errors'] == 0 else 'warnings',
                'errors_today': log_analysis['errors'],
                'warnings_today': log_analysis['warnings']
            }
            
            # Overall health
            all_healthy = all(
                comp.get('status') == 'healthy' 
                for comp in status_info['components'].values()
            )
            
            status_info['overall_health'] = 'healthy' if all_healthy else 'issues'
            status_info['summary'] = (
                'All systems operational' if all_healthy 
                else 'Some components have issues'
            )
            
            return status_info
            
        except Exception as e:
            logger.error("Failed to get system status", exception=e)
            return {
                'summary': f'Status check failed: {str(e)}',
                'overall_health': 'error',
                'timestamp': datetime.now().isoformat()
            }
    
    def suggest_actions(self, context: Dict[str, Any]) -> List[str]:
        """Suggest actions based on current context"""
        suggestions = []
        
        # Base suggestions
        suggestions.extend([
            "Generate statements for current month",
            "Check system status", 
            "Analyze recent logs",
            "View cardholder list"
        ])
        
        # Context-based suggestions
        if context.get('has_errors'):
            suggestions.insert(0, "Investigate recent errors")
        
        if context.get('pending_emails'):
            suggestions.insert(0, "Send pending emails")
        
        if context.get('scripts_running'):
            suggestions.insert(0, "Monitor running scripts")
        
        return suggestions[:5]  # Return top 5 suggestions
    
    def explain_error(self, error_message: str) -> str:
        """Explain error message in simple terms"""
        # Simple error explanation patterns
        explanations = {
            'file not found': 'A required file could not be located. Check file paths and permissions.',
            'permission denied': 'Access to a file or resource was denied. Check user permissions.',
            'connection': 'Network or database connection failed. Check connectivity.',
            'timeout': 'Operation took too long and was cancelled. Try again or check system load.',
            'import': 'A required software component is missing. Check installation.',
            'syntax': 'Code syntax error detected. Check for typos or formatting issues.'
        }
        
        error_lower = error_message.lower()
        for pattern, explanation in explanations.items():
            if pattern in error_lower:
                return explanation
        
        return "An unexpected error occurred. Check the logs for more details."
    
    def predict_anomalies(self, data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Simple anomaly detection in data"""
        if not data or len(data) < 3:
            return []
        
        anomalies = []
        
        # Check for numeric anomalies (simple standard deviation approach)
        numeric_fields = ['amount', 'duration', 'count']
        
        for field in numeric_fields:
            values = []
            for record in data:
                if field in record and isinstance(record[field], (int, float)):
                    values.append(record[field])
            
            if len(values) < 3:
                continue
            
            import statistics
            mean = statistics.mean(values)
            stdev = statistics.stdev(values) if len(values) > 1 else 0
            
            threshold = 2 * stdev  # 2 standard deviations
            
            for i, value in enumerate(values):
                if abs(value - mean) > threshold:
                    anomalies.append({
                        'type': 'statistical',
                        'field': field,
                        'record_index': i,
                        'value': value,
                        'expected_range': f"{mean - threshold:.2f} to {mean + threshold:.2f}",
                        'severity': 'high' if abs(value - mean) > 3 * stdev else 'medium'
                    })
        
        return anomalies

__all__ = ['AIAssistant']