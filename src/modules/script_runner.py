"""
Script Runner Module
====================

Advanced script execution with monitoring, scheduling,
and output streaming.
"""

import os
import sys
import subprocess
import threading
import queue
import time
from pathlib import Path
from typing import Dict, List, Optional, Callable, Any
from datetime import datetime, timedelta
from dataclasses import dataclass
from enum import Enum

from ..core.logger import logger
from ..core.database import get_db_manager

class ScriptStatus(Enum):
    """Script execution status"""
    PENDING = "pending"
    RUNNING = "running"
    SUCCESS = "success"
    FAILED = "failed"
    CANCELLED = "cancelled"
    TIMEOUT = "timeout"

@dataclass
class ScriptInfo:
    """Script information"""
    name: str
    path: str
    description: str = ""
    category: str = "general"
    timeout: int = 300  # 5 minutes default
    retry_count: int = 0
    environment: Dict[str, str] = None
    working_directory: str = None

@dataclass
class ScriptExecution:
    """Script execution instance"""
    script_info: ScriptInfo
    execution_id: str
    status: ScriptStatus = ScriptStatus.PENDING
    start_time: Optional[datetime] = None
    end_time: Optional[datetime] = None
    process: Optional[subprocess.Popen] = None
    output: List[str] = None
    error_output: List[str] = None
    exit_code: Optional[int] = None
    
    def __post_init__(self):
        if self.output is None:
            self.output = []
        if self.error_output is None:
            self.error_output = []

class ScriptRunner:
    """Advanced script execution handler"""
    
    def __init__(self, config):
        self.config = config
        self.db_manager = get_db_manager(config.database.url)
        
        # Script registry
        self.scripts: Dict[str, ScriptInfo] = {}
        self.running_executions: Dict[str, ScriptExecution] = {}
        
        # Output queues for streaming
        self.output_queues: Dict[str, queue.Queue] = {}
        
        # Monitoring
        self.monitor_thread = None
        self.monitor_active = False
        
        # Auto-discovery of scripts
        self._discover_scripts()
        
        # Start monitoring thread
        self._start_monitoring()
        
        logger.info("Script runner initialized")
    
    def cleanup(self):
        """Clean up resources"""
        self.monitor_active = False
        
        # Stop all running scripts
        for execution_id, execution in list(self.running_executions.items()):
            try:
                self.stop_script(execution_id)
            except Exception as e:
                logger.warning(f"Failed to stop script {execution_id}: {e}")
        
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=5)
        
        logger.info("Script runner cleaned up")
    
    def _discover_scripts(self):
        """Auto-discover Python scripts in the repository"""
        script_files = [
            "Create Statements.py",
            "Bulk Mail.py", 
            "process_delegation.py",
            "gui.py"
        ]
        
        for script_file in script_files:
            if Path(script_file).exists():
                self.register_script(ScriptInfo(
                    name=script_file.replace('.py', '').replace(' ', '_'),
                    path=script_file,
                    description=f"Legacy script: {script_file}",
                    category="legacy"
                ))
        
        # Discover PowerShell scripts
        ps_files = Path('.').glob("*.ps1")
        for ps_file in ps_files:
            self.register_script(ScriptInfo(
                name=ps_file.stem,
                path=str(ps_file),
                description=f"PowerShell script: {ps_file.name}",
                category="powershell"
            ))
        
        logger.info(f"Discovered {len(self.scripts)} scripts")
    
    def register_script(self, script_info: ScriptInfo):
        """Register a script for execution"""
        self.scripts[script_info.name] = script_info
        logger.debug(f"Registered script: {script_info.name}")
    
    def list_scripts(self) -> List[ScriptInfo]:
        """Get list of all registered scripts"""
        return list(self.scripts.values())
    
    def get_script_by_category(self, category: str) -> List[ScriptInfo]:
        """Get scripts by category"""
        return [script for script in self.scripts.values() if script.category == category]
    
    def run_script(self, script_name: str, args: List[str] = None, 
                   output_callback: Callable[[str], None] = None) -> str:
        """Run a script and return execution ID"""
        if script_name not in self.scripts:
            raise ValueError(f"Script '{script_name}' not found")
        
        script_info = self.scripts[script_name]
        execution_id = f"{script_name}_{int(time.time())}"
        
        # Create execution instance
        execution = ScriptExecution(
            script_info=script_info,
            execution_id=execution_id
        )
        
        # Register execution
        self.running_executions[execution_id] = execution
        
        # Create output queue
        if output_callback:
            self.output_queues[execution_id] = queue.Queue()
        
        # Log execution start to database
        db_execution = self.db_manager.log_script_start(script_name)
        
        # Build command
        command = self._build_command(script_info, args)
        
        try:
            # Start process
            execution.start_time = datetime.now()
            execution.status = ScriptStatus.RUNNING
            
            env = os.environ.copy()
            if script_info.environment:
                env.update(script_info.environment)
            
            working_dir = script_info.working_directory or os.getcwd()
            
            execution.process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True,
                bufsize=1,
                cwd=working_dir,
                env=env
            )
            
            # Start output monitoring thread
            if output_callback or execution_id in self.output_queues:
                self._start_output_monitoring(execution, output_callback)
            
            logger.info(f"Started script {script_name} with PID {execution.process.pid}")
            return execution_id
            
        except Exception as e:
            execution.status = ScriptStatus.FAILED
            execution.end_time = datetime.now()
            
            # Log to database
            self.db_manager.log_script_end(
                db_execution.id,
                "failed",
                error_output=str(e)
            )
            
            logger.error(f"Failed to start script {script_name}", exception=e)
            raise
    
    def _build_command(self, script_info: ScriptInfo, args: List[str] = None) -> List[str]:
        """Build command for script execution"""
        args = args or []
        
        if script_info.path.endswith('.py'):
            command = [sys.executable, script_info.path] + args
        elif script_info.path.endswith('.ps1'):
            command = ['powershell.exe', '-ExecutionPolicy', 'Bypass', '-File', script_info.path] + args
        elif script_info.path.endswith('.bat'):
            command = [script_info.path] + args
        else:
            # Assume it's an executable
            command = [script_info.path] + args
        
        return command
    
    def _start_output_monitoring(self, execution: ScriptExecution, 
                                callback: Callable[[str], None] = None):
        """Start monitoring script output"""
        def monitor_output():
            try:
                # Monitor stdout
                while execution.process.poll() is None:
                    line = execution.process.stdout.readline()
                    if line:
                        line = line.strip()
                        execution.output.append(line)
                        
                        if callback:
                            callback(line)
                        
                        if execution.execution_id in self.output_queues:
                            self.output_queues[execution.execution_id].put(line)
                
                # Get remaining output
                remaining_stdout, remaining_stderr = execution.process.communicate()
                
                if remaining_stdout:
                    for line in remaining_stdout.strip().split('\n'):
                        if line:
                            execution.output.append(line)
                            if callback:
                                callback(line)
                
                if remaining_stderr:
                    for line in remaining_stderr.strip().split('\n'):
                        if line:
                            execution.error_output.append(line)
                            if callback:
                                callback(f"ERROR: {line}")
                
            except Exception as e:
                logger.error(f"Output monitoring error for {execution.execution_id}", exception=e)
        
        thread = threading.Thread(target=monitor_output, daemon=True)
        thread.start()
    
    def get_script_output(self, execution_id: str) -> List[str]:
        """Get script output lines"""
        if execution_id not in self.running_executions:
            raise ValueError(f"Execution {execution_id} not found")
        
        execution = self.running_executions[execution_id]
        return execution.output.copy()
    
    def get_script_status(self, execution_id: str) -> ScriptStatus:
        """Get script execution status"""
        if execution_id not in self.running_executions:
            raise ValueError(f"Execution {execution_id} not found")
        
        return self.running_executions[execution_id].status
    
    def stop_script(self, execution_id: str) -> bool:
        """Stop a running script"""
        if execution_id not in self.running_executions:
            raise ValueError(f"Execution {execution_id} not found")
        
        execution = self.running_executions[execution_id]
        
        if execution.process and execution.process.poll() is None:
            try:
                # Try graceful termination first
                execution.process.terminate()
                
                # Wait a bit for graceful shutdown
                try:
                    execution.process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    # Force kill if necessary
                    execution.process.kill()
                    execution.process.wait()
                
                execution.status = ScriptStatus.CANCELLED
                execution.end_time = datetime.now()
                execution.exit_code = execution.process.returncode
                
                logger.info(f"Stopped script {execution.script_info.name}")
                return True
                
            except Exception as e:
                logger.error(f"Failed to stop script {execution_id}", exception=e)
                return False
        
        return False
    
    def _start_monitoring(self):
        """Start the monitoring thread"""
        def monitor():
            while self.monitor_active:
                try:
                    # Check running executions
                    for execution_id, execution in list(self.running_executions.items()):
                        if execution.process and execution.process.poll() is not None:
                            # Process has finished
                            execution.end_time = datetime.now()
                            execution.exit_code = execution.process.returncode
                            
                            if execution.exit_code == 0:
                                execution.status = ScriptStatus.SUCCESS
                            else:
                                execution.status = ScriptStatus.FAILED
                            
                            # Log to database
                            self._log_execution_complete(execution)
                            
                            # Clean up
                            if execution_id in self.output_queues:
                                del self.output_queues[execution_id]
                        
                        # Check for timeouts
                        elif execution.status == ScriptStatus.RUNNING:
                            runtime = datetime.now() - execution.start_time
                            if runtime.total_seconds() > execution.script_info.timeout:
                                logger.warning(f"Script {execution.script_info.name} timed out")
                                self.stop_script(execution_id)
                                execution.status = ScriptStatus.TIMEOUT
                    
                    # Clean up completed executions older than 1 hour
                    cutoff_time = datetime.now() - timedelta(hours=1)
                    completed_executions = [
                        exec_id for exec_id, execution in self.running_executions.items()
                        if execution.status in [ScriptStatus.SUCCESS, ScriptStatus.FAILED, 
                                              ScriptStatus.CANCELLED, ScriptStatus.TIMEOUT]
                        and execution.end_time and execution.end_time < cutoff_time
                    ]
                    
                    for exec_id in completed_executions:
                        del self.running_executions[exec_id]
                    
                    time.sleep(1)  # Check every second
                    
                except Exception as e:
                    logger.error("Monitoring thread error", exception=e)
                    time.sleep(5)  # Wait longer if there's an error
        
        self.monitor_active = True
        self.monitor_thread = threading.Thread(target=monitor, daemon=True)
        self.monitor_thread.start()
        
        logger.info("Script monitoring started")
    
    def _log_execution_complete(self, execution: ScriptExecution):
        """Log completed execution to database"""
        try:
            # Find the database execution record
            with self.db_manager.get_session() as session:
                from ..core.database import ScriptExecution as DBScriptExecution
                
                db_execution = session.query(DBScriptExecution).filter(
                    DBScriptExecution.script_name == execution.script_info.name,
                    DBScriptExecution.start_time >= execution.start_time - timedelta(seconds=5)
                ).first()
                
                if db_execution:
                    self.db_manager.log_script_end(
                        db_execution.id,
                        execution.status.value,
                        exit_code=execution.exit_code,
                        output='\\n'.join(execution.output),
                        error_output='\\n'.join(execution.error_output)
                    )
        except Exception as e:
            logger.error("Failed to log execution completion", exception=e)
    
    def get_execution_history(self, limit: int = 50) -> List[Dict[str, Any]]:
        """Get recent script execution history"""
        try:
            with self.db_manager.get_session() as session:
                from ..core.database import ScriptExecution as DBScriptExecution
                
                executions = session.query(DBScriptExecution).order_by(
                    DBScriptExecution.start_time.desc()
                ).limit(limit).all()
                
                history = []
                for execution in executions:
                    history.append({
                        'script_name': execution.script_name,
                        'start_time': execution.start_time,
                        'end_time': execution.end_time,
                        'status': execution.status,
                        'exit_code': execution.exit_code,
                        'duration_seconds': execution.duration_seconds
                    })
                
                return history
        except Exception as e:
            logger.error("Failed to get execution history", exception=e)
            return []
    
    def get_running_scripts(self) -> List[Dict[str, Any]]:
        """Get currently running scripts"""
        running = []
        for execution_id, execution in self.running_executions.items():
            if execution.status == ScriptStatus.RUNNING:
                runtime = datetime.now() - execution.start_time if execution.start_time else timedelta(0)
                running.append({
                    'execution_id': execution_id,
                    'script_name': execution.script_info.name,
                    'start_time': execution.start_time,
                    'runtime_seconds': runtime.total_seconds(),
                    'pid': execution.process.pid if execution.process else None
                })
        return running
    
    def schedule_script(self, script_name: str, schedule_time: datetime, 
                       args: List[str] = None) -> str:
        """Schedule a script to run at a specific time"""
        # This would integrate with a scheduler like APScheduler
        # For now, just log the intent
        logger.info(f"Schedule request: {script_name} at {schedule_time}")
        return f"scheduled_{script_name}_{int(schedule_time.timestamp())}"

__all__ = ['ScriptRunner', 'ScriptInfo', 'ScriptStatus', 'ScriptExecution']