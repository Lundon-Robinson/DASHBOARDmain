"""
Legacy Script Repair Module
===========================

Automatically scan, repair, and validate existing scripts in the codebase.
Fixes hardcoded paths, improves error handling, and integrates with the dashboard.
"""

import os
import re
import shutil
from pathlib import Path
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass

from ..core.logger import logger

@dataclass
class ScriptIssue:
    """Represents an issue found in a script"""
    file_path: str
    line_number: int
    issue_type: str
    description: str
    suggested_fix: str
    severity: str  # critical, high, medium, low

@dataclass 
class RepairResult:
    """Results of script repair operation"""
    file_path: str
    issues_found: List[ScriptIssue]
    issues_fixed: List[ScriptIssue]
    backup_created: bool
    success: bool
    error_message: Optional[str] = None

class LegacyScriptRepairer:
    """Repairs and modernizes legacy scripts"""
    
    def __init__(self, config):
        self.config = config
        self.backup_dir = config.paths.data_dir / "script_backups"
        self.backup_dir.mkdir(parents=True, exist_ok=True)
    
    def scan_all_scripts(self, directory: str = ".") -> List[ScriptIssue]:
        """Scan all Python scripts in directory for issues"""
        issues = []
        python_files = Path(directory).glob("*.py")
        
        for py_file in python_files:
            if py_file.name.startswith("main_dashboard"):
                continue  # Skip our new files
            
            logger.info(f"Scanning {py_file}")
            file_issues = self._scan_script(str(py_file))
            issues.extend(file_issues)
        
        logger.info(f"Found {len(issues)} total issues across all scripts")
        return issues
    
    def _scan_script(self, file_path: str) -> List[ScriptIssue]:
        """Scan a single script for issues"""
        issues = []
        
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()
        except Exception as e:
            logger.error(f"Could not read {file_path}: {e}")
            return issues
        
        for line_num, line in enumerate(lines, 1):
            # Check for hardcoded Windows paths
            if re.search(r'[rRfF]?["\']C:\\\\', line):
                issues.append(ScriptIssue(
                    file_path=file_path,
                    line_number=line_num,
                    issue_type="hardcoded_path",
                    description="Hardcoded Windows path found",
                    suggested_fix="Use Path from pathlib or config-based paths",
                    severity="high"
                ))
            
            # Check for poor exception handling
            if "except:" in line or "except Exception:" in line:
                issues.append(ScriptIssue(
                    file_path=file_path,
                    line_number=line_num,
                    issue_type="poor_exception_handling", 
                    description="Generic exception handling",
                    suggested_fix="Use specific exception types and proper logging",
                    severity="medium"
                ))
            
            # Check for missing imports
            if "import tkinter" in line and "try:" not in lines[max(0, line_num-3):line_num+2]:
                issues.append(ScriptIssue(
                    file_path=file_path,
                    line_number=line_num,
                    issue_type="missing_optional_import",
                    description="GUI import without optional handling",
                    suggested_fix="Wrap GUI imports in try/except blocks",
                    severity="medium"
                ))
            
            # Check for hardcoded credentials or sensitive data
            if re.search(r'(password|secret|api_key)\s*=\s*["\'][^"\']+["\']', line.lower()):
                issues.append(ScriptIssue(
                    file_path=file_path,
                    line_number=line_num,
                    issue_type="hardcoded_credentials",
                    description="Hardcoded credentials detected",
                    suggested_fix="Use environment variables or secure config",
                    severity="critical"
                ))
            
            # Check for old-style string formatting
            if "%" in line and any(fmt in line for fmt in ["%s", "%d", "%f"]):
                issues.append(ScriptIssue(
                    file_path=file_path,
                    line_number=line_num,
                    issue_type="old_string_formatting",
                    description="Old-style string formatting",
                    suggested_fix="Use f-strings or .format()",
                    severity="low"
                ))
        
        return issues
    
    def repair_script(self, file_path: str, backup: bool = True) -> RepairResult:
        """Repair a single script"""
        logger.info(f"Starting repair of {file_path}")
        
        # Scan for issues
        issues = self._scan_script(file_path)
        
        if not issues:
            logger.info(f"No issues found in {file_path}")
            return RepairResult(
                file_path=file_path,
                issues_found=[],
                issues_fixed=[],
                backup_created=False,
                success=True
            )
        
        # Create backup
        backup_created = False
        if backup:
            backup_path = self.backup_dir / f"{Path(file_path).name}.backup"
            shutil.copy2(file_path, backup_path)
            backup_created = True
            logger.info(f"Backup created: {backup_path}")
        
        # Read original file
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except Exception as e:
            return RepairResult(
                file_path=file_path,
                issues_found=issues,
                issues_fixed=[],
                backup_created=backup_created,
                success=False,
                error_message=f"Could not read file: {e}"
            )
        
        # Apply fixes
        fixed_content = content
        fixed_issues = []
        
        for issue in issues:
            if issue.issue_type == "hardcoded_path":
                fixed_content, was_fixed = self._fix_hardcoded_paths(fixed_content, issue)
                if was_fixed:
                    fixed_issues.append(issue)
            
            elif issue.issue_type == "poor_exception_handling":
                fixed_content, was_fixed = self._fix_exception_handling(fixed_content, issue)
                if was_fixed:
                    fixed_issues.append(issue)
            
            elif issue.issue_type == "missing_optional_import":
                fixed_content, was_fixed = self._fix_optional_imports(fixed_content, issue)
                if was_fixed:
                    fixed_issues.append(issue)
            
            elif issue.issue_type == "old_string_formatting":
                fixed_content, was_fixed = self._fix_string_formatting(fixed_content, issue)
                if was_fixed:
                    fixed_issues.append(issue)
        
        # Write repaired file
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(fixed_content)
            
            logger.info(f"Repaired {len(fixed_issues)}/{len(issues)} issues in {file_path}")
            
            return RepairResult(
                file_path=file_path,
                issues_found=issues,
                issues_fixed=fixed_issues,
                backup_created=backup_created,
                success=True
            )
            
        except Exception as e:
            return RepairResult(
                file_path=file_path,
                issues_found=issues,
                issues_fixed=[],
                backup_created=backup_created,
                success=False,
                error_message=f"Could not write repaired file: {e}"
            )
    
    def _fix_hardcoded_paths(self, content: str, issue: ScriptIssue) -> Tuple[str, bool]:
        """Fix hardcoded Windows paths"""
        # Replace common hardcoded paths with config-based alternatives
        replacements = {
            r'r"C:\\Users\\NADLUROB\\Desktop\\Dash\\log\.txt"': 'str(config.paths.logs_dir / "dashboard.log")',
            r"r'C:\\Users\\NADLUROB\\Desktop\\Dash\\log\.txt'": 'str(config.paths.logs_dir / "dashboard.log")',
            r'r"C:\\Users\\NADLUROB\\Desktop\\test\\': 'str(config.paths.data_dir / "test" / ',
            r"r'C:\\Users\\NADLUROB\\Desktop\\test\\": 'str(config.paths.data_dir / "test" / ',
        }
        
        original_content = content
        for pattern, replacement in replacements.items():
            content = re.sub(pattern, replacement, content)
        
        # Add import for config if paths were replaced
        if content != original_content and "from src.core.config import config" not in content:
            # Find the best place to add the import
            lines = content.split('\n')
            import_line_index = 0
            for i, line in enumerate(lines):
                if line.startswith(('import ', 'from ')):
                    import_line_index = i + 1
            
            lines.insert(import_line_index, "from src.core.config import config")
            content = '\n'.join(lines)
        
        return content, content != original_content
    
    def _fix_exception_handling(self, content: str, issue: ScriptIssue) -> Tuple[str, bool]:
        """Fix poor exception handling"""
        # This is complex, so we'll make basic improvements
        original_content = content
        
        # Replace bare except: with except Exception:
        content = re.sub(r'except\s*:', 'except Exception:', content)
        
        return content, content != original_content
    
    def _fix_optional_imports(self, content: str, issue: ScriptIssue) -> Tuple[str, bool]:
        """Fix missing optional imports"""
        original_content = content
        
        # Wrap tkinter imports in try/except
        tkinter_pattern = r'(import tkinter.*)'
        replacement = '''try:
    \\1
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False
    print("Tkinter not available, GUI features disabled")'''
        
        content = re.sub(tkinter_pattern, replacement, content)
        
        return content, content != original_content
    
    def _fix_string_formatting(self, content: str, issue: ScriptIssue) -> Tuple[str, bool]:
        """Fix old-style string formatting"""
        # This is complex to do automatically, so we'll skip for now
        return content, False
    
    def repair_all_scripts(self, directory: str = ".") -> List[RepairResult]:
        """Repair all Python scripts in directory"""
        results = []
        python_files = [f for f in Path(directory).glob("*.py") 
                       if not f.name.startswith("main_dashboard")]
        
        logger.info(f"Starting repair of {len(python_files)} scripts")
        
        for py_file in python_files:
            result = self.repair_script(str(py_file))
            results.append(result)
        
        # Summary
        total_issues = sum(len(r.issues_found) for r in results)
        total_fixed = sum(len(r.issues_fixed) for r in results)
        successful_repairs = sum(1 for r in results if r.success)
        
        logger.info(f"Repair summary: {successful_repairs}/{len(results)} files processed successfully")
        logger.info(f"Fixed {total_fixed}/{total_issues} total issues")
        
        return results
    
    def create_integration_module(self, script_path: str) -> str:
        """Create integration module for legacy script"""
        script_name = Path(script_path).stem
        integration_content = f'''"""
Integration module for {script_name}
Generated by LegacyScriptRepairer
"""

import sys
import os
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.core.config import config
from src.core.logger import logger
from src.core.database import get_db_manager

# Import the original script
try:
    from {script_name.replace(" ", "_")} import *
    
    logger.info("Successfully loaded legacy script: {script_name}")
    
    def run_with_dashboard_integration():
        """Run the script with dashboard integration"""
        try:
            if hasattr(sys.modules["{script_name.replace(" ", "_")}"], "main"):
                # Call main function if it exists
                return main()
            else:
                logger.warning("No main() function found in {script_name}")
                return None
        except Exception as e:
            logger.error(f"Error running {script_name}", exception=e)
            raise
            
except ImportError as e:
    logger.error(f"Could not import {script_name}: {{e}}")
    
    def run_with_dashboard_integration():
        raise ImportError(f"Could not import {script_name}")
'''
        
        # Write integration module
        integration_path = f"src/modules/{script_name.replace(' ', '_')}_integration.py"
        with open(integration_path, 'w', encoding='utf-8') as f:
            f.write(integration_content)
        
        logger.info(f"Created integration module: {integration_path}")
        return integration_path

__all__ = ['LegacyScriptRepairer', 'ScriptIssue', 'RepairResult']