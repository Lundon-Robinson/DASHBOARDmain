"""
Configuration Management
========================

Centralized configuration system with environment variables,
file-based settings, and runtime configuration management.
"""

import os
import json
import yaml
from pathlib import Path
from typing import Any, Dict, Optional
from pydantic import BaseModel, Field
from dotenv import load_dotenv

class DatabaseConfig(BaseModel):
    """Database configuration settings"""
    url: str = Field(default="sqlite:///data/dashboard.db")
    pool_size: int = Field(default=10)
    max_overflow: int = Field(default=20)
    echo: bool = Field(default=False)

class EmailConfig(BaseModel):
    """Email configuration settings"""
    smtp_server: str = Field(default="smtp.outlook.com")
    smtp_port: int = Field(default=587)
    use_tls: bool = Field(default=True)
    username: str = Field(default="")
    password: str = Field(default="")
    
class AIConfig(BaseModel):
    """AI assistant configuration"""
    openai_api_key: str = Field(default="")
    model: str = Field(default="gpt-3.5-turbo")
    max_tokens: int = Field(default=2048)
    temperature: float = Field(default=0.7)

class UIConfig(BaseModel):
    """UI configuration settings"""
    theme: str = Field(default="dark")
    window_width: int = Field(default=1400)
    window_height: int = Field(default=900)
    font_family: str = Field(default="Segoe UI")
    font_size: int = Field(default=10)

class PathConfig(BaseModel):
    """File path configurations"""
    data_dir: Path = Field(default_factory=lambda: Path("data"))
    logs_dir: Path = Field(default_factory=lambda: Path("logs"))
    temp_dir: Path = Field(default_factory=lambda: Path("temp"))
    exports_dir: Path = Field(default_factory=lambda: Path("exports"))
    templates_dir: Path = Field(default_factory=lambda: Path("templates"))

class DashboardConfig(BaseModel):
    """Main dashboard configuration"""
    debug: bool = Field(default=False)
    version: str = Field(default="1.0.0")
    database: DatabaseConfig = Field(default_factory=DatabaseConfig)
    email: EmailConfig = Field(default_factory=EmailConfig)
    ai: AIConfig = Field(default_factory=AIConfig)
    ui: UIConfig = Field(default_factory=UIConfig)
    paths: PathConfig = Field(default_factory=PathConfig)
    
    def __init__(self, config_path: Optional[str] = None, **kwargs):
        # Load environment variables
        load_dotenv()
        
        # Load from file if provided
        config_data = {}
        if config_path and os.path.exists(config_path):
            config_data = self._load_config_file(config_path)
        
        # Override with environment variables
        config_data.update(self._load_from_env())
        
        # Override with kwargs
        config_data.update(kwargs)
        
        super().__init__(**config_data)
        
        # Create directories
        self._create_directories()
    
    def _load_config_file(self, config_path: str) -> Dict[str, Any]:
        """Load configuration from file"""
        path = Path(config_path)
        
        if path.suffix.lower() == '.json':
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        elif path.suffix.lower() in ['.yml', '.yaml']:
            with open(path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f) or {}
        else:
            raise ValueError(f"Unsupported config file format: {path.suffix}")
    
    def _load_from_env(self) -> Dict[str, Any]:
        """Load configuration from environment variables"""
        env_config = {}
        
        # Database
        if os.getenv('DATABASE_URL'):
            env_config.setdefault('database', {})['url'] = os.getenv('DATABASE_URL')
        
        # Email
        if os.getenv('SMTP_USERNAME'):
            env_config.setdefault('email', {})['username'] = os.getenv('SMTP_USERNAME')
        if os.getenv('SMTP_PASSWORD'):
            env_config.setdefault('email', {})['password'] = os.getenv('SMTP_PASSWORD')
        
        # AI
        if os.getenv('OPENAI_API_KEY'):
            env_config.setdefault('ai', {})['openai_api_key'] = os.getenv('OPENAI_API_KEY')
        
        # Debug mode
        if os.getenv('DEBUG'):
            env_config['debug'] = os.getenv('DEBUG').lower() in ['true', '1', 'yes']
        
        return env_config
    
    def _create_directories(self):
        """Create necessary directories"""
        for path in [self.paths.data_dir, self.paths.logs_dir, 
                     self.paths.temp_dir, self.paths.exports_dir,
                     self.paths.templates_dir]:
            path.mkdir(parents=True, exist_ok=True)
    
    def save_to_file(self, config_path: str):
        """Save configuration to file"""
        path = Path(config_path)
        config_dict = self.dict()
        
        # Convert Path objects to strings for serialization
        for key, value in config_dict.items():
            if isinstance(value, dict):
                for sub_key, sub_value in value.items():
                    if isinstance(sub_value, Path):
                        value[sub_key] = str(sub_value)
        
        if path.suffix.lower() == '.json':
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(config_dict, f, indent=2)
        elif path.suffix.lower() in ['.yml', '.yaml']:
            with open(path, 'w', encoding='utf-8') as f:
                yaml.safe_dump(config_dict, f, default_flow_style=False)
        else:
            raise ValueError(f"Unsupported config file format: {path.suffix}")

# Global configuration instance
config = DashboardConfig()