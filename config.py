"""
Configuration management for ExcelBot Pro
"""

import os
from typing import Optional


class Config:
    """Application configuration."""
    
    # API Keys
    GITHUB_TOKEN: Optional[str] = os.getenv("GITHUB_TOKEN", "")
    OPENAI_API_KEY: Optional[str] = os.getenv("OPENAI_API_KEY", "")
    
    # Application Settings
    APP_NAME: str = "ExcelBot Pro"
    APP_VERSION: str = "1.0.0"
    DEFAULT_PORT: int = 7860
    DEFAULT_HOST: str = "0.0.0.0"
    
    # OpenAI Settings
    OPENAI_MODEL: str = "gpt-4"
    OPENAI_TEMPERATURE: float = 0.3
    OPENAI_MAX_TOKENS: int = 2000
    
    # File Settings
    MAX_FILE_SIZE_MB: int = 50
    ALLOWED_EXTENSIONS: list = [".xlsx", ".xls", ".xlsm"]
    
    @classmethod
    def validate(cls) -> dict:
        """Validate configuration and return status."""
        status = {
            "github_configured": bool(cls.GITHUB_TOKEN),
            "openai_configured": bool(cls.OPENAI_API_KEY),
            "warnings": []
        }
        
        if not status["github_configured"]:
            status["warnings"].append("GitHub token not configured. GitHub features disabled.")
        
        if not status["openai_configured"]:
            status["warnings"].append("OpenAI API key not configured. Using template-based generation.")
        
        return status
