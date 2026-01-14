"""
Configuration management module for SQL Data Importer.
Handles database credentials and application settings securely.
"""
import os
from dataclasses import dataclass
from typing import Optional
import logging

logger = logging.getLogger(__name__)


@dataclass
class DatabaseConfig:
    """Database configuration with secure credential handling."""
    server: str
    database: str
    uid: str
    pwd: str
    table_name: str = "AllTagslist"
    tag_column: str = "Tag Number"
    driver: str = "ODBC Driver 17 for SQL Server"
    
    @classmethod
    def from_env(cls) -> 'DatabaseConfig':
        """Load configuration from environment variables."""
        required_vars = ['DB_SERVER', 'DB_NAME', 'DB_USER', 'DB_PASSWORD']
        missing_vars = [var for var in required_vars if not os.getenv(var)]
        
        if missing_vars:
            raise ValueError(f"Missing required environment variables: {', '.join(missing_vars)}")
        
        return cls(
            server=os.getenv('DB_SERVER'),
            database=os.getenv('DB_NAME'),
            uid=os.getenv('DB_USER'),
            pwd=os.getenv('DB_PASSWORD'),
            table_name=os.getenv('DB_TABLE', 'AllTagslist'),
            tag_column=os.getenv('TAG_COLUMN', 'Tag Number')
        )
    
    @classmethod
    def from_file(cls, filepath: str) -> 'DatabaseConfig':
        """Load configuration from a config file (INI format)."""
        import configparser
        config = configparser.ConfigParser()
        config.read(filepath)
        
        if 'database' not in config:
            raise ValueError("Config file must contain [database] section")
        
        db_config = config['database']
        return cls(
            server=db_config['server'],
            database=db_config['database'],
            uid=db_config['uid'],
            pwd=db_config['pwd'],
            table_name=db_config.get('table_name', 'AllTagslist'),
            tag_column=db_config.get('tag_column', 'Tag Number')
        )
    
    def get_connection_string(self) -> str:
        """Generate ODBC connection string."""
        return (
            f"Driver={{{self.driver}}};"
            f"Server={self.server};"
            f"Database={self.database};"
            f"UID={self.uid};"
            f"PWD={self.pwd}"
        )
    
    def __repr__(self) -> str:
        """Safe string representation without exposing password."""
        return (
            f"DatabaseConfig(server='{self.server}', "
            f"database='{self.database}', uid='{self.uid}', pwd='***')"
        )


@dataclass
class AppConfig:
    """Application-level configuration."""
    batch_size: int = 2000
    tooltip_wait_time: int = 500
    float_comparison_threshold: float = 1e-6
    connection_timeout: int = 5
    
    @classmethod
    def from_env(cls) -> 'AppConfig':
        """Load app configuration from environment variables."""
        return cls(
            batch_size=int(os.getenv('BATCH_SIZE', 2000)),
            tooltip_wait_time=int(os.getenv('TOOLTIP_WAIT', 500)),
            float_comparison_threshold=float(os.getenv('FLOAT_THRESHOLD', 1e-6)),
            connection_timeout=int(os.getenv('CONNECTION_TIMEOUT', 5))
        )


def load_config(config_file: Optional[str] = None) -> tuple[DatabaseConfig, AppConfig]:
    """
    Load both database and application configuration.
    
    Args:
        config_file: Optional path to config file. If not provided, uses environment variables.
        
    Returns:
        Tuple of (DatabaseConfig, AppConfig)
    """
    try:
        if config_file and os.path.exists(config_file):
            logger.info(f"Loading database config from file: {config_file}")
            db_config = DatabaseConfig.from_file(config_file)
        else:
            logger.info("Loading database config from environment variables")
            db_config = DatabaseConfig.from_env()
        
        app_config = AppConfig.from_env()
        logger.info("Configuration loaded successfully")
        
        return db_config, app_config
    
    except Exception as e:
        logger.error(f"Failed to load configuration: {e}")
        raise
