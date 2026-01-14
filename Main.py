"""
SQL Data Importer - Main Entry Point

This application imports data from Excel files into a SQL Server database
with validation, error reporting, and dry-run capability.
"""
import sys
import os
import logging
from pathlib import Path

# Add the project directory to Python path
project_dir = Path(__file__).parent
sys.path.insert(0, str(project_dir))

from gui import run_gui


def setup_logging():
    """Configure application logging."""
    log_dir = project_dir / "logs"
    log_dir.mkdir(exist_ok=True)
    
    log_file = log_dir / "sql_importer.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    logger = logging.getLogger(__name__)
    logger.info("="*60)
    logger.info("SQL Data Importer Application Starting")
    logger.info("="*60)


def main():
    """Main application entry point."""
    setup_logging()
    
    # Optional: Specify config file path
    config_file = None
    if len(sys.argv) > 1:
        config_file = sys.argv[1]
        if not os.path.exists(config_file):
            print(f"Warning: Config file not found: {config_file}")
            config_file = None
    
    # Run GUI
    try:
        run_gui(config_file)
    except Exception as e:
        logging.exception("Fatal error")
        print(f"\nFatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
