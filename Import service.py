"""
Main import service that orchestrates the data import process.
"""
import pandas as pd
from typing import Callable, Optional, List, Tuple
import logging
import os
from dataclasses import dataclass

from config import DatabaseConfig, AppConfig
from database import DatabaseManager, DatabaseConnectionError, DatabaseQueryError
from data_processor import (
    ConfigLoader, DataValidator, SheetProcessor, 
    ReportGenerator, ProcessingConfig, SheetResult, ValidationError
)

logger = logging.getLogger(__name__)


@dataclass
class ImportResult:
    """Result of an import operation."""
    success: bool
    sheet_results: List[SheetResult]
    total_detected_updates: int
    total_committed_updates: int
    validation_errors: List[ValidationError]
    error_report_path: Optional[str] = None
    error_message: Optional[str] = None


class ImportService:
    """Main service for importing Excel data into SQL database."""
    
    def __init__(
        self,
        db_config: DatabaseConfig,
        app_config: AppConfig,
        log_callback: Optional[Callable[[str, Optional[str]], None]] = None
    ):
        """
        Initialize import service.
        
        Args:
            db_config: Database configuration
            app_config: Application configuration
            log_callback: Optional callback for logging (message, tag)
        """
        self.db_config = db_config
        self.app_config = app_config
        self.log_callback = log_callback or self._default_logger
        self.db_manager = DatabaseManager(db_config)
    
    def _default_logger(self, message: str, tag: Optional[str] = None):
        """Default logger if no callback provided."""
        if tag == "error":
            logger.error(message)
        elif tag == "success":
            logger.info(message)
        else:
            logger.info(message)
    
    def _log(self, message: str, tag: Optional[str] = None):
        """Internal logging method."""
        self.log_callback(message, tag)
    
    def test_connection(self) -> Tuple[bool, Optional[str]]:
        """
        Test database connection.
        
        Returns:
            Tuple of (success, error_message)
        """
        return self.db_manager.test_connection(self.app_config.connection_timeout)
    
    def run_import(
        self,
        excel_path: str,
        config_path: str,
        dry_run: bool = True,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ImportResult:
        """
        Execute the import process.
        
        Args:
            excel_path: Path to input Excel file
            config_path: Path to configuration Excel file
            dry_run: If True, only simulate (don't write to DB)
            progress_callback: Optional callback for progress updates
            
        Returns:
            ImportResult object with operation results
        """
        mode = "DRY-RUN MODE (NO CHANGES)" if dry_run else "REAL IMPORT MODE (UPDATING DB)"
        
        self._log("=" * 60, "header")
        self._log(f"{mode:^60}", "header")
        self._log("=" * 60, "header")
        self._log("")
        
        try:
            # Step 1: Load processing configuration
            self._log("Loading processing configuration...")
            processing_config = ConfigLoader.load_from_excel(config_path)
            self._log("Configuration loaded successfully.\n")
            
            # Step 2: Connect to database
            self._log(f"Connecting to SQL Server: {self.db_config.server}...")
            conn = self.db_manager.connect()
            self._log("SQL connection successful.\n")
            
            # Step 3: Get database metadata
            self._log("Fetching database metadata...")
            sql_columns, numeric_columns = self.db_manager.get_table_metadata(conn)
            self._log(f"Found {len(sql_columns)} columns ({len(numeric_columns)} numeric)\n")
            
            # Step 4: Get existing tags
            self._log("Fetching existing tags from database...")
            existing_tags = self.db_manager.get_existing_tags(conn)
            self._log(f"Found {len(existing_tags)} tags in database.\n")
            
            # Step 5: Initialize processors
            validator = DataValidator(
                numeric_columns, 
                self.app_config.float_comparison_threshold
            )
            
            sheet_processor = SheetProcessor(
                processing_config,
                sql_columns,
                validator,
                self.db_config.tag_column
            )
            
            # Step 6: Process Excel file
            self._log(f"Reading input file: {os.path.basename(excel_path)}...\n")
            
            result = self._process_excel_file(
                excel_path,
                sheet_processor,
                existing_tags,
                conn,
                dry_run,
                progress_callback
            )
            
            # Step 7: Generate summary
            self._generate_summary(result)
            
            # Close connection
            self.db_manager.close()
            
            self._log("\nPROCESS COMPLETED.", "header")
            
            return result
        
        except (DatabaseConnectionError, DatabaseQueryError) as e:
            error_msg = f"Database Error: {str(e)}"
            self._log(error_msg, "error")
            return ImportResult(
                success=False,
                sheet_results=[],
                total_detected_updates=0,
                total_committed_updates=0,
                validation_errors=[],
                error_message=error_msg
            )
        
        except ValueError as e:
            error_msg = f"Configuration Error: {str(e)}"
            self._log(error_msg, "error")
            return ImportResult(
                success=False,
                sheet_results=[],
                total_detected_updates=0,
                total_committed_updates=0,
                validation_errors=[],
                error_message=error_msg
            )
        
        except Exception as e:
            error_msg = f"Unexpected Error: {str(e)}"
            self._log(error_msg, "error")
            logger.exception("Unexpected error during import")
            return ImportResult(
                success=False,
                sheet_results=[],
                total_detected_updates=0,
                total_committed_updates=0,
                validation_errors=[],
                error_message=error_msg
            )
    
    def _process_excel_file(
        self,
        excel_path: str,
        sheet_processor: SheetProcessor,
        existing_tags: set,
        conn,
        dry_run: bool,
        progress_callback: Optional[Callable[[str], None]]
    ) -> ImportResult:
        """Process all sheets in the Excel file."""
        xls = pd.ExcelFile(excel_path)
        
        sheet_results = []
        all_updates = []
        all_validation_errors = []
        total_detected = 0
        total_committed = 0
        
        allowed_sheets = sheet_processor.config.allowed_sheets
        
        for sheet_name in xls.sheet_names:
            if sheet_name not in allowed_sheets:
                continue
            
            try:
                # Process sheet
                sheet_result, updates, validation_errors = sheet_processor.process_sheet(
                    excel_path,
                    sheet_name,
                    existing_tags,
                    pd.DataFrame(),  # Will fetch SQL data as needed
                    progress_callback
                )
                
                # Fetch SQL data for tags in this sheet if there are potential updates
                if updates:
                    tags_in_sheet = [tag for tag, _ in updates]
                    sql_data = self.db_manager.fetch_records_by_tags(
                        conn,
                        tags_in_sheet,
                        self.app_config.batch_size
                    )
                    
                    # Re-process with actual SQL data
                    sheet_result, updates, validation_errors = sheet_processor.process_sheet(
                        excel_path,
                        sheet_name,
                        existing_tags,
                        sql_data,
                        progress_callback
                    )
                
                # Handle updates
                if updates:
                    total_detected += len(updates)
                    
                    if not dry_run:
                        try:
                            committed = self.db_manager.batch_update_records(conn, updates)
                            total_committed += committed
                            sheet_result.status = "UPDATED"
                            self._log(f"✓ {sheet_name}: {committed} records updated", "success")
                        except DatabaseQueryError as e:
                            sheet_result.status = "ERROR"
                            self._log(f"✗ {sheet_name}: Update failed - {e}", "error")
                    else:
                        self._log(f"○ {sheet_name}: {len(updates)} changes detected (dry-run)")
                
                sheet_results.append(sheet_result)
                all_validation_errors.extend(validation_errors)
                
            except Exception as e:
                error_msg = f"Error processing sheet '{sheet_name}': {e}"
                self._log(error_msg, "error")
                logger.exception(f"Sheet processing error: {sheet_name}")
                sheet_results.append(
                    SheetResult(sheet_name, "Processing Error", "ERROR")
                )
        
        # Generate error report if needed
        error_report_path = None
        if all_validation_errors:
            self._log(f"\nWARNING: {len(all_validation_errors)} Data Type Errors Found!", "error")
            error_report_path = ReportGenerator.generate_report_filename(
                excel_path, "ImportErrors"
            )
            
            if ReportGenerator.generate_error_report(all_validation_errors, error_report_path):
                self._log(f"Error report generated: {os.path.basename(error_report_path)}", "error")
        else:
            self._log("\nNo data type errors found.")
        
        return ImportResult(
            success=True,
            sheet_results=sheet_results,
            total_detected_updates=total_detected,
            total_committed_updates=total_committed,
            validation_errors=all_validation_errors,
            error_report_path=error_report_path
        )
    
    def _generate_summary(self, result: ImportResult):
        """Generate and log summary of import operation."""
        self._log("\n" + "=" * 60)
        
        # Header
        row_format = "{:<25} | {:<15} | {:<10}"
        self._log(row_format.format("Sheet Name", "Result", "Status"), "header")
        self._log("-" * 60)
        
        # Sheet results
        for sheet_result in result.sheet_results:
            status_tag = "success" if sheet_result.status == "UPDATED" else None
            
            # Truncate long sheet names
            display_name = sheet_result.sheet_name
            if len(display_name) > 24:
                display_name = display_name[:22] + '..'
            
            self._log(
                row_format.format(display_name, sheet_result.result, sheet_result.status),
                status_tag
            )
        
        self._log("=" * 60)
        
        # Totals
        self._log(f"\nTotal Detected Updates: {result.total_detected_updates}")
        if result.total_committed_updates > 0:
            self._log(
                f"Total Committed to DB:  {result.total_committed_updates}",
                "success"
            )
