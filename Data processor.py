"""
Data processing module for Excel file handling and validation.
"""
import pandas as pd
import numpy as np
from typing import Dict, List, Set, Tuple, Callable, Optional
from dataclasses import dataclass
import logging
import os
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

logger = logging.getLogger(__name__)


@dataclass
class ProcessingConfig:
    """Configuration for Excel processing."""
    column_mapping: Dict[str, str]
    allowed_sheets: List[str]
    ignored_headers: List[str]


@dataclass
class ValidationError:
    """Represents a data validation error."""
    sheet_name: str
    tag_number: str
    excel_header: str
    invalid_value: any
    error_reason: str


@dataclass
class SheetResult:
    """Result of processing a single sheet."""
    sheet_name: str
    result: str
    status: str  # SKIP, DRY-RUN, UPDATED, ERROR
    update_count: int = 0


class ConfigLoader:
    """Loads processing configuration from Excel file."""
    
    @staticmethod
    def load_from_excel(filepath: str) -> ProcessingConfig:
        """
        Load column mapping and settings from config Excel file.
        
        Args:
            filepath: Path to config Excel file
            
        Returns:
            ProcessingConfig object
            
        Raises:
            ValueError: If config file is invalid
        """
        try:
            logger.info(f"Loading processing config from: {filepath}")
            xls = pd.ExcelFile(filepath)
            
            # Load column mapping
            df_map = pd.read_excel(xls, "Column_Mapping").dropna(
                subset=["Excel_Header", "SQL_Column"]
            )
            column_mapping = dict(
                zip(
                    df_map["Excel_Header"].astype(str).str.strip(),
                    df_map["SQL_Column"].astype(str).str.strip()
                )
            )
            
            # Load allowed sheets
            allowed_sheets = (
                pd.read_excel(xls, "Allowed_Sheets")["Sheet_Name"]
                .dropna()
                .astype(str)
                .str.strip()
                .tolist()
            )
            
            # Load ignored headers
            ignored_headers = (
                pd.read_excel(xls, "Ignored_Headers")["Header_Name"]
                .dropna()
                .astype(str)
                .str.strip()
                .tolist()
            )
            
            logger.info(
                f"Config loaded: {len(column_mapping)} mappings, "
                f"{len(allowed_sheets)} allowed sheets"
            )
            
            return ProcessingConfig(
                column_mapping=column_mapping,
                allowed_sheets=allowed_sheets,
                ignored_headers=ignored_headers
            )
        
        except Exception as e:
            raise ValueError(f"Failed to load configuration file: {e}")


class DataValidator:
    """Validates data types and formats."""
    
    def __init__(self, numeric_columns: List[str], float_threshold: float = 1e-6):
        self.numeric_columns = numeric_columns
        self.float_threshold = float_threshold
    
    def validate_numeric_columns(
        self,
        df: pd.DataFrame,
        sheet_name: str,
        tag_column: str
    ) -> Tuple[pd.DataFrame, List[ValidationError]]:
        """
        Validate and convert numeric columns, collecting errors.
        
        Args:
            df: DataFrame to validate
            sheet_name: Name of the sheet (for error reporting)
            tag_column: Name of the tag column
            
        Returns:
            Tuple of (validated_df, list_of_errors)
        """
        errors = []
        validated_df = df.copy()
        raw_df = df.copy()
        
        for col in self.numeric_columns:
            if col not in df.columns:
                continue
            
            # Attempt conversion
            validated_df[col] = pd.to_numeric(df[col], errors="coerce")
            
            # Find conversion failures
            mask = (
                validated_df[col].isna() & 
                raw_df[col].notna() & 
                (raw_df[col].astype(str).str.strip() != "")
            )
            
            # Log errors
            for _, row in raw_df[mask].iterrows():
                errors.append(
                    ValidationError(
                        sheet_name=sheet_name,
                        tag_number=row[tag_column],
                        excel_header=col,
                        invalid_value=row[col],
                        error_reason="Type Mismatch - Expected numeric value"
                    )
                )
        
        logger.info(f"Validation found {len(errors)} type errors in sheet '{sheet_name}'")
        return validated_df, errors
    
    def compare_values(self, excel_val: any, sql_val: any) -> bool:
        """
        Compare two values considering data types.
        
        Args:
            excel_val: Value from Excel
            sql_val: Value from SQL database
            
        Returns:
            True if values are different, False if same
        """
        if excel_val is None:
            return False
        
        # Float comparison with threshold
        if isinstance(excel_val, float) and isinstance(sql_val, float):
            return abs(excel_val - sql_val) > self.float_threshold
        
        # Standard comparison
        return excel_val != sql_val


class SheetProcessor:
    """Processes individual Excel sheets."""
    
    def __init__(
        self,
        processing_config: ProcessingConfig,
        sql_columns: List[str],
        validator: DataValidator,
        tag_column: str
    ):
        self.config = processing_config
        self.sql_columns = sql_columns
        self.validator = validator
        self.tag_column = tag_column
    
    def process_sheet(
        self,
        excel_path: str,
        sheet_name: str,
        existing_tags: Set[str],
        sql_data: pd.DataFrame,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> Tuple[SheetResult, List[Tuple[str, Dict]], List[ValidationError]]:
        """
        Process a single Excel sheet and identify updates.
        
        Args:
            excel_path: Path to Excel file
            sheet_name: Name of sheet to process
            existing_tags: Set of tags that exist in database
            sql_data: DataFrame with current SQL data for matching tags
            progress_callback: Optional callback for progress updates
            
        Returns:
            Tuple of (SheetResult, list_of_updates, list_of_validation_errors)
        """
        if progress_callback:
            progress_callback(f"Processing sheet: {sheet_name}")
        
        logger.info(f"Processing sheet: {sheet_name}")
        
        # Load sheet
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        
        # Clean column names
        df.columns = (
            df.columns.str.strip()
            .str.replace('\xa0', ' ')
            .str.replace(r'[\n\r\t]', '', regex=True)
        )
        
        # Apply column mapping
        df.rename(columns=lambda x: self.config.column_mapping.get(x, x), inplace=True)
        
        # Validate tag column exists
        if self.tag_column not in df.columns:
            logger.warning(f"Sheet '{sheet_name}' missing tag column")
            return (
                SheetResult(sheet_name, "Tag Column Missing", "SKIP"),
                [],
                []
            )
        
        # Clean and filter tags
        df[self.tag_column] = df[self.tag_column].astype(str).str.strip()
        df = df[df[self.tag_column].isin(existing_tags)]
        
        if df.empty:
            logger.info(f"Sheet '{sheet_name}' has no valid tags")
            return (
                SheetResult(sheet_name, "No Valid Tags", "SKIP"),
                [],
                []
            )
        
        # Select valid columns
        valid_columns = [
            c for c in df.columns 
            if c in self.sql_columns and c not in self.config.ignored_headers
        ]
        df = df[valid_columns + [self.tag_column]]
        
        # Validate numeric columns
        df, validation_errors = self.validator.validate_numeric_columns(
            df, sheet_name, self.tag_column
        )
        
        # Clean up nulls and empty strings
        df = self._clean_dataframe(df)
        
        # Identify updates
        updates = self._identify_updates(df, sql_data, valid_columns)
        
        if updates:
            result_str = f"{len(updates)} changes detected"
            logger.info(f"Sheet '{sheet_name}': {result_str}")
        else:
            result_str = "No changes"
            logger.info(f"Sheet '{sheet_name}': No changes detected")
        
        return (
            SheetResult(sheet_name, result_str, "DRY-RUN", len(updates)),
            updates,
            validation_errors
        )
    
    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean DataFrame by replacing NaN and empty strings with None."""
        df = df.replace({np.nan: None})
        
        # Use map() for pandas >= 2.1, applymap() for older versions
        try:
            df = df.map(lambda x: x if x is not None and x != "" else None)
        except AttributeError:
            df = df.applymap(lambda x: x if x is not None and x != "" else None)
        
        return df
    
    def _identify_updates(
        self,
        excel_df: pd.DataFrame,
        sql_df: pd.DataFrame,
        columns_to_check: List[str]
    ) -> List[Tuple[str, Dict]]:
        """
        Identify which records need updates by comparing Excel vs SQL data.
        
        Args:
            excel_df: DataFrame from Excel
            sql_df: DataFrame from SQL
            columns_to_check: List of columns to compare
            
        Returns:
            List of (tag_number, {column: new_value}) tuples
        """
        updates = []
        
        for _, excel_row in excel_df.iterrows():
            tag = excel_row[self.tag_column]
            sql_match = sql_df[sql_df[self.tag_column] == tag]
            
            if sql_match.empty:
                continue
            
            sql_row = sql_match.iloc[0]
            changes = {}
            
            for col in columns_to_check:
                excel_val = excel_row[col]
                sql_val = sql_row[col]
                
                if self.validator.compare_values(excel_val, sql_val):
                    changes[col] = excel_val
            
            if changes:
                updates.append((tag, changes))
        
        return updates


class ReportGenerator:
    """Generates Excel reports for errors and summaries."""
    
    @staticmethod
    def generate_error_report(
        validation_errors: List[ValidationError],
        output_path: str
    ) -> bool:
        """
        Generate an Excel report of validation errors.
        
        Args:
            validation_errors: List of validation errors
            output_path: Path where report should be saved
            
        Returns:
            True if report was generated successfully
        """
        try:
            error_data = [
                {
                    "Sheet Name": e.sheet_name,
                    "Tag Number": e.tag_number,
                    "Excel Header": e.excel_header,
                    "Invalid Value": e.invalid_value,
                    "Error Reason": e.error_reason
                }
                for e in validation_errors
            ]
            
            df = pd.DataFrame(error_data)
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                ReportGenerator._format_excel_output(writer, df, "Errors")
            
            logger.info(f"Error report saved: {output_path}")
            return True
        
        except Exception as e:
            logger.error(f"Failed to generate error report: {e}")
            return False
    
    @staticmethod
    def _format_excel_output(writer, df: pd.DataFrame, sheet_name: str = "Errors"):
        """Apply formatting to Excel output."""
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        
        # Header styling
        header_font = Font(bold=True, color="000000")
        header_fill = PatternFill(
            start_color="D3D3D3", 
            end_color="D3D3D3", 
            fill_type="solid"
        )
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Format header row
        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # Auto-adjust column widths and add borders
        for col in worksheet.columns:
            max_length = 0
            column_letter = col[0].column_letter
            
            for cell in col:
                cell.border = thin_border
                try:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
                except:
                    pass
            
            worksheet.column_dimensions[column_letter].width = max_length + 2
    
    @staticmethod
    def generate_report_filename(input_path: str, report_type: str = "ImportErrors") -> str:
        """
        Generate a timestamped report filename.
        
        Args:
            input_path: Path to input Excel file
            report_type: Type of report (for filename)
            
        Returns:
            Full path for the report file
        """
        input_dir = os.path.dirname(input_path)
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        return os.path.join(input_dir, f"{base_name}_{report_type}_{timestamp}.xlsx")
