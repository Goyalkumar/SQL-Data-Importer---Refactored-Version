"""
Database operations module with secure query handling.
"""
import pyodbc
import pandas as pd
from typing import List, Dict, Set, Optional, Tuple
import logging
from contextlib import contextmanager

from config import DatabaseConfig

logger = logging.getLogger(__name__)


class DatabaseConnectionError(Exception):
    """Raised when database connection fails."""
    pass


class DatabaseQueryError(Exception):
    """Raised when database query fails."""
    pass


class DatabaseManager:
    """Manages database connections and operations securely."""
    
    def __init__(self, config: DatabaseConfig):
        self.config = config
        self._connection: Optional[pyodbc.Connection] = None
    
    def test_connection(self, timeout: int = 5) -> Tuple[bool, Optional[str]]:
        """
        Test database connection.
        
        Returns:
            Tuple of (success: bool, error_message: Optional[str])
        """
        try:
            conn = pyodbc.connect(self.config.get_connection_string(), timeout=timeout)
            conn.close()
            logger.info("Database connection test successful")
            return True, None
        except Exception as e:
            error_msg = f"Connection failed: {str(e)}"
            logger.error(error_msg)
            return False, error_msg
    
    def connect(self) -> pyodbc.Connection:
        """
        Establish database connection.
        
        Returns:
            Database connection object
            
        Raises:
            DatabaseConnectionError: If connection fails
        """
        try:
            self._connection = pyodbc.connect(
                self.config.get_connection_string(),
                autocommit=False
            )
            logger.info("Database connection established")
            return self._connection
        except Exception as e:
            raise DatabaseConnectionError(f"Failed to connect to database: {e}")
    
    def close(self):
        """Close database connection."""
        if self._connection:
            try:
                self._connection.close()
                logger.info("Database connection closed")
            except Exception as e:
                logger.warning(f"Error closing connection: {e}")
            finally:
                self._connection = None
    
    @contextmanager
    def get_connection(self):
        """Context manager for database connection."""
        try:
            if not self._connection:
                self.connect()
            yield self._connection
        except Exception as e:
            if self._connection:
                self._connection.rollback()
            raise
        finally:
            # Don't close here - let the caller manage connection lifecycle
            pass
    
    def get_table_metadata(self, conn: pyodbc.Connection) -> Tuple[List[str], List[str]]:
        """
        Get table column names and identify numeric columns.
        
        Args:
            conn: Database connection
            
        Returns:
            Tuple of (all_columns, numeric_columns)
            
        Raises:
            DatabaseQueryError: If query fails
        """
        try:
            query = """
                SELECT COLUMN_NAME, DATA_TYPE 
                FROM INFORMATION_SCHEMA.COLUMNS 
                WHERE TABLE_NAME = ?
            """
            df = pd.read_sql(query, conn, params=[self.config.table_name])
            
            all_columns = df["COLUMN_NAME"].tolist()
            
            numeric_types = [
                "decimal", "numeric", "float", "real", 
                "int", "bigint", "smallint", "tinyint"
            ]
            numeric_columns = df[
                df["DATA_TYPE"].isin(numeric_types)
            ]["COLUMN_NAME"].tolist()
            
            logger.info(f"Retrieved metadata: {len(all_columns)} columns, "
                       f"{len(numeric_columns)} numeric")
            
            return all_columns, numeric_columns
        
        except Exception as e:
            raise DatabaseQueryError(f"Failed to retrieve table metadata: {e}")
    
    def get_existing_tags(self, conn: pyodbc.Connection) -> Set[str]:
        """
        Retrieve all existing tag numbers from database.
        
        Args:
            conn: Database connection
            
        Returns:
            Set of tag numbers as strings
            
        Raises:
            DatabaseQueryError: If query fails
        """
        try:
            # Use parameterized query for table/column names is not standard,
            # but we control these values from config
            query = f"SELECT [{self.config.tag_column}] FROM {self.config.table_name}"
            df = pd.read_sql(query, conn)
            
            tags = set(
                df[self.config.tag_column]
                .astype(str)
                .str.strip()
            )
            
            logger.info(f"Retrieved {len(tags)} existing tags from database")
            return tags
        
        except Exception as e:
            raise DatabaseQueryError(f"Failed to retrieve existing tags: {e}")
    
    def fetch_records_by_tags(
        self, 
        conn: pyodbc.Connection, 
        tags: List[str],
        batch_size: int = 2000
    ) -> pd.DataFrame:
        """
        Fetch records matching given tags using batched queries.
        
        Args:
            conn: Database connection
            tags: List of tag numbers to fetch
            batch_size: Size of each batch for chunked queries
            
        Returns:
            DataFrame containing all matching records
            
        Raises:
            DatabaseQueryError: If query fails
        """
        try:
            result_dfs = []
            
            for i in range(0, len(tags), batch_size):
                chunk = tags[i:i + batch_size]
                placeholders = ",".join(['?'] * len(chunk))
                
                # Safe parameterized query
                query = f"""
                    SELECT * FROM {self.config.table_name} 
                    WHERE [{self.config.tag_column}] IN ({placeholders})
                """
                
                chunk_df = pd.read_sql(query, conn, params=chunk)
                result_dfs.append(chunk_df)
            
            if not result_dfs:
                return pd.DataFrame()
            
            combined_df = pd.concat(result_dfs, ignore_index=True)
            logger.info(f"Fetched {len(combined_df)} records for {len(tags)} tags")
            
            return combined_df
        
        except Exception as e:
            raise DatabaseQueryError(f"Failed to fetch records by tags: {e}")
    
    def batch_update_records(
        self,
        conn: pyodbc.Connection,
        updates: List[Tuple[str, Dict[str, any]]]
    ) -> int:
        """
        Perform batch updates on database records.
        
        Args:
            conn: Database connection
            updates: List of (tag_number, {column: value}) tuples
            
        Returns:
            Number of records updated
            
        Raises:
            DatabaseQueryError: If update fails
        """
        try:
            cursor = conn.cursor()
            updated_count = 0
            
            for tag, changes in updates:
                if not changes:
                    continue
                
                # Build SET clause with proper SQL injection protection
                set_parts = []
                values = []
                
                for col_name, value in changes.items():
                    # Escape column names by doubling up square brackets
                    safe_col = col_name.replace(']', ']]')
                    set_parts.append(f"[{safe_col}] = ?")
                    values.append(value)
                
                set_clause = ", ".join(set_parts)
                values.append(tag)  # Add tag for WHERE clause
                
                query = f"""
                    UPDATE {self.config.table_name} 
                    SET {set_clause} 
                    WHERE [{self.config.tag_column}] = ?
                """
                
                cursor.execute(query, values)
                updated_count += cursor.rowcount
            
            conn.commit()
            logger.info(f"Successfully updated {updated_count} records")
            
            return updated_count
        
        except Exception as e:
            conn.rollback()
            raise DatabaseQueryError(f"Failed to update records: {e}")
