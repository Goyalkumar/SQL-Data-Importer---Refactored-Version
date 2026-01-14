# SQL Data Importer - Refactored Version

A professional-grade Python application for importing Excel data into SQL Server databases with validation, error reporting, and dry-run capabilities.

## 🎯 What's New in This Refactored Version

### ✅ Security Improvements
- **No hardcoded credentials** - Uses environment variables or config files
- **SQL injection protection** - All queries use proper parameterization
- **Secure password handling** - Credentials never logged or displayed

### 🏗️ Architecture Improvements
- **Modular design** - Separated concerns into dedicated modules
- **Better error handling** - Specific exception types with proper recovery
- **Testable code** - Clean separation allows for unit testing
- **Logging system** - Comprehensive logging to files and console

### 🚀 Feature Enhancements
- **Progress indicators** - Real-time feedback during long operations
- **Better error messages** - More informative and actionable errors
- **Connection testing** - Automatic connection health checks
- **Improved tooltips** - Show full file paths and metadata
- **Batch operations** - More efficient database updates

### 📊 Code Quality
- **Type hints** - Full type annotations throughout
- **Docstrings** - Comprehensive documentation for all functions
- **PEP 8 compliant** - Follows Python style guidelines
- **Error recovery** - Graceful handling of edge cases

---

## 📋 Requirements

- Python 3.8 or higher
- SQL Server with ODBC Driver 17
- Required Python packages (see requirements.txt)

---

## 🚀 Installation

### 1. Install Python Dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure Database Credentials

You have two options:

#### Option A: Environment Variables (Recommended)

1. Copy the example file:
   ```bash
   cp .env.example .env
   ```

2. Edit `.env` and add your credentials:
   ```
   DB_SERVER=your_server_name
   DB_NAME=your_database_name
   DB_USER=your_username
   DB_PASSWORD=your_password
   ```

3. Install python-dotenv:
   ```bash
   pip install python-dotenv
   ```

4. Load environment variables before running:
   ```python
   from dotenv import load_dotenv
   load_dotenv()
   ```

#### Option B: Configuration File

1. Copy the example file:
   ```bash
   cp db_config.ini.example db_config.ini
   ```

2. Edit `db_config.ini` with your credentials

3. Pass the config file when running:
   ```bash
   python main.py db_config.ini
   ```

### 3. Set Up SQL Server

Ensure you have:
- ODBC Driver 17 for SQL Server installed
- Network access to your SQL Server
- Appropriate permissions on the target database

---

## 📖 Usage

### Running the Application

#### Basic (using environment variables):
```bash
python main.py
```

#### With config file:
```bash
python main.py db_config.ini
```

### Application Workflow

1. **Check Connection Status**
   - The app automatically tests the database connection on startup
   - Green checkmark = connected, red X = connection failed

2. **Select Input File**
   - Click "Browse" next to "Input Data"
   - Select your Excel file with data to import

3. **Select Configuration File**
   - Click "Browse" next to "Mapping Config"
   - Select your Script_Config.xlsx file
   - *Note: If in the same folder as input file, it's auto-detected*

4. **Run Simulation (Dry Run)**
   - Click "🔍 RUN SIMULATION" to see what would change
   - No data is written to the database
   - Review the log for detected changes

5. **Execute Import**
   - Click "▶️ RUN REAL IMPORT" to update the database
   - Confirm the warning dialog
   - Changes are committed to the database

---

## 🏛️ Project Structure

```
sql_importer_refactored/
│
├── main.py                    # Application entry point
├── config.py                  # Configuration management
├── database.py                # Database operations
├── data_processor.py          # Excel processing & validation
├── import_service.py          # Main import orchestration
├── gui.py                     # GUI components
│
├── requirements.txt           # Python dependencies
├── .env.example               # Example environment file
├── db_config.ini.example      # Example config file
├── README.md                  # This file
│
└── logs/                      # Application logs (auto-created)
    └── sql_importer.log
```

---

## 🔧 Module Documentation

### config.py
Handles all configuration management including:
- Database credentials from environment or file
- Application settings (batch size, timeouts, etc.)
- Secure credential handling

### database.py
Manages all database operations:
- Connection management with error handling
- Secure parameterized queries
- Batch operations for efficiency
- Transaction management

### data_processor.py
Handles Excel file processing:
- Configuration loading from Excel
- Data validation and type checking
- Change detection between Excel and SQL
- Error report generation

### import_service.py
Orchestrates the import process:
- Coordinates all components
- Manages workflow steps
- Provides progress updates
- Generates operation summaries

### gui.py
User interface components:
- File selection dialogs
- Progress indicators
- Formatted logging output
- Connection status display

---

## 🔐 Security Best Practices

1. **Never commit credentials**
   - Add `.env` and `db_config.ini` to `.gitignore`
   - Use `.example` files for templates

2. **Use environment variables in production**
   - More secure than config files
   - Easier to manage in deployment

3. **Limit database permissions**
   - Use accounts with minimum required permissions
   - Consider read-only accounts for dry-run mode

4. **Review error reports**
   - Error reports may contain sensitive data
   - Store securely and delete when no longer needed

---

## 🐛 Troubleshooting

### Connection Fails

**Symptom**: Red "❌ Offline" status

**Solutions**:
1. Verify server name and database name
2. Check network connectivity
3. Confirm ODBC Driver 17 is installed
4. Test credentials with SQL Server Management Studio
5. Check firewall rules

### Import Errors

**Symptom**: "Error" status in sheet results

**Solutions**:
1. Review the execution log for specific errors
2. Check the generated error report Excel file
3. Verify column mappings in Script_Config.xlsx
4. Ensure data types match SQL table schema

### Performance Issues

**Symptom**: Slow import process

**Solutions**:
1. Increase `BATCH_SIZE` in environment variables
2. Reduce number of sheets being processed
3. Index the tag column in SQL Server
4. Check network latency to SQL Server

---

## 📊 Configuration Files

### Script_Config.xlsx Structure

Your configuration Excel file should contain these sheets:

#### 1. Column_Mapping
Maps Excel headers to SQL column names:
```
Excel_Header    | SQL_Column
----------------|------------------
Equipment Tag   | Tag Number
Description     | Description
Location        | Physical Location
```

#### 2. Allowed_Sheets
Lists which sheets to process:
```
Sheet_Name
-----------
Equipment
Instruments
Valves
```

#### 3. Ignored_Headers
Columns to skip during import:
```
Header_Name
-----------
Notes
Comments
Internal ID
```

---

## 🧪 Testing Recommendations

### Before First Use
1. **Test with dry-run** - Always run simulation first
2. **Backup database** - Create backup before real import
3. **Test with small dataset** - Start with one sheet/few rows
4. **Verify mappings** - Check column mapping configuration

### Regular Use
1. Review error reports after each import
2. Monitor execution logs for warnings
3. Validate data in database after import
4. Keep configuration files up to date

---

## 📈 Performance Tips

1. **Batch Size**: Default is 2000, adjust based on your network and data
   ```
   BATCH_SIZE=5000  # For faster networks
   BATCH_SIZE=1000  # For slower networks
   ```

2. **Database Indexes**: Ensure tag column is indexed:
   ```sql
   CREATE INDEX IX_TagNumber ON AllClasses([Tag Number])
   ```

3. **Network**: Use wired connection for large imports

4. **Concurrent Users**: Limit to avoid lock conflicts

---

## 🔄 Migration from Old Version

To migrate from the original script:

1. **Export existing configuration**:
   - Note your server, database, user, password
   - Keep your Script_Config.xlsx file

2. **Set up new configuration**:
   - Create `.env` or `db_config.ini` with credentials
   - Copy Script_Config.xlsx to working directory

3. **Test import**:
   - Run dry-run mode first
   - Compare results with old script
   - Verify error reports are similar

4. **Full transition**:
   - Run real import
   - Monitor logs carefully
   - Keep old script as backup initially

---

## 📝 Changelog

### Version 2.0 (Refactored)
- ✨ Removed hardcoded credentials
- ✨ Modular architecture with separation of concerns
- ✨ Comprehensive error handling and logging
- ✨ Progress indicators and status updates
- ✨ SQL injection protection
- ✨ Batch update optimization
- ✨ Type hints and documentation
- ✨ Professional code structure

### Version 1.0 (Original)
- Basic import functionality
- Dry-run mode
- Error reporting
- GUI interface

---

## 👥 Support

For issues or questions:
1. Check the troubleshooting section
2. Review execution logs in `logs/sql_importer.log`
3. Examine error reports generated in the output directory

---

## 📄 License

[Add your license information here]

---

## 🙏 Acknowledgments

Refactored version maintains all original functionality while adding:
- Enterprise-grade security
- Professional architecture
- Enhanced user experience
- Comprehensive documentation

---

**Happy Importing! 🚀**
