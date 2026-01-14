# Quick Start Guide - SQL Data Importer

Get up and running in 5 minutes!

## 📦 Step 1: Install Dependencies (1 minute)

```bash
pip install pandas pyodbc openpyxl numpy python-dotenv
```

## 🔐 Step 2: Set Up Credentials (2 minutes)

### Option A: Using .env file (Recommended)

Create a file named `.env` in the project directory:

```bash
DB_SERVER=SRVEV357
DB_NAME=System_Architecture_Report
DB_USER=sa
DB_PASSWORD=your_actual_password
```

### Option B: Using config file

Create `db_config.ini`:

```ini
[database]
   DB_SERVER=your_server_name
   DB_NAME=your_database_name
   DB_USER=your_username
   DB_PASSWORD=your_password
```

## ▶️ Step 3: Run the Application (1 minute)

```bash
python main.py
```

Or with config file:
```bash
python main.py db_config.ini
```

## 🎯 Step 4: Import Your Data (1 minute)

1. **Check connection** - Should show green "✅ Online"
2. **Browse for your Excel file** - Click "Browse" next to "Input Data"
3. **Browse for Script_Config.xlsx** - Or let it auto-detect
4. **Run Simulation** - Click "🔍 RUN SIMULATION" to test first
5. **Review the log** - Check for any errors or warnings
6. **Run Real Import** - Click "▶️ RUN REAL IMPORT" when ready

## ✅ You're Done!

Check the execution log for results and any error reports generated.

---

## 🚨 Common First-Time Issues

### Issue: "Connection failed"
**Fix**: Check your credentials in `.env` or `db_config.ini`

### Issue: "ODBC Driver not found"
**Fix**: Install ODBC Driver 17 for SQL Server

### Issue: "Configuration error"
**Fix**: Make sure `.env` or `db_config.ini` exists and has correct format

### Issue: "No valid tags found"
**Fix**: Ensure your Excel tag column matches database tags

---

## 💡 Pro Tips

1. **Always test with dry-run first** - See what will change without risk
2. **Keep Script_Config.xlsx in same folder** - It will be auto-detected
3. **Check the error report** - If generated, review before re-running
4. **Backup your database** - Before first real import
5. **Review logs** - Check `logs/sql_importer.log` for detailed info

---

## 📞 Need Help?

1. Read the full README.md for detailed documentation
2. Check the Troubleshooting section
3. Review logs in the `logs/` directory
4. Verify your Script_Config.xlsx structure

Happy importing! 🎉
