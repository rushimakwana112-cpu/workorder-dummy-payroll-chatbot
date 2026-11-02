-- Run first (in SSMS, connected to the new SQL Server). Change the name if you like,
-- then set DATABASE = "payroll" at the top of payroll_chatbot.py (line ~28) to the same name.
IF DB_ID('payroll') IS NULL
    CREATE DATABASE payroll;
GO
USE payroll;
GO
-- Now run 01..05 in this database.
