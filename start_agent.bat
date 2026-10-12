@echo off
title VFSTR Academic Document Formatting Agent
echo ========================================================
echo   VFSTR Academic Document Formatting Agent (Localhost)
echo ========================================================
echo.
echo Starting agent server on http://localhost:8000 ...
echo.

:: Open default browser to localhost
start http://localhost:8000

:: Run the FastAPI application
python app.py

pause
