@echo off
REM Set Excel-DNA Logging Environment Variables Persistently

setx EXCELDNA_DIAGNOSTICS_SOURCE_LEVEL "Warning"
setx EXCELDNA_DIAGNOSTICS_LOGDISPLAY_LEVEL "Warning"
setx EXCELDNA_DIAGNOSTICS_DEBUGGER_LEVEL "Error"
setx EXCELDNA_DIAGNOSTICS_FILE_LEVEL "Verbose"
setx EXCELDNA_DIAGNOSTICS_FILE_NAME "ExcelDnaLog.txt"

echo Excel-DNA Logging Environment Variables are set.
pause