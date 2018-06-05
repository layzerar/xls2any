@ECHO OFF
CD "%~dp0"

@ECHO ON
.\xls2any.exe %*
@ECHO OFF

PAUSE
TIMEOUT 6
