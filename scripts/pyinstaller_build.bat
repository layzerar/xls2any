@ECHO OFF
CD "%~dp0\..\"

IF EXIST ".\build" RMDIR /S ".\build"
IF EXIST ".\dist" RMDIR /S ".\dist"

pyinstaller ^
    --clean ^
    --console ^
    --onefile ^
    --exclude-module=lib2to3 ^
    --exclude-module=ssl ^
    --exclude-module=win32api ^
    --exclude-module=win32com ^
    --exclude-module=win32ui ^
    --exclude-module=win32wnet ^
    --name=xls2any ^
    main_.py
IF %ERRORLEVEL% GEQ 1 (
    PAUSE
    EXIT 1
)

FOR /F "tokens=*" %%x in ('.\dist\xls2any.exe --version') DO SET VER=%%x

@ECHO ON
XCOPY /S .\examples .\dist\examples\
COPY /B /Y .\scripts\assoc_j2ext.bat .\dist\
COPY /B /Y .\scripts\console.bat .\dist\
COPY /B /Y .\scripts\xls2any.bat .\dist\
RMDIR /S /Q ".\xls2any-%VER%"
RENAME .\dist\ "xls2any-%VER%"
@ECHO OFF

PAUSE
