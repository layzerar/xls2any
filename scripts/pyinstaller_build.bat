@ECHO ON

CD %~dp0\..\

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

PAUSE
