@ECHO ON
ASSOC .j2=jinja2file
ASSOC .jinja=jinja2file
ASSOC .jinja2=jinja2file
@ECHO OFF

IF EXIST "%~dp0xls2any.bat" (
    @ECHO ON
    FTYPE jinja2file="%~dp0xls2any.bat" "%%1"
    @ECHO OFF
) ELSE (
    @ECHO ON
    FTYPE jinja2file="%~dp0xls2any.exe" "%%1"
    @ECHO OFF
)

TIMEOUT 6
