@ECHO OFF
ERASE .\*.trc > nul 2> nul
BWSSOTestTool %*
IF %ERRORLEVEL% EQU 1 (
    ECHO.
    ECHO.
    ECHO Connection attempt failed with exit code %ERRORLEVEL%
    ECHO.
    ECHO.
    ECHO Trace file listings:
    TYPE .\*.trc 2> nul
)
net config workstation | findstr /C:"Full Computer name" > BWSSOhostname.txt 2> nul