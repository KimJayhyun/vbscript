FOR /F " tokens=2* USEBACKQ" %%F IN (`nslookup myip.opendns.com resolver1.opendns.com`) DO (
SET getip=%%F
)

echo %getip% > myip.txt

@REM :checkForFile
@REM If exist myip.txt GOTO foundIt

@REM TIMEOUT /t 2 >nul

@REM GOTO checkForFile

@REM :foundIt
@REM TIMEOUT /t 3 >nul