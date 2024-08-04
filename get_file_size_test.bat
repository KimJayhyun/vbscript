@REM @echo off

set folderName=D:
set fileName=%folderName%\0126.sql

For /F "usebackq" %%A In ('%fileName%') Do set size=%%~zA

echo %size%