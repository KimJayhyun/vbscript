@echo off

set folderName=D:\RPA\DB_Backup\Backup\public_%1\%2
set fileName=%folderName%\dump-%2.sql

If not Exist %folderName% mkdir %folderName%
If Exist %fileName% del %fileName%

timeout /t 2

@REM mysqldump -u root -pP@ssw0rd --default-character-set utf8 backup_test test > "%fileName%"
@REM mysqldump -u public_super -ppublic_super12#$ --default-character-set utf8 public_%1 > "%fileName%"
mysqldump -u root -pP@ssw0rd public_%1 > "%fileName%"

	