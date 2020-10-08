@echo off

SET WIN_RAR_APP=c:\program files\winrar\rar.exe
SET BACKUP_LOGS_DIR=backup_logs
setlocal ENABLEDELAYEDEXPANSION

SET FILES=
for /f "delims=" %%i in ('dir /a-d/s/b "*.log"') do SET FILES=!FILES! "%%i"

mkdir "%BACKUP_LOGS_DIR%"

SET Today=%Date: =.%
SET Today=%Today:~-4%.%Date:~-7,2%.%Date:~-10,2%
SET Now=%Time: =0%
SET Now=%Now::=.%
set archivename="logs_%Today%-%Now%.rar"
set archivename="%archivename: =%"

"%WIN_RAR_APP%" m "%BACKUP_LOGS_DIR%\%archivename%" %files%