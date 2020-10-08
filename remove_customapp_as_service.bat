@echo off
echo Make sure to run this script in the correct directory! For example: R:\project\deploy\monitor\CBS-TST-APP1\remove_monitor_as_service.bat
SET MONITOR_NAME=
for %%* in (.) do set MONITOR_NAME=%%~n*
SET MONITOR_NAME=MONITOR-%MONITOR_NAME%


:STOP_SERVICE
NET STOP %MONITOR_NAME%
echo.
echo Service %MONITOR_NAME% stopped
echo.

:REMOVE_SERVICE
%INST_SERVICE_APP_PATH% %MONITOR_NAME% REMOVE
echo.
echo Service %MONITOR_NAME% removed
echo.

:UPDATE_REGISTRY
REG DELETE HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\%MONITOR_NAME%\ /va /f
echo.
echo Registry cleaned up fro service %MONITOR_NAME%
echo.

:EXIT
pause
@echo on