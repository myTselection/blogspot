rem add the full path to the baretail executable as a system variable
rem go to Start -> Run -> control Sysdm.cpl,,3 -> Environment variables -> new -> name: baretail, value: for example: C:\Program Files\BareTail\baretail.exe
rem if the env variable is not set, the default value will be used.

setlocal ENABLEDELAYEDEXPANSION

if "%baretail%" =="" set baretail=C:\Program Files\BareTail\baretail.exe

SET FILES=
for /f "delims=" %%i in ('dir /a-d/s/b "*.log"') do SET FILES=!FILES! "%%i"
start "tail" "%baretail%" %files%
