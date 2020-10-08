@echo off
set ECLIPSE_BIN=R:\tools\eclipse\eclipse.exe
set ECLIPSE_OPTIONS=-refresh -showlocation -Xmx512M -XX:MaxPermSize=512m
set PROJECT_FULL_PATH=%1%
set PROJECT_FOLDER=%PROJECT_FULL_PATH:.project=%
cd %PROJECT_FOLDER%
cd ..
set PROJECT_WORKSPACE="%CD%"

start "eclipse" "%ECLIPSE_BIN%" %ECLIPSE_OPTIONS% -data %PROJECT_WORKSPACE%
@echo on