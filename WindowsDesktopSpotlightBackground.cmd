@echo off
setlocal EnableDelayedExpansion
IF NOT EXIST "%HOMEPATH%\Pictures\Spotlight" mkdir "%HOMEPATH%\Pictures\Spotlight\"
IF NOT EXIST "%LOCALAPPDATA%\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets" goto :END
cd "%LOCALAPPDATA%\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets"
for /f "tokens=*" %%f in ('dir /b *.*') do (
  SET newname=%%f.jpg
  copy /Y "%%f" "%HOMEPATH%\Pictures\Spotlight\!newname!"
)
:END
exit