@echo off

:POWERSHELL
powershell.exe  -windowstyle hidden -File "spotlight.ps1"
GOTO :END

:OLD_BATCH_SCRIPT
setlocal EnableDelayedExpansion
if not exist "%LOCALAPPDATA%\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets\" goto :END
if not exist "%USERPROFILE%\Pictures\Spotlight\" mkdir "%USERPROFILE%\Pictures\Spotlight\"
cd "%LOCALAPPDATA%\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets\"
for /f "tokens=*" %%f in ('dir /b *.*') do (
  SET newname=%%f.jpg
  copy /Y "%%f" "%USERPROFILE%\Pictures\Spotlight\!newname!"
)
:END
exit
