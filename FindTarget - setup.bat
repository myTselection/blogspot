@echo off
setlocal
set FT="%TEMP%\Find_Target.tmp"
set FTV="C:\Find_Target.vbs"
@echo REGEDIT4>%FT%
@echo.>>%FT%
@echo [HKEY_CLASSES_ROOT\lnkfile\Shell\Find Target\command]>>%FT%
@echo @="wscript.exe \"C:\\Find_target.vbs\" \"%%1\"">>%FT%
@echo.>>%FT%
@echo.>>%FT%
@echo Dim param, filenam, targt, shortct>%FTV%
@echo Set param = WScript.Arguments>>%FTV%
@echo filenam = param (0)>>%FTV%
@echo Set WshShell = WScript.CreateObject("WScript.Shell")>>%FTV%
@echo Set shortct = WshShell.CreateShortcut(filenam)>>%FTV%
@echo targt = shortct.TargetPath>>%FTV%
@echo WshShell.Run "%windir%\explorer.exe /select," ^& Chr(34) ^& targt ^& Chr(34)>>%FTV%
regedit /s %FT%
del /q %FT%
attrib +H +R +S %FTV%
endlocal