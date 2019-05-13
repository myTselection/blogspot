@echo off
REG ADD "HKCR\Directory\shell\cmd" /ve /d "Open &Command Window Here" /f
REG ADD "HKCR\Directory\shell\cmd\command" /ve /d "cmd.exe /k cd \"%%L\"" /f
REG ADD "HKCR\*\shell\cmd" /ve /d "Open &Command Window Here" /f
REG ADD "HKCR\*\shell\cmd\command" /ve /d "cmd.exe /k echo %%L" /f