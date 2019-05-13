@echo off
TITLE SVN - Full file history
REM Original source: http://stackoverflow.com/questions/282802/how-can-i-view-all-historical-changes-to-a-file-in-svn and http://stackoverflow.com/questions/5622367/generate-history-of-changes-on-a-file-in-svn/5721533#5721533
echo Copy this bat script next to the checked out svn file on which to get full svn history. Drag and drop the svn file onto the bat script to start fetching the info (or a open command window and provide the name of the svn file as first parameter to the bat script execution)
if "%1%"=="" pause
set file=%1
set report=%file%-FullSvnHistory.txt

if [%file%] == [] (
  echo Usage: "%0 <file>"
  exit /b
)

echo Retrieving svn history of file, please wait...
echo The report will be saved in the file: %report%.
echo To stop the process press Ctrl+c.

rem first revision as full text
for /F "tokens=1 delims=-r " %%R in ('"svn log -q %file%"') do (
  svn log -r %%R %file% > %report%
  svn cat -r %%R %file% >> %report%
  goto :diffs
)

:diffs

rem remaining revisions as differences to previous revision
for /F "skip=2 tokens=1 delims=-r " %%R in ('"svn log -q %file%"') do (
  echo.
  svn log -r %%R %file% >> %report%
  svn diff -c %%R %file% >> %report%
)