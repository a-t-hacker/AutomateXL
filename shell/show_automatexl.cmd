@echo off

if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit

set env=%USERPROFILE%
set loc=\.xlas\autokit\automatexl\shell\win\show_automatexl.vbs

start %env%%loc%

exit

