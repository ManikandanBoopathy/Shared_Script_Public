@ECHO OFF
Echo.

:checkPrivileges 
rem to make sure net file works the Server service has to be running 
Rem net start server
NET FILE 1>NUL 2>NUL
if '%errorlevel%' == '0' ( goto gotPrivileges ) else ( goto getPrivileges ) 

:getPrivileges 
if '%1'=='ELEV' (shift & goto gotPrivileges)  
ECHO. 
ECHO **************************************
ECHO Invoking UAC for Privilege Escalation 
ECHO **************************************

setlocal DisableDelayedExpansion
set "batchPath=%~0"
setlocal EnableDelayedExpansion
ECHO Set UAC = CreateObject^("Shell.Application"^) > "%temp%\OEgetPrivileges.vbs" 
ECHO UAC.ShellExecute "!batchPath!", "ELEV", "", "runas", 1 >> "%temp%\OEgetPrivileges.vbs" 
"%temp%\OEgetPrivileges.vbs" 
exit /B 

:gotPrivileges 

REM #################################
REM Specific - This will not export the list from DFS and it will not upate "LimitedListComputers.txt". This will run against the list that you have updated
REM All - This will export the list from DFS and it will update the list of MDT members but it will exclude regional servers
REM #################################

PowerShell.exe -ExecutionPolicy Bypass -noprofile -command ""%~dp0MDT_Trigger-ScheduledTask.ps1" -Target 'ALL'"

pause