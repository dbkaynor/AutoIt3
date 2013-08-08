@echo off
set INPUT=
SetLocal 
cls


FOR /L %%i IN (1,1,10) DO (

  %windir%\System32\reg.exe QUERY HKLM\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}\000%%i /v "DriverDesc" 2>nul
  
)


FOR /L %%i IN (10,1,100) DO (

  %windir%\System32\reg.exe QUERY HKLM\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}\00%%i /v "DriverDesc" 2>nul

)


REM Set get card number from user
echo.
echo.
SET /P INPUT=Which card number do you want? 01-99: %=%

cls


%windir%\System32\reg.exe QUERY HKLM\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}\00"%INPUT%"
pause


