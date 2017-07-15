@echo off

REM Step1: remove old version
del ..\Installer\AddIns\Perfect_Lecture\*.py

del %APPDATA%\Microsoft\AddIns\Perfect_Lecture\*.py

REM Step2: copy files
xcopy /Y /C /R /Q *.py ..\Installer\AddIns\Perfect_Lecture\
xcopy /Y /C /R /Q *.py %APPDATA%\Microsoft\AddIns\\Perfect_Lecture\

pause