@ECHO OFF

ECHO === GL Account Clearing Setup ver. 1.1.20220616 ===
set /P choice="Do you wish to install the application? (Y/N):"

IF NOT %choice%==Y IF NOT %choice%==y (
  pause>nul|set/p=Installation aborted. Press any key to exit setup ...
  EXIT /b 0
)

ECHO Installing application...
ECHO:

IF NOT EXIST "env" (
  MKDIR env
) ELSE (

  ECHO Detecting python virtual environment ...
  SETLOCAL enabledelayedexpansion

  IF EXIST "env/Scripts/python.exe" (

    SET/P choice="A python virtual environment already exists. Do you wish to reinstall? (Y/N):"
    
    IF NOT !choice!==Y IF NOT !choice!==y (
      pause>nul|set/p=Installation aborted. Press any key to exit setup ...
      EXIT /b 0
    )

    ECHO Removing existing virtual environment ...
    RD /s /q env

  )

  ENDLOCAL
)

ECHO Creating virtual environment ...
C:\bia\_pyEnv\python.exe -m venv env
ECHO:

ECHO Updating virtual environment ...
env\Scripts\python.exe -m pip install --upgrade pip
env\Scripts\python.exe -m pip install --upgrade setuptools
ECHO:

ECHO Creating application folders ...
md temp\exports
ECHO:

ECHO Installing packages ...
env\Scripts\python.exe -m pip install -r reqs.txt
pause>nul|set/p=Installation completed. Press any key to exit setup ...
