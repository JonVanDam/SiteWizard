@echo off
rem ---------------------------------------------------------------
rem  SiteWizard launcher - double-click this file to start the app.
rem  Works from any location as long as it stays in the project folder.
rem ---------------------------------------------------------------
setlocal
cd /d "%~dp0"

set "ELECTRON=node_modules\electron\dist\electron.exe"

if not exist "%ELECTRON%" (
    echo SiteWizard: dependencies not installed yet.
    echo Running "npm install" - this only happens once and may take a few minutes.
    echo.
    call npm install
    if errorlevel 1 (
        echo.
        echo ---------------------------------------------------------------
        echo  npm install failed.
        echo.
        echo  Make sure Node.js ^(LTS^) is installed: https://nodejs.org/
        echo  Then double-click this file again.
        echo ---------------------------------------------------------------
        echo.
        pause
        exit /b 1
    )
)

if not exist "%ELECTRON%" (
    echo.
    echo SiteWizard: could not find Electron at "%ELECTRON%" even after install.
    echo Try deleting the node_modules folder and running this file again.
    echo.
    pause
    exit /b 1
)

start "SiteWizard" "%ELECTRON%" "."
exit /b 0
