@ECHO OFF
net session >nul 2>&1
    if %errorLevel% == 0 (
	PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell.exe -ArgumentList '-NoProfile -NoExit -ExecutionPolicy Bypass -WindowStyle Hidden -File ""%~dpn0.ps1""' -Verb RunAs}"
    exit
    ) else (
        echo You must be logged in as a member of the Adminstrators group and right-click
        echo this batch file then "Run as Administrator" for the PowerShell script
        echo to execute properly.
    )    
PAUSE
