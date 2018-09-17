@ECHO OFF
net session >nul 2>&1
    if %errorLevel% == 0 (
    if exist BRU-uninstall-helpers\streams.exe BRU-uninstall-helpers\streams.exe /accepteula -s -d *.bat
    if exist BRU-uninstall-helpers\streams.exe BRU-uninstall-helpers\streams.exe /accepteula -s -d *.ps1
    if exist BRU-uninstall-helpers\streams.exe BRU-uninstall-helpers\streams.exe /accepteula -s -d BRU-uninstall-helpers\*.exe
    if exist BRU-uninstall-helpers\streams.exe BRU-uninstall-helpers\streams.exe /accepteula -s -d BRU-uninstall-helpers\*.dll
    if exist BRU-uninstall-helpers\streams.exe BRU-uninstall-helpers\streams.exe /accepteula -s -d BRU-uninstall-helpers\*.vbs
	PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell.exe -ArgumentList '-NoProfile -NoExit -ExecutionPolicy Bypass -WindowStyle Hidden -File ""%~dpn0.ps1""' -Verb RunAs}"
    exit
    ) else (
        echo You must be logged in as a member of the Adminstrators group and right-click
        echo this batch file then "Run as Administrator" for the PowerShell script
        echo to execute properly.
    )    
PAUSE
