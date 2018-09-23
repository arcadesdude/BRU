@ECHO OFF

net session >nul 2>&1

    if %errorLevel% == 0 (

        if exist BRU-uninstall-helpers\streams.exe (
            BRU-uninstall-helpers\streams.exe /accepteula -s -d *.bat
            BRU-uninstall-helpers\streams.exe /accepteula -s -d *.ps1
            BRU-uninstall-helpers\streams.exe /accepteula -s -d BRU-uninstall-helpers\*.exe
            BRU-uninstall-helpers\streams.exe /accepteula -s -d BRU-uninstall-helpers\*.dll
            BRU-uninstall-helpers\streams.exe /accepteula -s -d BRU-uninstall-helpers\*.vbs
        )

	REM PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell.exe -ArgumentList '-NoProfile -NoExit -ExecutionPolicy Bypass -WindowStyle Hidden -File ""%~dpn0.ps1""' -Verb RunAs}"
    PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell.exe -ArgumentList '-NoProfile -NoExit -ExecutionPolicy Bypass -File ""%~dpn0.ps1""' -Verb RunAs}"
    
    PAUSE

    REM exit

    ) else (

        echo You must be logged in as a member of the Adminstrators group and right-click
        echo this batch file then "Run as Administrator" for the PowerShell script
        echo to execute properly.

    )    

PAUSE
