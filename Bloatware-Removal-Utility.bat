@ECHO OFF

net session >nul 2>&1

    if %errorLevel% == 0 (

        setlocal
        cd /d %~dp0

        if exist BRU-uninstall-helpers\streams.exe (

            BRU-uninstall-helpers\streams.exe /accepteula -s -d *.ps1
            BRU-uninstall-helpers\streams.exe /accepteula -s -d *.bat
            BRU-uninstall-helpers\streams.exe /accepteula -s -d BRU-uninstall-helpers\*.exe
            BRU-uninstall-helpers\streams.exe /accepteula -s -d BRU-uninstall-helpers\*.dll
            BRU-uninstall-helpers\streams.exe /accepteula -s -d BRU-uninstall-helpers\*.vbs
        )

    Set "bK=HKLM\SOFTWARE\Microsoft\PowerShell"
    Set "eK=PowerShellEngine"
    Set "kV=PowerShellVersion"

    For /F "Tokens=2*" %%A In ('Reg Query "%bK%\3\%eK%" /V "%kV%" 2^>Nul^ ^|^|Reg Query "%bK%\1\%eK%" /V "%kV%" 2^>Nul') Do Set "pV=%%~nB"

    If Not "%pV%" GEq "3" PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell.exe -ArgumentList '-NoProfile -NoExit -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File ""%~dpn0.ps1""' -Verb RunAs}"

    GOTO:eof

	PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell.exe -ArgumentList '-NoProfile -NoExit -ExecutionPolicy Bypass -WindowStyle Hidden -File ""%~dpn0.ps1""' -Verb RunAs}"
	exit

    ) else (

        echo You must be logged in as a member of the Adminstrators group and right-click
        echo this batch file then "Run as Administrator" for the PowerShell script
        echo to execute properly.

    )

PAUSE

