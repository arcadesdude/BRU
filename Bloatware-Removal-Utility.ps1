# BRU
# By Ricky Cobb
#
# Bloatware Removal Utility
# Removes common bloatware from HP, Dell, Lenovo, Sony, Etc
# Supports Powershell 2+, Windows 7/Server 2008 R2 (Winver 6.1+) and newer - including removing Win8/10+ UWP (metro/modern) Apps.
#
# Reboot before running this script and after running it (if anything is removed)
#
# Supporting files that are needed are found in the BRU-uninstall-helpers folder.

<#
Next step if desired to make it fully GUI:

---

add gui box (ok/cancel?) for system restore point and then with list of selected items for confirming removal

progress dialog in main uninstall loop (use current program / total num of programs)

modify/update gui for better UI/button placement/layout, consider additional winforms/dialogboxes

make this warning better by detecting service state and free disk space to state which it is (and if can read option where system protection is enabled/disabled that too to explicitly state the error)
  Write-Warning "A System Restore Point could not be created. Ensure the service is running and System Protection is enabled with enough disk space available.`n`n" | Out-Default

add more error handling
 check for null cases or "" empty strings
 check for edge cases

review error messages/log messages for clarity again, ensure up to date with changes made

review/update all comments for accuracy

add more afteruninstaller starts leftover cleanup for leftover files/registry keys
  can make clean VM then see what changes or possibly use sandboxie when installing to see changes

----

<#
Uninstall logic:

generate nondupped list of all programs installed, wmi/registry 32bit/registry 64bit
 if exists both uninstallstring and msi, take uninstallstring (exe uninstaller) first (better removal)

create list of items to remove, by name with wildcards and regex

filter items to be removed including special cases
 add everything else to removal list

add special cases back to removal list at the very end in a specific order

begin removal of removal-list

for each one

stop processes/services

start program removal

use quiet uninstall string if it exists, otherwise use regular uninstallstring/or msi guid

not msi uninstall? then do regular uninstall handling special cases

if identifying number is not null OR ( identifyingnumber is null AND uninstall string contains guid and contains msiexec)
 then use msiexec /x$id /norestart /qn to uninstall

run function after uninstaller started (clean up files/services/registry keys or start sendkeys method to remove if needed)

end removal

#>

# Parts of this script from:
# http://web.archive.org/web/20150318075245/http://www.gabrielcpinto.com/2014/08/hp-bloatware-removal
# The ideas from that script have been expanded and used to make the process more automated (more silent uninstalls)
# and more complete using the registry and wmi program list as to not miss any programs.
# Using both wmi and registry you get lots of duplicates but we just remove the duplicates after generating and aggregating the list
# (array) of all installed/registered software (doesn't count software not formally installed or not registered with Windows but it should target everything that appears in the Programs and Features (Add/Remove Programs) list and a little more).


# Command Line switches

# -silent (or -quiet, -s, -S), run without GUI, implies auto-confirmation, no warnings
#
# -ignoredefault (or -nd, -id, -ignoredefaults, -nodefaultsuggestions)
#     does not automatically remove default suggestions, you'll need to supply what is to be included and what
#     is to be excluded for removal.
#
# -reboot (or -rebootafterremoval), if option is specficied, will reboot after removal in -silent mode without confirmation
#
# -include, -includefirst, matched first for inclusion, useful if you want to specifically match something the default suggestions
#     do not match or specify your own list (using -ignoredefault options)
#
# -exclude, -filter, this will be filtered out from the includefirst + default suggestions list (if using that), specify what you do NOT want to match (ex. -exclude "Microsoft\ Office","Microsoft\.Office")
#
# -includelast, -specialcases, these will be added after the default matching list and after exclusions are filtered out. Useful for software that needs to go AFTER other software to be removed properly. This list will be parsed in order so you can put what you want to remove in the order it should be removed in.
#
# -includefile -selectionfile [File Path (default: c:\BRU\BRU-Saved-Selection.xml)]

#   This uses the saved file that is created in the GUI with the 'File, Export Selection' option to create the selection list used when running silently. If using this includefile option, the options ignoredefaults, include, exclude, includelast (specialcases) are all ignored and not applied. This also skips the default suggestions list. This assumes the file supplied has programs already chosen and ready to remove. Speeds up removal of bloatware for batches of the same selections (i.e. all same model with same installed bloatware).
#
# -win10leaverecommendedappsdownloadon, with this option given it will not automatically set the registry keys to stop the suggested auto downloaded UWP Windows10 Store apps (i.e. CandyCrush...etc)
#
# -win10leavestartmenuadson, -keepstartads, if this option is give it will NOT set the start menu for new users to get rid of the ContentDeliveryManager Ads. It doesn't affect existing user accounts.
#
# -norestorepoint, -skiprestorepoint, -nr, this will not make a System Restore Point if the option is given.
#
# -dry, -dr, -dryrun, -detect, -detectonly, -whatif, this option will do a Dry Run and not actually remove anything. This is useful with the silent option to see what the applications will be selected. Recommended you use this first to know what will be removed when you run this silently. To actually uninstall/remove the bloatware don't specify this option. This can be useful when refining your includes and excludes to get the matches correct and the order correct for the specialcases.
#


Param(

        [Parameter(Mandatory=$False)]
            [Alias("silent", "quiet", "s")]
            [switch]$Global:isSilent,

        [Parameter(Mandatory=$False)]
            [Alias("nd", "id", "ignoredefault", "ignoredefaults", "ignoredefaultsuggestions", "nodefaultsuggestions")]
            [switch]$Global:isIgnoreDefaultSuggestionListSilentOption,

        [Parameter(Mandatory=$False)]
            [Alias("reboot", "rebootafterremoval")]
            [switch]$Global:isRebootAfterRemovalswitchSilentOption,

        [Parameter(Mandatory=$False)]
            [Alias("include", "includefirst")]
            [string[]]$Global:bloatwareIncludeFirstSilentOption,

        [Parameter(Mandatory=$False)]
            [Alias("exclude", "filter")]
            [string[]]$Global:bloatwareExcludeSilentOption,

        [Parameter(Mandatory=$False)]
            [Alias("includelast", "specialcases")]
            [string[]]$Global:bloatwareIncludeLastSilentOption,

        [Parameter(Mandatory=$False)]
            [Alias("includefile", "selectionfile")]
            [string]$Global:usingSavedSelectionFileSilentOption,

        [Parameter(Mandatory=$False)]
            [Alias("win10leaverecommendedappsdownloadon")]
            [switch]$Global:isWin10RecommendedDownloadsOffSilentOption,

        [Parameter(Mandatory=$False)]
            [Alias("win10leavestartmenuadson", "keepstartads")]
            [switch]$Global:isWin10StartMenuAdsSilentOption,

        [Parameter(Mandatory=$False)]
            [Alias("norestorepoint", "skiprestorepoint", "nr")]
            [switch]$Global:isrequireSystemRestorePointBeforeRemovalSilentOption,

        [Parameter(Mandatory=$False)]
            [Alias("dry", "dr", "dryrun", "detect", "detectonly", "whatif")]
            [switch]$Global:isDetectOnlyDryRunSilentOption
     )



# Go read BEGIN block to follow program flow then come back here to PROCESS
PROCESS {

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator" )) {
    Write-Output 'You must be logged in as a member of the Adminstrators group for this script to execute properly.'  | Out-Default
    Write-Output "" | Out-Default
    Write-Output "This window is safe to close." | Out-Default
#    Write-Output "Press any key to continue..." | Out-Default
#    $HOST.UI.RawUI.Flushinputbuffer()
#    $HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | Out-Null
    break;
}

$Script:isConsoleShowing = $false
if (!($Global:isSilent)) {
    hideConsole | Out-Null
}

#Save path that the script is run from (e.g. "usbflashdriveletter:\" )
$scriptPath = (Split-Path -Parent $MyInvocation.MyCommand.Definition).TrimEnd('\')
$scriptFullLaunchCmd = ($MyInvocation.Line)
$scriptName = (Split-Path -Leaf $MyInvocation.MyCommand.Definition)


$Script:dest = "C:\BRU" # no trailing slash, for iss response files created using Set-Content and copying uninstall helper files so we can remove flash drive or removable media the script is run from if needed
if ( !(Test-path $Script:dest) ) { md -Path $Script:dest | Out-Null }
$savedPathLocation = Get-Location # save to be restored later
Set-Location $Script:dest # if using removable media like a flash drive, will be safe to remove later

$Script:logfile = [string]$Script:dest+$("\Bloatware-Removal-"+$(get-date -uformat %d-%b-%Y-%H-%M)+".log")

Start-Transcript $Script:logfile

Write-Output "Bloatware-Removal Initializing..." | Out-Default

Write-Output "PowerShell Version: $($PSVersionTable.PSVersion.Major)" | Out-Default

try {
    [float]$Script:winVer = (((Get-CimInstance Win32_OperatingSystem).Version.Split('.') | Select -first 2) -join '.')
} catch {
    [float]$Script:winVer = ((([Environment]::OSVersion.Version).ToString().Split('.') | Select -first 2) -join '.')
}

[string]$Script:osArch = (@( "64" , "86" )[32 -eq (8*[IntPtr]::Size)])

$scriptName -match '(.*)\..*$|^(.*)$' | Out-Null # match filename not extension
if ( $matches[2] ) {
    $scriptNameNoExtension = $matches[2]
} elseif ( $matches[1] ) {
    $scriptNameNoExtension = $matches[1]
} else {
    Write-Warning "No match for (Split-Path -Leaf `$MyInvocation.MyCommand.Definition) found." | Out-Default
    Write-Warning "Please rename this program in the format programname.ext" | Out-Default
    return
}

$Global:statusupdate =  $null

# Set default options which are saved to file (as [scriptname].ini)
# Command line parameter switches/values are runtime only and not saved here or in the file
$Global:globalSettings = @{  # can be any type not just booleans
    "requireSystemRestorePointBeforeRemoval" = $true
    "requireConfirmationBeforeRemoval" = $true
    "optionsWin10RecommendedDownloadsOff" = $true
    "optionsWin10StartMenuAds" = $true
    "rebootAfterRemoval" = $false
    "showMicrosoftPublished" = $true
    "showUWPapps" = $true
}

function globalSettingsSetIndividualVariables( ) {
   For( $i = 0; $i -lt (($Global:globalSettings.Keys).Count); $i++ ) {
        $settingname = ($Global:globalSettings.Keys | Select-Object -Index $i)
        Set-Variable -Name $settingname -Value $Global:globalSettings[$settingname] -Scope Global -Force
    }
}

function importSettings( ) {
    try {
        $settings = Import-CliXml "$($scriptPath)\$($scriptNameNoExtension).ini" -ErrorAction Stop
        $Global:globalSettings = $settings
    } catch {
        $Global:statusupdate = "Loading Settings Failed. Check settings file: $($scriptPath)\$($scriptNameNoExtension).ini Continuing with default settings."
        Write-Warning $Global:statusupdate | Out-Default
    }
    globalSettingsSetIndividualVariables
}

function saveSettings( ) {
    For( $i = 0; $i -lt (($Global:globalSettings.Keys).Count); $i++) {
        $settingname = ($Global:globalSettings.Keys | Select-Object -Index $i)
        $Global:globalSettings[$settingname] = (Get-Variable -Name $settingname -ValueOnly -Scope Global)
    }
    try {
        $Global:globalSettings | Export-CliXML -Force "$($scriptPath)\$($scriptNameNoExtension).ini" -ErrorAction Stop
        return $false
    } catch {
        $Global:statusupdate = "Saving Settings Failed. Check permissions on settings file: $($scriptPath)\$($scriptNameNoExtension).ini and ensure there is sufficient disk space available."
        Write-Warning $Global:statusupdate | Out-Default
        return $true
    }
}

# in case there are saved settings but a newer version of script has new global variable default to be set
globalSettingsSetIndividualVariables

if ( (!(Test-Path "$($scriptPath)\$($scriptNameNoExtension).ini") -and (!($Global:isSilent))) ) {
    globalSettingsSetIndividualVariables
    saveSettings | Out-Null
    Write-Output "Settings file not found. Creating one and using default settings." | Out-Default
}

if ( $Global:isSilent ) { # Running silently ignores the saved preferences file
    Write-Output "`nRunning silently using -silent switch."
}

Write-Output "`nFull command line:"
Write-Output "$scriptFullLaunchCmd"
if ($($PSVersionTable.PSVersion.Major) -gt 2) {
    Write-Output "`$PSCommandPath:"
    Write-Output $PSCommandPath
}
if ( !($Global:isSilent) ) {
    importSettings
}

# If Running Silently set up options based on the command line options
if ( $Global:isSilent ) { # -silent always implies no confirmation prompts
    if ($Global:usingSavedSelectionFileSilentOption) {
        if (!(Test-Path "$($Global:usingSavedSelectionFileSilentOption)")) {
            Write-Warning "`nSaved selection file $($Global:usingSavedSelectionFileSilentOption) not found or doesnt exist!"
            Write-Warning "Please check the filename and provide the full path to the file`nThe default location and filename is (c:\BRU\BRU-Saved-Selection.xml)"
            Write-Output "" | Out-Default
            stopTranscript
            Set-Location $savedPathLocation # Restore working directory path
            Return
        }
    }
    $Global:requireConfirmationBeforeRemoval = $false
    $Global:globalSettings["requireConfirmationBeforeRemoval"] = $Global:requireConfirmationBeforeRemoval
    if ( $Global:isRebootAfterRemovalswitchSilentOption ) {
        $Global:rebootAfterRemoval = $true
        $Global:globalSettings["rebootAfterRemoval"] = $Global:rebootAfterRemoval
    }

    if ( $Global:isWin10RecommendedDownloadsOffSilentOption ) {
        $Global:optionsWin10RecommendedDownloadsOff = $false
        $Global:globalSettings["optionsWin10RecommendedDownloadsOff"] = $Global:optionsWin10RecommendedDownloadsOff
    }

    if ( $Global:isWin10StartMenuAdsSilentOption ) {
        $Global:optionsWin10StartMenuAds = $false
        $Global:globalSettings["optionsWin10StartMenuAds"] = $Global:optionsWin10StartMenuAds

    }

    if ( $Global:isrequireSystemRestorePointBeforeRemovalSilentOption ) {
        $Global:requireSystemRestorePointBeforeRemoval = $false
        $Global:globalSettings["requireSystemRestorePointBeforeRemoval"] = $Global:requireConfirmationBeforeRemoval
    }
}


Write-Output "`nUsing Options:" | Out-Default
For( $i = 0; $i -lt (($Global:globalSettings.Keys).Count); $i++) {
    $settingname = ($Global:globalSettings.Keys | Select-Object -Index $i)
    Write-Output "$($settingname): $($Global:globalSettings[$($settingname)])" | Out-Default
}
Write-Output "`n" | Out-Default

# Start GUI Init if not silent
if ( !($Global:isSilent) ) {
    Add-Type -AssemblyName "System.Windows.Forms" | Out-Null
    #[System.Windows.Forms.Application]::EnableVisualStyles()

    $Script:LastColumnClicked = 0
    $Script:LastColumnAscending = $false

    # GUI code based on examples by ZeroSevenCodes, https://www.youtube.com/watch?v=GQSQesbw3B0
    $mainUI = New-Object System.Windows.Forms.Form
    $mainUI.Text = "Bloatware Removal Utility"
    $mainUI.Size = New-Object System.Drawing.Size(832,528)
    $mainUI.MinimumSize = New-Object System.Drawing.Size(480,280)
    $mainUI.FormBorderStyle = "Sizable"#FixedDialog"
    $mainUI.SizeGripStyle = "Hide"
    $mainUI.TopMost  = $false
    $mainUI.MaximizeBox  = $true
    $mainUI.MinimizeBox  = $true
    $mainUI.ControlBox = $true
    $mainUI.StartPosition = "CenterScreen"
    $mainUI.WindowState = "Normal" #"Maximized"
    $mainUI.Font = "Segoe UI"
    # base64 icon code from http://www.alkanesolutions.co.uk/2013/04/19/embedding-base64-image-strings-inside-a-powershell-application/
    # get base64 string from here: http://www.base64-image.de
    $mainUIicon = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAmlJREFUeNq0VkEoRFEUnRkjImYSJTVNJjsaOwtqFnazYCc2kxU7Vuws7VhhY6w0G7JQLKYUC0WRDaFsFEpqMs2YaBA5OnV7vf/+92fM3H7z37z737333Xvuec8bab32VFN8niqLn6+D0wgHzQHf1cX7RiqX3inIR6KFPNx9qtpQuHZ9K4TBYN+tccYvK6HgoD/WgGchnEkuZ61aDKBNRrML8xmZgVc1anXGNkWzc21i1CqTUy09vfUl1+D48A17F+fx4SarNp/74t/RRKCEGlDmZp5ofXsvjAAHYo1qlkSLFCNL0JaPos1UngU3aleXn5noQLCmTAeXF0X82q2XHNpF8N8+gAP6KH8HFKln5Tu5J1qvpsIqTI6bCMwOCFDg0qh1WV4DTOcX2xEUAAoTGKR3C8Y1BCgr4dCMBgdAt5QRqDdmABbRxr9IOy+WvAPVCvahpYgUJiGD8kousuCPXKSxDUyLdYRvVyEnB+MjD+BYPNy+HdsgiOmJx/JhivXJlawUUyU77g/uHRDsCqbcvpoTkh3zDiy4h6nProeNEOJBBuvd0br/nskv+W8VMxoKXHL132RnPbaODl/VjhHmsEua2QG+ZuzWYrI8cEyL3JNqnWPJsNmBpNhaCYG/tglZwsiYYbMDhIAu4xhXGGv96UPlFTWZnL+/+zA4WFrrwGm8f9LJr4F3I5ZYBqkzu3IsEWRi40O/THx1/m7gIq2kPJmduwThwx/oD+Ozmy4tCCcUAfJ2dCadLMeGhgV1xqdhHA/W49amsY1mAgFihgcfRL3dQOTSB/FW6naNGvCcAKOoROut9vX9R4ABAAJJVwDVgzCOAAAAAElFTkSuQmCC"
    $iconimageBytes = [Convert]::FromBase64String($mainUIicon)
    $imageStream = New-Object IO.MemoryStream($iconimageBytes, 0, $iconimageBytes.Length)
    $imageStream.Write($iconimageBytes, 0, $iconimageBytes.Length)
    [System.Drawing.Image]::FromStream($imageStream, $true) | Out-Null
    $mainUI.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $imageStream).GetHIcon())

    $mainMenuStrip = New-Object System.Windows.Forms.MenuStrip
    $mainMenuStrip.Location = New-Object System.Drawing.Point(0, 0)
    $mainMenuStrip.Name = "mainMenuStrip"
    $mainMenuStrip.Size = New-Object System.Drawing.Size(320, 22)
    #$mainMenuStrip.Padding = 0
    ###
    $fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $fileMenu.Name = "fileMenu"
    $fileMenu.Size = New-Object System.Drawing.Size(35, 22)
    $fileMenu.Text = "&File"
    $fileMenu.TextAlign = "MiddleLeft"
    ###
    $fileExportMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $fileExportMenu.Name = "fileExportMenu"
    $fileExportMenu.Size = New-Object System.Drawing.Size(152, 22)
    $fileExportMenu.Text = "&Export Selection..."
    $fileExportMenu.TextAlign = "MiddleLeft"
    function doFileExportMenu($Sender,$e){
        Write-Output "File, Export Selection choosen..." | Out-Default
        $fileExportSaveDiaglog = New-Object System.Windows.Forms.SaveFileDialog
        $fileExportSaveDiaglog.FileName = "BRU-Saved-Selection" # Default file name
        $fileExportSaveDiaglog.InitialDirectory = $Script:dest
        $fileExportSaveDiaglog.DefaultExt = ".xml" # Default file extension
        $fileExportSaveDiaglog.Filter = "XML Document (.xml)|*.xml" # Filter files by extension
        $fileExportSaveDiaglog.OverwritePrompt = $false
        $result = $fileExportSaveDiaglog.ShowDialog()
        if ($result -ne 'Cancel') {
            $selectionExportPath = $fileExportSaveDiaglog.FileName
            Write-Output "" | Out-Default
            $Global:statusupdate = "Exporting selection list..."
            Write-Host $Global:statusupdate | Out-Default
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
            Start-Sleep -Milliseconds 500
            $progslistSelectedforExport = selectedProgsListviewtoArray $programsListview
            if ( $progslistSelectedforExport -ne $null ) {
                $removeOrderedSelectedListforExport = ,@()
                # sort $progslistSelectedforExport against $Script:progslisttoremove
                ForEach ($prog in $Script:progslisttoremove) {
                    ForEach ( $selectedprog in $progslistSelectedforExport) {
                        if ( isObjectEqual $selectedprog $prog ) {
                            $removeOrderedSelectedListforExport += $selectedProg
                            $progslistSelectedforExport = $progslistSelectedforExport | Where { $_ -ne $selectedProg }
                        }
                    }
                }
                # add items selected that weren't in progslisttoremove but were selected to $progslistSelectedforExport
                $removeOrderedSelectedListforExport += $progslistSelectedforExport | Sort-Object UninstallString
                Write-Host "" | Out-Default
                try {
                    $removeOrderedSelectedListforExport | Select-Object -Skip 1 | Export-Clixml -Path $selectionExportPath -ErrorAction Stop
                    $Global:statusupdate = "Selection list exported to $($selectionExportPath)"
                } catch {
                    $Global:statusupdate = "Failed to export selection list to $($selectionExportPath) Check File permissions and free space."
                }
            } else {
                $Global:statusupdate = "Nothing is selected for export."
            }
        } else {
            $Global:statusupdate = "No file chosen for export."
        }
        Write-Host $Global:statusupdate | Out-Default
        $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
        $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
    }
    $fileMenu.DropDownItems.Add($fileExportMenu) | Out-Null
    $fileExportMenu.Add_Click( { doFileExportMenu $fileExportMenu $EventArgs} )
    ###
    $fileExitMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $fileExitMenu.Name = "fileExitMenu"
    $fileExitMenu.Size = New-Object System.Drawing.Size(152, 22)
    $fileExitMenu.Text = "&Exit"
    $fileExitMenu.TextAlign = "MiddleLeft"
    function doFileExitMenu($Sender,$e){
        $mainUI.Close()
        Return
        #[Environment]::Exit(4)
    }
    $fileMenu.DropDownItems.Add($fileExitMenu) | Out-Null
    $fileExitMenu.Add_Click( { doFileExitMenu $fileExitMenu $EventArgs} )
    ###
    $optionsMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $optionsMenu.Name = "optionsMenu"
    $optionsMenu.Size = New-Object System.Drawing.Size(35, 22)
    $optionsMenu.Text = "&Options"
    $optionsMenu.TextAlign = "MiddleCenter"
    ###
    $optionsRequireSystemRestorePointMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $optionsRequireSystemRestorePointMenu.Name = "optionsRequireSystemRestorePointMenu"
    $optionsRequireSystemRestorePointMenu.Size = New-Object System.Drawing.Size(152, 22)
    $optionsRequireSystemRestorePointMenu.Text = "Require System Restore &Point before removal"
    $optionsRequireSystemRestorePointMenu.TextAlign = "MiddleLeft"
    $optionsRequireSystemRestorePointMenu.Checked = $Global:requireSystemRestorePointBeforeRemoval
    function doOptionsRequireSystemRestorePointMenu($Sender,$e){
        $optionsRequireSystemRestorePointMenu.Checked = !($optionsRequireSystemRestorePointMenu.Checked)
        $Global:requireSystemRestorePointBeforeRemoval = $optionsRequireSystemRestorePointMenu.Checked
        $hasStatusUpdate = saveSettings
        if ( $hasStatusUpdate ) {
            Write-Output "" | Out-Default
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
        }
    }
    $optionsRequireSystemRestorePointMenu.Add_Click( { doOptionsRequireSystemRestorePointMenu $optionsRequireSystemRestorePointMenu $EventArgs} )
    $optionsMenu.DropDownItems.Add($optionsRequireSystemRestorePointMenu) | Out-Null
    ###
    $optionsRequireConfirmationMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $optionsRequireConfirmationMenu.Name = "optionsRequireConfirmationMenu"
    $optionsRequireConfirmationMenu.Size = New-Object System.Drawing.Size(152, 22)
    $optionsRequireConfirmationMenu.Text = "Require &Confirmation before removal"
    $optionsRequireConfirmationMenu.TextAlign = "MiddleLeft"
    $optionsRequireConfirmationMenu.Checked = $Global:requireConfirmationBeforeRemoval
    function doOptionsRequireConfirmationMenu($Sender,$e){
        $optionsRequireConfirmationMenu.Checked = !($optionsRequireConfirmationMenu.Checked)
        $Global:requireConfirmationBeforeRemoval = $optionsRequireConfirmationMenu.Checked
        $hasStatusUpdate = saveSettings
        if ( $hasStatusUpdate ) {
            Write-Output "" | Out-Default
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
        }
    }
    $optionsRequireConfirmationMenu.Add_Click( { doOptionsRequireConfirmationMenu $optionsRequireConfirmationMenu $EventArgs} )
    $optionsMenu.DropDownItems.Add($optionsRequireConfirmationMenu) | Out-Null
    ###
    $optionsWin10RecommendedDownloadsOffMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $optionsWin10RecommendedDownloadsOffMenu.Name = "optionsWin10RecommendedDownloadsOffMenu"
    $optionsWin10RecommendedDownloadsOffMenu.Size = New-Object System.Drawing.Size(152, 22)
    $optionsWin10RecommendedDownloadsOffMenu.Text = "Win10+ only - After removal set `"recommended`" UWP app &auto-downloads off"
    $optionsWin10RecommendedDownloadsOffMenu.TextAlign = "MiddleLeft"
    $isWin10 = [bool]($Script:winVer -ge 10)
    $optionsWin10RecommendedDownloadsOffMenu.Enabled = $isWin10
    $optionsWin10RecommendedDownloadsOffMenu.Checked = $isWin10 -and $Global:optionsWin10RecommendedDownloadsOff
    function doOptionsWin10RecommendedDownloadsOffMenu($Sender,$e){
        $optionsWin10RecommendedDownloadsOffMenu.Checked = !($optionsWin10RecommendedDownloadsOffMenu.Checked)
        $Global:optionsWin10RecommendedDownloadsOff = $optionsWin10RecommendedDownloadsOffMenu.Checked
        $hasStatusUpdate = saveSettings
        if ( $hasStatusUpdate ) {
            Write-Output "" | Out-Default
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
        }
    }
    $optionsWin10RecommendedDownloadsOffMenu.Add_Click( { doOptionsWin10RecommendedDownloadsOffMenu $optionsWin10RecommendedDownloadsOffMenu $EventArgs} )
    $optionsMenu.DropDownItems.Add($optionsWin10RecommendedDownloadsOffMenu) | Out-Null
    ###
    $optionsWin10StartMenuAdsMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $optionsWin10StartMenuAdsMenu.Name = "optionsWin10StartMenuAdsMenu"
    $optionsWin10StartMenuAdsMenu.Size = New-Object System.Drawing.Size(152, 22)
    $optionsWin10StartMenuAdsMenu.Text = "Win10+ only - After removal set default &Start Menu layout for new users"
    $optionsWin10StartMenuAdsMenu.TextAlign = "MiddleLeft"
    $optionsWin10StartMenuAdsMenu.Enabled = $isWin10
    $optionsWin10StartMenuAdsMenu.Checked = $isWin10 -and $Global:optionsWin10StartMenuAds
    function doOptionsWin10StartMenuAdsMenu($Sender,$e){
        $optionsWin10StartMenuAdsMenu.Checked = !($optionsWin10StartMenuAdsMenu.Checked)
        $Global:optionsWin10StartMenuAds = $optionsWin10StartMenuAdsMenu.Checked
        $hasStatusUpdate = saveSettings
        if ( $hasStatusUpdate ) {
            Write-Output "" | Out-Default
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
        }
    }
    $optionsWin10StartMenuAdsMenu.Add_Click( { doOptionsWin10StartMenuAdsMenu $optionsWin10StartMenuAdsMenu $EventArgs} )
    $optionsMenu.DropDownItems.Add($optionsWin10StartMenuAdsMenu) | Out-Null
    ###
    $optionsRebootAfterRemovalMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $optionsRebootAfterRemovalMenu.Name = "optionsRebootAfterRemovalMenu"
    $optionsRebootAfterRemovalMenu.Size = New-Object System.Drawing.Size(152, 22)
    $optionsRebootAfterRemovalMenu.Text = "&Reboot after removal"
    $optionsRebootAfterRemovalMenu.TextAlign = "MiddleLeft"
    $optionsRebootAfterRemovalMenu.Checked = $Global:rebootAfterRemoval
    function doOptionsRebootAfterRemovalMenu($Sender,$e){
        $optionsRebootAfterRemovalMenu.Checked = !($optionsRebootAfterRemovalMenu.Checked)
        $Global:rebootAfterRemoval = $optionsRebootAfterRemovalMenu.Checked
        $hasStatusUpdate = saveSettings
        if ( $hasStatusUpdate ) {
            Write-Output "" | Out-Default
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
        }
    }
    $optionsRebootAfterRemovalMenu.Add_Click( { doOptionsRebootAfterRemovalMenu $optionsRebootAfterRemovalMenu $EventArgs} )
    $optionsMenu.DropDownItems.Add($optionsRebootAfterRemovalMenu) | Out-Null
    ###
    $viewMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $viewMenu.Name = "viewMenu"
    $viewMenu.Size = New-Object System.Drawing.Size(35, 22)
    $viewMenu.Text = "&View"
    $viewMenu.TextAlign = "MiddleCenter"
    ###
    $viewShowMicrosoftPublished = New-Object System.Windows.Forms.ToolStripMenuItem
    $viewShowMicrosoftPublished.Name = "viewShowMicrosoftPublished"
    $viewShowMicrosoftPublished.Size = New-Object System.Drawing.Size(152, 22)
    $viewShowMicrosoftPublished.Text = "Show Published by &Microsoft"
    $viewShowMicrosoftPublished.TextAlign = "MiddleLeft"
    $viewShowMicrosoftPublished.Checked = $Global:showMicrosoftPublished
    function doviewShowMicrosoftPublished($Sender,$e){
        $viewShowMicrosoftPublished.Checked = !($viewShowMicrosoftPublished.Checked)
        $Global:showMicrosoftPublished = $viewShowMicrosoftPublished.Checked
        $hasStatusUpdate = saveSettings
        if ( $hasStatusUpdate ) {
            Write-Output "" | Out-Default
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
        }
        Write-Host "View status of $($viewShowMicrosoftPublished.Text.Replace('&','')) changed to: $($Global:showMicrosoftPublished)" | Out-Default
        $currentSelectedProgs = selectedProgsListviewtoArray $programsListview
        refreshAlreadyGeneratedProgramsList $currentSelectedProgs
    }
    $viewShowMicrosoftPublished.Add_Click( { doviewShowMicrosoftPublished $viewShowMicrosoftPublished $EventArgs} )
    $viewMenu.DropDownItems.Add($viewShowMicrosoftPublished) | Out-Null
    ###
    $viewShowUWPapps = New-Object System.Windows.Forms.ToolStripMenuItem
    $viewShowUWPapps.Name = "viewShowUWPapps"
    $viewShowUWPapps.Size = New-Object System.Drawing.Size(152, 22)
    $viewShowUWPapps.Text = "Win8+ only - Show &UWP/Metro/Modern (Win8/10+) Apps"
    $viewShowUWPapps.TextAlign = "MiddleLeft"
    $viewShowUWPapps.Checked = $Global:showUWPapps -and ($winVer -gt 6.1)
    $viewShowUWPapps.Enabled = ($winVer -gt 6.1)
    function doviewShowUWPapps($Sender,$e){
        $viewShowUWPapps.Checked = !($viewShowUWPapps.Checked)
        $Global:showUWPapps = $viewShowUWPapps.Checked
        $hasStatusUpdate = saveSettings
        if ( $hasStatusUpdate ) {
            Write-Output "" | Out-Default
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
        }
        Write-Host "View status of $($viewShowUWPapps.Text.Replace('&','')) changed to: $($Global:showUWPapps)" | Out-Default
        $currentSelectedProgs = selectedProgsListviewtoArray $programsListview
        refreshAlreadyGeneratedProgramsList $currentSelectedProgs
    }
    $viewShowUWPapps.Add_Click( { doviewShowUWPapps $viewShowUWPapps $EventArgs} )
    $viewMenu.DropDownItems.Add($viewShowUWPapps) | Out-Null
    ###
    $viewShowSuggestedBloatware = New-Object System.Windows.Forms.ToolStripMenuItem
    $viewShowSuggestedBloatware.Name = "viewShowSuggestedBloatware"
    $viewShowSuggestedBloatware.Size = New-Object System.Drawing.Size(152, 22)
    $viewShowSuggestedBloatware.Text = "Show &Suggested Bloatware"
    $viewShowSuggestedBloatware.TextAlign = "MiddleLeft"
    $viewShowSuggestedBloatware.Checked = $true
    function doviewShowSuggestedBloatware($Sender,$e){
        toggleSuggestedBloatware
        Write-Host "View status of $($viewShowSuggestedBloatware.Text.Replace('&','')) changed to: $($Script:showSuggestedtoRemove)" | Out-Default
    }
    $viewShowSuggestedBloatware.Add_Click( { doviewShowSuggestedBloatware $viewShowSuggestedBloatware $EventArgs} )
    $viewMenu.DropDownItems.Add($viewShowSuggestedBloatware) | Out-Null
    ###
    $viewShowConsoleWindow = New-Object System.Windows.Forms.ToolStripMenuItem
    $viewShowConsoleWindow.Name = "viewShowConsoleWindow"
    $viewShowConsoleWindow.Size = New-Object System.Drawing.Size(152, 22)
    $viewShowConsoleWindow.Text = "Show &Console Window"
    $viewShowConsoleWindow.TextAlign = "MiddleLeft"
    $viewShowConsoleWindow.Checked = $false
    function doviewShowConsoleWindow($Sender,$e){
        toggleConsoleWindow
        Write-Host "View status of $($viewShowConsoleWindow.Text.Replace('&','')) changed to: $($Script:isConsoleShowing)" | Out-Default
    }
    $viewShowConsoleWindow.Add_Click( { doviewShowConsoleWindow $viewShowConsoleWindow $EventArgs} )
    $viewMenu.DropDownItems.Add($viewShowConsoleWindow) | Out-Null
    ###
    $viewRefreshPrograms = New-Object System.Windows.Forms.ToolStripMenuItem
    $viewRefreshPrograms.Name = "viewRefreshPrograms"
    $viewRefreshPrograms.Size = New-Object System.Drawing.Size(152, 22)
    $viewRefreshPrograms.Text = "&Refresh Programs List"
    $viewRefreshPrograms.TextAlign = "MiddleLeft"
    function doviewRefreshPrograms($Sender,$e){
        Write-Output "" | Out-Default
        $Global:statusupdate = "Programs List Refreshing..."
        Write-Host $Global:statusupdate | Out-Default
        $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
        $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
        $currentSelectedProgs = selectedProgsListviewtoArray $programsListview
        refreshProgramsList
        $Script:programsListviewWasJustRecreated = $true
        refreshAlreadyGeneratedProgramsList $currentSelectedProgs # restore matching selected items
        Write-Host "Programs List Refreshed." | Out-Default
    }
    $viewRefreshPrograms.Add_Click( { doviewRefreshPrograms $viewRefreshPrograms $EventArgs} )
    $viewMenu.DropDownItems.Add($viewRefreshPrograms) | Out-Null
    ###
    $helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $helpMenu.Name = "helpMenu"
    $helpMenu.Size = New-Object System.Drawing.Size(51, 22)
    $helpMenu.Text = "&Help"
    $helpMenu.TextAlign = "MiddleCenter"
    ###
    $helpAboutMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $helpAboutMenu.Name = "helpAboutMenu"
    $helpAboutMenu.Size = New-Object System.Drawing.Size(152, 22)
    $helpAboutMenu.Text = "&About"
    $helpAboutMenu.TextAlign = "MiddleLeft"
    function showHelpAboutMenu($Sender,$e){
        [void][System.Windows.Forms.MessageBox]::Show("Bloatware Removal Utility by Ricky Cobb (c) $((Get-Date).Year).`n`nIntended use for removing bloatware from new`nfactory image systems.`n`nCarefully check the selection list before`nremoving any selected programs.`n`nUse at your own risk!`n`nhttp://github.com/arcadesdude/BRU","About Bloatware Removal Utility (BRU)")
    }
    $helpMenu.DropDownItems.Add($helpAboutMenu) | Out-Null
    $helpAboutMenu.Add_Click( { showHelpAboutMenu $helpAboutMenu $EventArgs} )
    ###
    $mainMenuStrip.Items.AddRange(@($fileMenu,$optionsMenu,$viewMenu,$helpMenu))
    $mainUI.Controls.Add($mainMenuStrip)
    $mainUI.MainMenuStrip = $mainMenuStrip
    ###
    <#
    $allPrograms = New-Object System.Windows.Forms.Label
    $allPrograms.Location = New-Object System.Drawing.Size(8,16)
    $allPrograms.Size = New-Object System.Drawing.Size(320,32)
    $allPrograms.TextAlign = "MiddleLeft"
    $allPrograms.Text = "Please select the applications you wish to remove:"
    $mainUI.Controls.Add($allPrograms)
    #>
    ###
    $Global:programsListview = New-Object System.Windows.Forms.ListView # Global to be accessible within functions
    $programsListview.Location = New-Object System.Drawing.Size(8,26)
    $programsListview.Size = New-Object System.Drawing.Size(800,381)
    $programsListview.MinimumSize = New-Object System.Drawing.Size(100,100)
    $programsListview.CheckBoxes = $true
    $programsListview.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor
    [System.Windows.Forms.AnchorStyles]::Right -bor
    [System.Windows.Forms.AnchorStyles]::Top -bor
    [System.Windows.Forms.AnchorStyles]::Left
    $programsListview.View = "Details"
    $programsListview.FullRowSelect = $true
    $programsListview.MultiSelect = $true
    $programsListview.Sorting = "None"
    $programsListview.AllowColumnReorder = $true
    $programsListview.GridLines = $true
    $programsListview.Add_ColumnClick({sortprogramsListview $_.Column})
    $programsListview.Add_ItemCheck({updateSelectedProgsStatus $_})
    $mainUI.Controls.Add($programsListview)

    $buttonToggleSuggested = New-Object System.Windows.Forms.Button
    $buttonToggleSuggested.Location = New-Object System.Drawing.Size(8,420)
    $buttonToggleSuggested.Size = New-Object System.Drawing.Size(120,32)
    $buttonToggleSuggested.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor
    [System.Windows.Forms.AnchorStyles]::Left
    $buttonToggleSuggested.TextAlign = "MiddleCenter"
    $buttonToggleSuggested.Text = "Toggle Suggested Bloatware"
    $buttonToggleSuggested.Add_Click({doviewShowSuggestedBloatware}) # Reusing View Menu Function
    $buttonToggleSuggested.Enabled = $false
    $mainUI.Controls.Add($buttonToggleSuggested)

    $buttonToggleConsole = New-Object System.Windows.Forms.Button
    $buttonToggleConsole.Location = New-Object System.Drawing.Size(136,420)
    $buttonToggleConsole.Size = New-Object System.Drawing.Size(90,32)
    $buttonToggleConsole.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor
    [System.Windows.Forms.AnchorStyles]::Left
    $buttonToggleConsole.TextAlign = "MiddleCenter"
    $buttonToggleConsole.Text = "Show/Hide Console"
    $buttonToggleConsole.Add_Click({doviewShowConsoleWindow}) # Reusing View Menu Function
    $buttonToggleConsole.Enabled = $false
    $mainUI.Controls.Add($buttonToggleConsole)

    $buttonConfirmedSelectedforRemoval = New-Object System.Windows.Forms.Button
    $buttonConfirmedSelectedforRemoval.Location = New-Object System.Drawing.Size(591,420)
    $buttonConfirmedSelectedforRemoval.Size = New-Object System.Drawing.Size(120,32)
    $buttonConfirmedSelectedforRemoval.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor
    [System.Windows.Forms.AnchorStyles]::Right
    $buttonConfirmedSelectedforRemoval.TextAlign = "MiddleCenter"
    $buttonConfirmedSelectedforRemoval.Text = "Remove Selected"
    $buttonConfirmedSelectedforRemoval.Add_Click({$Script:button = $buttonConfirmedSelectedforRemoval.Text;$mainUI.Close()})
    $buttonConfirmedSelectedforRemoval.Enabled = $false
    $mainUI.Controls.Add($buttonConfirmedSelectedforRemoval)

    $buttonCancelRemoval = New-Object System.Windows.Forms.Button
    $buttonCancelRemoval.Location = New-Object System.Drawing.Size(719,420)
    $buttonCancelRemoval.Size = New-Object System.Drawing.Size(90,32)
    $buttonCancelRemoval.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor
    [System.Windows.Forms.AnchorStyles]::Right
    $buttonCancelRemoval.TextAlign = "MiddleCenter"
    $buttonCancelRemoval.Text = "Cancel"
    $buttonCancelRemoval.Add_Click({$Script:button = $buttonCancelRemoval.Text;$mainUI.Close()})
    $buttonCancelRemoval.Enabled = $false
    $mainUI.Controls.Add($buttonCancelRemoval)

    $Global:statusBarTextBox = New-Object System.Windows.Forms.StatusBar
    #$statusBarTextBox.Text = "  Status Bar"
    $statusBarTextBox.Width = 830
    $statusBarTextBox.Height = 20
    $statusBarTextBox.Location = New-Object System.Drawing.Size(0,465)
    $statusBarTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor
    [System.Windows.Forms.AnchorStyles]::Right -bor
    [System.Windows.Forms.AnchorStyles]::Left
    #$programsListview.Size = New-Object System.Drawing.Size(4,288)
    #$programsListview.MinimumSize = New-Object System.Drawing.Size(4,100)
    $statusBarTextBox.Font = "Microsoft Sans Serif,10"
    $statusBarTextBox.ShowPanels = $true
    $mainUI.controls.Add($statusBarTextBox)

    $statusBarTextBoxStatusText = New-Object System.Windows.Forms.StatusBarPanel
    $statusBarTextBoxStatusTextIndex = $statusBarTextBox.Panels.Add($statusBarTextBoxStatusText)
    $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Alignment = "Left"
    $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].AutoSize = "Spring"
    $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].MinWidth = 600
    #$statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = '  Status Bar'
    #$statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Tooltip = 'Status Bar Tooltip'

    $statusBarTextBoxSelectedProgs = New-Object System.Windows.Forms.StatusBarPanel
    $statusBarTextBoxSelectedProgsIndex = $statusBarTextBox.Panels.Add($statusBarTextBoxSelectedProgs)
    $statusBarTextBox.Panels[$statusBarTextBoxSelectedProgsIndex].Alignment = "Right"
    $statusBarTextBox.Panels[$statusBarTextBoxSelectedProgsIndex].AutoSize = "Spring"
    $statusBarTextBox.Panels[$statusBarTextBoxSelectedProgsIndex].Text = "Selected: 0"

    $statusBarTextBoxTotalProgs = New-Object System.Windows.Forms.StatusBarPanel
    $statusBarTextBoxTotalProgsIndex = $statusBarTextBox.Panels.Add($statusBarTextBoxTotalProgs)
    $statusBarTextBox.Panels[$statusBarTextBoxTotalProgsIndex].Alignment = "Right"
    $statusBarTextBox.Panels[$statusBarTextBoxTotalProgsIndex].AutoSize = "Spring"

    $Script:showSuggestedtoRemove = $true
    $progslistSelected = $null
    $button = "Cancel" # default if closing window

    ################## GUI is Activated and Shown #####################################################################

    $mainUI.add_Shown({

        if ( $Global:statusupdate ) {
            Write-Output "" | Out-Default
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
            Start-Sleep -Seconds 2
        }


        refreshProgramsList


    }) # end $mainUI.add_Shown({

    $buttonToggleSuggested.Enabled = $true
    $buttonToggleConsole.Enabled = $true
    $buttonCancelRemoval.Enabled = $true
    $buttonConfirmedSelectedforRemoval.Enabled = $true
    $mainUI.Activate()
    $mainUI.ShowDialog() | Out-Null

    ################## GUI is Closed ##################################################################################

} else { # if running with -silent command line switch

    refreshProgramsList # get default $Script:progslisttoremove

    if ($Global:usingSavedSelectionFileSilentOption) {
        if (!(Test-Path "$($Global:usingSavedSelectionFileSilentOption)")) {
            Write-Warning "Saved selection file $($Global:usingSavedSelectionFileSilentOption) not found or doesnt exist!"
            Write-Warning "Please check the filename and provide the full path to the file`nif not using the default location and filename (c:\BRU\BRU-Saved-Selection.xml)"
            Write-Output "" | Out-Default
            stopTranscript
            Set-Location $savedPathLocation # Restore working directory path
            Return
        } else {
            try {
                $progslistSelected = Import-CliXml -Path "$($Global:usingSavedSelectionFileSilentOption)" -ErrorAction Stop
                Write-Output "" | Out-Default
                Write-Verbose "Using file $($Global:usingSavedSelectionFileSilentOption) for the selection list. Import successful." -Verbose
                Write-Output "" | Out-Default
            } catch {
                Write-Output "" | Out-Default
                Write-Warning "File $($Global:usingSavedSelectionFileSilentOption) failed to import for the selection list. Import Failed."
                Write-Output "" | Out-Default
                stopTranscript
                Set-Location $savedPathLocation # Restore working directory path
                Return
            }
        }
    }

    $Script:proglistviewColumnsArray = @('DisplayName','Name','Version','Publisher','UninstallString','QuietUninstallString','IdentifyingNumber','PackageFullName','PackageName')

    if (!($Global:usingSavedSelectionFileSilentOption)) {
        $progslistSelected = @( $Script:progslisttoremove | Select-Object -Property $proglistviewColumnsArray -ExcludeProperty DisplayName | Sort-Object Name )
    }

    if ( $Script:winVer -gt 6.1 ) {

        $Global:UWPappsAUtoRemove = $Global:UWPappsAUtoRemove | Select-Object -Property $proglistviewColumnsArray -ExcludeProperty DisplayName | Sort Name
        $Global:UWPappsProvisionedAppstoRemove = $Global:UWPappsProvisionedAppstoRemove | Select-Object -Property $proglistviewColumnsArray -ExcludeProperty Name | Select-Object @{Name="Name";Expression={$_.DisplayName}},* -ExcludeProperty DisplayName | Sort Name
        if (!($Global:usingSavedSelectionFileSilentOption)) {
            $progslistSelected += @( $Global:UWPappsAUtoRemove )
            $progslistSelected += @( $Global:UWPappsProvisionedAppstoRemove )
        }
    }

    [int]$Script:numofSelectedProgs = @('0',($progslistSelected | Measure-Object).Count)[($progslistSelected | Measure-Object).Count -gt 0]


} # end if ( !($Global:isSilent) )

if ( ($button -ne "Cancel") -or ($Global:isSilent) ) {

    if ( !($Global:isSilent) ) {
        showConsole | Out-Null
    }

    if ( ($Global:isSilent) ) {
        Write-Output "`nRunning silently using -silent switch."
    }
    Write-Output "`nUsing Options Chosen:" | Out-Default
    For( $i = 0; $i -lt (($Global:globalSettings.Keys).Count); $i++) {
        $settingname = ($Global:globalSettings.Keys | Select-Object -Index $i)
        Write-Output "$($settingname): $($Global:globalSettings[$($settingname)])" | Out-Default
    }
    Write-Output "`n" | Out-Default

    if ( !($Global:isSilent) ) {
        # Create array from the selected listview items
        # might be a way to use .CopyTo() method for object System.array and then change that to powershell array but this works
        $progslistSelected = selectedProgsListviewtoArray $programsListview
    }

    # At this poing, both silent or GUI based selections have been made

    if ( $progslistSelected -ne $null ) {

        # $progslistSelectedOriginal = $progslistSelected # save selection?

        Write-Output "Sorting programs selected in uninstall order..." | Out-Default

        $removeOrderedSelectedList = ,@()

        # sort $progslistSelected against $Script:progslisttoremove
        ForEach ($prog in $Script:progslisttoremove) {
            ForEach ( $selectedprog in $progslistSelected) {
                if ( isObjectEqual $selectedprog $prog ) {
                    $removeOrderedSelectedList += $selectedProg
                    $progslistSelected = $progslistSelected | Where { $_ -ne $selectedProg }
                }
            }
        }

        # add items selected that weren't in progslisttoremove but were selected to $progslistSelected
        $removeOrderedSelectedList += @(@($progslistSelected | Where { $_.Uninstallstring } | Sort-Object UninstallString) + @($progslistSelected | Where { !($_.Uninstallstring) } | Sort-Object UninstallString )) #split up into two parts then recombined to put the non-msi uninstallers first, then the msi ones last to reduce msi uninstall errors, needed because importing the xml saved selection file has all the properties added to the items already, so this reorders them correctly again.

        # pull UWP apps out of full list and back into their own variables
        $removeOrderedSelectedUWPappsAU = $removeOrderedSelectedList | Where { $_.PackageFullName }
        $removeOrderedSelectedUWPappsProvisioned = $removeOrderedSelectedList | Where { $_.PackageName }
        # remove the UWP apps from the ordered selected progslist
        $removeOrderedSelectedList = $removeOrderedSelectedList | Where { !($_.PackageFullName) -and !($_.PackageName) } | Select-Object -Skip 1

        Write-Output "" | Out-Default
        Write-Verbose -Verbose "Selected and ordered programs to be removed:"
        Write-Output $removeOrderedSelectedList | Out-Default

        if ( $Script:winVer -gt 6.1 ) {
            Write-Output "" | Out-Default
            Write-Verbose -Verbose "Selected and ordered UWPappsAU (Installed Win8/10+ apps) to be removed:"
            Write-Output $removeOrderedSelectedUWPappsAU | Select-Object Name, Version, Publisher, PackageFullName | Format-List | Out-Default

            Write-Output "" | Out-Default
            Write-Verbose -Verbose "Selected and ordered UWPappsProvisionedApps (Provisioned Win8/10+ apps) to be removed:"
            Write-Output $removeOrderedSelectedUWPappsProvisioned | Select-Object DisplayName, Version, PackageName | Format-List | Out-Default
        }

        Write-Output "" | Out-Default
        Write-Output "Total number of selected programs to be removed: $($Script:numofSelectedProgs)" | Out-Default

        # save original detections
        #$Script:progslisttoremoveOriginal = $Script:progslisttoremove

        $Script:progslisttoremove = $removeOrderedSelectedList


        if ( $Script:winVer -gt 6.1 ) {
            #$Global:UWPappsAUtoRemoveOrginal = $Global:UWPappsAUtoRemove
            #$Global:UWPappsProvisionedAppstoRemoveOrginal = $Global:UWPappsProvisionedAppstoRemove
            $Global:UWPappsAUtoRemove = $removeOrderedSelectedUWPappsAU
            $Global:UWPappsProvisionedAppstoRemove = $removeOrderedSelectedUWPappsProvisioned
        }

        Write-Output "" | Out-Default
        Write-Output "Up to this point no changes have been made. Below this point is where software starts being removed.`n" | Out-Default

        if ( !($Global:isDetectOnlyDryRunSilentOption) ) {

            $isConfirmed = systemRestorePointIfRequired

            if ( $Global:requireConfirmationBeforeRemoval -and $isConfirmed ) { # options chosen, no system restore point but wants confirmation
                Write-Host "`nProceed with Bloatware Removal?`n" | Out-Default
                [bool]$isConfirmed = doOptionsRequireConfirmation
            }

        }

        if ( !($isConfirmed) -or $Global:isDetectOnlyDryRunSilentOption ) {
            if ( $Global:isDetectOnlyDryRunSilentOption ) {
                Write-Verbose -Verbose "***DryRun DetectOnly WhatIf Mode***"
            }
            Write-Output "You have chosen to not proceed with removal. No changes will be made."
        }

        if ( ($Script:progslisttoremove -and $isConfirmed) -and !($Global:isDetectOnlyDryRunSilentOption) ) {
            Write-Output "" | Out-Default
            Write-Output "Copying Helper Files to $($Script:dest)\ ..." | Out-Default
            Write-Output "If running from removable media like a flash drive`nPlease do not remove it yet." | Out-Default

            if ( ($Script:progslisttoremove -match "HP Client Security Manager") -or ($Script:progslisttoremove -match "ProtectTools Security Manager") ) {
                if ( Test-Path "$($scriptPath)\BRU-uninstall-helpers\devcon_x$($Script:osArch).exe" ) {
                    Write-Output "" | Out-Default
                    Copy-Item -Verbose -Path "$($scriptPath)\BRU-uninstall-helpers\devcon_x$($Script:osArch).exe" -Destination $Script:dest
                } else {
                    Write-Warning "HP Client Security Manager uninstall helper devcon_x$($Script:osArch).exe not found in $($scriptPath)\BRU-uninstall-helpers\"  | Out-Default
                    Write-Warning "See: https://networchestration.wordpress.com/2016/07/11/how-to-obtain-device-console-utility-devcon-exe-without-downloading-and-installing-the-entire-windows-driver-kit-100-working-method/" | Out-Default
                }
            } # end if ( ($Script:progslisttoremove -match "HP Client Security Manager") -or ($Script:progslisttoremove -match "ProtectTools Security Manager") )

            if ( $Script:progslisttoremove -match "HP\ JumpStart\ Apps|VIP\ Access.*|Lenovo\ App\ Explorer" ) {
                if ( Test-Path "$($scriptPath)\BRU-uninstall-helpers\WASP.dll" ) {
                    Copy-Item -Verbose -Path "$($scriptPath)\BRU-uninstall-helpers\WASP.dll" -Destination $Script:dest
                } else {
                    Write-Warning "WASP uninstall helper WASP.dll (Windows Automation Snapin for PowerShell) not found in $($scriptPath)\BRU-uninstall-helpers\"  | Out-Default
                    Write-Warning "See: https://web.archive.org/web/20210701003321/https://archive.codeplex.com/?p=wasp" | Out-Default
                }
            } # end if ( $Script:progslisttoremove -match "HP\ JumpStart\ Apps|VIP\ Access.*|Lenovo\ App\ Explorer" )

            if ( $Script:progslisttoremove -match "Microsoft\ Office|Microsoft\ 365" ) {
                # Updated OffScrubc23.vbs for 2013/2016: https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Deploy-OfficeClickToRun/OffScrubc2r.vbs
                if ( Test-Path "$($scriptPath)\BRU-uninstall-helpers\OffScrubc2r.vbs" ) {
                    Copy-Item -Verbose -Path "$($scriptPath)\BRU-uninstall-helpers\OffScrubc2r.vbs" -Destination $Script:dest
                } else {
                    Write-Warning "Microsoft Office Click2Run (Trial/OEM) uninstall helper OffScrubc2r.vbs not found in $($scriptPath)\BRU-uninstall-helpers\"  | Out-Default
                    Write-Warning "See: https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Deploy-OfficeClickToRun/OffScrubc2r.vbs/" | Out-Default
                }
            } # end if ( $Script:progslisttoremove -match "Microsoft Office" )

            if ( $Script:progslisttoremove -match "McAfee" ) {
                if ( Test-Path "$($scriptPath)\BRU-uninstall-helpers\MCPR-10-5-374-0.exe" ) { # Version 10.5.374.0 works with mcclean version 10.5.128.0 which works silently be sure to provide both
                    Copy-Item -Verbose -Path "$($scriptPath)\BRU-uninstall-helpers\MCPR-10-5-374-0.exe" -Destination $Script:dest
                } else {
                    Write-Warning "McAfee uninstall helper MCPR-10-5-374-0.exe (McAfee Consumer Product Removal Tool) not found in $($scriptPath)\BRU-uninstall-helpers\"  | Out-Default
                    #Write-Warning "See: http://us.mcafee.com/apps/supporttools/mcpr/mcpr.asp" | Out-Default
                    Write-Warning "See: https://mcprtool.com/ (https://download.mcafee.com/molbin/iss-loc/SupportTools/MCPR/mcpr.exe)" | Out-Default
                }
                if ( Test-Path "$($scriptPath)\BRU-uninstall-helpers\mccleanup-10-5-128-0.exe" ) {
                    Copy-Item -Verbose -Path "$($scriptPath)\BRU-uninstall-helpers\mccleanup-10-5-128-0.exe" -Destination $Script:dest
                } else {
                    Write-Warning "McAfee uninstall helper mcclean.exe (mcclean version 10.5.128.0, part of McAfee Consumer Product Removal Tool) not found in $($scriptPath)\BRU-uninstall-helpers\"  | Out-Default
                    #Write-Warning "See: http://us.mcafee.com/apps/supporttools/mcpr/mcpr.asp" | Out-Default
                    Write-Warning "See: https://mcprtool.com/ (https://download.mcafee.com/molbin/iss-loc/SupportTools/MCPR/mcpr.exe)" | Out-Default
                }
            } # end if ( $Script:progslisttoremove -match "McAfee" )

            if ( !($Global:isSilent) ) { # ensure literal silence (no sounds :)
                #http://scriptolog.blogspot.com/2007/09/playing-sounds-in-powershell.html
                $soundloc = "c:\Windows\Media\Speech On.wav"
                if (Test-Path $soundloc) {
                    $sound = New-Object System.Media.SoundPlayer;
                    $sound.SoundLocation = $soundloc;
                    $sound.Play();
                }
            }

            Write-Output "" | Out-Default
            Write-Verbose "If you are running from a flash drive or removable media and need to remove it, it is now safe to do so." -Verbose


            #return

            # match the process's full path (group 1) and the argumentList (group 2)
            # https://regex101.com/ is a good tool to visualize the regex groups
            $uninstallstringmatchstring = "^(.*?)((\ +?[\/\-_].*)|([`"`']\ +?.*))$"
            $procnamelist = @()
            $MCPRalreadyran = $null
            $microsoftofficeC2Ralreadyran = $null

################# Main Uninstallation Loop ########################################################################

            ForEach( $prog in $Script:progslisttoremove ) {

                $functionAfterUninstallerStarted = $null # no additional cleanup by default unless defined for each specific program
                $waitForExitAfterUninstallerStarted = 1 # default is to wait for uninstaller to exit before continuing unless modified below, useful if needed to start the uninstaller but not wait on it to exit, such as when using SendKeys to to uninstaller

                $uninstallpath = $null
                $uninstallarguments = $null

                if ( $prog.QuietUninstallString ) { # Use QuietUninstallString if it exists
                    $proguninstallstring = $prog.QuietUninstallString
                    $returned = parseUninstallString $proguninstallstring $uninstallstringmatchstring
                    $uninstallpath = $returned[0]
                    $uninstallarguments = $returned[1]

                    Write-Output "`n$($prog.Name) is being removed. (Using QuietUninstallString)" | Out-Default

                    # Special Cases when using QuietUninstallString here

                    # VIP Access SDK (comes with Norton 2012) doesn't actually uninstall quietly, have to click close
                    if ( $prog.Name -like "VIP Access*" ) {
                       $waitForExitAfterUninstallerStarted = 0
                        $uninstallprocname = "uninstall"
                        function VIPAccessAfterUninstallerStarted {
                            $waittimeoutexitcode = waitForProcessToStartOrTimeout 'Au_' 15
                            if ( $waittimeoutexitcode -eq 0 ) {
                                sleepProgress (@{"Seconds" = 5})
                                # Using WASP.dll commands (Windows Automation Snapin for PowerShell)

                                $scriptblock = {

                                    param($Script:dest)
                                    Start-Sleep -Seconds 1

                                    if ( Test-Path "$($Script:dest)\WASP.dll" ) {
                                        $loadWASP = "$($Script:dest)\WASP.dll"
                                        [void][System.Reflection.Assembly]::LoadFrom("$($loadWASP)")
                                        Import-Module "$($Script:dest)\WASP.dll"
                                        if ( Select-Window "VIP Access*" | Where { $_.Title -match "VIP.*Uninstall" } ) {
                                            $a = Select-Window "VIP Access*" | Where { $_.Title -match "VIP.*Uninstall" } | Set-WindowActive
                                            $a = Select-Window "VIP Access*" | Select -First 1 | Select-Control -Title "&Close" | Send-Click
                                        }
                                    } else {
                                        Write-Warning "$($Script:dest)\WASP.dll wasn't found or couldn't be loaded." | Out-Default
                                    }
                                } # end $scriptblock

                                ###### LoadingWASP #########################################################
                                Start-Job $scriptblock -ArgumentList $Script:dest | Out-Null
                                Get-Job | Wait-Job | Receive-Job
                                ###### UnLoadingWASP #######################################################

                            } else {
                                Write-Warning "Uninstall of $($prog.Name) aborted due to wait timeout of uninstaller process name: $($uninstallprocname)" | Out-Default
                                Write-Warning "Reboot and manually remove program." | Out-Default
                            }
                            sleepProgress (@{"Seconds" = 20})
                        }
                            $functionAfterUninstallerStarted = "VIPAccessAfterUninstallerStarted"
                    } # end if ( $prog.Name -like "VIP Access*" )


                } else { # no QuietUninstallString exists

                    if ( $prog.IdentifyingNumber -ne $null `
                         -or ( ($prog.IdentifyingNumber -eq $null) `
                               -and ( $($prog.UninstallString) -match "msiexec" -and $($prog.UninstallString) -match $Global:guidmatchstring ) `
                             ) ) { # it is MSI uninstaller
                        if ( $prog.IdentifyingNumber) {
                           $id = $prog.IdentifyingNumber
                        } else {
                            $id = $matches[0] # $($prog.UninstallString) -match $Global:guidmatchstring
                        }
                        $uninstallpath = "msiexec.exe"
                        $uninstallarguments = "/x$($id) /qn /norestart"

                        # Special Case MSI Uninstallers

                        if ( $prog.Name -match "HP Auto" `
                             -or $prog.Name -match "HP Odometer" ) {
                            $procnamelist = @('hpsysdrv') # HP Odometer/Auto "telemetry data"
                            stopProcesses( $procnamelist )
                        }

                        if ( $prog.Name -match "HP Client Security Manager" `
                             -or $prog.Name -match "ProtectTools Security Manager" ) { # msi uninstaller
                            $procnamelist = @('dpcardengine',
                                              'dphostw',
                                              'DpAgent',
                                              'DPAdminWizard',
                                              'DPClientWizard',
                                              'sidebar')  # Native Windows sidebar process, stop to remove HP Gadget
                            stopProcesses( $procnamelist )
                            Write-Output "" | Out-Default
                            Write-Verbose -Verbose "Disabling DVD/CD Drive so HP Client Security Manager/HP ProtectTools won't throw an error 1325 about non-valid short file name."
                            Write-Verbose -Verbose "It will be reenabled after Client Security Manager (ProtectTools) is uninstalled."
                            if ( Test-Path "$($Script:dest)\devcon_x$($Script:osArch).exe" ) {
                                # disable DVD/CD drive first to prevent uninstall bug with HP Client Security Manager
                                # http://h30434.www3.hp.com/t5/Notebook-Operating-System-and-Recovery/Error-1325-upon-trying-to-uninstall-HP-ProtectTools-Security/td-p/5253850
                                # BUG in HP's uninstaller still exists in 2017 Client Security Manager. Disabling, Uninstalling, then Reenabling DVD drive works. Usually happens with laptop HP DVD drives.
                                & cmd /c "$($Script:dest)\devcon_x$($Script:osArch).exe" disable =CDROM 2>&1 | Out-Default
                            } else {
                                Write-Output "" | Out-Default
                                Write-Warning "devcon_x$($Script:osArch).exe not found in $($Script:dest)\`n" | Out-Default
                                Write-Warning "See: https://networchestration.wordpress.com/2016/07/11/how-to-obtain-device-console-utility-devcon-exe-without-downloading-and-installing-the-entire-windows-driver-kit-100-working-method/" | Out-Default
                                Write-Warning "If you need to you can manually disable the DVD/CD drive and uninstall HP Client Security Manager (or ProtectTools) then reenable the DVD/CD drive." | Out-Default
                            } # end else (devcon not found)
                            function HPClientSecurityManagerAfterUninstallerStarted {
                                if ( Test-Path "$($Script:dest)\devcon_x$($Script:osArch).exe" ) {
                                    #reenable DVD/CD drive
                                    Write-Output "" | Out-Default
                                    Write-Verbose -Verbose "Reenabling DVD/CD Drive..."
                                    & cmd /c "$($Script:dest)\devcon_x$($Script:osArch).exe" enable =CDROM 2>&1 | Out-Default
                                }
                                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "C:\Program Files (x86)\HP\HP ProtectTools Security Manager"
                                # Remove HP ProtectTools/Client Security Sidebar Gadget
                                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "C:\Program Files\Windows Sidebar\Gadgets\DPIDCard.Gadget"
                                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Security and Protection"
                            }
                            $functionAfterUninstallerStarted = "HPClientSecurityManagerAfterUninstallerStarted"
                        } # end  if ( $prog.Name -match "HP Client Security Manager" -or $prog.Name -match "ProtectTools Security Manager" )

                        if ( $prog.Name -match "HP Customer Experience Enhancements" ) {
                            function functionHPFeedbackUninstallerStarted {
                                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "C:\Program Files (x86)\Hewlett-Packard\HP Customer Feedback"
                            }
                            $functionAfterUninstallerStarted = "functionHPFeedbackUninstallerStarted"
                        } # end HP Customer Experience Enhancements msiexec uninstaller

                        if ( $prog.Name -match "HP JumpStart Bridge" ) {
                            $procnamelist = @('HPJumpStartBridge')
                            stopProcesses( $procnamelist )
                        }

                        if ( $prog.Name -match "HP JumpStart Launch" ) {
                            $procnamelist = @('hpjumpstartprovider')
                            stopProcesses( $procnamelist )
                        }

                        if ( $prog.Name -match "HP Notifications" ) {
                            $procnamelist = @('hpnotifications')
                            stopProcesses( $procnamelist )
                        }

                        if ( $prog.Name -match "HP Registration Service" ) {
                            $procnamelist = @('HPRegistrationService')
                            stopProcesses( $procnamelist )
                        }

                        if ( $prog.Name -match "HP Support Assistant" ) { # HPSA msi uninstaller
                            $uninstallarguments = "/x$($id) /qn /norestart UninstallKeepPreferences=TRUE" # Only way to skip prompt asking to save preferences in version 8.1.52.1, also works silently for older and newer versions of HPSA
                            function functionHPSAMSIAfterUninstallerStarted {
                                Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($id)" -Force -ErrorAction SilentlyContinue
                                Remove-Item "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($id)" -Force -ErrorAction SilentlyContinue
                            }
                            $functionAfterUninstallerStarted = "functionHPSAMSIAfterUninstallerStarted"
                        } # end HPSA msi uninstaller

                        if ( $prog.Name -match "HP Support Solutions Framework" ) {
                            $procnamelist = @('HPSupportSolutionsFrameworkService')
                            stopProcesses( $procnamelist )
                        }

                        # end Special Case MSI Uninstallers

                        Write-Output "`n$($prog.Name) is being removed. (msi)" | Out-Default

                    } else { # non-msi such as exe, cmd, bat, etc

                        #$prog.Name = "McAfee something or other"
                        #$prog.Name = "HP JumpStart Apps"
                        #$prog.UninstallString = "c:\BRU\uninstall.exe"
                        #$prog.name = "HP Client Security Manager"
                        #$prog.UninstallString = "`"C:\Program Files (x86)\InstallShield Installation Information\{6468C4A5-E47E-405F-B675-A70A70983EA6}\setup.exe`" -runfromtemp -l0x0409  -removeonly"

                        Write-Output "`n$($prog.Name) is being removed. (non-msi)" | Out-Default

                        $proguninstallstring = $prog.UninstallString
                        $returned = parseUninstallString $proguninstallstring $uninstallstringmatchstring

                        $uninstallpath = $returned[0]
                        $uninstallarguments = $returned[1]

                        # Special Case Non-MSI Uninstallers
                        if ( $prog.Name -match "Adobe Air" ) {
                            $uninstallarguments = "-uninstall"
                        }

                        if (( $prog.Name -match "Dell Optimizer Service|Dell Optimizer Core|Dell Optimizer|Dell Precision Optimizer" )) {
                            $uninstallarguments = "-silent"+" "+$uninstallarguments
                        }

                        if ( $prog.Name -match "Dell SupportAssist" ) {
                            $uninstallarguments = "/S"+" "+$uninstallarguments
                        }

                        if ( $prog.Name -match "Dropbox.*" ) {
                            $uninstallarguments = "/S"+" "+$uninstallarguments
                        }

                        if ( $prog.Name -match "HP Client Security Manager" `
                             -or $prog.Name -match "ProtectTools Security Manager" ) {
                            Continue # skip exe installer
                        }

                        if ( $prog.Name -match "HP Collaboration Keyboard" ) {
                            Continue
                        }

                        # HP ePrint likely has a QuietUninstallString but keeping this in case it doesn't.
                        if ( $prog.Name -match "HP ePrint" ) {
                            $procnamelist = @('HP.DeliveryAndStatus.Desktop.App') # HP ePrint
                            stopProcesses( '$procnamelist' )
                            $uninstallarguments = $uninstallarguments+" "+"/quiet"
                        }

                        if ( $prog.Name -match "HP JumpStart Apps" ) {
                            $waitForExitAfterUninstallerStarted = 0
                            $uninstallprocname = "uninstall"
                            function HPJumpStartAppsAfterUninstallerStarted {
                                $waittimeoutexitcode = waitForProcessToStartOrTimeout $uninstallprocname 20
                                if ( $waittimeoutexitcode -eq 0 ) {
                                    sleepProgress (@{"Seconds" = 5})
                                    # Using WASP.dll commands (Windows Automation Snapin for PowerShell)

                                    $scriptblock = {

                                        param($Script:dest)
                                        Start-Sleep -Seconds 1

                                        Write-Output "$($Script:dest)\WASP.dll" | Out-Default

                                        if ( Test-Path "$($Script:dest)\WASP.dll" ) {
                                            $loadWASP = "$($Script:dest)\WASP.dll"
                                            [void][System.Reflection.Assembly]::LoadFrom("$($loadWASP)")
                                            Import-Module "$($Script:dest)\WASP.dll"
                                            $a = Select-Window $uninstallprocname | Where { $_.Title -match "HP JumpStart Apps"} | Set-WindowActive
                                            $a = Select-Window $uninstallprocname | Select -First 1 | Select-Control -Title "OK" | Send-Click
                                        } else {
                                            Write-Warning "$($Script:dest)\WASP.dll wasn't found or couldn't be loaded." | Out-Default
                                        }
                                    } # end of $scriptblock

                                    ###### LoadingWASP #########################################################
                                    Start-Job $scriptblock -ArgumentList $Script:dest | Out-Null
                                    Get-Job | Wait-Job | Receive-Job
                                    ###### UnLoadingWASP #######################################################

                                } else {
                                    Write-Warning "Uninstall of $($prog.Name) aborted due to wait timeout of uninstaller process name: $($uninstallprocname)" | Out-Default
                                    Write-Warning "Reboot and manually remove program." | Out-Default
                                }
                                sleepProgress (@{"Seconds" = 30})
                                Remove-Item "$Script:dest\Connecting" -Force -Verbose -ErrorAction SilentlyContinue
                                Remove-Item "$Script:dest\1" -Force -Verbose -ErrorAction SilentlyContinue
                            }
                            $functionAfterUninstallerStarted = "HPJumpStartAppsAfterUninstallerStarted"
                        }

                        if ( $prog.Name -match "HP Setup" ) {
                            $procnamelist = @('HPTCS')
                            stopProcesses( $procnamelist )
                            $prog.UninstallString -match $Global:guidmatchstring | Out-Null
                            if ( $matches ) {
                                $id = $matches[0]
                                Set-Content -Path "$($Script:dest)\hpsetup.iss" -Value "[InstallShield Silent]`r`nVersion=v7.00`r`nFile=Response File`r`n[File Transfer]`r`nOverwrittenReadOnly=NoToAll`r`n[$($id)-DlgOrder]`r`nDlg0=$($id)-SdWelcomeMaint-0`r`nCount=3`r`nDlg1=$($id)-MessageBox-0`r`nDlg2=$($id)-SdFinishReboot-0`r`n[$($id)-SdWelcomeMaint-0]`r`nResult=303`r`n[$($id)-MessageBox-0]`r`nResult=6`r`n[Application]`r`nName=HP Setup`r`nVersion=$($prog.Version)`r`nCompany=Hewlett-Packard Company`r`nLang=0009`r`n[$($id)-SdFinishReboot-0]`r`nResult=1`r`nBootOption=0`r`n"
                                $uninstallarguments = "-s -SMS -w -clone_wait -f1`"$($Script:dest)\hpsetup.iss`""+" "+"-removeonly"
                                function functionHPSetupAfterUninstallerStarted {
                                    Remove-Item "$Script:dest\hpsetup.iss" -Force -Verbose -ErrorAction SilentlyContinue
                                    Remove-Item "$Script:dest\setup.log" -Force -Verbose -ErrorAction SilentlyContinue
                                }
                                $functionAfterUninstallerStarted = "functionHPSetupAfterUninstallerStarted"
                            }
                        } # end HP Setup

                        if ( $prog.Name -match "HP Support Assistant" ) {
                            $procnamelist = @('HPSA_Service') # HP Support Assistant
                            stopProcesses( $procnamelist )
                            $HPSAuninstallpaths = @("C:\ProgramData\Hewlett-Packard\UninstallHPSA.exe",
                                                    "C:\Program Files\Hewlett-Packard\HP Health Check\Tools\UninstallHPSA.exe",
                                                    "C:\Program Files (x86)\Hewlett-Packard\HP Health Check\Tools\UninstallHPSA.exe",
                                                    "C:\Program Files (x86)\Hewlett-Packard\HP Support Framework\UninstallHPSA.exe" )
                            $HPSAuninstallpaths += @( (Get-ChildItem -Recurse "C:\SWSETUP\*\UninstallHPSA.exe").FullName )

                            $prog.UninstallString -match $Global:guidmatchstring | Out-Null
                            $id = $matches[0]

                            $ErrorActionPreference = "SilentlyContinue"

                            ForEach ( $path in $HPSAuninstallpaths ) {
                                if ( Test-Path $path -ErrorAction SilentlyContinue ) { # HPSA exe uninstaller
                                    $uninstallpath = $path
                                    Write-Output "Found HPSA Uninstaller: $path" | Out-Default
                                    $uninstallarguments = "/ProductCode $($id) UninstallKeepPreferences=TRUE"
                                    break # just run the first uninstaller that is found, they're all identical copies in multiple locations
                                }
                            }
                            if ( !(Test-Path $path) ) {
                                Write-Warning "No valid HPSA uninstall paths. Search filesystem for UninstallHPSA.exe and update this script with new paths in `$HPSAuninstallpaths" | Out-Default
                                $ErrorActionPreference = "Continue"
                                Continue # skip running this non-msi uninstaller
                            }
                            $ErrorActionPreference = "Continue"
                            function functionHPSAAfterUninstallerStarted {
                                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue $uninstallpath # Remove exe uninstaller
                                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "C:\ProgramData\Hewlett-Packard\UninstallHPSA.exe"
                                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "C:\Program Files\Hewlett-Packard\HP Health Check\Tools\UninstallHPSA.exe"
                                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "C:\Program Files (x86)\Hewlett-Packard\HP Health Check\Tools\UninstallHPSA.exe"
                                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "C:\users\Public\Desktop\HP Support Assistant.lnk"
                                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "$($Script:dest)\InstallUtil.InstallLog"
                                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Hewlett-Packard\HP Support Assistant"
                            }
                            $functionAfterUninstallerStarted = "functionHPSAAfterUninstallerStarted"
                        } # end HP Support Assistant

                        if ( $prog.Name -match "HP Sure Connect" -or $prog.Name -match "HP Connection Optimizer" -or $prog.Name -match "HP Wireless Rescue Tool" ) { # Works with version 1.0.0.29 likely newer as well
                            $procnamelist = @('HPCommRecovery') # HP Sure Connect
                            stopProcesses( $procnamelist )
                            $prog.UninstallString -match $Global:guidmatchstring | Out-Null
                            if ( $matches ) {
                                $id = $matches[0]
                                Set-Content -Path "$($Script:dest)\hpsureconnect.iss" -Value "[InstallShield Silent]`r`nVersion=v7.00`r`nFile=Response File`r`n[File Transfer]`r`nOverwrittenReadOnly=NoToAll`r`n[$($id)-DlgOrder]`r`nDlg0=$($id)-MessageBox-0`r`nCount=2`r`nDlg1=$($id)-SdFinish-0`r`n[$($id)-MessageBox-0]`r`nResult=6`r`n[Application]`r`nName=$($prog.Name)`r`nVersion=$($prog.Version)`r`nCompany=HP Inc.`r`nLang=0409`r`n[$($id)-SdFinish-0]`r`nResult=1`r`nbOpt1=0`r`nbOpt2=0`r`n"
                                $uninstallarguments = "-s -SMS -w -clone_wait -f1`"$($Script:dest)\hpsureconnect.iss`""+" "+"-removeonly"
                                function functionHPSureConnectAfterUninstallerStarted {
                                    Remove-Item "$Script:dest\hpsureconnect.iss" -Force -Verbose -ErrorAction SilentlyContinue
                                    Remove-Item "$Script:dest\setup.log" -Force -Verbose -ErrorAction SilentlyContinue
                                    Remove-Item -Recurse -Force -Verbose -ErrorAction SilentlyContinue "C:\Program Files\HPCommRecovery\"
                                }
                                $functionAfterUninstallerStarted = "functionHPSureConnectAfterUninstallerStarted"
                            } # end if ( $matches )
                        } # end HP Sure Connect/HP Connection Optimizer

                        if ( $prog.Name -match "HP Theft Recovery" `
                             -or $prog.Name -match "Theft Recovery for HP ProtectTools" ) {
                            # HP Theft Recovery for HP ProtectTools
                            # AKA Computrace
                            $procnamelist = @('CTService') # HP TheftRecovery "Computrace"
                            stopProcesses( $procnamelist )
                            $prog.UninstallString -match $Global:guidmatchstring | Out-Null
                            if ( $matches ) {
                                $id = $matches[0]
                                Set-Content -Path "$($Script:dest)\theftrecovery.iss" -Value "[$($id)-DlgOrder]`r`nDlg0=$($id)-MessageBox-0`r`nCount=2`r`nDlg1=$($id)-SdFinish-0`r`n[$($id)-MessageBox-0]`r`nResult=6`r`n[$($id)-SdFinish-0]`r`nResult=1`r`nbOpt1=0`r`nbOpt2=0`r`n"
                                # uninstallarguments needs to remain short as too many arguments will cause IS to not start uninstall
                                $uninstallarguments = "-s -f1`"$($Script:dest)\theftrecovery.iss`""+" "+"-removeonly"#" -l0x0409 -removeonly"
                                function functionHPTheftRecoveryAfterUninstallerStarted {
                                    Remove-Item "$Script:dest\theftrecovery.iss" -Force -Verbose -ErrorAction SilentlyContinue
                                    Remove-Item "$Script:dest\setup.log" -Force -Verbose -ErrorAction SilentlyContinue
                                }
                                $functionAfterUninstallerStarted = "functionHPTheftRecoveryAfterUninstallerStarted"
                                Stop-Process -Name "CTService" -Force -ErrorAction SilentlyContinue # make sure not running
                            }
                        } # end HP Theft Recovery

                        if ( $prog.Name -match "HP Velocity" ) {
                            $uninstallarguments = "/S"+" "+$uninstallarguments
                        }

                        if ( $prog.Name -match "HP WorkWise" ) {
                            $uninstallarguments = "-silent"+" "+$uninstallarguments
                       }

                        # INSTALLSHIELD
                        # The following should catch most non special cases of InstallShield that respect the /S switch.
                        # Also try -silent switch for newer InstallShield cases.
                        # If it doesn't it will need to be made into a special case and likely an iss file recorded.
                        # For InstallShield uninstalls that do not respect the /S (or -silent) silent switch we create an InstallShield Response
                        # file for silent uninstallation by running the InstallShield -r (record uninstall)
                        # Then we play back that recorded iss file.
                        # http://publib.boulder.ibm.com/tividd/td/framework/GC32-0804-00/en_US/HTML/instgu25.htm
                        # http://helpnet.flexerasoftware.com/installshield19helplib/helplibrary/IHelpSetup_EXECmdLine.htm
                        if ( $prog.UninstallString -match "InstallShield" `
                            -and ( ($prog.Name -notmatch $Global:specialcasestoremovesinglestring `
                                    -and $prog.Name -notmatch "CyberLink\ Media.*Suite" ) `
                            -or    ($prog.Name -match $Global:specialcasestoremovesinglestring `
                                    -and $prog.Name -match "CyberLink\ Media.*Suite")) `
                            -and ( $prog.Name -notmatch "HP\ Collaboration\ Keyboard" ) `
                            -and ( $prog.Name -notmatch "HP\ Connection\ Optimizer" ) `
                            -and ( $prog.Name -notmatch "Dell\ Optimizer\ Service|Dell\ Optimizer|Dell\ Precision\ Optimizer" ) `
                            -and ( $prog.Name -notmatch "MyDell" ) `
                            ) {
                                $uninstallarguments = "/S"+" "+$uninstallarguments
                        }

                        if ( $prog.Name -match "Lenovo App Explorer" ) {
                            $waitForExitAfterUninstallerStarted = 0
                            $uninstallprocname = "Un_A"
                            function LenovoAppExplorerAfterUninstallerStarted {
                                $waittimeoutexitcode = waitForProcessToStartOrTimeout $uninstallprocname 20
                                if ( $waittimeoutexitcode -eq 0 ) {
                                    sleepProgress (@{"Seconds" = 5})
                                    # Using WASP.dll commands (Windows Automation Snapin for PowerShell)

                                    $scriptblock = {

                                        param($Script:dest)
                                        Start-Sleep -Seconds 1

                                        #Write-Output "$($Script:dest)\WASP.dll" | Out-Default
                                        Write-Output "Running GUI automation tool 'Windows Automation Snapin for PowerShell'..."

                                        if ( Test-Path "$($Script:dest)\WASP.dll" ) {
                                            $loadWASP = "$($Script:dest)\WASP.dll"
                                            [void][System.Reflection.Assembly]::LoadFrom("$($loadWASP)")
                                            Import-Module "$($Script:dest)\WASP.dll"
                                            $a = Select-Window $uninstallprocname | Where { $_.Title -match "Uninstall Lenovo App Explorer"} | Set-WindowActive
                                            $a = Select-Window $uninstallprocname | Select -First 1 | Select-Control -Title "Uninstall Lenovo App Explorer" | Send-Click
                                        } else {
                                            Write-Warning "$($Script:dest)\WASP.dll wasn't found or couldn't be loaded." | Out-Default
                                        }
                                    } # end of $scriptblock

                                    ###### LoadingWASP #########################################################
                                    Start-Job $scriptblock -ArgumentList $Script:dest | Out-Null
                                    Get-Job | Wait-Job | Receive-Job
                                    ###### UnLoadingWASP #######################################################

                                } else {
                                    Write-Warning "Uninstall of $($prog.Name) aborted due to wait timeout of uninstaller process name: $($uninstallprocname)" | Out-Default
                                    Write-Warning "Reboot and manually remove program." | Out-Default
                                }

                                # Uninstall Started, wait for it to finish...
                                $waittimeoutexitcode = waitForProcessToStopOrTimeout $uninstallprocname 20
                            }
                            $functionAfterUninstallerStarted = "LenovoAppExplorerAfterUninstallerStarted"
                        }

                        if ( $prog.Name -match "McAfee Security Scan Plus" ) {
                            $uninstallarguments = "/S /inner"
                        }

                        if ( $prog.Name -match "McAfee" -and $prog.Name -notmatch "McAfee Security Scan Plus" ) {

                            if ( $MCPRalreadyran ) {
                                Write-Output "$($prog.Name) has already been removed." | Out-Default
                                Continue # skip if already ran it for a previous McAfee product
                            }
                            # Remove all McAfee consumer products
                            $MCPRalreadyran = 1
                            if ( (Test-Path "$($Script:dest)\MCPR-10-5-374-0.exe") -and (Test-Path "$($Script:dest)\mccleanup-10-5-128-0.exe")) {

                                Start-Process "$($Script:dest)\MCPR-10-5-374-0.exe" # Starting it will extract its contents to a temporary folder
                                $waittimeoutexitcode = waitForProcessToStartOrTimeout "McClnUI" 45

                                $mcprDir = $null
                                $mcprProcessPath = (Get-Process mcclnui -ErrorAction SilentlyContinue | Select-Object Path).Path
                                if ($mcprProcessPath) {
                                    $mcprDir = Split-Path $mcprProcessPath
                                }
                                if ($mcprDir) {
                                    if (!(Test-Path "$($mcprDir)")) {
                                        $mcprDir = $null
                                    }
                                }
                                if (!($mcprDir)) {
                                    $recentTempDirs = Get-ChildItem "$($env:temp)" | Where-Object { $_.PSIsContainer -and $_.LastWriteTime -gt (Get-Date).AddHours(-1) }
                                    $mcprTempDirs = (Get-ChildItem -Recurse $recentTempDirs.FullName -Filter "McClnUI.exe").FullName
                                    if ($mcprTempDirs) {
                                        $mcprDirs = Split-Path $mcprTempDirs
                                    }
                                    if ($mcprDirs) {
                                        $mcprDir = ($mcprDirs | Get-Item | Sort-Object LastWriteTime | Select-Object -Last 1).FullName
                                    }
                                }
                                Start-Sleep -Seconds 10
                                Start-Process "taskkill.exe" -ArgumentList "/im `"McClnUI.exe`" /f"
                                if ( Test-Path "$mcprDir" ) {
                                    Set-Location "$mcprDir"
                                    if ( Test-Path "$($mcprDir)\mccleanup.exe" ) {
                                        $uninstallpath = "$($mcprDir)\mccleanup.exe"
                                        Rename-Item -Path "$($mcprDir)\mccleanup.exe" -NewName "mccleanup.exe.new" -Force
                                        #copy old one 10.5.128.0
                                        Copy-Item -Path "$($Script:dest)\mccleanup-10-5-128-0.exe" -Destination "$($mcprDir)" -Force
                                        if ( Test-Path "$($Script:dest)\mccleanup-10-5-128-0.exe" ) {
                                            Rename-Item -Path "$($Script:dest)\mccleanup-10-5-128-0.exe" -NewName "mccleanup.exe" -Force
                                        }

                                        #read version from file
                                        $mcprVersion = (Get-Item "$uninstallpath" | Select-Object VersionInfo).VersionInfo.ProductVersion

                                        $mcprKeys = @{ # key needed to run silently
                                            "10.5.128.0" = "83D598B28704805D599AE0512AB8066E31DCE48D6BD9691F304FD895B191EECBCD86900CFF18E603CFFE8D7C27F362E53B70ACFDA1D37F101A7CFB3856D2D8825AAEE68DAF7C988D46EDC54D2D26ECC333F6BA0D7D22873D45ECEFA2BB9FC87DC244F5B791222430B708DF22895F8AC3145986F07477A545A0E80A5556B372E7DEB7CD959C715F65034C68D657F62A9206582AB1244FA8AAA6917BEAD2E85F18D66F3A66FF4282EFA57AD3A594E18060C0A2025D9EA8B1D9877CE83C7BEC08C05486C99308FF967895F3324D082669D00BBDFE004B9D4243580A0C7103664DF768D10E6BCC479553D3159E4DD51BC81418368B3C790155552AE8817015FE2DEF9C77ABFF09AC18A9CD0F01D7871B9CA9D2D15B6F047CB043A9B201F730BA20B70AD88344AFE9E03151C5C700B9E1C1C4";
                                            "10.4.194.0" = "83D598B28704805D599AE0512AB8066E31DCE48D6BD9691F304FD895B191EECBCD86900CFF18E603CFFE8D7C27F362E53B70ACFDA1D37F101A7CFB3856D2D882BC9078509FAA05370EA70FA186EC44AE3A657B43EC9559FDEF33C6E8AAF0D7BDB71F264D419E66EBCA2045AF9717434E8A4AAE1FA6F7F2A6EE6EE4F37FA199298DDAFF1F1E3124F4837EAA344CA44ADC129C0C9A1C112CA77050705A304AA3428E264FF96942728C839D4B675753DCFED36D95CF1E5FA3F0F8DFA7C5FEF32C481D8160BA8A96CE44BDC1E3B3F3B198456633E83E467775AD0BBF0E8FC09C94150F1F2FE79E13247DD89EF520425269A557765E64EE0F73208A078FCAE244F317CCE7006FBFCB354401D044FF08FBF800477F0BB5415682DE406DF0BADF6624761F76E0EFAB9543BBB924149A64B9BB4A";

                                        }

                                        if ($mcprVersion) {
                                            $mcprKey = $mcprKeys[$mcprVersion]
                                        }


                                        $uninstallarguments = "-p $mcprKey -silent StopServices,MFSY,PEF,MXD,CSP,Sustainability,MOCP,MFP,APPSTATS,Auth,EMproxy,FWdiver,HW,MAS,MAT,MBK,MCPR,McProxy,McSvcHost,VUL,MHN,MNA,MOBK,MPFP,MPFPCU,MPS,SHRED,MPSCU,MQC,MQCCU,MSAD,MSHR,MSK,MSKCU,MWL,NMC,RedirSvc,VS,REMEDIATION,MSC,YAP,TRUEKEY,LAM,PCB,Symlink,SafeConnect,MGS,WMIRemover,RESIDUE"
                                        Write-Output "MCPR may take quite a while to run. Please wait..." | Out-Default
                                        #unpin McAfee LiveSafe from taskbar
                                        function UnPinFromTaskbar { param( [string]$appname )
                                            # must be run from user context (not system)
                                            Try {
                                                ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ? { $_.Name -eq $appname }).Verbs() | ? { $_.Name -like 'Unpin from*' } | % { $_.DoIt() }
                                            } Catch {
                                                Write-Output "Did not UnPin $appname, it may not be pinned."
                                            }
                                        }
                                        Start-Sleep -Seconds 4
                                        UnPinFromTaskbar 'McAfee LiveSafe'
                                        function functionMcAfeeAfterUninstallerStarted {
                                            Set-Location $Script:dest
                                            Remove-Item "$mcprDir" -Force -Recurse -ErrorAction SilentlyContinue
                                            Remove-Item "$($Script:dest)\MCPR-10-5-374-0.exe" -Force -ErrorAction SilentlyContinue
                                            Remove-Item "$($Script:dest)\mccleanup-10-5-128-0.exe" -Force -ErrorAction SilentlyContinue
                                        }
                                        $functionAfterUninstallerStarted = "functionMcAfeeAfterUninstallerStarted"
                                    }
                                } else {
                                    Write-Warning "No directory in $($env:temp) found with McClnUI.exe." | Out-Default
                                    Write-Warning "Probably low disk space." | Out-Default
                                    Write-Warning "Try clearing room, check write permissions and then manually uninstall $($prog.Name) or run MCPR.exe and reboot." | Out-Default
                                    Continue
                                }
                            } else {
                                Write-Warning "McAfee uninstall helper MCPR.exe (McAfee Consumer Product Removal Tool) not found in $($Script:dest)\"  | Out-Default
                                Write-Warning "MCPR will remove all McAfee consumer products and likely $($prog.Name)." | Out-Default
                                #Write-Warning "Download from McAfee: http://us.mcafee.com/apps/supporttools/mcpr/mcpr.asp" | Out-Default
                                Write-Warning "Download from McAfee: https://mcprtool.com (https://download.mcafee.com/molbin/iss-loc/SupportTools/MCPR/mcpr.exe)" | Out-Default
                                Continue
                            }
                        } # end if ( $prog.Name -match "McAfee" )

                        if ( $prog.Name -match "Microsoft\ Office|Microsoft\ 365" ) {
                            if ( $microsoftofficeC2Ralreadyran ) {
                                Write-Output "$($prog.Name) has already been removed." | Out-Default
                                Continue # skip if already ran it as running the remover once gets all languages/installs
                            }
                            $microsoftofficeC2Ralreadyran = 1 # Skip next time so this only is run once
                            $uninstallpath = "wscript.exe"
                            $uninstallarguments = " //B //NoLogo `"$($Script:dest)\OffScrubc2r.vbs`" ALL /Quiet /NoCancel"
                            Write-Output "Using OffScrubc2r Office Click To Run (C2R) Remover. This will take a few minutes. Please wait.`n" | Out-Default
                            Write-Output "If it returns exitcode 42 from OffScrubc2r that is normal and means the program was removed sucessfully." | Out-Default
                            function functionMicrosoftOfficeAfterUninstallerStarted {
                                Remove-Item "C:\Users\Public\Desktop\Microsoft Office 2010.lnk" -Force -Verbose -ErrorAction SilentlyContinue

                            }
                            $functionAfterUninstallerStarted = "functionMicrosoftOfficeAfterUninstallerStarted"

                        }

                        if ( $prog.Name -like "Microsoft Security Essentials*" ) {
                            $procnamelist = @('epplauncher') # MSE installer provided through Windows Updates
                            stopProcesses( $procnamelist )
                            $uninstallarguments = "/x /s /u"
                        }


                        if (( $prog.Name -match "MyDell" )) {
                            $uninstallarguments = "-silent"+" "+$uninstallarguments
                        }


                        if ( $prog.Name -like "NewBlue Video Essentials" ) {
                        # Bundled with CyberLink Media Suite Essentials
                            $uninstallarguments = "/S"+" "+$uninstallarguments
                        }

                        if ( $prog.Name -like "Norton Internet Security" ) {
                        # Works with Norton Internet Security 2012 version 19.0.0.128
                        function NortonInternetSecuritySAfterUninstallerStarted {
                            Start-Sleep -Seconds 4
                            if ( Get-Process instStub ) {
                                $uninstallerstarted = (New-Object -ComObject WScript.Shell).AppActivate((Get-Process instStub).MainWindowTitle)
                                #[void] [System.Reflection.Assembly]::LoadWithPartialName("'System.Windows.Forms" )
                                if ($uninstallerstarted) {
                                    Add-Type -AssemblyName System.Windows.Forms
                                    [System.Windows.Forms.SendKeys]::SendWait("%{TAB}%{TAB}" )

                                    <#
                                    # For Beta version of NIS2012, 'Thank you for being a BETA tester screen'
                                    Start-Sleep -Seconds 4
                                    [System.Windows.Forms.SendKeys]::SendWait("%{TAB}" )
                                    [System.Windows.Forms.SendKeys]::SendWait("{TAB}{TAB}" )
                                    Start-Sleep -Seconds 4
                                    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}" )
                                    Start-Sleep -Seconds 4
                                    #>

                                    [System.Windows.Forms.SendKeys]::SendWait("{TAB}{TAB}{ENTER}" )
                                    Start-Sleep -Seconds 4
                                    [System.Windows.Forms.SendKeys]::SendWait("{TAB}{TAB}{ENTER}" )
                                    Start-Sleep -Seconds 4
                                    [System.Windows.Forms.SendKeys]::SendWait("{TAB}{TAB}{ENTER}" )
                                    Start-Sleep -Seconds 4
                                    [System.Windows.Forms.SendKeys]::SendWait("{TAB}{TAB}{ENTER}" )
                                    sleepProgress (@{"Seconds" = 180})
                                    if ( Get-Process instStub ) {
                                        (New-Object -ComObject WScript.Shell).AppActivate((Get-Process instStub).MainWindowTitle) | Out-Null
                                        [System.Windows.Forms.SendKeys]::SendWait("{TAB}{TAB}{TAB}{ENTER}" )
                                    } # end if ( Get-Process instStub )


                                } # end if ($uninstallerstarted)
                            } # end if ( Get-Process instStub )
                            $proc.WaitForExit()
                            Remove-Item "C:\Users\Public\Desktop\SafeWeb.url" -Force -ErrorAction SilentlyContinue
                        } # end function NortonInternetSecuritySAfterUninstallerStarted
                        $functionAfterUninstallerStarted = "NortonInternetSecuritySAfterUninstallerStarted"
                        $waitForExitAfterUninstallerStarted = 0
                        } # end if ( $prog.Name -like "Norton Internet Security" )

                        # NSIS
                        if ( $prog.UninstallString -match "NSIS" ) {
                            if ( $prog.Name -match "CyberLink Media.*Suite" ) {
                                Continue # It will be removed from the other InstallShield Installer for Media Suite
                            }
                            if ( ($prog.Name -match "CyberLink") -and ($Script:progslisttoremove -match "CyberLink Media.*Suite") ) {
                                $uninstallarguments = "/S"
                                function CyberLinkNSISAfterUninstallerStarted {
                                    Write-Output "Waiting for CyberLink uninstaller to finish..." | Out-Default
                                    $waittimeoutexitcode = waitForProcessToStartOrTimeout 'Au_' 15
                                    if ( Get-Process 'Au_' -ErrorAction SilentlyContinue ) {
                                        Wait-Process 'Au_' -ErrorAction SilentlyContinue # Wait for it to end
                                    } else {
                                        Write-Output "Process Au_.exe (for the NSIS version of the CyberLink Uninstaller) wasn't running." | Out-Default
                                    }
                                }
                                $functionAfterUninstallerStarted = "CyberLinkNSISAfterUninstallerStarted"
                            } else {
                                $uninstallarguments = "/S"+" "+$uninstallarguments
                            }
                        } # end if ( $prog.UninstallString -match "NSIS" )

                        if ( $prog.Name -like "PDF Complete*" ) {
                            $uninstallarguments = "/S /x"+" "+$uninstallarguments
                        }

                        if ( $prog.Name -like "VIP Access*" ) {
                            $uninstallarguments = "/S"+" "+$uninstallarguments
                        }
                        # Special Case Non-MSI Uninstallers

                    } # end else (if non-msi uninstall)
                } # end else (no QuietUninstallString)

                Write-Output "Running:`n$($uninstallpath) $($uninstallarguments)" | Out-Default
                $ph = $null
                # Allow Powershell V2 to search the environment variable %PATH% for existing files (like cmd.exe, notepad.exe etc which are in system folders)
                # In Powershell V3+ you can just do Test-Path $path e.g. Test-Path "notepad.exe" and it will check the environment variable %PATH%
                # In V2, need to do get-command first
                try {
                    $isValidPath = Test-Path ( ( (Get-Command ($uninstallpath.TrimStart("`"`' " ).TrimEnd("`"`' " )) -ErrorAction SilentlyContinue) | Select -First 1).Definition  ) -ErrorAction SilentlyContinue
                } catch {
                    Write-Warning "Failed to detect valid commmand.`n`nProg: $($prog.Name)`nUninstallPath:`n$($uninstallpath)" | Out-Default
                }
                $ErrorActionPreference = "SilentlyContinue"
                if ( $uninstallarguments -and ($isValidPath) ) {
                    $proc = Start-Process $uninstallpath -ArgumentList $uninstallarguments -PassThru
                    $ph = $proc.Handle
                } elseif ( $isValidPath ) {
                    $proc = Start-Process $uninstallpath -PassThru
                    $ph = $proc.Handle
                } else {
                    Write-Warning "$($uninstallpath) not found!" | Out-Default
                }
                $ErrorActionPreference = "Continue"
                if ( $ph ) {
                    if ( $waitForExitAfterUninstallerStarted ) {
                        $proc.WaitForExit()
                    }

                    if ($functionAfterUninstallerStarted) {
                        Write-Output "Performing: $functionAfterUninstallerStarted" | Out-Default
                        & $functionAfterUninstallerStarted
                    }

                    if ( $proc.ExitCode -ne 0 ) {
                        $a = " $($prog.Name) might not have been removed. Check Programs list if uninstalled. If not reboot and try again or manually uninstall. Some applications give exit codes on uninstallation but still uninstall just fine."
                        if ( $proc.ExitCode -eq 3010 ) {
                            $a = ", Uninstalled. Reboot Required.`n"
                        }
                        if ( $proc.ExitCode -eq 1605 ) {
                            $a = ", This action is only valid for products that are currently installed. Program was already removed.`n"
                        }
                        Write-Warning "ExitCode: $($proc.ExitCode)$($a)" | Out-Default
                    } else {
                        Write-Output "Removed $($prog.Name)." | Out-Default
                    }
                } # end if ( $ph )

            } # end ForEach (Main Loop)


            #$ErrorActionPreference = "SilentlyContinue"
            # Clean up uninstall helpers
            Write-Output "" | Out-Default
            Write-Output "Removing uninstall helpers that were copied to $($Script:dest)\ ..." | Out-Default
            Remove-Item "$Script:dest\OffScrubc2r.vbs" -Force -Verbose -ErrorAction SilentlyContinue
            Remove-Item "$Script:dest\devcon_x$($Script:osArch).exe" -Force -Verbose -ErrorAction SilentlyContinue
            Remove-Item "$Script:dest\MCPR-10-5-374-0.exe" -Force -Verbose -ErrorAction SilentlyContinue
            Remove-Item "$Script:dest\mccleanup-10-5-128-0.exe" -Force -Verbose -ErrorAction SilentlyContinue
            Remove-Item "$Script:dest\WASP.dll" -Force -Verbose -ErrorAction SilentlyContinue
            $ErrorActionPreference = "Continue"

        } # end if ( ($Script:progslisttoremove -and $isConfirmed) -and !($Global:isDetectOnlyDryRunSilentOption) )

    ###############################################################################################################

        # Remove non Microsoft Metro/UWP/"Modern" Apps
        if ( ($Script:winVer -gt 6.1 -and $isConfirmed) -and !($Global:isDetectOnlyDryRunSilentOption) ) { # UWP apps only in Win 2012/8+

            # NOTE: Get-AppxProvisionedPackage uses PackageName and Get-AppxPackage uses PackageFullName

            if ( $Global:UWPappsAUtoRemove ) {

                Write-Output "" | Out-Default
                Write-Verbose -Verbose "Removing Matching All Users UWP Apps..."
                Write-Output "" | Out-Default

                # Unpin from start code adapted from: https://superuser.com/questions/1191143/how-to-unpin-windows-10-start-menu-ads-with-powershell

                ForEach ($removeitem in $Global:UWPappsAUtoRemove) {

                    if (Get-AppxPackage -AllUsers | Where {$_.PackageFullName -eq "$($removeitem.PackageFullName)"}) {
                        Write-Output "`nRemoving $($removeitem.Name)`nPackageFullName: $($removeitem.PackageFullName)" | Out-Default
                        try {
                            [void]$(Remove-AppxPackage -Allusers -Package "$($removeitem.PackageFullName)" -ErrorAction Ignore)
                        } catch {
                            try {
                                [void]$(Remove-AppxPackage -Package "$($removeitem.PackageFullName)" -ErrorAction Ignore)
                            } catch {
                            }
                        }
                    }

                    Start-Sleep -Seconds 4
                    Write-Output "Unpinning from Start Menu" | Out-Default
                    $unpinName = $removeitem.Name
                    try {
                        $currentLivetile = ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | Where {$_.Path -match $unpinName})
                        if ( $currentLivetile ) {
                            try {
                                $ErrorActionPreference = "Ignore"
                                $currentLivetile.Verbs() | Where { $_.Name.replace('&','') -match 'Unpin from Start' } | % { $_.DoIt() }
                                $ErrorActionPreference = SilentlyContinue
                            } catch {
                                Start-Sleep -Milliseconds 10
                            }                        }
                    } catch {
                        Write-Warning "Unable to Unpin $unpinName from Start Menu." | Out-Default
                    }
                } # end ForEach ($removeitem in $Global:UWPappsAUtoRemove)


            } # end if ( $Global:UWPappsAUtoRemove )


            if ( $Global:UWPappsProvisionedAppstoRemove ) {

                Write-Output "" | Out-Default
                Write-Verbose -Verbose "Removing Matching All Users `'Provisioned`' UWP Apps..."
                Write-Output "" | Out-Default

                ForEach ($removeProvisioneditem in $Global:UWPappsProvisionedAppstoRemove)  {

                    if (Get-AppxProvisionedPackage -Online | Where {$_.PackageName -eq "$($removeProvisioneditem.PackageName)"}) {
                        Write-Output "`nRemoving $($removeProvisioneditem.DisplayName)`nPackageName: $($removeProvisioneditem.PackageName)" | Out-Default
                        try {
                            [void]$(Remove-AppxProvisionedPackage -PackageName "$($removeProvisioneditem.PackageName)" -Online -Allusers -ErrorAction Ignore)
                        } catch {
                            try {
                                [void]$(Remove-AppxProvisionedPackage -PackageName "$($removeProvisioneditem.PackageName)" -Online -ErrorAction Ignore)
                            } catch {
                            }
                        }
                    }
                    Start-Sleep -Seconds 4
                    Write-Output "Unpinning from Start Menu" | Out-Default
                    $unpinName = $removeProvisioneditem.DisplayName
                    try {
                        $currentLivetile = ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | Where { $_.Path -match $unpinName })
                        if ( $currentLivetile ) {
                            try {
                                $ErrorActionPreference = "Ignore"
                                $currentLivetile.Verbs() | Where { $_.Name.replace('&','') -match 'Unpin from Start' } | % { $_.DoIt() }
                                $ErrorActionPreference = SilentlyContinue
                            } catch {
                                Start-Sleep -Milliseconds 10
                            }
                        }
                    } catch {
                        Write-Warning "Unable to Unpin $unpinName from Start Menu." | Out-Default
                    }
                } # end ForEach ($removeProvisioneditem in $Global:UWPappsProvisionedAppstoRemove)

            } # end if ( $Global:UWPappsProvisionedAppstoRemove )

            doWindows10Options # if Win10+ and Option(s) enabled

        } # end if ( $Script:winVer -gt 6.1 -and $isConfirmed )

    ###############################################################################################################

        if ( $isConfirmed -and !($Global:isDetectOnlyDryRunSilentOption) ) {

            if ( !($Global:isSilent) ) { # literal silence again
            $soundloc = "c:\Windows\Media\tada.wav"
                if (Test-Path $soundloc) {
                    $sound = New-Object System.Media.SoundPlayer;
                    $sound.SoundLocation = $soundloc;
                    $sound.Play();
                }
            }

            Write-Output "" | Out-Default
            Write-Verbose -Verbose "Finished removing bloatware. If any were removed please reboot before installing any new software.`n"
            Write-Output "" | Out-Default
            Write-Verbose -Verbose "Please review the above output or logfile and reboot now to complete the bloatware removal before installing any new software."

            if ( $Global:rebootAfterRemoval ) {
                Write-Output "" | Out-Default
                Write-Verbose -Verbose "You have chosen to reboot after removal."
                Write-Output "" | Out-Default
                sleepProgress (@{"Seconds" = 15})
                Write-Output "Rebooting now...`n" | Out-Default
            }
            stopTranscript
            if ( $Global:rebootAfterRemoval ) {
                Restart-Computer -Force
            }
        } # end if ( $isConfirmed )

    } else { # end if ( $progslistSelected -ne $null )

    #    if ( !($Script:progslisttoremove) -and !($Global:UWPappsAUtoRemove) -and !($Global:UWPappsProvisionedAppstoRemove) ) {
            Write-Output "" | Out-Default
            $nonefoundmessage = "No Bloatware was selected"
            if ( $Global:isSilent ) {
                $nonefoundmessage += " or matched"
            }
            $nonefoundmessage += "."
            Write-Output $nonefoundmessage | Out-Default
    #   } # end of if ( ($Script:progslisttoremove) -or ($Global:UWPappsAUtoRemove) -or ($Global:UWPappsProvisionedAppstoRemove) )

        $isConfirmed = $false
        if ( $Script:winVer -ge 10 ) {
            $isConfirmed = systemRestorePointIfRequired
            if ( !($isConfirmed) ) {
                Write-Output "You have chosen to not proceed with removal or change settings. No changes will be made."
            } else {
                doWindows10Options # even if no programs to remove or selected, apply selected Win10 tweaks
            }
        }



        Write-Output "" | Out-Default
        stopTranscript
        Set-Location $savedPathLocation # Restore working directory path
        Return
    }

} # end if ( ($button -ne "Cancel") -or ($Global:isSilent) ) # Cancel was clicked or window closed or not confirmed to continue or dry run / detect only

if (  !($isConfirmed) ) {
    Write-Output "" | Out-Default
    Stop-Transcript
}

if ( $button -eq "Cancel" -and !($Global:isSilent) ) {
    Write-Output "Removing Log file as removal was canceled or window closed." | Out-Default
    Remove-Item -Force $Script:logfile
    [Environment]::Exit(4) # User Canceled
}

Set-Location $savedPathLocation # Restore working directory path


Return


} # end PROCESS

END {




}


################################################################################################
BEGIN {


    function refreshProgramsList {

        Write-Output "" | Out-Default
        $Global:statusupdate = "Generating and filtering programs list, please wait..."
        Write-Verbose -Verbose "$($Global:statusupdate)`n"

        if ( !($Global:isSilent) ) {
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
        }

        # Get an initial list of installed programs and save it
        try {
            $Script:proglistwithdupes = Get-CIMInstance -class Win32_Product
        } catch {
            $Script:proglistwithdupes = Get-WMIobject -class Win32_Product
        }

        $a = @()
        $a = @(Get-ChildItem -Recurse HKLM:Software\Microsoft\Windows\CurrentVersion\Uninstall | gp | Where { $_.DisplayName -ne $null -and $_.UninstallString -ne $null } )
        $b = @()
        if (Test-Path HKLM:Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall) {
            $b = @(Get-ChildItem -Recurse HKLM:Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | gp | where { $_.Displayname -ne $null -and $_.UninstallString -ne $null } )
        }
        $c = @()
        if (Test-Path HKCU:Software\Microsoft\Windows\CurrentVersion\Uninstall) {
            $c = @(Get-ChildItem -Recurse HKCU:Software\Microsoft\Windows\CurrentVersion\Uninstall | gp | Where { $_.DisplayName -ne $null -and $_.UninstallString -ne $null  } )
        }

        $proglistwithdupes = @((@( $proglistwithdupes | Select-Object Name,IdentifyingNumber) + @(($a+$b+$c) | select @{Name="Name";Expression={$_."DisplayName"}},UninstallString,@{Name="Version";Expression={$_."DisplayVersion"}},Publisher,QuietUninstallString)) | sort UninstallString)


        ###############################################################################################################

        if ( $Script:winVer -gt 6.1) { # UWP apps only in Win 2012/8+
            #UWP Win8/Win10+ Apps

            Write-Output "" | Out-Default
            Write-Verbose -Verbose "All Users UWP Win8/Win10+ Apps:"
            Write-Output "" | Out-Default
            $Global:statusupdate = "Processing Windows8/10+ UWP Apps..."

            if ( !($Global:isSilent) ) {
                $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
                $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
            }

            try {
                $ErrorActionPreference = "SilentlyContinue"
                $Global:UWPappsAU = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue
                $Global:UWPappsAU | % { Write-Output "PackageFullName: $($_.PackageFullName)" } | Out-Default
                Write-Output "" | Out-Default
            } catch {
                Write-Warning "Service AppX Deployment Service (AppXSVC) must be running (and not disabled) for Get-AppxPackage to work." | Out-Default
            }

            Write-Output "" | Out-Default
            Write-Verbose -Verbose "All Users provisioned UWP Win8/Win10+ Apps:"
            Write-Output "" | Out-Default
            $Global:UWPappsProvisionedApps = Get-AppxProvisionedPackage -Online
            $Global:UWPappsProvisionedApps | % { Write-Output "PackageName: $($_.PackageName)" } | Out-Default
            Write-Output "" | Out-Default
            #Change Property: DisplayName to Name to be consistent
            $Global:UWPappsProvisionedApps = $Global:UWPappsProvisionedApps | Where-Object { $_.DisplayName } | Add-Member -MemberType AliasProperty -Name Name -Value DisplayName -PassThru | Select-Object Name, Architecture, Build, DisplayName, InstallLocation, LogLevel, LogPath, MajorVersion, MinorVersion, Online, PackageName, Path, PublisherId, Regions, ResourceId, RestartNeeded, Revision, ScratchDirectory, SysDrivePath, Version, WinPath


        } # end if ( $Script:winVer -gt 6.1)

        ###############################################################################################################


        <#
        Write-Output "" | Out-Default
        Write-Output "" | Out-Default
        Write-Output "Full List with Duplicates (excludes UWP Win8/10 metro/modern apps)...`n" | Out-Default
        $proglistwithdupes | Sort UninstallString | Out-Default | Format-List
        #>

        Write-Output "" | Out-Default
        $Global:statusupdate = "Deduplicating programs list..."
        Write-Verbose -Verbose "$($Global:statusupdate)`n"

        if ( !($Global:isSilent) ) {
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
        }

        <#
        Deduplication logic
        sort uninstallstring (exe uninstallers 1st, uninstallstring w/msiexec 2nd, regular msi 3rd)
        add those first (if identifyingnumber == null and not guid just add exe uninstaller to array)
        if identifyingnumber not null (i.e. it is a regular msi only entry) then add to array if not already in array
         and also add it if it is in the array but the matching uninstall string doesn't have 'msiexec' in it (such as installshield uninstallers)
         so if the one it match already had msiexec in the uninstall string we skip it, no point adding it again
        if uninstall string is msiexec uninstaller (isguid) and doesn't match a guid already in the array, add to array,
        repeat next item in array with dupes until completed, new list '$Global:proglist' contains only unique items
        (a mixture of registry uninstall entries and wmi/win32 entries)
        ------------------------------------------
        #>
        #Remove duplicates program listings
        $Global:guidmatchstring = "(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}"
        $Global:proglist = @()

        ForEach ($item in $proglistwithdupes) {
            $isguid = $item.UninstallString -match $Global:guidmatchstring
            #add non-guid/non-msi program to list
            if ($item.IdentifyingNumber -eq $null -and !($isguid)) { #non-msi uninstaller
                $Global:proglist += $item;
            }
            elseif ( $item.IdentifyingNumber -ne $null ) { # if from wmi, add any wmi ones that aren't on the list already or that match against an item on the list that doesn't have msiexec in the uninstallstring (like HPSA)
                if ( !($Global:proglist -match $item.IdentifyingNumber) `
                     -or ($Global:proglist | Where { $_.UninstallString -notmatch "msiexec" -and $_.UninstallString -ne $null } | Where { $_ -match $item.IdentifyingNumber }) ) {
                    $Global:proglist += $item;
                }
            }
            elseif ( $item.UninstallString -ne $null -and $isguid ) { #if guid is an uninstallstring
                  if ( !($Global:proglist -match $matches[0]) ) {
                    $Global:proglist += $item;
                  }
            }
        }

        #Put properties in order for later GUI display in columns, remove entries without Name property
        $Global:proglist = ($Global:proglist | Where { $_.Name } | Select-Object Name,Version,Publisher,UninstallString,QuietUninstallString,IdentifyingNumber)

        Write-Output "" | Out-Default
        Write-Verbose -Verbose "List after duplicates removed..."
        $Global:proglist| Sort-Object Name | Out-Default | Format-List

        if (!($Global:usingSavedSelectionFileSilentOption)) {
            Write-Output "" | Out-Default
            $Global:statusupdate = "Enumerating suggested bloatware to remove..."
            Write-Verbose -Verbose "$($Global:statusupdate)`n"
        } else {
            Write-Output "" | Out-Default
            $Global:statusupdate = "Checking installed list of programs to be used with selection list file..."
            Write-Verbose -Verbose "$($Global:statusupdate)`n"
        }

        if ( !($Global:isSilent) ) {
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
        }

        $bloatwarelike = (
        # include, in regular expression format
        # must be regex escaped here (and powershell escaped if needed)
        "ActivClient.*",
        "Adobe\ AIR",
        "Ai\ Meeting\ Manager\ Service",
        "^ASUS\ .*",
        "Bing\ Bar",
        "Corel\ .*",
        "Create\ Recovery\ Media",
        "^Cyberlink.*",
        "^Power2Go",
        "^PowerDVD",
        "^Dell.*",
        "Dropbox.*",
        "DVD\ Architect\ Studio.*",
        "Evernote.*",
        "Google\ Toolbar.*",
        "^HP\ Sure\ Run.*",
        "^HP\ .*",
        "HPInc.EnergyStar",
        "HPPrinterControl",
        "HPPrivacySettings",
        "HPSupportAssistant",
        "HPSystemEventUtility",
        "Discover\ HP",
        "HP\ Touchpoint",
        "^Hewlett\-Packard.*",
        "for\ HP\ ProtectTools",
        "Intel\ AppUp\(SM\)\ center",
        "Energy\ Star",
        "Foxit\ .*",
        "Kaspersky.*",
        "Lenovo",
        "Message\ Center\ Plus",
        "Microsoft\ Security\ Client",
        "Microsoft\ Security\ Essentials",
        "Movie\ Maker",
        "Multifactor Authentication Client.*",
        "NewBlue\ Video.*",
        "Nitro\ .*",
        "Norton\ Security",
        "PDF\ Complete.*",
        "Photo\ Gallery",
        "PlayMemories",
        "proDAD\ Adorage.*",
        "Reader\ for\ PC",
        "Softex.*",
        "ThinkVantage\ .*",
        "VAIO\ .*",
        "Windows\ Live\ .*",
        "WinZip.*",
        # Windows 8/10/11+ UWP apps ######################################################################################
        "3DBuilder",
        "ACGMediaPlayer",
        "ActiproSoftware",
        "AD2F1837", # All HP UWP apps (except those in the bloatwarenotmatch list below)
        "AdobeSystemsIncorporated\.AdobePhotoshopExpress",
        "ASUSGIFTBOX",
        "ASUSPCAssistant",
        "AutodeskSketchBook",
        "BingFinance",
        "BingNews",
        "BingSports",
        "BingWeather",
        "BlueEdge\.OneCalendar",
        "BubbleWitch",
        "CaesarsSlotsFreeCasino",
        "CandyCrush",
        "CommsPhone",
        "ConnectivityStore",
        "CookingFever",
        "CyberLinkMediaSuiteEssentials",
        "DellCustomerConnect",
        "DellHelpSupport",
        "DellInc\.DellSupportAssistforPCs",
        "DellProductRegistration",
        "DiscoverHPTouchpointManager",
        "DisneyMagicKingdoms",
        "DolbyAccess",
        "DragonManiaLegends",
        "DrawboardPDF",
        "Duolingo",
        "EclipseManager",
        "Facebook",
        "FalloutShelter",
        "FarmHeroesSaga",
        "FarmVille2CountryEscape",
        "Flipboard",
        "Getstarted",
        "HiddenCityMysteryofShadows",
        "iHeartRadio",
        "KeeperSecurity",
        "LenovoCompanion",
        "LenovoCorporation\.LenovoID",
        "LenovoCorporation\.LenovoSettings",
        "LenovoUtility",
        "LinkedInforWindows",
        "MarchofEmpires",
        "McAfeeSecurity",
        "McAfee\.Security",
        "MediaSuiteEssentials",
        "Messaging",
        "Microsoft3DViewer",
        "Microsoft\.Asphalt",
        "Microsoft\.Getstarted",
        "Microsoft\.Office\.Desktop", # Windows Store version of Office
        "MicrosoftOfficeHub",
        "MinecraftUWP",
        "MircastView",
        "MyASUS",
        "Netflix",
        "Norton",
        "OneConnect",
        "PandoraMediaInc",
        "ParadiseBay",
        "PhototasticCollage",
        "PicsArt",
        "Plex",
        "Power2Go",
        "PowerDirector",
        "PowerMediaPlayer",
        "RoyalRevolt",
        "Shazam",
        "SkypeApp",
        "SlingTV",
        "SpotifyAB",
        "Sway",
        "TheNewYorkTimes",
        "TuneInRadio",
        "Twitter",
        "Viber",
        "Wallet",
        "windowsphone",
        "WinZipUniversal",
        "Wunderlist",
        "XboxApp",
        "XboxGameOverlay",
        "XboxOneSmartGlass",
        "XboxSpeechToTextOverlay",
        "XINGAG\.XING",
        "ZuneMusic",
        "ZuneVideo"
        )
        $bloatwarenotmatch = (
        # skip these, do not remove at all
        # Matches any part of the string. Will be regex escaped later.
        "Dell Command | Update",
        "Dell ControlVault",
        "Dell Update",
        "Dell Digital Delivery",
        "Dell MD Storage",
        "Dell OpenManage Server Administrator",
        "Dell Unified Wireless Suite",
        "HP Battery Recall Utility",
        "HP Hotkey Support", # for brightness/media controls on some HP laptops and 2-in-1s
        "HP Pen Control", # stylus driver
        "HP USB Audio", # docking station audio driver
        "Lenovo Patch Utility",
        "Lenovo System Update",
        "Lenovo USB",
        "NetExtender",
        "Touchpad",
        "ThinkVantage System Update",
        "Driver",
        "SonicWALL",
        "Recovery Manager",
        "Hardware Diagnostic",
        "VAIO Care",
        "VAIO Control Center",
        "VAIO Movie Creator",
        "VAIO Update",
        # Windows 8/10+ UWP apps ######################################################################################
        "Appconnector",
        "CommsPhone",
        "HoloCamera",
        "HolographicFirstRun",
        "HoloItemPlayerApp",
        "HoloShell",
        "HPPCHardwareDiagnostics",
        "Messaging",
        "MicrosoftSolitaireCollection",
        "MicrosoftStickyNotes",
        "MSPaint",
        "OneConnect",
        "OneNote",
        "People",
        "Windows.Photos", #exclude MS app, but allow matching of AdobePHOTOShopExpress
        "Soundrecorder",
        "WindowsAlarms",
        "WindowsCalculator",
        "WindowsCamera",
        "windowscommunicationsapps",
        "WindowsFeedbackApp",
        "WindowsFeedbackHub",
        "WindowsMaps",
        "WindowsScan",
        "WindowsStore"
        )
        $specialcasestoremove = (
        # special cases below will be removed later after the matching list above are removed first
        # VERY IMPORTANT, the order in which these are listed is important as they'll be removed in that order. Some programs need to be removed before others. Matches any part of the string. Will be regex escaped later.
        "CyberLink Media Suite",
        "Dell Optimizer",
        "Dell Precision Optimizer", # Need to create iss file for this possibly or try -silent instead of /S switch
        "Dell SupportAssist",
        "Dell Data Vault",
        "HP Setup",
        "HP Support Assistant",
        "HP Support Solutions Framework",
        "HP Theft Recovery",
        "HPWorkWise64",
        "HP WorkWise64",
        "HP WorkWise",
        "MyDell",
        "Theft Recovery for HP ProtectTools",
        "ProtectTools Security Manager",
        "HP Client Security Manager",
        "McAfee", # ALL McAfee Consumer Software (runs MCPR.exe)
        "Norton Internet Security",
        "VIP Access",
        "Microsoft Office", # Trial/OEM Versions of MS Office
        "Microsoft 365",
        "Office.Click-to-Run"
        )

        # skip special cases, then add them back at the end later

        if ( $Global:isSilent ) { # set the command line modifications if running silently using switches

            if ( $Global:isIgnoreDefaultSuggestionListSilentOption ) { # no default suggestions if -nd or -ignoredefaults switch
                [string[]]$bloatwarelike = ""
                [string[]]$bloatwarenotmatch = ""
                [string[]]$specialcasestoremove = ""
            }

            if ( $Global:bloatwareIncludeFirstSilentOption ) {
                $Global:bloatwareIncludeFirstSilentOption = (([string[]]$Global:bloatwareIncludeFirstSilentOption | % { $_ }) -split ',' -join '|').TrimStart('|').TrimEnd('|')
                $Global:bloatwareIncludeFirstSilentOption = $Global:bloatwareIncludeFirstSilentOption -Replace '"',''
                #$Global:bloatwareIncludeFirstSilentOption = (([string[]]$Global:bloatwareIncludeFirstSilentOption | % { if ( $_ ) { "$([regex]::Escape($_))" } }) -split ',' -join '|').TrimStart('|').TrimEnd('|')
            }
            if ( $Global:bloatwareExcludeSilentOption ) {
                $Global:bloatwareExcludeSilentOption = (([string[]]$Global:bloatwareExcludeSilentOption | % { if ( $_ ) { "$([regex]::Escape($_))" } }) -split ',' -join '|' ).TrimStart('|').TrimEnd('|')
            }
            if ( $Global:bloatwareIncludeLastSilentOption ) { # special cases last
                $Global:bloatwareIncludeLastSilentOption = (([string[]]$Global:bloatwareIncludeLastSilentOption | % { if ( $_ ) { "$([regex]::Escape($_))" } }) -split ',' -join '|').TrimStart('|').TrimEnd('|')
            }
        }

        $Global:bloatwarelikesinglestring = (($bloatwarelike | % { if ( !([string]::IsNullorEmpty($_)) ) { $_ } })  -join '|').TrimStart('|').TrimEnd('|') #$bloatware like is not escaped here because it is already regex escaped
        #Write-Host '$Global:bloatwarelikesinglestring'
        #Write-Host $Global:bloatwarelikesinglestring

        $Global:specialcasestoremovesinglestring = ((($specialcasestoremove | % { if ( !([string]::IsNullorEmpty($_)) ) { ".*$([regex]::Escape($_)).*$" } }) -join '|') + '|' + $Global:bloatwareIncludeLastSilentOption).TrimStart('|').TrimEnd('|').Trim()
        #Write-Host "`$Global:specialcasestoremovesinglestring"
        #Write-Host $Global:specialcasestoremovesinglestring

        $Global:bloatwarenotmatchsinglestring = ((($bloatwarenotmatch | % { if ( !([string]::IsNullorEmpty($_)) ) { ".*$([regex]::Escape($_)).*$" } }) -join '|') + '|' + $Global:bloatwareExcludeSilentOption).TrimStart('|').TrimEnd('|').Trim() # turn into single string for regex excluding
        #Write-Host '$Global:bloatwarenotmatchsinglestring'
        #Write-Host $Global:bloatwarenotmatchsinglestring

        ############## Core regular expression matching magic regular programs ##############

        function matchAgainstProglist {
            Param(
                [Parameter(Position=0,Mandatory=$true)]
                    [array]$proglisttomatchagainst,
                [Parameter(Position=1,Mandatory=$true)]
                    [string[]]$matchpatterns,
                [Parameter(Position=2,Mandatory=$false)]
                    [array]$dontmatchagainstthislist
            )

            $proglisttoreturn = @()
            ForEach ($matchpattern in ($($matchpatterns -replace "\\\|","%%%%" -split "\|" -replace "%%%%","\|"))) {
                $proglisttoreturn += @($proglisttomatchagainst | Where {

                    if ($_.Name -match $matchpattern) {

                        #Write-Host "`$_.Name: $(($_).Name)"
                        #Write-Host "Matchpattern: $matchpattern"

                        if ( ($_.Name -match $Global:bloatwarenotmatchsinglestring) -and (!([string]::IsNullorEmpty($Global:bloatwarenotmatchsinglestring))) ) {
                            $false
                        } elseif ( ($dontmatchagainstthislist -match $_.Name) -and (!([string]::IsNullorEmpty($dontmatchagainstthislist))) ) {
                            $false
                        } else {
                            $true
                        }

                    }
                })
            }

            Return $proglisttoreturn
        } # end function matchAgainstProglist

        if (!([string]::IsNullOrEmpty($Global:bloatwareIncludeFirstSilentOption))) {
            $bloatwareincludefirstdeduped = matchAgainstProglist -proglisttomatchagainst $Global:proglist -matchpatterns $Global:bloatwareIncludeFirstSilentOption
        }
        #Write-Host "Bloatware include first deduped:"
        #Write-Host $bloatwareincludefirstdeduped

        if (!([string]::IsNullOrEmpty($Global:specialcasestoremovesinglestring))) {
            $specialcasesdeduped = matchAgainstProglist -proglisttomatchagainst $Global:proglist -matchpatterns $Global:specialcasestoremovesinglestring -dontmatchagainstthislist @($bloatwareincludefirstdeduped)
        }
        #Write-Host "Special Cases deduped:"
        #Write-Host $specialcasesdeduped

        if (!([string]::IsNullOrEmpty($Global:bloatwarelikesinglestring))) {
            $bloatwarelikededuped = matchAgainstProglist -proglisttomatchagainst $Global:proglist -matchpatterns $Global:bloatwarelikesinglestring -dontmatchagainstthislist $(@($bloatwareincludefirstdeduped) + @($specialcasesdeduped))
        }
        #Write-Host 'Bloatware like deduped:'
        #Write-Host $bloatwarelikededuped

        $Script:progslisttoremove = @(@($bloatwareincludefirstdeduped) + @($bloatwarelikededuped) + @($specialcasesdeduped))

        $ignoreDefaultSuggestionListMsg = "No default suggestions given because running with -ignoredefaultsuggestions switch."

        if (!($Global:usingSavedSelectionFileSilentOption)) {
            Write-Output "" | Out-Default
            Write-Output "Bloatware suggested for removal (non UWP Win8/Win10+ Apps):`n" | Out-Default
            if ( $Global:isSilent -and $Global:isIgnoreDefaultSuggestionListSilentOption ) {
                Write-Output $ignoreDefaultSuggestionListMsg | Out-Default
            }
            $Script:progslisttoremove | Out-Default | Format-List
        }

        ###############################################################################################################

        if ( $Script:winVer -gt 6.1) { # UWP apps only in Win 2012/8+

            ############## Core regular expression matching magic UWP / Windows Store Programs ##############

            if (!([string]::IsNullOrEmpty($Global:bloatwareIncludeFirstSilentOption))) {
                $UWPbloatwareincludefirstdeduped = matchAgainstProglist -proglisttomatchagainst $Global:UWPappsAU -matchpatterns $Global:bloatwareIncludeFirstSilentOption
            }
            #Write-Host "Bloatware include first deduped:"
            #Write-Host $UWPbloatwareincludefirstdeduped

            if (!([string]::IsNullOrEmpty($Global:specialcasestoremovesinglestring))) {
                $UWPspecialcasesdeduped = matchAgainstProglist -proglisttomatchagainst $Global:UWPappsAU -matchpatterns $Global:specialcasestoremovesinglestring -dontmatchagainstthislist @($UWPbloatwareincludefirstdeduped)
            }
            #Write-Host "Special Cases deduped:"
            #Write-Host $UWPspecialcasesdeduped

            if (!([string]::IsNullOrEmpty($Global:bloatwarelikesinglestring))) {
                $UWPbloatwarelikededuped = matchAgainstProglist -proglisttomatchagainst $Global:UWPappsAU -matchpatterns $Global:bloatwarelikesinglestring -dontmatchagainstthislist $(@($UWPbloatwareincludefirstdeduped) + @($UWPspecialcasesdeduped))
            }
            #Write-Host 'Bloatware like deduped:'
            #Write-Host $UWPbloatwarelikededuped

            $Global:UWPappsAUtoRemove = @(@($UWPbloatwareincludefirstdeduped) + @($UWPbloatwarelikededuped) + @($UWPspecialcasesdeduped))

            if (!([string]::IsNullOrEmpty($Global:bloatwareIncludeFirstSilentOption))) {
                $UWPProvisionedbloatwareincludefirstdeduped = matchAgainstProglist -proglisttomatchagainst $Global:UWPappsProvisionedApps -matchpatterns $Global:bloatwareIncludeFirstSilentOption
            }
            #Write-Host "UWPProvisioned Bloatware include first deduped:"
            #Write-Host $UWPProvisionedbloatwareincludefirstdeduped

            if (!([string]::IsNullOrEmpty($Global:specialcasestoremovesinglestring))) {
                $UWPProvisionedspecialcasesdeduped = matchAgainstProglist -proglisttomatchagainst $Global:UWPappsProvisionedApps -matchpatterns $Global:specialcasestoremovesinglestring -dontmatchagainstthislist @($UWPProvisionedbloatwareincludefirstdeduped)
            }
            #Write-Host "UWPProvisioned Special Cases deduped:"
            #Write-Host $UWPProvisionedspecialcasesdeduped

            if (!([string]::IsNullOrEmpty($Global:bloatwarelikesinglestring))) {
                $UWPProvisionedbloatwarelikededuped = matchAgainstProglist -proglisttomatchagainst $Global:UWPappsProvisionedApps -matchpatterns $Global:bloatwarelikesinglestring -dontmatchagainstthislist $(@($UWPProvisionedbloatwareincludefirstdeduped) + @($UWPProvisionedspecialcasesdeduped))
            }
            #Write-Host "UWPProvisioned Bloatware like deduped:"
            #Write-Host $UWPProvisionedbloatwarelikededuped

            $Global:UWPappsProvisionedAppstoRemove = @(@($UWPProvisionedbloatwareincludefirstdeduped) + @($UWPProvisionedbloatwarelikededuped) + @($UWPProvisionedspecialcasesdeduped))

            if (!($Global:usingSavedSelectionFileSilentOption)) {
                Write-Output "" | Out-Default
                Write-Verbose -Verbose "All Users UWP Win8/Win10+ Apps Suggested for Removal:"
                if ( $Global:isSilent -and $Global:isIgnoreDefaultSuggestionListSilentOption ) {
                    Write-Output $ignoreDefaultSuggestionListMsg | Out-Default
                }
                Write-Output "" | Out-Default
                $Global:UWPappsAUtoRemove | % { $_.PackageFullName | Out-Default }
                Write-Output "" | Out-Default
                Write-Verbose -Verbose "All Users Provisioned UWP Win8/Win10+ Apps Suggested for Removal:"
                if ( $Global:isSilent -and $Global:isIgnoreDefaultSuggestionListSilentOption ) {
                    Write-Output $ignoreDefaultSuggestionListMsg | Out-Default
                }
                Write-Output "" | Out-Default
                $Global:UWPappsProvisionedAppstoRemove | % { $_.PackageName | Out-Default }
            }
        }

        ###############################################################################################################

        # At this point:
        #$Global:proglist exists
        #$Global:UWPappsAU exists if ( $Script:winVer -gt 6.1 )
        #$Global:UWPappsProvisionedApps exists if ( $Script:winVer -gt 6.1 )


        # following 3 variables are modified by the GUI selection list or command line options
        #$Script:progslisttoremove exists
        #$Global:UWPappsAUtoRemove exists if ( $Script:winVer -gt 6.1 )
        #$Global:UWPappsProvisionedAppstoRemove exists if ( $Script:winVer -gt 6.1 )

        $Script:proglistviewColumnsArray = @('DisplayName','Name','Version','Publisher','UninstallString','QuietUninstallString','IdentifyingNumber','PackageFullName','PackageName')
        $Global:progslisttodisplay = $Global:proglist | Select-Object -Property $proglistviewColumnsArray -ExcludeProperty DisplayName | Sort-Object Name
        #$Global:orignalprogslisttodisplay = $Global:progslisttodisplay
        # Add in the UWP Win8/10+ apps to the list

        if ( $Script:winVer -gt 6.1 ) {
            $Global:UWPappsAUlisttodisplay = $Global:UWPappsAU | Select-Object -Property $proglistviewColumnsArray -ExcludeProperty DisplayName | Sort Name
            $Global:UWPappsProvisionedAppslisttodisplay = $Global:UWPappsProvisionedApps | Select-Object -Property $proglistviewColumnsArray -ExcludeProperty Name | Select-Object @{Name="Name";Expression={$_.DisplayName}},* -ExcludeProperty DisplayName | Sort Name
            $Global:progslisttodisplay = @($Global:progslisttodisplay) + @($Global:UWPappsAUlisttodisplay)
            $Global:progslisttodisplay = @($Global:progslisttodisplay) + @($Global:UWPappsProvisionedAppslisttodisplay)
        }

        [int]$Global:numofprogs = @('0',($Global:progslisttodisplay | Measure-Object).Count)[($Global:progslisttodisplay | Measure-Object).Count -gt 0]

        if ( !($Global:isSilent) ) {

            # the Where { $_ }  in the following statement is useful if the Get-AppxPackage or Get-AppxProvisionedPackage services were off and the result returned to them was a bunch of boolean false statements
            $Global:progslisttoshowchecked = @($Script:progslisttoremove | Where { $_ } | Select-Object $proglistviewColumnsArray -ExcludeProperty DisplayName) + @($Global:UWPappsAUtoRemove | Where { $_ } | Select-Object $proglistviewColumnsArray -ExcludeProperty DisplayName) + @($Global:UWPappsProvisionedAppstoRemove | Where { $_ } | Select-Object -Property $proglistviewColumnsArray -ExcludeProperty Name | Select-Object @{Name="Name";Expression={$_.DisplayName}},* -ExcludeProperty DisplayName) | Sort Name

            generateProgListView $Global:progslisttodisplay
            if ( $Script:showSuggestedtoRemove ) {
                generateProgListViewChecked $Global:progslisttoshowchecked
            }

            Write-Output "" | Out-Default
            Write-Verbose -Verbose "Please make your selection in the GUI list."

            $Global:statusupdate = "Programs list generated and suggested bloatware preselected."
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
            $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate

            Write-Output "" | Out-Default

        } else { # if running silently

            Write-Output "" | Out-Default
            Write-Output "Total number of programs: $($Global:numofprogs)" | Out-Default

        } # end if ( !($Global:isSilent) )

    } # End function refreshProgramsList

#############

    function stopProcesses( [array]$procnamelist ) {
        # Stop processes used by bloatware before removal
        # takes an array of processes and attempts to stop them

    Write-Host "" | Out-Default
        Write-Verbose -Verbose "Stopping processes used by bloatware..."

        ForEach ($procname in $procnamelist) {
            if (Get-Process -Name $procname -ErrorAction SilentlyContinue) {
                Write-Host "Stopping process: $($procname)" | Out-Default
                Stop-Process -Name $procname -Force -ErrorAction SilentlyContinue
                Write-Verbose -Verbose "Process $($procname) stopped."
            }
            else {
                if ( (Get-Process $procname -ErrorAction SilentlyContinue) -ne $null ) {
                    Write-Warning "Process $($procname) could not be stopped." | Out-Default
                }
            }
        }
    } # end function stopProcesses

#############

    function parseUninstallString( [string]$proguninstallstring, [string]$uninstallstringmatchstring ) {
    # parseUninstallString takes the uninstallstring, and the uninstallstringmatchstring then
    # then returns the path and the arguments seperately in a form that Start-Process can use
        # Reset $uninstallarguments and $matches each loop iteration
        $uninstallpath = $proguninstallstring

        $uninstallarguments = $null
        $matches = $null

        $uninstallpath = $uninstallpath -replace "^cmd \/c", ""
        $uninstallpath = $uninstallpath -replace "^RunDll32.*LaunchSetup\ ", ""
        $uninstallpath = $uninstallpath.TrimStart(" ").TrimEnd(" ")


        $matched = [RegEx]::Match($uninstallpath, $uninstallstringmatchstring)

        if ( $matched.Success ) { # only matches if arguments exist
            $uninstallpath = $matched.Groups[1].Value
            $uninstallarguments = $matched.Groups[2].Value
        }

        #remove spaces, single and double quotes from the process path and aurgument list at the begining and end of each
        $uninstallpath = $uninstallpath.TrimStart("`"`' " ).TrimEnd("`"`' " )
        if ( $uninstallarguments -ne $null ) {
            $uninstallarguments = $uninstallarguments.TrimStart("`"`' " ).TrimEnd("`"`' " )
        }

        $uninstallpath = "`""+$uninstallpath+"`""
        $returned = @($uninstallpath,$uninstallarguments)
        return $returned

    } # end function parseUninstallString( $proguninstallstring, $uninstallstringmatchstring)

#############

    function systemRestorePointIfRequired( ) {
        if ( $Global:requireSystemRestorePointBeforeRemoval ) {
            Write-Host "`nProceed with System Restore Point Creation and Continue?`n" | Out-Default

            [bool]$isConfirmed = doOptionsRequireConfirmation

            if ( $isConfirmed ) {
                Write-Host "" | Out-Default
                Write-Host "Creating a System Restore Point called BeforeBloatwareRemoval...`n" | Out-Default
                if ( $Script:winVer -gt 6.2) { # Win8+
                    $savedKey = $null
                    $key = Get-Item "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\SystemRestore"
                    if ( $key.SystemRestorePointCreationFrequency -ne $null ) {
                        $savedKey = $key
                    }
                    # Ensure restore point is created ignoring default 1 restore point per day setting in Windows 8+
                    Set-ItemProperty -Path $key.PsPath -Name 'SystemRestorePointCreationFrequency' -value '0' -Type DWord
                    $key = Get-Item "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\SystemRestore"
                }
                try {
                    Checkpoint-Computer -Description "BeforeBloatwareRemoval" -RestorePointType "APPLICATION_UNINSTALL" -ErrorAction Stop
                } catch {
                    if ( $Script:winVer -gt 6.2 -and ( $savedKey -ne $null ) ) { # Win8+
                        Set-ItemProperty $key.PsPath -name 'SystemRestorePointCreationFrequency' -value $savedKey.SystemRestorePointCreationFrequency -Type DWord
                    }

                    Write-Progress -Activity "Creating a system restore point ..." -Status "Not Completed" -Completed # Clears the progress bar at the top

                    Write-Warning "A System Restore Point could not be created. Ensure the service is running and System Protection is enabled with enough disk space available.`n`n" | Out-Default
                    Write-Host "Do you want to continue (without the Restore Point)?`n" | Out-Default

                    [bool]$isConfirmed = doOptionsRequireConfirmation
                } # end catch

                if ( $Script:winVer -gt 6.2 -and ( $savedKey -ne $null )) { # Win8+
                    Set-ItemProperty $key.PsPath -name 'SystemRestorePointCreationFrequency' -value $savedKey.SystemRestorePointCreationFrequency -Type DWord
                }
                # Turn off additional confirmation prompt
                $Global:requireConfirmationBeforeRemoval = $false
                return $isConfirmed
            } # end if ( $isConfirmed )
            return $false # chose not to confirm at first prompt
        } # end if ( $Global:requireSystemRestorePointBeforeRemoval )
        return $true
    } # end function systemRestorePointIfRequired( )

#############

    function sleepProgress( [hashtable]$SleepHash ) {
    # https://rcmtech.wordpress.com/2013/03/13/powershell-sleep-with-progress-bar/
        [int]$SleepSeconds = 0
        ForEach($Key in $SleepHash.Keys){
            switch($Key){
                "Seconds" {
                    $SleepSeconds = $SleepSeconds + $SleepHash.Get_Item($Key)
                }
                "Minutes" {
                    $SleepSeconds = $SleepSeconds + ($SleepHash.Get_Item($Key) * 60)
                }
                "Hours" {
                    $SleepSeconds = $SleepSeconds + ($SleepHash.Get_Item($Key) * 60 * 60)
                }
            }
        }
        for($Count=0;$Count -lt $SleepSeconds;$Count++) {
            $SleepSecondsString = [convert]::ToString($SleepSeconds)
            Write-Progress -Activity "Please wait for $SleepSecondsString seconds" -Status "Waiting" -PercentComplete ($Count/$SleepSeconds*100)
            Start-Sleep -Seconds 1
        }
        Write-Progress -Activity "Please wait for $SleepSecondsString seconds" -Status "Completed" -Completed
    } # end function sleepProgress([hashtable]$SleepHash)

#############

    function waitForProcessToStartOrTimeout( [string]$procname, [int]$timeoutvalue ) {
    # Takes process name (without exe) and time value in seconds, waits until process starts or timesout
        [int]$waittries = 0
        Write-Host "Waiting until Process `'$procname`' is started or until $($timeoutvalue) seconds have passed..."
        while ( !(Get-Process $procname -ErrorAction SilentlyContinue) -and ($waittries -lt $timeoutvalue) )  {
            Start-Sleep -Seconds 1
            $waittries += 1
        }
        if ( $waittries -ge $timeoutvalue ) {
            Write-Host "Waiting for Process `'$procname`' to start timed out."
            return 258
        } else {
            Write-Host "Process `'$procname`' has started. Continuing..."
            return 0
        }
    }

#############

    function waitForProcessToStopOrTimeout( [string]$procname, [int]$timeoutvalue ) {
    # Takes process name (without exe) and time value in seconds, waits until process stops or timesout
        [int]$waittries = 0
        Write-Host "Waiting until Process `'$procname`' has exited or until $($timeoutvalue) seconds have passed..."
        while ( (Get-Process $procname -ErrorAction SilentlyContinue) -and ($waittries -lt $timeoutvalue) )  {
            Start-Sleep -Seconds 1
            $waittries += 1
        }
        if ( $waittries -ge $timeoutvalue ) {
            Write-Host "Waiting for Process `'$procname`' to exit timed out."
            return 258
        } else {
            Write-Host "Process `'$procname`' has exited. Continuing..."
            return 0
        }
    }

#############

    function doOptionsRequireConfirmation( ) {
    # Waits for input, returns $true if y or Y pressed, not case sensitive
    # or just returns $true when option for no confirmation required is set or running with -silent switch
        if ( !($Global:requireConfirmationBeforeRemoval) -or $Global:isSilent ) {
            if ( $Global:isSilent ) {
                Write-Verbose -Verbose "Running with -silent switch implies confirmation request is not required."
            }
            Write-Verbose -Verbose "You have chosen option to not require confirmation, continuing..."
            return $true
        } else {
            Write-Verbose -Verbose "Press 'Y' to continue or press any other key to stop..."
            $key = [Console]::ReadKey($true)
            return ($key.Key -eq 'y')
        }
    }

#############

    function doOptionsWin10RecommendedDownloadsOff( ) {
        Write-Host "" | Out-Default
        Write-Verbose -Verbose "Setting Registry Keys to turn off `"recommended`" downloads (ads) of applications that would have automaticaly downloaded."
        # Prevent "recommended" downloads of apps automatically by the OS
        Write-Host "Setting HKLM\Software\Policies\Microsoft\Windows\CloudContent\DisableWindowsConsumerFeatures to 1 (REG_DWORD)" | Out-Default
        & reg add "HKLM\Software\Policies\Microsoft\Windows\CloudContent" /v DisableWindowsConsumerFeatures /d 1 /t REG_DWORD /f 2>&1 | Out-Default
        Write-Host "Setting HKCU\Software\Policies\Microsoft\Windows\CloudContent\DisableWindowsConsumerFeatures to 1 (REG_DWORD)" | Out-Default
        & reg add "HKCU\Software\Policies\Microsoft\Windows\CloudContent" /v DisableWindowsConsumerFeatures /d 1 /t REG_DWORD /f 2>&1 | Out-Default
        Write-Host "Setting HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager\ContentDeliveryAllowed to 0 (REG_DWORD)" | Out-Default
        & reg add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v ContentDeliveryAllowed /d 0 /t REG_DWORD /f  2>&1 | Out-Default
        Write-Host "Setting HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager\SilentInstalledAppsEnabled to 0 (REG_DWORD)" | Out-Default
        & reg add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v SilentInstalledAppsEnabled /d 0 /t REG_DWORD /f  2>&1 | Out-Default
        Write-Host "Setting HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager\SystemPaneSuggestionsEnabled to 0 (REG_DWORD)" | Out-Default
        & reg add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager\" /v SystemPaneSuggestionsEnabled /d 0 /t REG_DWORD /f  2>&1 | Out-Default
    }

#############

    function doOptionsWin10StartMenuAds( ) {
        Write-Host "" | Out-Default
        Write-Verbose -Verbose "Exporting Start Menu tiles layout."

        try {
            Export-StartLayout "$($Script:dest)\exported-startlayout.xml"
        } catch {
            Write-Warning "Export-StartLayout did not complete. Try updating Windows and rebooting first, then re-running this with this option enabled again."
        }


        Write-Host "" | Out-Default
        Write-Verbose -Verbose "Removing Advertisements (Windows ContentDeliveryManager) Ads from exported Layout for new users."
        $startlayout = Get-Content "$($Script:dest)\exported-startlayout.xml" -Raw
        $noCDMadsstartlayout = $($startlayout -Replace ".*<start:SecondaryTile\ AppUserModelID=`"Microsoft\.Windows\.ContentDeliveryManager.*\ />.*\n.*?")
        Write-Host "" | Out-Default
        Write-Verbose -Verbose "Setting default Start Menu tiles layout for new users only (doesn't apply to any current user or existing account)."
        Set-Content -Path "$($Script:dest)\exported-startlayout-noCDMads.xml" -Value $noCDMadsstartlayout

        mkdir "$env:LOCALAPPDATA\Microsoft\Windows\Shell" -Force | Out-Null
        Import-StartLayout -LayoutPath "$($Script:dest)\exported-startlayout-noCDMads.xml" -MountPath "$($env:SystemDrive)\"

        Remove-Item "$($Script:dest)\exported-startlayout.xml" -Force -ErrorAction SilentlyContinue
        Remove-Item "$($Script:dest)\exported-startlayout-noCDMads.xml" -Force -ErrorAction SilentlyContinue
    }

#############

    function doWindows10Options( ) {
        if ( $Script:winVer -ge 10 ) {
            if ( $Global:optionsWin10RecommendedDownloadsOff ) {
                Write-Host "`nConfirm setting Win10 UWP `"recommended`" apps (ads) auto-download to be off." | Out-Default
            }

            if ( $Global:optionsWin10StartMenuAds ) {
                Write-Host "Confirm setting default Start Menu for new user accounts (doesn't affect existing accounts)." | Out-Default
            }

            if ( $Global:optionsWin10RecommendedDownloadsOff -or $Global:optionsWin10StartMenuAds ) {
                [bool]$isConfirmed = doOptionsRequireConfirmation
            }

            if ( $Global:optionsWin10RecommendedDownloadsOff -and $isConfirmed ) {
                Write-Host "Setting Win10 UWP `"recommended`" apps (ads) auto-download to be off." | Out-Default
                doOptionsWin10RecommendedDownloadsOff
            } else { # end if ( $Global:optionsWin10RecommendedDownloadsOff )
                Write-Verbose -Verbose "You have chosen to not set Win10 `"recommended`" auto-downloads of UWP apps off."
            }
            if ( $Global:optionsWin10StartMenuAds -and $isConfirmed ) {
                Write-Host "Setting default Start Menu for new user accounts (doesn't affect existing accounts)." | Out-Default
                doOptionsWin10StartMenuAds
            } else { # end if ( $Global:optionsWin10StartMenuAds )
                Write-Verbose -Verbose "You have chosen to not remove the Win10 ContentDeliveryManager Ads in the Start Menu for new user accounts (doesn't affect existing accounts)"
            }
        } # end if ( $Script:winVer -ge 10 )
    } # end function doWindows10Options( )

#############

        # https://stackoverflow.com/questions/40617800/opening-powershell-script-and-hide-command-prompt-but-not-the-gui
        # .Net methods for hiding/showing the console in the background
        Add-Type -Name Window -Namespace Console -MemberDefinition '
        [DllImport("Kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    '
    function showConsole {

        $consolePtr = [Console.Window]::GetConsoleWindow()

        # Hide = 0,
        # ShowNormal = 1,
        # ShowMinimized = 2,
        # ShowMaximized = 3,
        # Maximize = 3,
        # ShowNormalNoActivate = 4,
        # Show = 5,
        # Minimize = 6,
        # ShowMinNoActivate = 7,
        # ShowNoActivate = 8,
        # Restore = 9,
        # ShowDefault = 10,
        # ForceMinimized = 11

        [Console.Window]::ShowWindow($consolePtr, 3)

    }

    function hideConsole {
        $consolePtr = [Console.Window]::GetConsoleWindow()
        [Console.Window]::ShowWindow($consolePtr, 0)
    }

#############

    function stopTranscript( ) {
        Write-Output "" | Out-Default
        Stop-Transcript
        # format \n as \r\n for viewing log in notepad with proper line breaks
        $logfilecontent = Get-Content $Script:logfile
        $logfilecontent > $Script:logfile
    }

############## GUI FUNCTIONS ##################################################################################################

    function selectedProgsListviewtoArray( $programsListview ) {
        $progslistSelected = @()
        $programsListview.items | % {
            if ($_.Checked) {
                # take checked items and put into array of format similar to $Script:progslisttoremove to work with easier
                $proglistSelectedItem = New-Object –TypeName System.Management.Automation.PSObject
                $i = 0; $_.SubItems | % {
                    if ( $proglistviewColumnsArray[$i] -ne "" ) {
                        $currentSubItem = $_
                        if ( $currentSubItem.Text -eq "" ) {
                            $proglistselectedItem | Add-Member –MemberType NoteProperty –Name $proglistviewColumnsArray[$i] -Value $null
                        } else {
                            $proglistselectedItem | Add-Member –MemberType NoteProperty –Name $proglistviewColumnsArray[$i] -Value $currentSubItem.Text
                        }
                    }
                    $i++
                }


                if ( $proglistselectedItem.PackageName ) {
                    $proglistselectedItem  = $proglistselectedItem | Select-Object -Property $proglistviewColumnsArray -ExcludeProperty DisplayName | Select-Object @{Name="DisplayName";Expression={$_.Name}},* -ExcludeProperty Name
                    $progslistSelected += @( $proglistselectedItem | Select-Object * -ExcludeProperty Name )
                } else {
                    $progslistSelected += @( $proglistselectedItem | Select-Object * -ExcludeProperty DisplayName )
                }

            } # end if ($_.Checked)
        } # end $programsListview.items | %



        return $progslistSelected
    } # end function selectedProgsListviewtoArray( $programsListview )

#############

    function generateProgListView( $progslisttodisplay ) {
        $programsListview.BeginUpdate()
        $programsListview.Items.Clear()
        $programsListview.Columns.Clear()
        $programsListview.Columns.Add(" ") | Out-Null # for checkbox
        $proglistviewColumnsArray | Select-Object -Skip 1 | % { $programsListview.Columns.Add("$($_)") | Out-Null }

        $backgroundhighlight = 1 # alternate color each row, first row off

        [int]$Script:numofSelectedProgs = 0

        # Modification (show/hide) options of what to display in list here
        if ( !($Global:showMicrosoftPublished) ) {
           # $progslisttodisplay = $progslisttodisplay | Where { $_.Name -notmatch "Microsoft" -and $_.Publisher -notmatch "Microsoft" -and $_.PackageName -notmatch "Microsoft" -and $_.PackageFullName -notmatch "Microsoft" }
            $progslisttodisplay = $progslisttodisplay | Where { $_ -notmatch "Microsoft" }
        }
        if ( !($Global:showUWPapps) ) {
            $progslisttodisplay = $progslisttodisplay | Where { ($_.PackageFullName -eq $null) -and ($_.PackageName -eq $null) }
        }

        [int]$Global:numofprogs = @('0',($progslisttodisplay | Measure-Object).Count)[($progslisttodisplay | Measure-Object).Count -gt 0]

        # Strange issue with PS 2.0 Here as ForEach ( $prog in $Global:progslisttodisplay ) had 2 extra loops, using for loop instead works
        For ($i=0; $i -lt $Global:numofprogs; $i++) {
            $prog = $progslisttodisplay[$i]
            $progListViewItem = New-Object System.Windows.Forms.ListViewItem( "" )
            $backgroundhighlight = $backgroundhighlight -xor 1
            if ( $backgroundhighlight ) {
                $progListViewItem.BackColor = [System.Drawing.Color]::Beige
            }
            $proglistviewColumnsArray | Select-Object -Skip 1 | % { $progListViewItem.SubItems.Add("$($prog.$($_))") | Out-Null }
            $programsListview.Items.Add($progListViewItem) | Out-Null
        }
        $programsListview.AutoResizeColumns("HeaderSize")
        Write-Output "" | Out-Default
        Write-Output "Total number of programs: $($Global:numofprogs)"
        $statusBarTextBox.Panels[$statusBarTextBoxTotalProgsIndex].Text = "Total: "+"$($Global:numofprogs)  "
        $programsListview.EndUpdate()
    }

#############

    function generateProgListViewChecked( $progslisttoshowchecked, [bool]$check=$true ) {
    # Applies/Reapplies the checked items to the ListView items
        if ( $progslisttoshowchecked ) {
            $progslisttoshowchecked | % {
            # Compare if 'items to show checked' matches current listview checked items and check them if needed
                $currentprogslistitemtoshowchecked = $_
                ForEach ( $item in $programsListview.Items ) {
                    $i = 1 # skip checkbox
                    ForEach ( $columnheader in $($proglistviewColumnsArray | Select-Object -Skip 1)) {
                        if ( ($item.SubItems[$i].Text) -notmatch [RegEx]::Escape($currentprogslistitemtoshowchecked.$columnheader) ) {
                            Break
                        }
                        $i++
                    } # end ForEach ( $columnheader in ($proglistviewColumnsArray | Select-Object -Skip 1))
                    if ( ($item.SubItems[$i-1].Text) -notmatch [RegEx]::Escape($currentprogslistitemtoshowchecked.$columnheader) ) {
                            Continue
                    }
                    $item.Checked = $check #this also triggers registered event to update status bar selected count text
                    Break
                } # end ForEach ( $item in $programsListview.Items )
            } # end $progslisttoshowchecked | %
        } # end if ( $progslisttoshowchecked )
    } # end function generateProgListViewChecked( $progslisttoshowchecked )

#############

    function refreshAlreadyGeneratedProgramsList( $currentSelectedProgs ) {
        $Global:statusupdate = "Refreshing list of programs to display..."
        Write-Output "" | Out-Default
        $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
        $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
        $Script:programsListview.items | % {
            if ($_.Checked) {
                $_.Checked =  $false
            }
        }
        if ( !($Script:programsListviewWasJustRecreated) ) {
            generateProgListView $Global:progslisttodisplay
        }
        if ( $Script:showSuggestedtoRemove ) {
            generateProgListViewChecked $Global:progslisttoshowchecked
        }
        if ( $currentSelectedProgs ) {
            generateProgListViewChecked $currentSelectedProgs
        }
        $Global:statusupdate = "List updated."
        Write-Output "" | Out-Default
        $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].Text = "  "+$Global:statusupdate
        $statusBarTextBox.Panels[$statusBarTextBoxStatusTextIndex].ToolTipText = $Global:statusupdate
        $Script:programsListviewWasJustRecreated = $false
    }

#############

    function sortprogramsListview {
    # Uses https://etechgoodness.wordpress.com/2014/02/25/sort-a-windows-forms-programsListview-in-powershell-without-a-custom-comparer/ by Eric Siron
    # slightly modified to support sorting on .checked property
        param( [parameter(Position=0)][UInt32]$Column )
        $Numeric = $true # determine how to sort
        # if the user clicked the same column that was clicked last time, reverse its sort order. otherwise, reset for normal ascending sort
        if ( $Script:LastColumnClicked -eq $Column ) {
            $Script:LastColumnAscending = -not $Script:LastColumnAscending
        } else {
            $Script:LastColumnAscending = $false
        }

        [int]$Script:numofSelectedProgs = 0

        $Script:LastColumnClicked = $Column
        $ListItems = @(@(@())) # three-dimensional array; column 1 indexes the other columns, column 2 is the value to be sorted on, and column 3 is the System.Windows.Forms.ListViewItem object
        ForEach( $ListItem in $programsListview.Items ) {
            # if all items are numeric, can use a numeric sort
            if ( $Numeric -ne $false ) {
                try {
                    $Test = [Double]$ListItem.SubItems[[int]$Column].Text
                } catch {
                    $Numeric = $false # a non-numeric item was found, so sort will occur as a string
                }
            }
            if ( $Script:LastColumnClicked -eq '' ) { # if sorting by checked/selected items
                $ListItems += ,@($ListItem.Checked,$ListItem)
            } else {
                $ListItems += ,@($ListItem.SubItems[[int]$Column].Text,$ListItem)
            }
        }
        # create the expression that will be evaluated for sorting
        $EvalExpression = {
            if ( $Numeric ) {
                return [Double]$_[0]
            } else {
                $_[0]
            }
        }
        # all information is gathered; perform the sort
        $ListItems = $ListItems | Sort-Object -Property @{Expression=$EvalExpression; Ascending=@($Script:LastColumnAscending,!($Script:LastColumnAscending))[($Script:LastColumnClicked -eq '')]}
        ## the list is sorted; display it in the programsListview
        $programsListview.BeginUpdate()
        $programsListview.Items.Clear()
        ForEach( $ListItem in $ListItems ) {
            $programsListview.Items.Add($ListItem[1])
        }
        $programsListview.EndUpdate()
    }

#############

    function toggleSuggestedBloatware( ) {
        $viewShowSuggestedBloatware.Checked = !($viewShowSuggestedBloatware.Checked)
        $Script:showSuggestedtoRemove = $viewShowSuggestedBloatware.Checked
        generateProgListViewChecked $Global:progslisttoshowchecked $Script:showSuggestedtoRemove
        $programsListview.Focus()
    }

#############

    function toggleConsoleWindow( ) {
        if ( $Script:isConsoleShowing ) {
            $Script:isConsoleShowing = $false
            hideConsole
        } else {
            $Script:isConsoleShowing = $true
            showConsole
        }
        $viewShowConsoleWindow.Checked = $Script:isConsoleShowing
    }

#############

    function isObjectEqual( $refobj, $diffobj ) {
    # takes two objects (assumes both have same properties), returns boolean true if all properties match
        $refobjPropertiesArray = ,@()
        $refobj.PSObject.Properties | % { $refobjPropertiesArray += $_.Name }
        $numofproperties = ($refobjPropertiesArray | Measure-Object).Count
        For ($i=0; $i -le $numofproperties; $i++) {
            if ( $refobj.($refobjPropertiesArray[$i]) -ne $diffobj.($refobjPropertiesArray[$i]) ) {
                Return $false
            }
        }
        Return $true
    }

#############

    function updateSelectedProgsStatus( $itemChecked ) {
        if ( $itemChecked.NewValue -eq "Checked" ) {
            $Script:numofSelectedProgs += 1
        } else {
            $Script:numofSelectedProgs -= 1
        }
        $statusBarTextBox.Panels[$statusBarTextBoxSelectedProgsIndex].Text = "Selected: "+@('0',[string]$Script:numofSelectedProgs)[[int]$Script:numofSelectedProgs -gt 0]
    }

############## END GUI FUNCITIONS ##############################################################################################


<#
    function unZip( $zipfilename ) { # extracts to same directory
    # https://blogs.iis.net/steveschofield/unzip-several-files-with-powershell
    # Works for zip compressed files but not PE SXE Archives (.exe self extracting)
        $shellApplication = New-Object -Com Shell.Application
        $zipPackage = $shellApplication.NameSpace($zipfilename)
        $destinationFolder = $shellApplication.NameSpace($zipfilename.DirectoryName)
        # CopyHere vOptions Flag # 4 - Do not display a progress dialog box.
        # 16 - Respond with "Yes to All" for any dialog box that is displayed.
        $destinationFolder.CopyHere($zipPackage.Items(),20)
    }

#############

    function deleteNowOrOnReboot( [string]$path2 ) {
    # Adapted from https://gist.github.com/marnix/7565364 to support nested folders
    # Everything in the folder will be deleted before the folder is deleted, recursively

        Add-Type @'
            using System;
            using System.Text;
            using System.Runtime.InteropServices;
            public class Posh {
                public enum MoveFileFlags {
                    MOVEFILE_DELAY_UNTIL_REBOOT = 0x00000004
                }
                [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
                static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, MoveFileFlags dwFlags);
                public static bool MarkFileDelete (string sourcefile) {
                    return MoveFileEx(sourcefile, null, MoveFileFlags.MOVEFILE_DELAY_UNTIL_REBOOT);
                }
            }
'@
        $path = (Get-Item (Resolve-Path $path2 -ErrorAction Stop).Path)

        if ( ($path.PSIsContainer) -and (Get-ChildItem $path) ) {
        #folders with files
            $allitems = (Get-Childitem $path -Force | Select FullName, @{Name="FolderDepth";Expression={$_.DirectoryName.Split('\').Count}} | Sort -Descending FolderDepth,FullName | Select FullName)
            ForEach ($item in $allitems) {
                #Write-Verbose -Verbose $item.FullName
                deleteNowOrOnReboot $item.FullName
            }
            # after marking the children for deletion, mark that folder
            try {
                & takeown /f $path /a /r /d y
                Remove-Item $path -Force -Recurse -ErrorAction Stop
            } catch {
                $deleteResult = [Posh]::MarkFileDelete($path)
                if ($deleteResult -eq $false) {
                    Write-Warning "Could not schedule $($path) for deletion." | Out-Default
                    #throw (New-Object ComponentModel.Win32Exception)
                } else {
                    #Write-Output "Scheduled $($path) for delete on next reboot." | Out-Default
                }
            }
        } else {
        #file or empty folder
            try {
                if ( ($path.PSIsContainer) ) {
                    & takeown /f $path /a /d y
                } else {
                    & takeown /f $path /a
                }
                Remove-Item $path -Force -Recurse -ErrorAction Stop
            } catch {
                $deleteResult = [Posh]::MarkFileDelete($path)
                if ($deleteResult -eq $false) {
                    #throw (New-Object ComponentModel.Win32Exception)
                    Write-Warning "Could not schedule $($path) for deletion." | Out-Default
                } else {
                    #Write-Output "Scheduled $($path) for delete on next reboot." | Out-Default
                }
            }
        }
    } # end function deleteNowOrOnReboot
#>

#############

    <#
    function killPopup([string] $matchstring) {
        if ($matchstring) {
            # Kill Survey Popup (from browser window after HPSA uninstall)
            $a = $null
            $timer = $null
            $timer = [Diagnostics.Stopwatch]::StartNew() # timeout time to wait for Survey Popup, if takes too long or offline just times out and continues script
            Write-Output "`nWaiting to kill popup that matches $($matchstring)..." | Out-Default
            While ( $a -eq $null -and $timer.ElapsedMilliseconds -lt 60000 ){
                $a = ( Get-Process | where { $_.mainWindowTitle -match $matchstring } )
                if ( $a -ne $null) {
                    Stop-Process -Id $a.Id -Force | Out-Null
                    Write-Output "Popup matching $($matchstring) killed." | Out-Default
                }
            }
            if ( $a -eq $null) { Write-Output "`nWaiting for popup timed out." } | Out-Default
            $timer.Stop()
        }
    } # end function killPopup([string] $matchstring)

    #>

} # end BEGIN (define functions before PROCESS block of code starts)

################################################################################################

<#
SendKeys example if needed
Best to use silent command line switches because it isn't timing dependent and also you can handle the exit codes easily.

SendKeys type uninstallation will be a last resort if other methods such as silent command line switches or install shield iss recording of installation do not work.

Write-Verbose -Verbose "Please do not use the keyboard or mouse during this time for the script to be able to correctly activate the uninstaller."
Write-Host "`n"
Write-Host "Waiting 45 seconds for uninstaller to load..."
#Start-Sleep -Seconds 45
$SleepTime = @{"Seconds" = 45}
sleepProgress $SleepTime
$uninstallerstarted = $null
if ( Get-Process McUIHost -ErrorAction SilentlyContinue ) {
    $uninstallerstarted = (New-Object -ComObject WScript.Shell).AppActivate((Get-Process McUIHost).MainWindowTitle)
    if ($uninstallerstarted) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}{TAB}{TAB}{TAB}")
        Start-Sleep -Seconds 1
        [System.Windows.Forms.SendKeys]::SendWait(" ")
        Start-Sleep -Seconds 1
        [System.Windows.Forms.SendKeys]::SendWait("{TAB} ")
        Start-Sleep -Seconds 1
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}{TAB} ")
        Start-Sleep -Seconds 3
        [System.Windows.Forms.SendKeys]::SendWait(" ")
        Write-Host "Waiting 7 minutes for uninstall to finish before continuing..."
        #Start-Sleep -Seconds 420
        $SleepTime = @{"Minutes" = 7}
        sleepProgress $SleepTime
        $uninstallerstarted = (New-Object -ComObject WScript.Shell).AppActivate((Get-Process McUIHost).MainWindowTitle)
        if ($uninstallerstarted) {
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}{TAB}{TAB} ")
            $proc.WaitForExit() #check if this exists in function scope
        }
    }
}
#>
