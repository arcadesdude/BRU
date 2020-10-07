# Bloatware Removal Utility (BRU)    ![BRU](BRU.PNG?raw=true "BRU Icon")

Bloatware Removal Utility, for automating removal of pre-installed, factory bloatware from devices running Windows 7-10 and newer. Silently removes items selected if possible. Preselects common bloatware. Can remove Win10 UWP/Metro/Modern/Windows Store apps and provisioned apps.

Bloatware Removal Utility
Removes common bloatware from HP, Dell, Lenovo, Sony, Etc
Supports Powershell 2+, Windows 7/Server 2008 R2 (Winver 6.1+) and newer - including removing Win8/10+ UWP (metro/modern) Apps.
Reboot before running this script and after running it (if anything is removed)

# Intended use

MSPs (Managed Service Providers), IT Professionals, Computer Repair shops and those who need to uninstall bloatware on a machine may find this useful. Careful! It will pre-select Microsoft Office and other applications you may want to keep. Review the list prior to clicking on 'Remove Selected' as it will be uninstalled and not recoverable. Use at your own risk!

Imaging would be a good way to set up multiple computers of the same model as there are sometimes issues with restoring images to dissimilar hardware. Also if you happen to have many different models of machines or would like to quickly and quietly remove the default bloatware that comes with many HP, Dell, and a few others this script supports this can automate that process to save you time and free you up for other more pressing concerns. This script will bring up the uninstallers and remove items silently in most cases. I've used it mostly for HP ProBook/EliteBook/ProDesk and varients and also Dell Insprion/Lattitude/OptiPlex/Precision. HP is by far the worst when it comes to preinstalled bloatware.

# History/Inspiration

I've credited many of the original ideas and parts that helped make up this script inside it with comments on the relevant sections. There were some 'HP bloatware' removal scripts out there but they didn't get everything and weren't totally automated. I've tried to make this as automated as possible but still feel those scripts were valuable in getting the right approach to solve this problem and contributed to my work so they are appropriately attributted as well.

# Creating the Uninstall Helpers folder

Supporting files that are needed should be saved to the "BRU-uninstall-helpers" folder (named exactly that without the quotes).

You'll need to create that folder and get the appropriate uninstall helper files to support removal of programs like:
McAfee products
HP JumpStart Apps
HP Client Security Manager
Office Click-2-Run apps (Preinstalled O365 which prevents Business licensed versions from installing)

The folder 'BRU-Uninstall-Helpers' should be in the same location as the PS1/BAT files:

![BRU-Uninstall-Helpers-Folder-Layout](BRU-Uninstall-Helpers-Folder-Layout.PNG?raw=true "Folder Layout")

The contents of the 'BRU-Uninstall-Helpers' folder.

![BRU-Uninstall-Helpers-Folder-Contents](BRU-Uninstall-Helpers-Folder-Contents.PNG?raw=true "Folder Contents")

# Obtaining Specific Bloatware Uninstall Helpers

streams.exe (for unblocking files and preventing script from getting closed without warning by Windows SmartScreen)
Streams v1.6 By Mark Russinovich may be downloaded from:
https://docs.microsoft.com/en-us/sysinternals/downloads/streams
Place the streams.exe in the BRU-uninstall-helpers folder (streams64.exe is not needed). When the .Bat file is run as administrator it will check for streams.exe and run the commands to remove the Zone.Identifier info that it was downloaded from the internet. If it isn't removed, Windows Smartscreen may suddenly close the Powershell script before it is able to run.


devcon
HP Client Security Manager
Uses DevCon to disable/reenable DVD/CD drive during uninstallation to prevent an HP uninstaller bug.
https://networchestration.wordpress.com/2016/07/11/how-to-obtain-device-console-utility-devcon-exe-without-downloading-and-installing-the-entire-windows-driver-kit-100-working-method/


WASP
HP JumpStart Apps or 'VIP Access' (Comes with old Norton)
Uses the WASP uninstall helper
https://wasp.codeplex.com/
It is a dll file called WASP.dll.

To get the WASP.dll file Go to https://archive.codeplex.com/?p=wasp

Click on the download archive button (on that page not here ;)
![Download Archive at Codeplex](https://user-images.githubusercontent.com/14213202/45259103-dda56000-b392-11e8-9723-bd4cdb59192e.png?raw=true)

In the archive
https://codeplexarchive.blob.core.windows.net/archive/projects/WASP/WASP.zip

Go to the releases\4\55453160-4bf6-41a4-be7f-7cacc781b9b6 file and rename it .zip

![image](https://user-images.githubusercontent.com/14213202/45259074-6ff93400-b392-11e8-8bd4-3514069a80d4.png)

![image](https://user-images.githubusercontent.com/14213202/45259088-9fa83c00-b392-11e8-8e84-e8fbf9d9d9fe.png)

The file you need is inside that as WASP.dll, (ver 1.2.0.0, 42kb).

![image](https://user-images.githubusercontent.com/14213202/45259092-bc447400-b392-11e8-8de2-b4e0e0ac7db1.png)

The snap-ins aren't needed just the dll file.



OffScrub23.vbs
Microsoft Office/C2R Office365 Preinstalled Apps
Uses updated OffScrubc23.vbs for 2013/2016/2017/2018
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Deploy-OfficeClickToRun/OffScrubc2r.vbs


MCRP.exe
McAfee (Consumer) Applications
Uses MCRP.exe
http://us.mcafee.com/apps/supporttools/mcpr/mcpr.asp

# Usage

For silent / command line usage see the next section. Below is for GUI usage (default).

Right Click and run as administrator on the BAT file (not the PS1) file.

![Run as Administrator](run-as-admin.PNG?raw=true "Run As Administrator")

The program will get all installed software and show you a list you can pick from what you want to be removed (silently if possible).
It may take 30 seconds to 3-4 minutes to show depending on the speed of the device it is running on and the total number of installed programs.

![GUI](GUI1.PNG?raw=true "List of programs GUI")

Carefully review the selected items. Remember this is intended for factory fresh systems to remove bloatware and prepare them for your specific setup -- intended for preparers of computers and IT professionals. Don't simply click remove selected without reviewing the program list.

Disclaimers aside, The pre-selected items are built from fuzzy, regex patterns which you can modify in the script.
It matches bloatware against lists of items, and doesn't match other items (like drivers and such). Special cases are handled after the general list and done in a specific order (as some programs require others to be removed first (HP Client Security Manager is one such program that needs several programs removed prior to its removal, for example)).

Check the Options if you want to change automatic reboot after uninstall of all programs, confirmation prompts and System restore point options. There are some specific Windows 10 options as well.

The setting recommended UWP apps auto download off option is supposed to stop UWP and windows store 'recommended' applications from being automatically pushed and installed. In newer windows versions 1703 and on, it may not actually work (or on non-educational/enterprise versions of Windows 10). Note that whenever a windows update is installed, Windows tends to reinstall these UWP/suggested/recommended applications which end up being games or promotional content.

The other Windows 10 specific option 'set default start menu layout for new users' will not affect any existing accounts or current users. If a new user profile is created it will though. What this does is once the bloatware UWP apps are removed, they're also unpinned from the start menu so the new user won't see the uninstalled UWP bloatware applicaitons. This doesn't always seems to work and may give an error about the tiledatabase unless windows is updated first. So for setting up the computer, create your setup admin account first, update windows completely, then run this script to remove the bloatware and set the default start menu layout, then create the new user account which should start off without the default tiles pinned to the start menu.

![BRU-Script-Running](BRU-2.PNG?raw=true "BRU Script Running")

# Silent / Command line usage

To run from command line launch either an admin command prompt and type powershell or launch an administrator powershell.

You may have to set your execution policy to allow scripts to run. If you have Windows SmartScreen on you may have to right click the ps1 file and click Properties then Unblock file and OK. Or use the PS3+ command `Unblock-file`.

`Set-ExecutionPolicy -ExecutionPolicy RemoteSigned`

The following command line options are supported.

`-silent (or -quiet or -s)`

  Silent mode. Without this switch the GUI will run and manual user input will be required.

`-nd (or -id or -ignoredefault or -ignoredefaults or -ignoredefaultsuggestions or -nodefaultsuggestions)`

  This will not reference the built in suggestions lists so you'll need to use this with `-include, -exclude and/or -includelast (-specialcases)`

`-reboot -rebootafterremoval`

  Reboots after running silently. You can check the log (see next section) for details after script runs.

`-include -includefirst`

  This will allow you to choose what you want to include. This comes after the default list if that is used or, if you want to not use the built in suggestions be sure to use the -nd switch (or other above aliases) to prevent the default detection list of including what you don't want. You would include using *Regular Expressions* (escaped and case-INsensitive). The list to include is separated by | if you need to use | in the program name you can escape it with a preceeding backslash \ Here is an example:

`   -include "PROGRAM\ NAME|Something-else|HP\ .*"`

`-exclude -filter`

  This will allow you to exclude (not detect) items you don't want to match. This matches text in Regular Expressions but it is escaped in the program so you would enter examples such as:

`  -exclude '"keyboard","driver"'`

If you have more than one item and are using Powershell Version 2, you'll need to wrap the strings into a single quoted string (like in the example above). If you're using newer Powershell versions you don't have to do that and can just put in the items to match in quotes separated by commas without having to wrap the entire string in single quotes.

 What you put into each "string" above will turn into a single Regex escaped string like `".*keyboard|driver.*"` That is done automatically by the program so you don't have to escape it here.

`-includelast -specialcases`

  This is for programs you want uninstalled AFTER everything else. Useful for stuff that needs to come after other stuff to be removed properly (`*cough*` HP Client Security Manager `*cough*`). This matches text in Regular Expressions but it is escaped in the program so you would enter examples such as:

`  -includelast '"HP Client Security Manager","HP Support Assistant"'`

If you have more than one item and are using Powershell Version 2, you'll need to wrap the strings into a single quoted string (like in the example above). If you're using newer Powershell versions you don't have to do that and can just put in the items to match in quotes separated by commas without having to wrap the entire string in single quotes.

 What you put into each "string" above will turn into a single Regex escaped string like `".*HP\ Client\ Security\ Manager|HP\ Support\ Assistant.*"` You don't have to do that but it is good to know that happens in the program automatically.

`-win10leaverecommendedappsdownloadon`

  This will keep the default Windows 10 option for Windows to download those random, recommended game apps like Candy Crush and such. If you don't include this option this script will set the registry keys to stop that unwanted behavior (default).

`-win10leavestartmenuadson -keepstartads`

  This option will keep the Win10 ContentDeliveryManager Ads that appear in the start menu. Not having this option will remove the ads and export the default option so new default accounts won't see them. It doesn't affect existing Windows accounts.

`-norestorepoint -skiprestorepoint -nr`

  This skips the Windows Restore Point creation attempt which is on by default.

`-dry -dr -dryrun -detect -detectonly -whatif`

  Dry Run / Detect Only / WhatIf mode will not remove anything but show you what your -include and -exclude (and -specialcases) filters will target if you're working on trying to target just specific software to be removed.

# Full example from Powershell admin prompt:

Remove All HP apps and do the Client Security Manager and Support Assistant last:
`.\Bloatware-Removal-Utility.ps1 -silent -nd -includelast '"HP Client Security Manager","HP Support Assistant"' -include '"HP\ .*"'`

If you find a setup that works for you you can modify the batch script to specify the options. The current batch script will also run the streams.exe program if you've included it in the uninstall helpers folder to remove the download zone information from the PS1/VBS/BAT/EXE files so Windows SmartScreen doesn't stop the script from running when launching.


# Log

Logfile will be saved in c:\BRU (or you can edit script to suit your needs). If running with the automatic reboot option this is handy to see if something did not automatically uninstall and what error message was given.

# After running

Be sure to reboot after running this as some programs need a reboot when uninstalling. Also you can compare the programs and features list of currently installed programs and see if there is anything left you would need to manually uninstall.

# Version History
10/07/2020
- Fixed Matching issues. Rewrote core matching and fixed issues matching when exclude list blank in silent/cli options.

09/20/2020
- Fixed Matching issues. Rewrote core matching and fixed out of order or match/not match issues with command line options.
- Updated Inclusion/Exclusion default suggestions.
- Excluded "Dell MD Storage" from default suggestions
- Added "HPSureShield" UWP app to suggested apps
- Added "HPSupportAssistant" UWP app to suggested apps
- Added "HPPrivacySettings" UWP app to suggested apps
- Added "FarmHeroesSaga" UWP app to suggested apps
- Added "Norton" UWP app to suggested apps
- Added "Norton Security" UWP app to suggested apps
- Minor display fixes

01/28/2020
- Added LenovoUtility (Vantage) UWP app detection to the list of suggested apps to remove.

11/27/2019
- Added Added ASUS software including ASUSGiftBox, ASUSPCAssistant (MyASUS) and McAfee Security (UWP) app detection to the list of suggested apps to remove.

11/23/2019
- Added logging full command line options if run silently
- Fixed https://github.com/arcadesdude/BRU/issues/5 "-includelast or -specialcases not working"

10/17/2019
- Added "HPInc.EnergyStar" UWP app to suggested apps
- Added "HPPrinterControl" UWP app to suggested apps
- Added "HPPrivacySettings" UWP app to suggested apps
- Added "HPSupportAssistant" UWP app to suggested apps
- Added "HPSystemEventUtility" UWP app to suggested apps

08/29/2019
- Fixed minor bug in programs list generation

06/02/2019
- Fixed bug in function refreshProgramsList when adding registry results from multiple keys

03/24/2019
- Fixed detection of MS Office UWP apps
- Fixed selection bug when generating default detected list of bloatware when running in command line options mode with more than one match
- Added silent removal support for "Lenovo App Explorer"
- Updated documentation


10/06/2018
- Added silent command line options, custom include/exclude lists, dry run/WhatIf option, etc - see Silent / Command line usage section
- Added URL in GUI about window to the github link
- Fixed match detection bugs and updated comments to match what they are. Only the included items need to be manually Regex escaped
- Fixed Batch file launcher to fix SmartScreen issues (to prevent Windows SmartScreen from closing script window)
- Fixed GUI list checked items when refreshing programs list
- Updated McAfee uninstall helper launch args
- Updated documentation, fixed typos, etc


9/16/2018
- Added streams.exe command in batch file (from Sysinternals) to remove Zone.Identifier so scripts won't get closed without warning by Windows SmartScreen. You'll need to download that separately and put streams.exe in the BRU-uninstall-helpers folder. See "Obtaining Specific Bloatware Uninstall Helpers"
- Changed Windows Store version of Office detection for preinstalled UWP Office
- Added "CookingFever" UWP app to suggested apps
- Added "DragonManiaLegends" UWP app to suggested apps
- Added "HPBusinessSlimKeyboard" UWP app to suggested apps


6/05/2018
- Added "Viber" UWP app to suggested apps
- Added "ACGMediaPlayer" UWP app to suggested apps
- Added "BlueEdge.OneCalendar" UWP app to suggested apps
- Added "HiddenCityMysteryofShadows" UWP app to suggested apps
- Added "LenovoCompanion" UWP app to suggested apps
- Added "LenovoCorporation.LenovoID" UWP app to suggested apps
- Added "LenovoCorporation.LenovoSettings" UWP app to suggested apps


4/13/2018
- Added "McAfeeSecurity" UWP app to suggested apps
- Added "LinkedInforWindows" UWP app to suggested apps
- Added "MediaSuiteEssentials" UWP app to suggested apps
- Added "Power2Go" UWP app to suggested apps
- Added "PowerDirector" UWP app to suggested apps
- Added "PowerMediaPlayer" UWP app to suggested apps
- Added "DellCustomerConnect" UWP app to suggested apps
- Added "DellHelpSupport" UWP app to suggested apps
- Added "DellProductRegistration" UWP app to suggested apps
- Added "Microsoft.Office.Desktop" UWP app to suggested apps (Windows Store version of Office)


3/01/2018
- Added "HPWorkWise64" UWP app to suggested apps
- Excluded "HP Battery Recall Utility" from suggested apps
- Fixed - Increased delay in between removing UWP apps from 2 seconds to 4 seconds
- Added Screenshot of script running/removing bloatware


2/28/2018
- Added "SpotifyAB" UWP app to suggested apps
- Added "CaesarsSlotsFreeCasino" UWP app to suggested apps
- Added "DisneyMagicKingdoms" UWP app to suggested apps
- Added "DolbyAccess" UWP app to suggested apps
- Added "Duolingo" UWP app to suggested apps
- Added "PhototasticCollage" UWP app to suggested apps
- Added "PicsArt" UWP app to suggested apps
- Added "TheNewYorkTimes" UWP app to suggested apps
- Added "TuneInRadio" UWP app to suggested apps
- Added "WinZipUniversal" UWP app to suggested apps
- Added "Wunderlist" UWP app to suggested apps
- Added "XINGAG\.XING" UWP app to suggested apps
- Option 'After removal set "recommended" UWP app auto-downloads off' now also sets HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager\SystemPaneSuggestionsEnabled to 0
