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

devcon
HP Client Security Manager
Uses DevCon to disable/reenable DVD/CD drive during uninstallation to prevent an HP uninstaller bug.
https://networchestration.wordpress.com/2016/07/11/how-to-obtain-device-console-utility-devcon-exe-without-downloading-and-installing-the-entire-windows-driver-kit-100-working-method/

WASP
HP JumpStart Apps or 'VIP Access' (Comes with old Norton)
Uses the WASP uninstall helper
https://wasp.codeplex.com/

OffScrub23.vbs
Microsoft Office/C2R Office365 Preinstalled Apps
Uses updated OffScrubc23.vbs for 2013/2016/2017/2018
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Office-ProPlus-Deployment/Deploy-OfficeClickToRun/OffScrubc2r.vbs

MCRP.exe
McAfee (Consumer) Applications
Uses MCRP.exe
http://us.mcafee.com/apps/supporttools/mcpr/mcpr.asp

# Usage

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

# Log

Logfile will be saved in c:\BRU (or you can edit script to suit your needs). If running with the automatical reboot option this is handy to see if something did not automatically uninstall and what error message was given.

# After running

Be sure to reboot after running this as some programs need a reboot when uninstalling. Also you can compare the programs and features list of currently installed programs and see if there is anything left you would need to manually uninstall.

# Version History

2/28/2018 - Added SpotifyAB UWP app to Windows 1709 suggested apps
