# Teams Auto Attendant Backup and Restore

## Description
There is no native Backup and Restore of Auto Attendants configuration in Teams.
The script uses smart object manipulation in order to create Auto Attendants in Teams.

Objective:
* Allow organisations to migrate Auto Attendants in between Tenants
* Backup and restore Auto Attendant configuration in case unintended changes have been made

### Solution and dependencies

Backup 
The script downloads the Auto Attendant configuration (output of Get-CSAutoattendant) to the defined working directory.
If Audio prompts have been found in Auto Attendants also the Audio files will be downloaded.

Restore
The script will restore the Auto Attendant configuration including the Audio Prompts.
For limitation and future work please read the limitations section.

Prequisite:
* Installed Teams Powershell
* Teams Admin Role

### Installing

Download the Powershell Code
Make sure Teams Powerhsell is installed

### Executing program

Define the working folder. This folder holds the configuration files but also Auto Attendant audio prompts.

Define working folder
$path="C:\tmp\backup\"

This function runs the backup operations and stores all files in a subfolder AABackup_YYYYMMDD (YYYYMMDD reflects the current date in Year, Month, Date format)

BackupAAConfig -Path "C:\tmp\backup\"

This function restores the Auto Attendants.
Please make sure you are loggin in to the right tenant for restore operation.

RestoreAAConfig -Path "C:\tmp\backup\AABackup_20230827\"

## Known Issues or future work items

-Ressource Accounts will not be created or connected to the Auto Attendant instance
-In a tenant to tenant migration scenario user objects do not exist in the destinantion tenant.
 Therefore Users,Operators, Groups, nested Auto Attendants, Call Queues need to be created and reassigned

## Authors

Mario Moeller

## Version History

* 1
    * First release
