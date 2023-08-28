# Teams Auto Attendant Backup and Restore

## Description
There is no native Backup and Restore of Auto Attendants configuration in Teams.
The script uses smart object manipulation in order to create Auto Attendants in Teams.

Objective:
* Allow organisations to migrate Auto Attendants in between Tenants
* Backup and restore Auto Attendant configuration in case unintended changes have been made

### Solution and dependencies

A key element of the solution is the Azure Audit log where all group membership events are logged. The Power Automate flow queries the Audit log via Graph API to filter users which have been added to a team recently. The flow runs every 30min per default to find user with a matching domain. If removed from the team via Graph API

Prequisite:
* Installed Teams Powershell
* Teams Admin Role

### Installing

Import the Power Automate flow, add email connector and adjust the variables.

Variables: 
Runcycle
GraphTenanid
GraphClientSecret
GraphClientID

### Executing program

Adjust the runcycle variable to match with the Power Automate schedule. 
Default is 30min 

## Help

Contact the author

## Known Issues or future work items

Contributors names and contact info

ex. Mario Möller

## Authors

Contributors names and contact info

ex. Mario Möller


## Version History

* 1
    * First release

## Acknowledgments

Inspiration, code snippets, etc.
* [Send Email on School Data Sync Errors](https://emea.flow.microsoft.com/en-us/galleries/public/templates/ffec9fa3101e4a8281a2b2f7425ef0f1/send-email-on-school-data-sync-errors/)
