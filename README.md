# AD GPO Backup

![gpo_backup_tool_r1 1](https://user-images.githubusercontent.com/55346298/120389418-adfdc980-c32c-11eb-9dfb-bc43ff5fd311.gif)

When editing Group Policy Objects (GPO) you should always be aware of taking backups before changing anything. But from time to time you need to backup multiple GPO. Sometimes you need to backup all GPOs that are linked to a Organizational Unit and all units below. There is no GUI tool from Microsoft supporting this case, therefore I’ve built a small PowerShell script that will support you.
The script is easy to handle. Just start it and select the OU your policies are linked to. The script will start backup of your linked GPOs (recursively).
I’ve used the “Active Directory OU picker” from [MicaH’s IT blog](https://itmicah.wordpress.com/2016/03/29/active-directory-ou-picker-revisited/) for easy OU selection.

## Installation

Download the latest [Release](https://github.com/beneckecloud/AD_GPO_Backup/releases) and unzip the content to a folder. Start the script by running 

```powershell
.\AD_Backup_GPO.ps1
```

## Usage

The script will automatically check if all requirements are met.

![image](https://user-images.githubusercontent.com/55346298/120388528-80fce700-c32b-11eb-8dba-08c53f932609.png)

Select the OU you like to Backup. The script will also backup all linked group policy opjects beneath your selection when selecting "Auto Select Sub OU" (default). 
If not selected only group policy objects linked to the selected OU will be part of the backup.
When selecting "ZIP Backup" (default) the backup will be compressed as zip file.

![image](https://user-images.githubusercontent.com/55346298/120388577-9540e400-c32b-11eb-9030-2c4fe49d1c01.png)

A summary will be shown when everything is done.

![image](https://user-images.githubusercontent.com/55346298/120388651-b0abef00-c32b-11eb-8607-bd8f812d3080.png)

Your backup will be located in AD_GPO_Backup\01_GPO_Backups\

![image](https://user-images.githubusercontent.com/55346298/120388751-d933e900-c32b-11eb-826f-643b86c16096.png)

## License
[MIT](https://choosealicense.com/licenses/mit/)
