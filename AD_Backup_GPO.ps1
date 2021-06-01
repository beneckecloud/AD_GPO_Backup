<#
.Synopsis
   Backup Group Policy Objects (GPO) linked to an Oranizational Unit (OU)
.DESCRIPTION
   Starts Backup of all GPOs linked to an defined OU including all child OUs. In Addition to
   the general backups, HTML reports will be generated and saved.   
   Backups and reports are located in the folder "GPO_Backups\%Date%_%Time%"   
   After the backup process a grid view will report the result of the process in a PowerShell GUI
   and will save all the displayed information to "GPO_Backups\%Date%_%Time%\%Date%_%Time%.csv"
.EXAMPLE
   .\AD_Backup_GPO.ps1
   
   Will start a prerequirement check and a graphical user interface to select the OU you like to backup.
   After selecting the OU the script will start the backup by executing "Backup-GPOCB -RootOU OU=Selected_OU"
   
.OUTPUTS
   Backup Folder:	GPO_Backups\%Date%_%Time%\
   
   GPO Backups:		GPO_Backups\%Date%_%Time%\{2C1BAA1B-E393-409F-852B-1CF745F644B8}\
					GPO_Backups\%Date%_%Time%\{2C370DBE-9789-4755-A573-2C97CAE32F1B}\
					GPO_Backups\%Date%_%Time%\...\
   
   Reports:			GPO_Backups\%Date%_%Time%\GPO_Name_1.html
					GPO_Backups\%Date%_%Time%\GPO_Name_2.html
					GPO_Backups\%Date%_%Time%\...
   
   Summary:			GPO_Backups\%Date%_%Time%\%Date%_%Time%.csv 
   
   Log:				Log\%Date%_%Time%_All.log 
.NOTES
   Author : Christopher	Benecke
   Version: 1.0
   Date   : 3/8/2019
.LINK
   https://www.benecke.cloud
#>

function Get-ScriptDirectory{
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
	$Invocation.PSScriptRoot
}

function Get-Requirements{
    
    Clear-Host

    Write-Host ""
    Write-Host "< Check Requirements >"
    Write-Host ""
    
    Start-Transcript -path "$(Get-ScriptDirectory)\Log\$(get-date -format 'yyyyMMdd_HHmmss')_All.log" -Force | Out-Null
    
    #PSVersion 5 Required
    $FullVersion = $PSVersionTable.PSVersion

    if( $PSVersionTable.PSVersion.Major -ge 5 ) {write-host "[OK] Powershell V$FullVersion" -ForegroundColor Green}
    elseif( $PSVersionTable.PSVersion.Major -le 4 ) {
        write-host "[ERROR] PowerShell V$FullVersion is not supported. At least V5.0 is required." -ForegroundColor Red
        Sleep -Seconds 5
        Exit
    } 

    #Module GroupPolicy Required
    Import-Module GroupPolicy -Verbose:$false -ErrorAction SilentlyContinue

    $m = Get-Module -ListAvailable GroupPolicy

    if($m){
        Write-Host "[OK] Module Import 'GroupPolicy'" -ForegroundColor Green       
    }
    else{
        write-host "[ERROR] Module Import 'GroupPolicy'" -ForegroundColor Red
        Sleep -Seconds 5
        Exit   
    }

    #DOT Source

    try{
        $ExecuteTSTpath = $(Get-ScriptDirectory) + "\Backup-GPO.ps1"
        . $ExecuteTSTpath
        Write-Host "[OK] Import Backup Logic 'Backup-GPO.ps1'" -ForegroundColor Green
    }catch{
        Write-Host "[ERROR] Import Backup Logic 'Backup-GPO.ps1'" -ForegroundColor Red
        Sleep -Seconds 5
        Exit
    }

    #DOT Source

    try{
        $ExecuteTSTpath = $(Get-ScriptDirectory) + "\ChooseADOrganizationalUnit.ps1"
        . $ExecuteTSTpath
        Write-Host "[OK] Import GUI 'ChooseADOrganizationalUnit.ps1'" -ForegroundColor Green
    }catch{
        Write-Host "[ERROR] Import GUI 'ChooseADOrganizationalUnit.ps1'" -ForegroundColor Red
        Sleep -Seconds 5
        Exit
    }

    Write-Host ""

    #Select Root OU
    Write-Host "[INFO] Start OU Selection"
    Write-Host "[INFO] Select the OU you like to Backup. Sub OUs are included in Backup automatically!"

    $selRootOU = Choose-ADOrganizationalUnit
    if($selRootOU -eq $null){
            Write-Host "[INFO] Selection was cancelled by user. Closing Program."
            Sleep -Seconds 5
            Exit
    }

    $ChildOU = Get-ADOrganizationalUnit -SearchBase $selRootOU.DistinguishedName -SearchScope Subtree -Filter *

    #OU Test 
    Write-Host ""
    Write-Host "< SELECTED OU >"
    
    foreach ($source in $ChildOU.DistinguishedName){
        if([adsi]::Exists("LDAP://$source"))
        {
            Write-Host "[OK] $source" -ForegroundColor Green
        }
        else{
            Write-Host "[ERROR] $source" -ForegroundColor Red
            Sleep -Seconds 5
            Exit
        }
    }


    Write-Host ""
    Write-Host "< Start Backup >"


    #Start Backup
    Backup-GPOCB -RootOU $selRootOU.DistinguishedName

    Stop-Transcript

}


Get-Requirements