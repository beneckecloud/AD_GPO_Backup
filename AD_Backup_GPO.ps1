<#
.Synopsis
   Start script including prerequirement checkt to backup Group Policy Objects (GPO) linked to Oranizational Units (OU)
.DESCRIPTION
   Starts Backup of all GPOs linked to an defined OU including all child OUs when selected in GUI. 
   In Addition to backups, HTML reports will be generated and saved.
   Backups and reports are located in the folder "01_GPO_Backups\%Date%_%Time%_GPO-BACKUP"   
   After the backup process a grid view will report the result of the process in a PowerShell GUI
   and will save all the displayed information to "01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\%Date%_%Time%.csv"
.EXAMPLE
   .\AD_Backup_GPO.ps1
   
   Will start a prerequirement check and a graphical user interface to select the OU you like to backup.
   After selecting the OU the script will start the backup by executing "Backup-GPOCB" with selected parameters from UI.
   
.OUTPUTS
   Backup Folder:	01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\
   
   GPO Backups:		01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\{2C1BAA1B-E393-409F-852B-1CF745F644B8}\
					01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\{2C370DBE-9789-4755-A573-2C97CAE32F1B}\
					01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\...\
   
   Reports:			01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\GPO_Name_1.html
					01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\GPO_Name_2.html
					01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\...
   
   Summary:			01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\%Date%_%Time%.csv 
   
   Log:				02_Log\%Date%_%Time%_All.log

   ------------------------------------------------------------------------------------------------------

   Zip Compression: When zip compression is enabled there will be only one compressed zip file as output

                    01_GPO_Backups\%Date%_%Time%_GPO-BACKUP.zip

.NOTES

   Author : Christopher Benecke
   Version: 1.1
   Date   : 31-05-2021

   New in this version: 

   * Added switch for zip Compression
   * Added switch for sub OU structures
   * New file structure

   --------------------------------------------------------------------------------------------

   Author : Christopher	Benecke
   Version: 1.0
   Date   : 03-08-2019
.LINK
   https://www.benecke.cloud
#>

function Get-ScriptDirectory{
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
	$Invocation.PSScriptRoot
}

function Test-OU{

    [CmdletBinding()]
    param(
    	[Parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true)]
    	[Alias("OU")]
    	[Array]$SubOU
	)
    
    Begin{}      
    Process	{

        foreach ($source in $SubOU.DistinguishedName){
            if([adsi]::Exists("LDAP://$source"))
            {
                Write-Host "[OK] $source" -ForegroundColor Green
            }
            else{
                Write-Host "[ERROR] $source" -ForegroundColor Red
                Start-Sleep -Seconds 5
                Exit
            }  
        }

	}
    End	{}

}

function Get-Requirements{
    
    Clear-Host

    Write-Host ""
    Write-Host "< Check Requirements >"
    Write-Host ""
    
    Start-Transcript -path "$(Get-ScriptDirectory)\02_Log\$(get-date -format 'yyyyMMdd_HHmmss')_All.log" -Force | Out-Null
    
    #PSVersion 5 Required
    $FullVersion = $PSVersionTable.PSVersion

    if( $PSVersionTable.PSVersion.Major -ge 5 ) {write-host "[OK] Powershell V$FullVersion" -ForegroundColor Green}
    elseif( $PSVersionTable.PSVersion.Major -le 4 ) {
        write-host "[ERROR] PowerShell V$FullVersion is not supported. At least V5.0 is required." -ForegroundColor Red
        Start-Sleep -Seconds 5
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
        Start-Sleep -Seconds 5
        Exit   
    }

    #DOT Source

    try{
        $ExecuteTSTpath = $(Get-ScriptDirectory) + "\03_Functions\Backup-GPO.ps1"
        . $ExecuteTSTpath
        Write-Host "[OK] Import Backup Logic 'Backup-GPO.ps1'" -ForegroundColor Green
    }catch{
        Write-Host "[ERROR] Import Backup Logic 'Backup-GPO.ps1'" -ForegroundColor Red
        Start-Sleep -Seconds 5
        Exit
    }

    #DOT Source

    try{
        $ExecuteTSTpath = $(Get-ScriptDirectory) + "\03_Functions\ChooseADOrganizationalUnit.ps1"
        . $ExecuteTSTpath
        Write-Host "[OK] Import GUI 'ChooseADOrganizationalUnit.ps1'" -ForegroundColor Green
    }catch{
        Write-Host "[ERROR] Import GUI 'ChooseADOrganizationalUnit.ps1'" -ForegroundColor Red
        Start-Sleep -Seconds 5
        Exit
    }

    Write-Host ""

    #Select Root OU
    Write-Host "[INFO] Start OU Selection"
    Write-Host "[INFO] Select the OU you like to Backup."

    $selRootOU = Choose-ADOrganizationalUnit -HideNewOUFeature

    #Check if OU was selected by User
    if($null -eq $selRootOU){
            Write-Host "[INFO] Selection was cancelled by user. Closing Program."
            Start-Sleep -Seconds 5
            Exit
    }

    #Include Sub OU structure?
    if($selRootOU -match "SubOU=YES"){
    
        $ChildOU = Get-ADOrganizationalUnit -SearchBase $selRootOU.DistinguishedName -SearchScope Subtree -Filter *

        #OU Test 
        Write-Host "[INFO] Include Sub OU structure selected"
        Write-Host ""
        Write-Host "< SELECTED OU >"
        
        Test-OU $ChildOU

        #Start Backup
        Write-Host ""
        Write-Host "< Start Backup >"

        #Start Backup
        if($selRootOU -match "ZIP=YES"){Backup-GPOCB -RootOU $selRootOU.DistinguishedName -Zip $true -Subs $true}else{Backup-GPOCB -RootOU $selRootOU.DistinguishedName -Zip $false -Subs $true}


        }else{

        #OU Test 
        Write-Host "[INFO] Exlude Sub OU structure selected"
        Write-Host ""
        Write-Host "< SELECTED OU >"

        Test-OU $selRootOU

        #Start Backup
        Write-Host ""
        Write-Host "< Start Backup >"

        #Start Backup
        if($selRootOU -match "ZIP=YES"){Backup-GPOCB -RootOU $selRootOU.DistinguishedName -Zip $true -Subs $false}else{Backup-GPOCB -RootOU $selRootOU.DistinguishedName -Zip $false -Subs $false}
    
    }

    Stop-Transcript

}


Get-Requirements