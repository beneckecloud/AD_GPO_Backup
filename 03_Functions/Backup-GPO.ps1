function Get-ScriptName
{
	if ($null -ne $hostinvocation){$hostinvocation.MyCommand.Path}
	else{$script:MyInvocation.MyCommand.Path}
}

[string]$ScriptName = Get-ScriptName
[string]$ScriptDirectory = Split-Path $ScriptName  

$dateTimeNow = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

function Backup-GPOCB {   
<#
.Synopsis
   Backup Group Policy Objects (GPO) linked to an Oranizational Unit (OU)
.DESCRIPTION
   Starts Backup of all GPOs linked to an defined OU including all child OUs. In Addition to
   the general backups, HTML reports will be generated and saved.   
   Backups and reports are located in the folder "01_GPO_Backups\%Date%_%Time%"   
   After the backup process a grid view will report the result of the process in a PowerShell GUI
   and will save all the displayed information to "01_GPO_Backups\%Date%_%Time%\%Date%_%Time%.csv"
.EXAMPLE
   Backup-GPODN -RootOU OU=Computers_W10,OU=Tst_Environment,DC=intranet,DC=customer,DC=local -Zip $false -Subs $true
   This command will start a backup of all GPO linked beneath the above OU "Computers_W10"
.EXAMPLE
   Backup-GPODN -RootOU OU=Computers_W10,OU=Tst_Environment,DC=intranet,DC=customer,DC=local -Zip $false -Subs $false
   This command will start a backup of all GPO linked to OU "Computers_W10"
.EXAMPLE
   Backup-GPODN -RootOU OU=Computers_W10,OU=Tst_Environment,DC=intranet,DC=customer,DC=local -Zip $true -Subs $false
   This command will start a backup of all GPO linked to OU "Computers_W10" and compress the result to one zip file.
.PARAMETER RootOU
   Define the RootOU you like to backup:
   OU=Computers_W10,OU=Tst_Environment,DC=intranet,DC=customer,DC=local
.PARAMETER Zip
   Activates zip compression when set to $true.
.PARAMETER Subs
   Starts backup of all GPOs linked to an defined OU including all child OUs when set to $true.
   Starts backup of all GPOs linked to an defined OU when set to $false.
.OUTPUTS
   Backup Folder:	01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\
   
   GPO Backups:		01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\{2C1BAA1B-E393-409F-852B-1CF745F644B8}\
					01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\{2C370DBE-9789-4755-A573-2C97CAE32F1B}\
					01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\...\
   
   Reports:			01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\GPO_Name_1.html
					01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\GPO_Name_2.html
					01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\...
   
   Summary:			01_GPO_Backups\%Date%_%Time%_GPO-BACKUP\%Date%_%Time%.csv 
   
   ------------------------------------------------------------------------------------------------------

   Zip Compression: When zip compression is enabled there will be only one compressed zip file as output

                    01_GPO_Backups\%Date%_%Time%_GPO-BACKUP.zip

.NOTES

   Author : Christopher Benecke
   Version: 1.1
   Date   : 31-05-2021

   New in this version: 

   * Added zip compression as function Compress-Path
   * Added switch for zip compression to Backup-GPOCB
   * Added switch for sub OU structures to Backup-GPOCB

   --------------------------------------------------------------------------------------------

   Author : Christopher	Benecke
   Version: 1.0
   Date   : 03-08-2019

.LINK
   https://www.benecke.cloud
#>

    [CmdletBinding()]
    param(
    	[Parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true,HelpMessage="Example: OU=Computers_W10,OU=Tst,DC=intranet,DC=local")]
    	[Alias("OU")]
    	[String]$RootOU,
        [Parameter(Position=1,Mandatory=$false)]
    	[Alias("Sub")]
    	[Boolean]$Subs,
        [Parameter(Position=2,Mandatory=$false)]
    	[Alias("Compression")]
    	[Boolean]$Zip
	)
    
    Begin
	{

        #START - Create Backup Folder
        $BackupFolder = $ScriptDirectory + "\01_GPO_Backups\" + $dateTimeNow + "_GPO-BACKUP"
    
        try {        
            New-Item -ItemType directory -Path $BackupFolder | Out-Null
            Write-Host "[OK] Create Backup Folder $BackupFolder" -ForegroundColor Green
        }catch{
            Write-Host "[ERROR] Create Backup Folder $BackupFolder" -ForegroundColor Red
            Write-Host $_.Exception.Message
            Break
        }
        #END - Create Backup Folder

        #Get Child OUs        
        switch ($Subs) {
            "$true" { $ChildOU = Get-ADOrganizationalUnit -SearchBase $RootOU -SearchScope Subtree -Filter *}
            "$false" { $ChildOU = Get-ADOrganizationalUnit -SearchBase $RootOU -SearchScope Base -Filter *}
            default { $ChildOU = Get-ADOrganizationalUnit -SearchBase $RootOU -SearchScope Subtree -Filter * }
        }        

        #Create Custom Object
        $DataInfo = New-Object System.Collections.Generic.List[object]
	}
      
    Process	{

        foreach ($target in $ChildOU.DistinguishedName){
            $linked = (Get-GPInheritance -Target $target).gpolinks
            foreach ($link in $linked){            
                $total++
            }

        }
        
		if($total){
		
			foreach ($target in $ChildOU.DistinguishedName){
			
				# Get the linked GPOs from target
				$linked = (Get-GPInheritance -Target $target).gpolinks

					foreach ($link in $linked){

						$i++

						$displayName = $link.DisplayName
						
						Write-Host ""
						Write-Host "[$i/$total]" $link.DisplayName  
						
						#Backup
						try{
							Backup-GPO -Name $link.DisplayName -Path $BackupFolder | Out-Null #-WhatIf
							$bkSuccess = "SUCCESSFUL"
							Write-Host "[OK] Backup" $link.DisplayName -ForegroundColor Green
						}catch{
						   $bkSuccess = "ERROR"
						   Write-Host "[ERROR] $_.Exception.Message"  -ForegroundColor Red
						   Write-Host "[ERROR]" $link.DisplayName -ForegroundColor Red
						}
						
						#HTML Report
						try{
							$DisplayName = $link.DisplayName
							Get-GPOReport -Name $link.DisplayName -ReportType Html -Path "$BackupFolder\$DisplayName.html"
							Write-Host "[OK] HTML Report $DisplayName.html" -ForegroundColor Green
						}catch{
							Write-Host "[ERROR] $_.Exception.Message"  -ForegroundColor Red
							Write-Host "[ERROR] HTML Report $DisplayName.html" -ForegroundColor Red
						}
						
						#Extend Object
						$Obj = New-Object Psobject -Property @{
											BackupStatus = $bkSuccess
											Order = $link.Order
											DisplayName = $link.DisplayName
											ID = $link.GpoId
											OU = $Target
											Linked = $link.Enabled
											Enforced = $link.Enforced
											} | Select-Object BackupStatus, Order, DisplayName, ID, OU, Linked, Enforced
						
						$DataInfo.add($obj)

					}
			}
		
		} else {
		
			Write-Host ""
			Write-Host "[WARNING] There are no linked GPOs in your selected OU structure" -ForegroundColor Yellow
			Write-Host ""
		
		}	
	}      
        
    End	{
    
        #Write to csv file.
        $DataInfo | ConvertTo-Csv -NoTypeInformation -Delimiter ";" | Out-File -Append -FilePath ($BackupFolder + "\" + $dateTimeNow + ".csv")

        ##If the zip option was configured, copy the working files to a temp dir and create a zip file, then remove the temp dir and move the zip file
        if($Zip){Compress-Path -zip $BackupFolder -dest "$ScriptDirectory\01_GPO_Backups\"}

        $DataInfo | Sort-Object OU,Order | Select-Object BackupStatus, Order, DisplayName, ID, OU, Linked, Enforced | Out-GridView -Title "BACKUP Store: $BackupFolder" -Wait        
        

    }

}

function Compress-Path{

    [CmdletBinding()]
    param(
    	[Parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true)]
    	[Alias("SourcePath")]
    	[String]$zip,
        [Parameter(Mandatory=$true)]
        [Alias("DestinationPath")]
        [String]$dest
	)
    
    Begin{}      
    Process	{

        Write-Host ""
        Write-Host "< Start ZIP compression >"

        Add-Type -AssemblyName "system.io.compression.filesystem"

        try {        
            New-Item -Path $zip\temp -ItemType Directory | Out-Null
            Write-Host "[OK] Create temp Folder $zip\temp" -ForegroundColor Green
        }catch{
            Write-Host "[ERROR] Create temp Folder $zip\temp" -ForegroundColor Red
            Write-Host $_.Exception.Message
            Break
        }

        try {        
            Get-ChildItem $zip -Exclude temp | Copy-Item -Destination $zip\temp -Recurse -Force #-Verbose
            Write-Host "[OK] Copy files to temp Folder $zip\temp" -ForegroundColor Green
        }catch{
            Write-Host "[ERROR] Copy files to temp Folder $zip\temp" -ForegroundColor Red
            Write-Host $_.Exception.Message
            Break
        }

        try {        
            [io.compression.zipfile]::CreateFromDirectory("$zip\temp", "$zip\{0:yyyy-MM-dd_HH-mm-ss}_GPO-BACKUP.zip" -f (Get-Date))
            Write-Host "[OK] Compress Folder $zip\temp" -ForegroundColor Green
        }catch{
            Write-Host "[ERROR] Compress Folder $zip\temp" -ForegroundColor Red
            Write-Host $_.Exception.Message
            Break
        }

        try {        
            Remove-Item $zip\temp -Recurse -Force #-Verbose
            Write-Host "[OK] Delete temp Folder $zip\temp" -ForegroundColor Green
        }catch{
            Write-Host "[ERROR] Delete temp Folder $zip\temp" -ForegroundColor Red
            Write-Host $_.Exception.Message
            Break
        }     
    
        try {        
            Move-Item $zip\*_GPO-BACKUP.zip $dest -Force #-Verbose
            Write-Host "[OK] Move *_GPO-BACKUP.zip to $dest" -ForegroundColor Green
        }catch{
            Write-Host "[ERROR] Move *_GPO-BACKUP.zip to $dest" -ForegroundColor Red
            Write-Host $_.Exception.Message
            Break
        }

        try {        
            Remove-Item -Path $zip -Recurse -Force #-Verbose
            Write-Host "[OK] Remove folder $zip" -ForegroundColor Green
        }catch{
            Write-Host "[ERROR] Remove folder $zip" -ForegroundColor Red
            Write-Host $_.Exception.Message
            Break
        }    
 

	}
    End	{}

}