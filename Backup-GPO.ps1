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
   Backup-GPODN -RootOU OU=Computers_W10,OU=Tst_Environment,DC=intranet,DC=customer,DC=local
   This command will start a backup of all GPos linked beneath the above OU "Computers_W10"
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

function Get-ScriptName
{
	if ($hostinvocation -ne $null){$hostinvocation.MyCommand.Path}
	else{$script:MyInvocation.MyCommand.Path}
}

[string]$ScriptName = Get-ScriptName
[string]$ScriptDirectory = Split-Path $ScriptName  

$dateTimeNow = Get-Date -Format "yyyyMMdd_HHmmss"

function Backup-GPOCB
{           

    [CmdletBinding()]
    param(
    	[Parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true,HelpMessage="Example: OU=Computers_W10,OU=_NordLB-Tst,DC=kbk,DC=NordLB,DC=local")]
    	[Alias("Environment")]
    	[String]$RootOU
	)
    
    Begin
	{
        
        #START - Create Backup Folder
        $BackupFolder = $ScriptDirectory + "\GPO_Backups\" + $dateTimeNow
    
        try {        
            New-Item -ItemType directory -Path $BackupFolder | Out-Null
        }catch{
            Write-Host $_.Exception.Message
            Break
        }
        #END - Create Backup Folder

        #Get Child OUs
        $ChildOU = Get-ADOrganizationalUnit -SearchBase $RootOU -SearchScope Subtree -Filter *

        #Create Custom Object
        $DataInfo = New-Object System.Collections.Generic.List[object]
	}
      
    Process	{

        #Count GPO
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
						
						write-host ""
						write-host "[$i/$total]" $link.DisplayName  
						
						#Backup
						try{
							Backup-GPO -Name $link.DisplayName -Path $BackupFolder | Out-Null #-WhatIf
							$bkSuccess = "SUCCESSFUL"
							write-host "[OK] Backup" $link.DisplayName -ForegroundColor Green
						}catch{
						   $bkSuccess = "ERROR"
						   Write-Host "[ERROR] $_.Exception.Message"  -ForegroundColor Red
						   write-host "[ERROR]" $link.DisplayName -ForegroundColor Red
						}
						
						#HTML Report
						try{
							$DisplayName = $link.DisplayName
							Get-GPOReport -Name $link.DisplayName -ReportType Html -Path "$BackupFolder\$DisplayName.html"
							write-host "[OK] HTML Report $DisplayName.html" -ForegroundColor Green
						}catch{
							Write-Host "[ERROR] $_.Exception.Message"  -ForegroundColor Red
							write-host "[ERROR] HTML Report $DisplayName.html" -ForegroundColor Red
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
											} | select BackupStatus, Order, DisplayName, ID, OU, Linked, Enforced
						
						$DataInfo.add($obj)

					}
			}
		
		} else {
		
			Write-Host ""
			Write-Host "[WARNING] There are no linked GPOs to Backup" -ForegroundColor Yellow
			Write-Host ""
		
		}	
	}      
        
    
    
    End	{
    
        #Write to csv file.
        $DataInfo | ConvertTo-Csv -NoTypeInformation -Delimiter ";" | Out-File -Append -FilePath ($BackupFolder + "\" + $dateTimeNow + ".csv")
        $DataInfo | Sort-Object OU,Order | select BackupStatus, Order, DisplayName, ID, OU, Linked, Enforced | Out-GridView -Title "BACKUP Store: $BackupFolder" -Wait
        
    }

}