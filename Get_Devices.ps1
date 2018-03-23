<# 
	.SYNOPSIS
    Obtain all devices.
	
	.Description
	This script will pull device list from AD and then store the devices into specified text files.
	
    .NOTES
		Author: Marcellus Seamster Jr
		
	.EXAMPLE
	.\Get_Devices.ps1

#>


#region Get control of powershell UI
$a = (Get-Host).UI.RawUI
$a.WindowTitle = "Gather Device List"
$a.BackgroundColor = "Black"
$a.ForegroundColor = "White"
cls
#EndRegion

#Region Set startup variables and load modules

Import-Module VirtualMachineManager
Import-Module ActiveDirectory

#EndRegion Set startup variables and load modules and XML file

#Region Initialize Variables

$FileDirectory = " "
$Date = Get-Date -Format 'MM_dd_yyyy'
$HostFile = $FileDirectory + "BOS Host List " + $Date + ".txt"
$VMFile = $FileDirectory + "BOS VM List " + $Date + ".txt"
$DCFile = $FileDirectory + "DC Device List " + $Date + ".txt"
#$AllFile = $FileDirectory + "All Device List " + $Date + ".txt"
$ComputerStaleDate = (Get-Date).AddDays(-90)

#EndRegion Initialize Variables



#Region Functions

# Check if the directory exists if not creates it
Function Check_Directory   {
	Write-Host "                 Checking if the Directory $FileDirectory on the local computer exists ..."
	if (!(Test-Path -Path $FileDirectory))   {
		Write-Host "                            Directory $FileDirectory on the local computer was NOT found." -ForegroundColor Red
		Write-Host "                            Creating Directory $FileDirectory on the local computer." -ForegroundColor Yellow
		New-Item $FileDirectory -type directory
	}
	else   {
		Write-Host "                 Directory $FileDirectory on the local computer was found." -ForegroundColor Magenta
	}
	Check_File
}

# If the file exists it removes the old file
Function Check_File   {
#<#
	If (Test-Path $HostFile)   {
		Remove-Item $HostFile
		New-Item $HostFile -type file
	}
	If (Test-Path $VMFile)   {
		Remove-Item $VMFile
		New-Item $VMFile -type file
	}
#>
	If (Test-Path $DCFile)   {
		Remove-Item $DCFile
		New-Item $DCFile -type file
	}
	Get_Devices
}	

Function Get_Devices   {	
	Write-Host "                            Obtaining Devices" -ForegroundColor Yellow
#	$HostNames = Get-ADComputer -Filter {(Name -like "*") -or (Name -like "*") -or (Name -like "*") -or (Name -like "*") -or (Name -like "*") -or (Name -like "**")} -Properties * | Sort Name
#	$VMNames = get-vm -vmmserver CO1RTPMVMMR2 | where-object { $_.name -ne $null -and $_.name -notlike "CO1*" -and $_.name -notlike "RTG*" -and $_.name -notlike "**" -and $_.name -notlike "*" -and $_.name -notlike "*" } | Sort-Object -Property Name -Unique
#	$DCNames = Get-ADComputer -Filter {(Name -like "*") -or (Name -like "*") -and (Name -notlike "**") -and (passwordLastSet -ge $ComputerStaleDate)} -Properties * | Sort Name
	#$DCNames = Get-ADComputer -Filter {(Name -like "*") -or (Name -like "*") -and (passwordLastSet -ge $ComputerStaleDate)} -Properties * | Sort Name
	#$DCNames = Get-ADComputer -Filter {(Name -like "*") -and (passwordLastSet -ge $ComputerStaleDate)} -Properties * | Sort Name
	$HostNames = Get-ADComputer -Filter {(Name -like "*") -or (Name -like "*") -or (Name -like "*") -or (Name -like "*") -or (Name -like "*") -and (passwordLastSet -ge $ComputerStaleDate)} -Properties * | Sort Name
	Write-Host "                            Devices Obtained" -ForegroundColor Yellow
	Create_Output
}

Function Create_Output   {

	Write-host "                 Creating $HostFile ...." -ForegroundColor Yellow
	Foreach ($HostN in $HostNames)   {
		$HostName = $HostN.CN
		$HostName | Out-File $HostFile -Append
		#$HostName | Out-File $AllFile -Append
	}
	Write-host "                 $HostFile has been created ...." -ForegroundColor Green
<#
	Write-host "                 Creating $VMFile ...." -ForegroundColor Yellow
	Foreach ($VMN in $VMNames)   {
		$VMName = $VMN.Name
		$VMName | Out-File $VMFile -Append
		#$VMName | Out-File $AllFile -Append
	}
	Write-host "                 $VMFile has been created ...." -ForegroundColor Green
	
	Write-host "                 Creating $DCFile ...." -ForegroundColor Yellow
	Foreach ($DCN in $DCNames)   {
		$Error.Clear()
		$Computer = $DCN.CN
		$WmiBios = Get-WmiObject Win32_Bios -ComputerName $Computer -ErrorAction SilentlyContinue
		IF (!($Error)) { 
			IF ($WmiBios.Version -like "*VRTUAL*")   {
				#Skip
			}
			else   { 
				$DCName = $DCN.Name
				$DCName | Out-File $DCFile -Append
				#$DCName | Out-File $AllFile -Append
			}
		}
	}		
	Write-host "                 $DCFile has been created ...." -ForegroundColor Green      
#>
}

Function Test-Administrator   {  
	$user = [Security.Principal.WindowsIdentity]::GetCurrent();
	$Admin = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)  
	Verify_Admin $Admin
}

Function Verify_Admin   {
	$Verify = $args[0]
	If (!$Verify)   {
		Write-Host "`tClose down this window and run script from Administrator prompt" -ForegroundColor Magenta
    	Exit
    }
}



#EndRegion Functions



#Region Main Processing

Test-Administrator
Check_Directory




#EndRegion Main Processing