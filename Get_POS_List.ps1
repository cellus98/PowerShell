<# 
	.SYNOPSIS
    Obtain all SPOS and TPOS devices.
	
	.Description
	This script will pull device list from AD and then store the devices into specified text files.
	
    .NOTES
		Author: Marcellus Seamster Jr
	
	.EXAMPLE
	.\Get_POS_List.ps1

#>


#Region Get control of powershell UI
	$a = (Get-Host).UI.RawUI
	$a.WindowTitle = "Get POS List"
	$a.BackgroundColor = "Black"
	$a.ForegroundColor = "White"
	cls
#EndRegion

#Region Set startup variables and load modules

Import-Module VirtualMachineManager
Import-Module ActiveDirectory

#EndRegion Set startup variables and load modules

#Region Initialize Variables

	$FileDirectory = " "
	$Date = Get-Date -Format 'MM_dd_yyyy'
	$ComputerStaleDate = (Get-Date).AddDays(-90)
	$count = @()

	$SPOSFile = $FileDirectory + "SPOS List " + $Date + ".csv"
	$TPOSFile = $FileDirectory + "TPOS List " + $Date + ".csv"
	
	$SPOSOutput = @()
	$SPOSOutputCSV = @()
	$TPOSOutput = @()
	$TPOSOutputCSV = @()
	$SPOSNames = @()
	$TPOSNames = @()

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
	If (Test-Path $TPOSFile)   {
		Remove-Item $TPOSFile
	}
	If (Test-Path $SPOSFile)   {
		Remove-Item $SPOSFile
	}
}	

Function Get_SPOS_Devices   {	
	$SPOSNames = Get-ADComputer -Filter {(Name -like "* *") -and (passwordLastSet -ge $ComputerStaleDate)} -Properties * | Sort Name
	return $SPOSNames
}

Function Get_TPOS_Devices   {	
	$TPOSNames = Get-ADComputer -Filter {(Name -like "* *") -and (passwordLastSet -ge $ComputerStaleDate)} -Properties * | Sort Name
	return $TPOSNames
}

# Obtains type of control PowerShell window has
Function Test-Administrator   {  
	$user = [Security.Principal.WindowsIdentity]::GetCurrent();
	$Admin = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)  
	Verify_Admin $Admin
}

# Checks if PowerShell window opened as admin, if not stops script
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

Write-Host "                            Obtaining Devices" -ForegroundColor Yellow
$SPOSNames = Get_SPOS_Devices
$TPOSNames = Get_TPOS_Devices
Write-Host "                            Devices Obtained" -ForegroundColor Yellow


Write-host "                 Creating $SPOSFile ...." -ForegroundColor Yellow
Foreach ($SPOS in $SPOSNames)   {
	$SPOSOutput = " " | Select-Object Name , OU , AuthenticatedDC
	$SPOSOutput.Name = $SPOS.Name
	$SPOSOutput.OU = $SPOS.CanonicalName
	$SPOSOutput.AuthenticatedDC = $SPOS."msDS-AuthenticatedAtDC" | Out-String
	$SPOSOutputCSV += $SPOSOutput
}
$SPOSOutputCSV | Export-Csv -Path $SPOSFile -Delimiter "`t" -NoTypeInformation -Encoding unicode -Force
Write-host "                 $SPOSFile has been created ...." -ForegroundColor Green

Write-host "                 Creating $TPOSFile ...." -ForegroundColor Yellow
Foreach ($TPOS in $TPOSNames)   {
	$TPOSOutput = " " | Select-Object Name , OU , AuthenticatedDC
	$TPOSOutput.Name = $TPOS.Name
	$TPOSOutput.OU = $TPOS.CanonicalName
	$TPOSOutput.AuthenticatedDC = $TPOS."msDS-AuthenticatedAtDC" | Out-String
	$TPOSOutputCSV += $TPOSOutput
}
$TPOSOutputCSV | Export-Csv -Path $TPOSFile -Delimiter "`t" -NoTypeInformation -Encoding unicode -Force
Write-host "                 $TPOSFile has been created ...." -ForegroundColor Green


#EndRegion Main Processing
