<# 
	.SYNOPSIS
    Obtain all devices.
	
	.Description
	This script will pull device list from AD and then store the devices into specified text files.
	
    .NOTES
		Author: Marcellus Seamster Jr
	
	.EXAMPLE
	.\Get_VM_List.ps1

#>


#Region Get control of powershell UI
	$a = (Get-Host).UI.RawUI
	$a.WindowTitle = "Get VM List"
	$a.BackgroundColor = "Black"
	$a.ForegroundColor = "White"
	cls
#EndRegion

#Region Set startup variables and load modules

Import-Module VirtualMachineManager
Import-Module ActiveDirectory

#EndRegion Set startup variables and load modules and XML file

#Region Initialize Variables

	$FileDirectory = "filename"
	$Date = Get-Date -Format 'MM_dd_yyyy'
	$HBIFile = $FileDirectory + "HBI List " + $Date + ".csv"
	$MBIFile = $FileDirectory + "MBI List " + $Date + ".csv"
	$HK1File = $FileDirectory + "HK1 List " + $Date + ".csv"
		
	$HBIoutput = @()
	$HBIoutputCSV = @()
	$MBIoutput = @()
	$MBIoutputCSV = @()
	$HK1output = @()
	$HK1outputCSV = @()	

	$HBINames = @()
	$MBINames = @()
	$HK1Names = @()


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
	If (Test-Path $HBIFile)   {
		Remove-Item $HBIFile
	}
	If (Test-Path $MBIFile)   {
		Remove-Item $MBIFile
	}
	If (Test-Path $HK1File)   {
		Remove-Item $HK1File
	}
}	

Function Get_HBI   {
#	$HBINames = get-vm -vmmserver "serverName"| where-object { $_.HostGroupPath -like " *" } | Sort-Object -Property Name -Unique
	$HBINames = get-vm -vmmserver "serverName"| where-object { $_.HostGroupPath -like " *" } | Sort-Object -Property Name -Unique
	return $HBINames
}

Function Get_MBI   {	
#	$MBINames = get-vm -vmmserver "serverName"| where-object { $_.HostGroupPath -like " *" } | Sort-Object -Property Name -Unique
	$MBINames = get-vm -vmmserver "serverName"| where-object { $_.HostGroupPath -like " *" } | Sort-Object -Property Name -Unique
	return $MBINames
}

Function Get_HK1   {	
	$HK1Names = get-vm -vmmserver "serverName"| where-object { $_.HostGroupPath -like " *" } | Sort-Object -Property Name -Unique	
	return $HK1Names
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
	
	Write-host "                 Creating $HBIFile ...." -ForegroundColor Yellow
	$HBINames = Get_HBI
	Foreach ($HBI in $HBINames)   {
		$HBIoutput = " " | Select-Object VM , Host
		$HBIoutput.VM = $HBI.Name
		$HBIoutput.Host = $HBI.HostName.split(".")[0]
		$HBIoutputCSV += $HBIoutput		
	}
	$HBIoutputCSV | Export-Csv -Path $HBIFile -Delimiter "`t" -NoTypeInformation -Encoding unicode -Force
	Write-host "                 $HBIFile has been created ...." -ForegroundColor Green
	$MBINames = Get_MBI
	Write-host "                 Creating $MBIFile ...." -ForegroundColor Yellow	
	Foreach ($MBI in $MBINames)   {
		$MBIoutput = " " | Select-Object VM , Host
		$MBIoutput.VM = $MBI.Name
		$MBIoutput.Host = $MBI.HostName.split(".")[0]
		$MBIoutputCSV += $MBIoutput
	}
	$MBIoutputCSV | Export-Csv -Path $MBIFile -Delimiter "`t" -NoTypeInformation -Encoding unicode -Force	
	Write-host "                 $MBIFile has been created ...." -ForegroundColor Green	
	$HK1Names = Get_HK1
	Write-host "                 Creating $HK1File ...." -ForegroundColor Yellow
	Foreach ($HK1 in $HK1Names)   {
			$HK1output = " " | Select-Object VM , Host
		$HK1output.VM = $HK1.Name
		$HK1output.Host = $HK1.HostName.split(".")[0]
		$HK1outputCSV += $HK1output
	}	
	Write-host "                 $HK1File has been created ...." -ForegroundColor Green
	$HK1outputCSV | Export-Csv -Path $HK1File -Delimiter "`t" -NoTypeInformation -Encoding unicode -Force

#EndRegion Main Processing