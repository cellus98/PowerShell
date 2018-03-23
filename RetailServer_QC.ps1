<# 
	.SYNOPSIS
	    QC Retail Servers based on data obtained from a CSV file.
	.Description
		This script will QC Retail Server devices obtained from a CSV file.
			The CSV will obtain a list of devices along with items to be verified.
	.PARAMETER inputCSVP
		The functional definition of which CSV file to use for QC.
	.PARAMETER Platform
		The functional definition of where the device is located (On-Premise or Azure).
	.INPUTS
		None. You cannot pipe objects to RetailServer_QC.ps1.
	.OUTPUTS
		Console and excel file
    .NOTES
		Author: Marcellus Seamster Jr
	.EXAMPLE
		.\RetailServer_QC.ps1 -inputCSVP .\test-RS-QC.csv -Platform Azure
#>

Param( $inputCSVP, [ValidateSet('On-Premise','Azure')]$Platform)
	
#Region Get control of powershell UI
	$a = (Get-Host).UI.RawUI
	$a.WindowTitle = "Retail Server QC"
	$a.BackgroundColor = "Black"
	$a.ForegroundColor = "White"
	cls
#Endregion Get control of powershell UI	

#Region Variable declartion
	[int]$totalProc = "0"
	$errorComment = @()
	$TempSession = @()
	$PatchLevel = @()
	$AzureHost = @()
	$PremiseServerName = @()
	$OnPremiseTimeZone = @()
	$AzureTimeZone = @()
	$totalMem = @()
	$cleanMemMB = @()
	$maxMem = @()
	$minMem = @()
	$colItems = @()
	$DeviceID = @()
	$DriveSize = @()
	$DiskSize = @()
	$DriveLetter = @()
	$MDOPValue = @()
	$BackupCompressionValue = @()
	$i = 0
	$servers = @()
	
#Endregion Variable declaration

#Region Set Startup Variables
	$version = "1.0"
	$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()
	$location = "fileName"
	$Date = Get-Date
	$ErrorActionPreference = "SilentlyContinue"
	$logdate = "{0:ddmmyyyy}" -f (Get-Date)
	$inputCSV = Import-Csv $inputCSVP
	$User = $env:username
	$Domain = $env:userdomain
	$UserName = $Domain + "\" + $User
	$cred = Get-Credential -Credential $UserName
	$Line = "================================================================================"

	If(Test-Path $inputCSVP)   {
		Split-Path $inputCSVp -Leaf | Out-Null
		$fileName = Split-Path $inputCSVp -Leaf
		$StoreCode = $fileName.Substring(0,3)
	}
	Else   {
		Write-Warning "Input CSV not found.  VerIfy path then retry."
		Write-Warning "Exiting Script"
		Exit
	}	

	$WarningColor = (Get-Host).PrivateData
	$WarningColor.WarningBackgroundColor = "yellow"
	$WarningColor.WarningForegroundColor = "blue"

#Endregion Set Startup Variables
	
#Region Functions

Function Write-Header   {
	param ([int]$count)
	
	Write-Host $Line -ForegroundColor Green
	Write-Host "`tRetail Server QC Version $version script Started at " -NoNewline
	Write-Host $Date -ForegroundColor Cyan
	Write-Host
	Write-Host "`t`t`t`tScript Parameters:"
	Write-Host "`t`t`tUser:" -NoNewline
	Write-Host "`t`t$UserName" -ForegroundColor Cyan
	Write-Host "`t`t`tPlatform:" -NoNewline
	Write-Host "`t$Platform" -ForegroundColor Cyan
	Write-Host "`t`t`tServer Count:" -NoNewline
	Write-Host "`t$count" -ForegroundColor Cyan
	Write-Host $Line -ForegroundColor Green
	Write-Host
}

Function Write-Footer   {
	param ([int]$count)
	
	Write-Host
	Write-Host $Line -ForegroundColor Green
	Write-Host "`t`t`tScript Completed at " -NoNewline
	Write-Host "`t$(get-date)" -ForegroundColor Cyan	
	$Message = "Total Elapsed Time:{0:00}:{1:00}:{2:00}" -F $ElapsedTime.Elapsed.Hours, $ElapsedTime.Elapsed.Minutes, $ElapsedTime.Elapsed.Seconds	
	$Message = $Message.Replace("Total Elapsed Time:","")
	Write-Host "`t`t`tTotal Elapsed Time:" -NoNewline
	Write-Host "`t$Message" -ForegroundColor Cyan
	Write-Host "`t`t`tUser:" -NoNewline
	Write-Host "`t`t`t$UserName" -ForegroundColor Cyan
	$Message = "`tUser: $UserName"
	Write-Host "`t`t`tPlatform:" -NoNewline
	Write-Host "`t`t$Platform" -ForegroundColor Cyan
	Write-Host "`t`t`tServer Count:" -NoNewline
	Write-Host "`t`t$count" -ForegroundColor Cyan
	Write-Host $Line -ForegroundColor Green
	Write-Host
}

Function Test-Administrator   {  
    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
    $Admin = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)  
    VerIfy_Admin $Admin
}

Function VerIfy_Admin   {
    $VerIfy = $args[0]
    If (!$VerIfy)   {
        Write-Host $Line -ForegroundColor Yellow
		Write-Host "`tClose down this window and run script from Administrator prompt" -ForegroundColor Magenta
        Write-Host $Line -ForegroundColor Yellow
		Write-Host
	 	Exit
    }
}

Function Get_Password   {
	$bstr = "default password"
	$Password =  "password"
	return $Password
}

Function Create-ExcelQC   {
##### 										Excel headers 										#####
	#"Azure_Computer", "On-Premise_Computer", "Processors", "Max_Memory", "Min_Memory", "Time_Zone"
	#"A_Drive", "B_Drive", "X_Drive", "W_Drive", "J_Drive", "K_Drive", "P_Drive"
	#"SQL_Version", "MDOP", "BackupCompression", "Recovery"
	#"M_Size", "M_Growth", "Mlog_Size", "Mlog_Growth"
	#"MSData_Size", "MSData_Growth", "MS_Size", "MS_Growth"
	#"M2dev_Size", "M2dev_Growth", "M2log_Size", "M2log_Growth", "Temp_Size"
	#"Temp_Growth", "Temp_Size", "Temp_Growth", "Temp1_Size", "Temp1_Growth", "Temp2_Size", "Temp2_Growth", "Temp3_Size"
	#"Temp3_Growth", "Temp4_Size", "Temp4_Growth", "Temp5_Size", "Temp5_Growth", "Temp6_Size", "Temp6_Growth", "Temp7_Size"
	#"Temp7_Growth", "Temp8_Size", "Temp8_Filegrowth", "Temp9_Size", "Temp9_Filegrowth", "Temp10_Size", "Temp10_Filegrowth"
	#"Temp11_Size", "Temp11_Filegrowth", "Temp12_Size", "Temp12_Filegrowth", "Temp13_Size", "Temp13_Filegrowth"
	#"Temp14_Size", "Temp14_Filegrowth", "Temp15_Size", "Temp15_Filegrowth"
	
	#New Excel Application
	$script:Excel = New-Object -Com Excel.Application
	$script:Excel.visible = $True 

	$script:Excel = $Excel.Workbooks.Add()
	$script:Sheet1 = $Excel.Worksheets.Item(1)
	$script:Sheet1.Name = $StoreCode + " VM QC"

	#Create Heading for General Sheet
	$Sheet1.Cells.Item(1,1) = "Azure Computer"
	$Sheet1.Cells.Item(1,2) = "On-Premise Computer"
	$Sheet1.Cells.Item(1,3) = "Processors"
	$Sheet1.Cells.Item(1,4) = "Maximum Memory"
	$Sheet1.Cells.Item(1,5) = "Minimum Memory"
	$Sheet1.Cells.Item(1,6) = "Time Zone"
	$Sheet1.Cells.Item(1,7) = "A Drive"
	$Sheet1.Cells.Item(1,8) = "B Drive"
	$Sheet1.Cells.Item(1,9) = "X Drive"
	$Sheet1.Cells.Item(1,10) = "W Drive"
	$Sheet1.Cells.Item(1,11) = "J Drive"
	$Sheet1.Cells.Item(1,12) = "K Drive"
	$Sheet1.Cells.Item(1,13) = "P Drive"
	$Sheet1.Cells.Item(1,14) = "SQL Version"
	$Sheet1.Cells.Item(1,15) = "MDOP"
	$Sheet1.Cells.Item(1,16) = "BackUp Compression"
	$Sheet1.Cells.Item(1,17) = "M2 Recovery"
	$Sheet1.Cells.Item(1,18) = "M Size"
	$Sheet1.Cells.Item(1,19) = "M Filegrowth"
	$Sheet1.Cells.Item(1,20) = "MLog Size"
	$Sheet1.Cells.Item(1,21) = "MLog Filegrowth"
	$Sheet1.Cells.Item(1,22) = "MS Size"
	$Sheet1.Cells.Item(1,23) = "MS Filegrowth"
	$Sheet1.Cells.Item(1,24) = "MS Size"
	$Sheet1.Cells.Item(1,25) = "MS Filegrowth"
	$Sheet1.Cells.Item(1,26) = "M2 Size"
	$Sheet1.Cells.Item(1,27) = "M2 Filegrowth"
	$Sheet1.Cells.Item(1,28) = "M2Log Size"
	$Sheet1.Cells.Item(1,29) = "M2Log Filegrowth"
	$Sheet1.Cells.Item(1,30) = "Temp Size"
	$Sheet1.Cells.Item(1,31) = "Temp Filegrowth"
	$Sheet1.Cells.Item(1,32) = "Temp Size"
	$Sheet1.Cells.Item(1,33) = "Temp Filegrowth"
	$Sheet1.Cells.Item(1,34) = "Temp1 Size"
	$Sheet1.Cells.Item(1,35) = "Temp1 Filegrowth"
	$Sheet1.Cells.Item(1,36) = "Temp2 Size"
	$Sheet1.Cells.Item(1,37) = "Temp2 Filegrowth"
	$Sheet1.Cells.Item(1,38) = "Temp3 Size"
	$Sheet1.Cells.Item(1,39) = "Temp3 Filegrowth"
	$Sheet1.Cells.Item(1,40) = "Temp4 Size"
	$Sheet1.Cells.Item(1,41) = "Temp4 Filegrowth"
	$Sheet1.Cells.Item(1,42) = "Temp5 Size"
	$Sheet1.Cells.Item(1,43) = "Temp5 Filegrowth"
	$Sheet1.Cells.Item(1,44) = "Temp6 Size"
	$Sheet1.Cells.Item(1,45) = "Temp6 Filegrowth"
	$Sheet1.Cells.Item(1,46) = "Temp7 Size"
	$Sheet1.Cells.Item(1,47) = "Temp7 Filegrowth"
	$Sheet1.Cells.Item(1,48) = "Temp8 Size"
	$Sheet1.Cells.Item(1,49) = "Temp8 Filegrowth"
	$Sheet1.Cells.Item(1,50) = "Temp9 Size"
	$Sheet1.Cells.Item(1,51) = "Temp9 Filegrowth"
	$Sheet1.Cells.Item(1,52) = "Temp10 Size"
	$Sheet1.Cells.Item(1,53) = "Temp10 Filegrowth"
	$Sheet1.Cells.Item(1,54) = "Temp11 Size"
	$Sheet1.Cells.Item(1,55) = "Temp11 Filegrowth"
	$Sheet1.Cells.Item(1,56) = "Temp12 Size"
	$Sheet1.Cells.Item(1,57) = "Temp12 Filegrowth"
	$Sheet1.Cells.Item(1,58) = "Temp13 Size"
	$Sheet1.Cells.Item(1,59) = "Temp13 Filegrowth"
	$Sheet1.Cells.Item(1,60) = "Temp14 Size"
	$Sheet1.Cells.Item(1,61) = "Temp14 Filegrowth"
	$Sheet1.Cells.Item(1,62) = "Temp15 Size"
	$Sheet1.Cells.Item(1,63) = "Temp15 Filegrowth"

	#add header color
	$colSheets = $script:Sheet1
	Foreach ($colorItem in $colSheets)   {
		$intRow = 2
		$intRowCPU = 2
		$intRowMem = 2
		$intRowDisk = 2
		$intRowNet = 2
		$script:WorkBook = $colorItem.UsedRange
		$script:WorkBook.Interior.ColorIndex = 11
		$script:WorkBook.Font.ColorIndex = 2
		$script:WorkBook.Font.Bold = $True
	}
	$script:WorkBook = $colorItem.UsedRange															
	$script:WorkBook.EntireColumn.AutoFit() | Out-Null
}

Function Get-Proccessors   {
	param ([String] $Computer, [String] $FQDN, [int]$IntRow)
	
	[int]$totalProc = (Get-WmiObject -class "Win32_computersystem" -ComputerName $Computer).numberoflogicalprocessors	
	[String]$Processors = $inputVM[0].Processors
	$errorComment = "Expected Value: `n" + $Processors

	If($totalProc -ne $inputVM[0].Processors)   {
		$Sheet1.Cells.Item($intRow, 3) = $totalProc
		$Sheet1.Cells.item($intRow, 3).font.colorIndex=3
		$Sheet1.Cells.item($intRow, 3).font.bold=$true
		$Sheet1.Cells.Item($intRow, 3).AddComment($errorComment) | Out-Null
		Write-Host "`t`tNumber of Incorrect Processors: " -ForegroundColor White -NoNewline
		Write-Host $totalProc -ForegroundColor Red
	}
	Else   {
		$Sheet1.Cells.Item($intRow, 3) = $totalProc
		$Sheet1.Cells.item($intRow, 3).font.colorIndex=10
		
		Write-Host "`t`tNumber of Processors: " -ForegroundColor White -NoNewline
		Write-Host $totalProc -ForegroundColor Green
	}
}

Function Get-TimeZone   {
	param ([String] $Computer, [String] $FQDN, [int]$IntRow)

	[String]$OnPremiseServerName = $inputVM[0].On_Premise_Computer
	$PremiseServerName = ("{0}.domain" -f $OnPremiseServerName)
	$OnPremiseTimeZone = Invoke-Command -Computer $OnPremiseServerName -Command {tzutil /g}
	
	$AzureTimeZone = Invoke-Command -Computer $FQDN -Command {tzutil /g}
	
	$errorComment = "Expected Value: `n" + $OnPremiseTimeZone
		
	If($AzureTimeZone -ne $OnPremiseTimeZone)   {
		$Sheet1.Cells.Item($intRow, 6) = $AzureTimeZone
		$Sheet1.Cells.item($intRow, 6).font.colorIndex=3
		$Sheet1.Cells.item($intRow, 6).font.bold=$true
		$Sheet1.Cells.Item($intRow, 6).AddComment($errorComment) | Out-Null
		Write-Host "`t`t`tIncorrect Time Zone: " -ForegroundColor White -NoNewline
		Write-Host $AzureTimeZone -ForegroundColor Red
	}
	Else   {    
		$Sheet1.Cells.Item($intRow, 6) = $AzureTimeZone
		$Sheet1.Cells.item($intRow, 6).font.colorIndex=10
		$Message = "`tTime Zone: $AzureTimeZone"
		Write-Host "`t`tTime Zone: " -ForegroundColor White -NoNewline
		Write-Host $AzureTimeZone -ForegroundColor Green
	}
}

Function Get-Memory   {
	param ([String] $Computer, [String] $FQDN, [int]$IntRow)
	
	$totalMem = ((Get-WmiObject -class "Win32_OperatingSystem" -computername $Computer).TotalVisibleMemorySize)
	$cleanMemMB = ([math]::round($totalMem / 1024, 0))
	
	Write-Host "`t`tRounded memory in MB: " -NoNewline
	Write-Host $cleanMemMB -ForegroundColor Cyan
	If ($cleanMemMB -gt 24400)   {
		$maxMem = ([math]::Round($cleanMemMB - 8192, 0))
		$minMem = ([math]::Round($maxMem * .8, 0))
		[String]$ExpectedMaxMemory = $inputVM[0].Max_Memory
		$errorComment = "Expected Value: `n" + $ExpectedMaxMemory		
		If($maxMem -ne $inputVM[0].Max_Memory)   {
			$Sheet1.Cells.Item($intRow, 4) = $maxMem
			$Sheet1.Cells.item($intRow, 4).font.colorIndex=3
			$Sheet1.Cells.item($intRow, 4).font.bold=$true
			$Sheet1.Cells.Item($intRow, 4).AddComment($errorComment) | Out-Null

			Write-Host "`t`t`tIncorrect Maximum Memory: " -ForegroundColor White -NoNewline
			Write-Host $maxMem -ForegroundColor Red -NoNewline
			Write-Host " MB" -ForegroundColor White
		}
		Else   {    
			$Sheet1.Cells.Item($intRow, 4) = $maxMem
			$Sheet1.Cells.item($intRow, 4).font.colorIndex=10
			Write-Host "`t`tMaximum Memory: " -ForegroundColor White -NoNewline
			Write-Host $maxMem -ForegroundColor Green -NoNewline
			Write-Host " MB" -ForegroundColor White
		}
		[String]$ExpectedMinMemory = $inputVM[0].Min_Memory	
		$errorComment = "Expected Value: `n" + $ExpectedMinMemory
		If($minMem -ne $inputVM[0].Min_Memory)   {
			$Sheet1.Cells.Item($intRow, 5) = $minMem
			$Sheet1.Cells.item($intRow, 5).font.colorIndex=3
			$Sheet1.Cells.item($intRow, 5).font.bold=$true
			$Sheet1.Cells.Item($intRow, 5).AddComment($errorComment) | Out-Null
			Write-Host "`t`t`tIncorrect Minimum Memory: " -ForegroundColor White -NoNewline
			Write-Host $minMem -ForegroundColor Red -NoNewline
			Write-Host " MB" -ForegroundColor White
		}
		Else   {    
			$Sheet1.Cells.Item($intRow, 5) = $minMem
			$Sheet1.Cells.item($intRow, 5).font.colorIndex=10
			Write-Host "`t`tMinimum Memory: " -ForegroundColor White -NoNewline
			Write-Host $minMem -ForegroundColor Green -NoNewline
			Write-Host " MB" -ForegroundColor White
		}
	}
	Else   {
		$maxMem = ([math]::Round($cleanMemMB * .8, 0))
		$minMem = ([math]::Round($maxMem * .8, 0))
		[String]$ExpectedMaxMemory = $inputVM[0].Max_Memory
		$errorComment = "Expected Value: `n" + $ExpectedMaxMemory
		If($maxMem -ne $ExpectedMaxMemory)   {
			$Sheet1.Cells.Item($intRow, 4) = $maxMem
			$Sheet1.Cells.item($intRow, 4).font.colorIndex=3
			$Sheet1.Cells.item($intRow, 4).font.bold=$true
			$Sheet1.Cells.Item($intRow, 4).AddComment($errorComment) | Out-Null
			Write-Host "`t`t`tIncorrect Maximum Memory: " -ForegroundColor White -NoNewline
			Write-Host $maxMem -ForegroundColor Red -NoNewline
			Write-Host " MB" -ForegroundColor White
		}
		Else   {    
			$Sheet1.Cells.Item($intRow, 4) = $maxMem
			$Sheet1.Cells.item($intRow, 4).font.colorIndex=10
			Write-Host "`t`tMaximum Memory: " -ForegroundColor White -NoNewline
			Write-Host $maxMem -ForegroundColor Green -NoNewline
			Write-Host " MB" -ForegroundColor White
		}
		[String]$ExpectedMinMemory = $inputVM[0].Min_Memory	
		$errorComment = "Expected Value: `n" + $ExpectedMinMemory
		If($minMem -lt $inputVM[0].Phy_Memory)   {
			$Sheet1.Cells.Item($intRow, 5) = $minMem
			$Sheet1.Cells.item($intRow, 5).font.colorIndex=3
			$Sheet1.Cells.item($intRow, 5).font.bold=$true
			$Sheet1.Cells.Item($intRow, 5).AddComment($errorComment) | Out-Null
			Write-Host "`t`t`tIncorrect Minimum Memory: " -ForegroundColor White -NoNewline
			Write-Host $minMem -ForegroundColor Red -NoNewline
			Write-Host " MB" -ForegroundColor White			
		}
		Else   {    
			$Sheet1.Cells.Item($intRow, 5) = $minMem
			$Sheet1.Cells.item($intRow, 5).font.colorIndex=10
			Write-Host "`t`tMinimum Memory: " -ForegroundColor White -NoNewline
			Write-Host $minMem -ForegroundColor Green -NoNewline
			Write-Host " MB" -ForegroundColor White
		}
	}
}

Function Get-Disk {
	param ([String] $Computer, [String] $FQDN, [int]$IntRow)
	
	#get new VM Drive configuration
	$colItems = get-WmiObject win32_logicaldisk -Computername $Computer

	If ($Platform -eq "On-Premise")   {
		If($colItems.DeviceID -contains "A:")   {
			$DeviceID = $colItems | Where DeviceID -eq "A:"
			$DriveSize = $DeviceID.Size
			$DiskSize = [Math]::Round($DriveSize/1GB)
			$DriveLetter = $DeviceID.DeviceID
			If($DiskSize -ne $inputVM[0].A_Drive)   {
				$errorComment = "Expected Value: `n" + $inputVM[0].A_Drive
				$Sheet1.Cells.Item($intRow, 7) = $DiskSize
				$Sheet1.Cells.item($intRow, 7).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 7).font.bold=$true
				$Sheet1.Cells.Item($intRow, 7).AddComment($errorComment) | Out-Null
				Write-Host "`t`t`tIncorrect Drive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Red -NoNewline
				Write-Host " GB" -ForegroundColor White
				}
			Else   {
				$Sheet1.Cells.Item($intRow, 7) = $DiskSize
				$Sheet1.Cells.item($intRow, 7).font.colorIndex=10
				Write-Host "`t`tDrive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Green -NoNewline
				Write-Host " GB" -ForegroundColor Green
			}
		}
		
		ElseIf($colItems.DeviceID -notcontains "A:")   {
			$errorComment = "Expected Value: `n" + $inputVM[0].A_Drive
			If($inputVM[0].A_Drive -eq ".")   {
			
			}
			Else   {
				$DeviceID = $colItems | Where DeviceID -eq "A:"
				$DriveSize = $DeviceID.Size
				$DiskSize = [Math]::Round($DriveSize/1GB)
				$Sheet1.Cells.Item($intRow, 7) = $DiskSize
				$Sheet1.Cells.item($intRow, 7).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 7).font.bold=$true
				$Sheet1.Cells.Item($intRow, 7).AddComment($errorComment) | Out-Null
				}
		}
			
		If($colItems.DeviceID -contains "B:")   {
			$DeviceID = $colItems | Where DeviceID -eq "B:"
			$DriveSize = $DeviceID.Size
			$DiskSize = [Math]::Round($DriveSize/1GB)
			$DriveLetter = $DeviceID.DeviceID
			If($DiskSize -ne $inputVM[0].B_Drive)   {
				$errorComment = "Expected Value: `n" + $inputVM[0].B_Drive
				$Sheet1.Cells.Item($intRow, 8) = $DiskSize
				$Sheet1.Cells.item($intRow, 8).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 8).font.bold=$true
				$Sheet1.Cells.Item($intRow, 8).AddComment($errorComment) | Out-Null
				Write-Host "`t`t`tIncorrect Drive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Red -NoNewline
				Write-Host " GB" -ForegroundColor White
			}
			Else   {
				$Sheet1.Cells.Item($intRow, 8) = $DiskSize
				$Sheet1.Cells.item($intRow, 8).font.colorIndex=10
				Write-Host "`t`tDrive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Green -NoNewline
				Write-Host " GB" -ForegroundColor Green
			}
		}
		
		ElseIf($colItems.DeviceID -notcontains "B:")   {
			$errorComment = "Expected Value: `n" + $inputVM[0].B_Drive
			If($inputVM[0].B_Drive -eq ".")   {
					
			}
			Else   {
				$DeviceID = $colItems | Where DeviceID -eq "B:"
				$DriveSize = $DeviceID.Size
				$DiskSize = [Math]::Round($DriveSize/1GB)
				$Sheet1.Cells.Item($intRow, 8) = $DiskSize
				$Sheet1.Cells.item($intRow, 8).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 8).font.bold=$true
				$Sheet1.Cells.Item($intRow, 8).AddComment($errorComment) | Out-Null
			}
		}
			
		If($colItems.DeviceID -contains "P:")   {
			$DeviceID = $colItems | Where DeviceID -eq "P:"
			$DriveSize = $DeviceID.Size
			$DiskSize = [Math]::Round($DriveSize/1GB)
			$DriveLetter = $DeviceID.DeviceID
			If($DiskSize -ne $inputVM[0].P_Drive)   {
				$errorComment = "Expected Value: `n" + $inputVM[0].B_Drive
				$Sheet1.Cells.Item($intRow, 13) = $DiskSize
				$Sheet1.Cells.item($intRow, 13).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 13).font.bold=$true
				$Sheet1.Cells.Item($intRow, 13).AddComment($errorComment) | Out-Null
				Write-Host "`t`t`tIncorrect Drive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Red -NoNewline
				Write-Host " GB" -ForegroundColor White
			}
			Else   {
				$Sheet1.Cells.Item($intRow, 13) = $DiskSize
				$Sheet1.Cells.item($intRow, 13).font.colorIndex=10
				Write-Host "`t`tDrive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Green -NoNewline
				Write-Host " GB" -ForegroundColor Green
			}
		}
		
		ElseIf($colItems.DeviceID -notcontains "P:")   {
			$errorComment = "Expected Value: `n" + $inputVM[0].P_Drive
			If($inputVM[0].P_Drive -eq ".")   {
					
			}
			Else   {
				$DeviceID = $colItems | Where DeviceID -eq "P:"
				$DriveSize = $DeviceID.Size
				$DiskSize = [Math]::Round($DriveSize/1GB)
				$Sheet1.Cells.Item($intRow, 13) = $DiskSize
				$Sheet1.Cells.item($intRow, 13).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 13).font.bold=$true
				$Sheet1.Cells.Item($intRow, 13).AddComment($errorComment) | Out-Null
			}
		}			

		If($colItems.DeviceID -contains "W:")   {
			$DeviceID = $colItems | Where DeviceID -eq "W:"
			$DriveSize = $DeviceID.Size
			$DiskSize = [Math]::Round($DriveSize/1GB)
			$DriveLetter = $DeviceID.DeviceID
			If($DiskSize -ne $inputVM[0].W_Drive)   {
				$errorComment = "Expected Value: `n" + $inputVM[0].W_Drive
				$Sheet1.Cells.Item($intRow, 10) = $DiskSize
				$Sheet1.Cells.item($intRow, 10).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 10).font.bold=$true
				$Sheet1.Cells.Item($intRow, 10).AddComment($errorComment) | Out-Null
				Write-Host "`t`t`tIncorrect Drive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Red -NoNewline
				Write-Host " GB" -ForegroundColor White
			}
			Else   {
				$Sheet1.Cells.Item($intRow, 10) = $DiskSize
				$Sheet1.Cells.item($intRow, 10).font.colorIndex=10
				Write-Host "`t`tDrive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Green -NoNewline
				Write-Host " GB" -ForegroundColor Green
			}					
		}
		
		ElseIf($colItems.DeviceID -notcontains "W:")   {
			$errorComment = "Expected Value: `n" + $inputVM[0].W_Drive
			If($inputVM[0].W_Drive -eq ".")   {
				
			}
			Else   {
				$DeviceID = $colItems | Where DeviceID -eq "W:"
				$DriveSize = $DeviceID.Size
				$DiskSize = [Math]::Round($DriveSize/1GB)
				$Sheet1.Cells.Item($intRow, 10) = $DiskSize
				$Sheet1.Cells.item($intRow, 10).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 10).font.bold=$true
				$Sheet1.Cells.Item($intRow, 10).AddComment($errorComment) | Out-Null
			}
		}
			
		If($colItems.DeviceID -contains "J:")   {
			$DeviceID = $colItems | Where DeviceID -eq "J:"
			$DriveSize = $DeviceID.Size
			$DiskSize = [Math]::Round($DriveSize/1GB)
			$DriveLetter = $DeviceID.DeviceID
			If($DiskSize -ne $inputVM[0].J_Drive)   {
				$errorComment = "Expected Value: `n" + $inputVM[0].J_Drive
				$Sheet1.Cells.Item($intRow, 11) = $DiskSize
				$Sheet1.Cells.item($intRow, 11).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 11).font.bold=$true
				$Sheet1.Cells.Item($intRow, 11).AddComment($errorComment) | Out-Null
				Write-Host "`t`t`tIncorrect Drive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Red -NoNewline
				Write-Host " GB" -ForegroundColor White
			}
			Else   {
				$Sheet1.Cells.Item($intRow, 11) = $DiskSize
				$Sheet1.Cells.item($intRow, 11).font.colorIndex=10
				Write-Host "`t`tDrive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Green -NoNewline
				Write-Host " GB" -ForegroundColor Green
			}
		}
		
		ElseIf($colItems.DeviceID -notcontains "J:")   {
			$errorComment = "Expected Value: `n" + $inputVM[0].J_Drive
			If($inputVM[0].J_Drive -eq ".")   {
					
			}
			Else   {
				$DeviceID = $colItems | Where DeviceID -eq "J:"
				$DriveSize = $DeviceID.Size
				$DiskSize = [Math]::Round($DriveSize/1GB)
				$Sheet1.Cells.Item($intRow, 11) = $DiskSize
				$Sheet1.Cells.item($intRow, 11).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 11).font.bold=$true
				$Sheet1.Cells.Item($intRow, 11).AddComment($errorComment) | Out-Null
			}
		}

		If($colItems.DeviceID -contains "K:")   {
			$DeviceID = $colItems | Where DeviceID -eq "K:"
			$DriveSize = $DeviceID.Size
			$DiskSize = [Math]::Round($DriveSize/1GB)
			$DriveLetter = $DeviceID.DeviceID
			If($DiskSize -ne $inputVM[0].K_Drive)   {
				$errorComment = "Expected Value: `n" + $inputVM[0].K_Drive
				$Sheet1.Cells.Item($intRow, 12) = $DiskSize
				$Sheet1.Cells.item($intRow, 12).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 12).font.bold=$true
				$Sheet1.Cells.Item($intRow, 12).AddComment($errorComment) | Out-Null
				Write-Host "`t`t`tIncorrect Drive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Red -NoNewline
				Write-Host " GB" -ForegroundColor White
			}
			Else   {
				$Sheet1.Cells.Item($intRow, 12) = $DiskSize
				$Sheet1.Cells.item($intRow, 12).font.colorIndex=10
				Write-Host "`t`tDrive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Green -NoNewline
				Write-Host " GB" -ForegroundColor Green
			}
		}
		
		ElseIf($colItems.DeviceID -notcontains "K:")   {
			$errorComment = "Expected Value: `n" + $inputVM[0].K_Drive
			If($inputVM[0].K_Drive -eq ".")   {
					
			}
			Else   {
				$DeviceID = $colItems | Where DeviceID -eq "K:"
				$DriveSize = $DeviceID.Size
				$DiskSize = [Math]::Round($DriveSize/1GB)
				$Sheet1.Cells.Item($intRow, 12) = $DiskSize
				$Sheet1.Cells.item($intRow, 12).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 12).font.bold=$true
				$Sheet1.Cells.Item($intRow, 12).AddComment($errorComment) | Out-Null
			}
		}
	}
	
	ElseIf ($Platform -eq "Azure")   {
		If($colItems.DeviceID -contains "A:")   {
			$DeviceID = $colItems | Where DeviceID -eq "A:"
			$DriveSize = $DeviceID.Size
			$DiskSize = [Math]::Round($DriveSize/1GB)
			$DriveLetter = $DeviceID.DeviceID
			If($DiskSize -ne $inputVM[0].A_Drive)   {
				$errorComment = "Expected Value: `n" + $inputVM[0].A_Drive
				$Sheet1.Cells.Item($intRow, 7) = $DiskSize
				$Sheet1.Cells.item($intRow, 7).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 7).font.bold=$true
				$Sheet1.Cells.Item($intRow, 7).AddComment($errorComment) | Out-Null
				Write-Host "`t`t`tIncorrect Drive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Red -NoNewline
				Write-Host " GB" -ForegroundColor White
				}
			Else   {
				$Sheet1.Cells.Item($intRow, 7) = $DiskSize
				$Sheet1.Cells.item($intRow, 7).font.colorIndex=10
				Write-Host "`t`tDrive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Green -NoNewline
				Write-Host " GB" -ForegroundColor Green
			}
		}
		
		ElseIf($colItems.DeviceID -notcontains "A:")   {
			$errorComment = "Expected Value: `n" + $inputVM[0].A_Drive
			If($inputVM[0].A_Drive -eq ".")   {
			
			}
			Else   {
				$DeviceID = $colItems | Where DeviceID -eq "A:"
				$DriveSize = $DeviceID.Size
				$DiskSize = [Math]::Round($DriveSize/1GB)
				$Sheet1.Cells.Item($intRow, 7) = $DiskSize
				$Sheet1.Cells.item($intRow, 7).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 7).font.bold=$true
				$Sheet1.Cells.Item($intRow, 7).AddComment($errorComment) | Out-Null
				}
		}
			
		If($colItems.DeviceID -contains "B:")   {
			$DeviceID = $colItems | Where DeviceID -eq "B:"
			$DriveSize = $DeviceID.Size
			$DiskSize = [Math]::Round($DriveSize/1GB)
			$DriveLetter = $DeviceID.DeviceID
			If($DiskSize -ne $inputVM[0].B_Drive)   {
				$errorComment = "Expected Value: `n" + $inputVM[0].B_Drive
				$Sheet1.Cells.Item($intRow, 8) = $DiskSize
				$Sheet1.Cells.item($intRow, 8).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 8).font.bold=$true
				$Sheet1.Cells.Item($intRow, 8).AddComment($errorComment) | Out-Null
				Write-Host "`t`t`tIncorrect Drive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Red -NoNewline
				Write-Host " GB" -ForegroundColor White
			}
			Else   {
				$Sheet1.Cells.Item($intRow, 8) = $DiskSize
				$Sheet1.Cells.item($intRow, 8).font.colorIndex=10
				Write-Host "`t`tDrive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Green -NoNewline
				Write-Host " GB" -ForegroundColor Green
			}
		}
		
		ElseIf($colItems.DeviceID -notcontains "B:")   {
			$errorComment = "Expected Value: `n" + $inputVM[0].B_Drive
			If($inputVM[0].B_Drive -eq ".")   {
					
			}
			Else   {
				$DeviceID = $colItems | Where DeviceID -eq "B:"
				$DriveSize = $DeviceID.Size
				$DiskSize = [Math]::Round($DriveSize/1GB)
				$Sheet1.Cells.Item($intRow, 8) = $DiskSize
				$Sheet1.Cells.item($intRow, 8).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 8).font.bold=$true
				$Sheet1.Cells.Item($intRow, 8).AddComment($errorComment) | Out-Null
			}
		}
			
		If($colItems.DeviceID -contains "P:")   {
			$DeviceID = $colItems | Where DeviceID -eq "P:"
			$DriveSize = $DeviceID.Size
			$DiskSize = [Math]::Round($DriveSize/1GB)
			$DriveLetter = $DeviceID.DeviceID
			If($DiskSize -ne $inputVM[0].P_Drive)   {
				$errorComment = "Expected Value: `n" + $inputVM[0].B_Drive
				$Sheet1.Cells.Item($intRow, 13) = $DiskSize
				$Sheet1.Cells.item($intRow, 13).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 13).font.bold=$true
				$Sheet1.Cells.Item($intRow, 13).AddComment($errorComment) | Out-Null
				Write-Host "`t`t`tIncorrect Drive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Red -NoNewline
				Write-Host " GB" -ForegroundColor White
			}
			Else   {
				$Sheet1.Cells.Item($intRow, 13) = $DiskSize
				$Sheet1.Cells.item($intRow, 13).font.colorIndex=10
				Write-Host "`t`tDrive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Green -NoNewline
				Write-Host " GB" -ForegroundColor Green
			}
		}
		
		ElseIf($colItems.DeviceID -notcontains "P:")   {
			$errorComment = "Expected Value: `n" + $inputVM[0].P_Drive
			If($inputVM[0].P_Drive -eq ".")   {
					
			}
			Else   {
				$DeviceID = $colItems | Where DeviceID -eq "P:"
				$DriveSize = $DeviceID.Size
				$DiskSize = [Math]::Round($DriveSize/1GB)
				$Sheet1.Cells.Item($intRow, 13) = $DiskSize
				$Sheet1.Cells.item($intRow, 13).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 13).font.bold=$true
				$Sheet1.Cells.Item($intRow, 13).AddComment($errorComment) | Out-Null
			}
		}			

		If($colItems.DeviceID -contains "X:")   {
			$DeviceID = $colItems | Where DeviceID -eq "X:"
			$DriveSize = $DeviceID.Size
			$DiskSize = [Math]::Round($DriveSize/1GB)
			$DriveLetter = $DeviceID.DeviceID
			If($DiskSize -ne $inputVM[0].X_Drive)   {
				$errorComment = "Expected Value: `n" + $inputVM[0].X_Drive
				$Sheet1.Cells.Item($intRow, 9) = $DiskSize
				$Sheet1.Cells.item($intRow, 9).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 9).font.bold=$true
				$Sheet1.Cells.Item($intRow, 9).AddComment($errorComment) | Out-Null
				Write-Host "`t`t`tIncorrect Drive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Red -NoNewline
				Write-Host " GB" -ForegroundColor White
			}
			Else   {
				$Sheet1.Cells.Item($intRow, 9) = $DiskSize
				$Sheet1.Cells.item($intRow, 9).font.colorIndex=10
				Write-Host "`t`tDrive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Green -NoNewline
				Write-Host " GB" -ForegroundColor Green
			}					
		}
		
		ElseIf($colItems.DeviceID -notcontains "X:")   {
			$errorComment = "Expected Value: `n" + $inputVM[0].X_Drive
			If($inputVM[0].X_Drive -eq ".")   {
				
			}
			Else   {
				$DeviceID = $colItems | Where DeviceID -eq "X:"
				$DriveSize = $DeviceID.Size
				$DiskSize = [Math]::Round($DriveSize/1GB)
				$Sheet1.Cells.Item($intRow, 9) = $DiskSize
				$Sheet1.Cells.item($intRow, 9).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 9).font.bold=$true
				$Sheet1.Cells.Item($intRow, 9).AddComment($errorComment) | Out-Null
			}
		}
			
		If($colItems.DeviceID -contains "W:")   {
			$DeviceID = $colItems | Where DeviceID -eq "W:"
			$DriveSize = $DeviceID.Size
			$DiskSize = [Math]::Round($DriveSize/1GB)
			$DriveLetter = $DeviceID.DeviceID
			If($DiskSize -ne $inputVM[0].W_Drive)   {
				$errorComment = "Expected Value: `n" + $inputVM[0].W_Drive
				$Sheet1.Cells.Item($intRow, 10) = $DiskSize
				$Sheet1.Cells.item($intRow, 10).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 10).font.bold=$true
				$Sheet1.Cells.Item($intRow, 10).AddComment($errorComment) | Out-Null
				Write-Host "`t`t`tIncorrect Drive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Red -NoNewline
				Write-Host " GB" -ForegroundColor White
			}
			Else   {
				$Sheet1.Cells.Item($intRow, 10) = $DiskSize
				$Sheet1.Cells.item($intRow, 10).font.colorIndex=10
				Write-Host "`t`tDrive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Green -NoNewline
				Write-Host " GB" -ForegroundColor Green
			}
		}
		
		ElseIf($colItems.DeviceID -notcontains "W:")   {
			$errorComment = "Expected Value: `n" + $inputVM[0].W_Drive
			If($inputVM[0].W_Drive -eq ".")   {
					
			}
			Else   {
				$DeviceID = $colItems | Where DeviceID -eq "W:"
				$DriveSize = $DeviceID.Size
				$DiskSize = [Math]::Round($DriveSize/1GB)
				$Sheet1.Cells.Item($intRow, 10) = $DiskSize
				$Sheet1.Cells.item($intRow, 10).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 10).font.bold=$true
				$Sheet1.Cells.Item($intRow, 10).AddComment($errorComment) | Out-Null
			}
		}
					
		If($colItems.DeviceID -contains "J:")   {
			$DeviceID = $colItems | Where DeviceID -eq "J:"
			$DriveSize = $DeviceID.Size
			$DiskSize = [Math]::Round($DriveSize/1GB)
			$DriveLetter = $DeviceID.DeviceID
			If($DiskSize -ne $inputVM[0].J_Drive)   {
				$errorComment = "Expected Value: `n" + $inputVM[0].J_Drive
				$Sheet1.Cells.Item($intRow, 11) = $DiskSize
				$Sheet1.Cells.item($intRow, 11).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 11).font.bold=$true
				$Sheet1.Cells.Item($intRow, 11).AddComment($errorComment) | Out-Null
				Write-Host "`t`t`tIncorrect Drive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Red -NoNewline
				Write-Host " GB" -ForegroundColor White
			}
			Else   {
				$Sheet1.Cells.Item($intRow, 11) = $DiskSize
				$Sheet1.Cells.item($intRow, 11).font.colorIndex=10
				Write-Host "`t`tDrive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Green -NoNewline
				Write-Host " GB" -ForegroundColor Green
			}
		}
		
		ElseIf($colItems.DeviceID -notcontains "J:")   {
			$errorComment = "Expected Value: `n" + $inputVM[0].J_Drive
			If($inputVM[0].J_Drive -eq ".")   {
					
			}
			Else   {
				$DeviceID = $colItems | Where DeviceID -eq "J:"
				$DriveSize = $DeviceID.Size
				$DiskSize = [Math]::Round($DriveSize/1GB)
				$Sheet1.Cells.Item($intRow, 11) = $DiskSize
				$Sheet1.Cells.item($intRow, 11).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 11).font.bold=$true
				$Sheet1.Cells.Item($intRow, 11).AddComment($errorComment) | Out-Null
			}
		}

		If($colItems.DeviceID -contains "K:")   {
			$DeviceID = $colItems | Where DeviceID -eq "K:"
			$DriveSize = $DeviceID.Size
			$DiskSize = [Math]::Round($DriveSize/1GB)
			$DriveLetter = $DeviceID.DeviceID
			If($DiskSize -ne $inputVM[0].K_Drive)   {
				$errorComment = "Expected Value: `n" + $inputVM[0].K_Drive
				$Sheet1.Cells.Item($intRow, 12) = $DiskSize
				$Sheet1.Cells.item($intRow, 12).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 12).font.bold=$true
				$Sheet1.Cells.Item($intRow, 12).AddComment($errorComment) | Out-Null
				Write-Host "`t`t`tIncorrect Drive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Red -NoNewline
				Write-Host " GB" -ForegroundColor White
			}
			Else   {
				$Sheet1.Cells.Item($intRow, 12) = $DiskSize
				$Sheet1.Cells.item($intRow, 12).font.colorIndex=10
				Write-Host "`t`tDrive " -ForegroundColor White -NoNewline
				Write-Host $DriveLetter $DiskSize -ForegroundColor Green -NoNewline
				Write-Host " GB" -ForegroundColor Green
			}
		}
		
		ElseIf($colItems.DeviceID -notcontains "K:")   {
			$errorComment = "Expected Value: `n" + $inputVM[0].K_Drive
			If($inputVM[0].K_Drive -eq ".")   {
					
			}
			Else   {
				$DeviceID = $colItems | Where DeviceID -eq "K:"
				$DriveSize = $DeviceID.Size
				$DiskSize = [Math]::Round($DriveSize/1GB)
				$Sheet1.Cells.Item($intRow, 12) = $DiskSize
				$Sheet1.Cells.item($intRow, 12).font.colorIndex=3
				$Sheet1.Cells.item($intRow, 12).font.bold=$true
				$Sheet1.Cells.Item($intRow, 12).AddComment($errorComment) | Out-Null
			}
		}
	}
}

Function Get-SQLVersion   {
	param ([String] $Computer, [String] $FQDN, [int]$IntRow)
	
	$TempSession = New-PSSession -ComputerName $Computer -Credential $cred
	$PatchLevel = Invoke-Command -Session $TempSession -ScriptBlock {(Get-ItemProperty "RegLocation").Patchlevel}
	[String]$SQLVersion = $inputVM[0].SQL_Version
	$errorComment = "Expected Value: `n" + $SQLVersion

	$column	= 14
	If ($PatchLevel -ne $SQLVersion)   {
		$Sheet1.Cells.Item($intRow, $column) = $PatchLevel
		$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
		$Sheet1.Cells.item($intRow, $column).font.bold=$true
		$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
		Write-Host "`t`t`tIncorrect SQL Version: " -ForegroundColor White -NoNewline
		Write-Host $PatchLevel -ForegroundColor Red
	}
	Else   {
		$Sheet1.Cells.Item($intRow, $column) = $PatchLevel
		$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
		Write-Host "`t`tSQL Version: " -ForegroundColor White -NoNewline
		Write-Host $PatchLevel -ForegroundColor Green
	}
}

Function Get-MDOP   {
	param ([String] $Computer, [String] $FQDN, [int]$IntRow)
	
	[String]$Value = Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		$MDOPValue = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT value_in_use FROM sys.configurations  WHERE name = 'max degree of parallelism'").value_in_use
		return $MDOPValue
		}
		
		[String]$MDOP = $inputVM[0].MDOP
		$errorComment = "Expected Value: `n" + $MDOP
		$column = 15

		If($MDOPValue -ne $MDOP)   {
			$Sheet1.Cells.Item($intRow, $column) = $Value
			$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
			$Sheet1.Cells.item($intRow, $column).font.bold=$true
			$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
			Write-Host "`t`t`tIncorrect MDOP Value: " -ForegroundColor White -NoNewline
			Write-Host $Value -ForegroundColor Red
		}
		Else   {    
			$Sheet1.Cells.Item($intRow, $column) = $Value
			$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
			Write-Host "`t`tMDOP Value: " -ForegroundColor White -NoNewline
			Write-Host $Value -ForegroundColor Green
		}
}

Function Get-BackupCompression   {
	param ([String] $Computer, [String] $FQDN, [int]$IntRow)
	
	$BackupValue = Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		$BackupCompressionValue = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT value FROM sys.configurations WHERE name = 'backup compression default' select j.name as 'JobName',run_date,run_time From MS.dbo.sysjobs j INNER JOIN MS.dbo.sysjobhistory h ON j.job_id = h.job_id where j.enabled = 1 order by JobName").value
		return $BackupCompressionValue
		}
		[String]$BCValue = $BackupValue[0]
		[String]$CompressionValue = $inputVM[0].BackupCompression
		$errorComment = "Expected Value: `n" + $CompressionValue
		$column = 16

		If($BCValue -ne $CompressionValue)   {
			$Sheet1.Cells.Item($intRow, $column) = $BCValue
			$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
			$Sheet1.Cells.item($intRow, $column).font.bold=$true
			$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
			Write-Host "`t`t`tIncorrect Backup Compression Value: " -ForegroundColor White -NoNewline
			Write-Host $BCValue -ForegroundColor Red
		}
		Else   {    
			$Sheet1.Cells.Item($intRow, $column) = $BCValue
			$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
			Write-Host "`t`tBackup Compression Value: " -ForegroundColor White -NoNewline
			Write-Host $BCValue -ForegroundColor Green
		}
}

Function Get-Tempdb   {
	param ([String] $Computer, [String] $FQDN, [int]$IntRow)

	$TempdbValue = Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		[Hashtable]$Tempdb = @{}
		$Tempdb.Size = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT DB_NAME(database_id) AS DatabaseName, Name AS Logical_Name, Physical_Name, (size * 8) / 1024 SizeMB FROM sys.M_files WHERE DB_NAME(database_id) = 'dbName'")
		$Tempdb.Growth = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT DB_NAME(database_id) AS DatabaseName, Name AS Logical_Name, Physical_Name, (growth * 8) / 1024 SizeMB FROM sys.M_files WHERE DB_NAME(database_id) = 'dbName'")
		$Tempdb.Name = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT DB_NAME(database_id) AS DatabaseName, Name AS Logical_Name, Physical_Name, (growth * 8) / 1024 SizeMB FROM sys.M_files WHERE DB_NAME(database_id) = 'dbName'")
		return $Tempdb
	}
	
	$TempdbSize = ($TempdbValue.Size).SizeMB
	$TempdbGrowth = ($TempdbValue.Growth).SizeMB
	$TempdbName = ($TempdbValue.Name).Logical_Name	

		$i = 0
		While($i -le $TempdbSize.Length)   {
			$Name = $TempdbName[$i]
			$Size = $TempdbSize[$i]
			$Growth = $TempdbGrowth[$i]

			If($Name -eq "Temp")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp_Size
				If($Size -ne $inputVM[0].Temp_Size)   {
					$column = 30
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   { 
					$column = 30
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp_Growth
				If($Growth -ne $inputVM[0].Temp_Growth)   {
					$column = 31					
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 31
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}
			}
			
			ElseIf($Name -eq "temp")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp_Size
				If($Size -ne $inputVM[0].Temp_Size)   {
					$column = 32
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   {    
					$column = 32
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp_Growth
				If($Growth -ne $inputVM[0].Temp_Growth)   {
					$column = 33
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 33
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}			
			}
			
			ElseIf($Name -eq "Temp1")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp1_Size			
				If($Size -ne $inputVM[0].Temp1_Size)   {
					$column = 34					
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp1 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   {    
					$column = 34
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp1 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp1_Growth				
				If($Growth -ne $inputVM[0].Temp1_Growth)   {
					$column = 35
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp1 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 35
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp1 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}			
			}
			
			ElseIf($Name -eq "Temp2")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp2_Size
				If($Size -ne $inputVM[0].Temp2_Size)   {
					$column = 36
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp2 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   {    
					$column = 36
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp2 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp2_Growth
				If($Growth -ne $inputVM[0].Temp2_Growth)   {
					$column = 37
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp2 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 37
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp2 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}				
			}

			ElseIf($Name -eq "Temp3")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp3_Size
				If($Size -ne $inputVM[0].Temp3_Size)   {
					$column = 38
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp3 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   {    
					$column = 38
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp3 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp3_Growth
				If($Growth -ne $inputVM[0].Temp3_Growth)   {
					$column = 39
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp3 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 39
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp3 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}		
			}
			
			ElseIf($Name -eq "Temp4")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp4_Size			
				If($Size -ne $inputVM[0].Temp4_Size)   {
					$column = 40
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp4 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red					
				}
				Else   {    
					$column = 40
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp4 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp4_Growth
				If($Growth -ne $inputVM[0].Temp4_Growth)   {
					$column = 41
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp4 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red					
				}
				Else   {    
					$column = 41
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp4 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}			
			}
			
			ElseIf($Name -eq "Temp5")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp5_Size			
				If($Size -ne $inputVM[0].Temp5_Size)   {
					$column = 42
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp5 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red					
				}
				Else   {    
					$column = 42
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp5 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp5_Growth
				If($Growth -ne $inputVM[0].Temp5_Growth)   {
					$column = 43
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp5 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red					
				}
				Else   {    
					$column = 43
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp5 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}
			}
			
			ElseIf($Name -eq "Temp6")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp6_Size			
				If($Size -ne $inputVM[0].Temp6_Size)   {
					$column = 44
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp6 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red					
				}
				Else   {    
					$column = 44
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp6 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp6_Growth
				If($Growth -ne $inputVM[0].Temp6_Growth)   {
					$column = 45
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp6 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red					
				}
				Else   {    
					$column = 45
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp6 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}
			}
			
			ElseIf($Name -eq "Temp7")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp7_Size			
				If($Growth -ne $inputVM[0].Temp7_Size)   {
					$column = 46
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp7 size: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red					
				}
				Else   {    
					$column = 46
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp7 size: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp7_Growth
				If($Growth -ne $inputVM[0].Temp7_Growth)   {
					$column = 47
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp7 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red					
				}
				Else   {    
					$column = 47
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp7 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}
			}
			
			ElseIf($Name -eq "Temp8")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp8_Size			
				If($Size -ne $inputVM[0].Temp8_Size)   {
					$column = 48					
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp8 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red					
				}
				Else   {    
					$column = 48
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp8 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp8_Growth				
				If($Growth -ne $inputVM[0].Temp8_Growth)   {
					$column = 49
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp8 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red					
				}
				Else   {    
					$column = 49
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp8 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}			
			}
			
			ElseIf($Name -eq "Temp9")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp9_Size
				If($Size -ne $inputVM[0].Temp9_Size)   {
					$column = 50
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp9 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   {    
					$column = 50
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp9 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp9_Growth
				If($Growth -ne $inputVM[0].Temp9_Growth)   {
					$column = 51
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t Incorrect Temp9 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 51
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp9 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}				
			}

			ElseIf($Name -eq "Temp10")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp10_Size
				If($Size -ne $inputVM[0].Temp10_Size)   {
					$column = 52
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp10 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   {    
					$column = 52
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp10 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp10_Growth
				If($Growth -ne $inputVM[0].Temp10_Growth)   {
					$column = 53
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp10 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 53
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp10 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}		
			}
			
			ElseIf($Name -eq "Temp11")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp11_Size			
				If($Size -ne $inputVM[0].Temp11_Size)   {
					$column = 54
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp11 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   {    
					$column = 54
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp11 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp11_Growth
				If($Growth -ne $inputVM[0].Temp11_Growth)   {
					$column = 55
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp11 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 55
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp11 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}			
			}
			
			ElseIf($Name -eq "Temp12")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp12_Size			
				If($Size -ne $inputVM[0].Temp12_Size)   {
					$column = 56
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp12 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   {    
					$column = 56
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp12 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp12_Growth
				If($Growth -ne $inputVM[0].Temp12_Growth)   {
					$column = 57
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp12 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 57
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp12 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}
			}
			
			ElseIf($Name -eq "Temp13")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp13_Size			
				If($Size -ne $inputVM[0].Temp13_Size)   {
					$column = 58
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp13 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   {    
					$column = 58
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp13 size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp13_Growth
				If($Growth -ne $inputVM[0].Temp13_Growth)   {
					$column = 59
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp13 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 59
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp13 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}
			}
			
			ElseIf($Name -eq "Temp14")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp14_Size			
				If($Growth -ne $inputVM[0].Temp14_Size)   {
					$column = 60
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp14 size: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 60
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp14 size: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp14_Growth
				If($Growth -ne $inputVM[0].Temp14_Growth)   {
					$column = 61
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp14 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 61
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp14 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}
			}			
			
			ElseIf($Name -eq "Temp15")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp15_Size			
				If($Growth -ne $inputVM[0].Temp15_Size)   {
					$column = 62
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp15 size: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 62
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp15 size: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Temp15_Growth
				If($Growth -ne $inputVM[0].Temp15_Growth)   {
					$column = 63
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Temp15 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 63
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tTemp15 growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}
			}						
			$i++
		}
}

Function Get-Databases   {
	param ([String] $Computer, [String] $FQDN, [int]$IntRow)

	$Recovery = Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		$DBRecovery = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT 'M2' AS [Database Name],DATABASEPROPERTYEX('M2', 'RECOVERY')").Column1
		return $DBRecovery		
	}
		$errorComment = "Expected Value: `n" + $inputVM[0].Recovery
		$column	= 17
		If($Recovery -ne $inputVM[0].Recovery)   {
			$Sheet1.Cells.Item($intRow, $column) = $Recovery
			$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
			$Sheet1.Cells.item($intRow, $column).font.bold=$true
			$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
			Write-Host "`t`t`tIncorrect M2 Recovery: " -ForegroundColor White -NoNewline
			Write-Host $Recovery -ForegroundColor Red			
		}
		Else   {    
			$Sheet1.Cells.Item($intRow, $column) = $Recovery
			$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
			Write-Host "`t`tM2 Recovery: " -ForegroundColor White -NoNewline
			Write-Host $Recovery -ForegroundColor Green
		}

	$MValue = Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		[Hashtable]$M = @{}
		$M.Size = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT DB_NAME(database_id) AS DatabaseName, Name AS Logical_Name, Physical_Name, (size * 8) / 1024 SizeMB FROM sys.M_files WHERE DB_NAME(database_id) = 'M'")
		$M.Growth = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT DB_NAME(database_id) AS DatabaseName, Name AS Logical_Name, Physical_Name, (growth * 8) / 1024 SizeMB FROM sys.M_files WHERE DB_NAME(database_id) = 'M'")
		$M.Name = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT DB_NAME(database_id) AS DatabaseName, Name AS Logical_Name, Physical_Name, (growth * 8) / 1024 SizeMB FROM sys.M_files WHERE DB_NAME(database_id) = 'M'")
		return $M
	}
		$MSize = ($MValue.Size).SizeMB
		$MGrowth = ($MValue.Growth).SizeMB
		$MName = ($MValue.Name).Logical_Name
		$i = 0
		While($i -le $MSize.Length)   {
			$Name = $MName[$i]
			$Size = $MSize[$i]
			$Growth = $MGrowth[$i]
			
			If($Name -eq "M")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].M_Size
				If($Size -ne $inputVM[0].M_Size)   {
					$column = 18
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect M size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   { 
					$column = 18
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tM size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].M_Growth
				If($Growth -ne $inputVM[0].M_Growth)   {
					$column = 19					
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect M growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 19
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tM growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}
			}
			ElseIf($Name -eq "mastlog")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].Mlog_Size
				If($Size -ne $inputVM[0].Mlog_Size)   {
					$column = 20
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Mlog size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   {    
					$column = 20
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tMlog size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].Mlog_Growth
				If($Growth -ne $inputVM[0].Mlog_Growth)   {
					$column = 21
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect Mlog growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 21
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tMlog growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}			
			}
		$i++
		}

	$MSValue = Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		[Hashtable]$MS = @{}
		$MS.Size = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT DB_NAME(database_id) AS DatabaseName, Name AS Logical_Name, Physical_Name, (size * 8) / 1024 SizeMB FROM sys.M_files WHERE DB_NAME(database_id) = 'MS'")
		$MS.Growth = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT DB_NAME(database_id) AS DatabaseName, Name AS Logical_Name, Physical_Name, (growth * 8) / 1024 SizeMB FROM sys.M_files WHERE DB_NAME(database_id) = 'MS'")
		$MS.Name = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT DB_NAME(database_id) AS DatabaseName, Name AS Logical_Name, Physical_Name, (growth * 8) / 1024 SizeMB FROM sys.M_files WHERE DB_NAME(database_id) = 'MS'")		
		return $MS
	}
		$MSSize = ($MSValue.Size).SizeMB
		$MSGrowth = ($MSValue.Growth).SizeMB
		$MSName = ($MSValue.Name).Logical_Name		
		$i = 0
		While($i -le $MSSize.Length)   {
			$Name = $MSName[$i]
			$Size = $MSSize[$i]
			$Growth = $MSGrowth[$i]
			If($Name -eq "MSData")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].MSData_Size
				If($Size -ne $inputVM[0].MSData_Size)   {
					$column = 22
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect MSdata size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   { 
					$column = 22
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tMSdata size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].MSData_Growth
				If($Growth -ne $inputVM[0].MSData_Growth)   {
					$column = 23					
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect MSdata growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 23
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tMSdata growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}
			}
			ElseIf($Name -eq "MS")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].MS_Size
				If($Size -ne $inputVM[0].MS_Size)   {
					$column = 24
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect MS size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   {    
					$column = 24
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tMS size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].MS_Growth
				If($Growth -ne $inputVM[0].MS_Growth)   {
					$column = 25
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect MS growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 25
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tMS growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}			
			}
			$i++
		}

	$M2Value = Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		[Hashtable]$M2 = @{}
		$M2.Size = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT DB_NAME(database_id) AS DatabaseName, Name AS Logical_Name, Physical_Name, (size * 8) / 1024 SizeMB FROM sys.M_files WHERE DB_NAME(database_id) = 'M2'")
		$M2.Growth = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT DB_NAME(database_id) AS DatabaseName, Name AS Logical_Name, Physical_Name, (growth * 8) / 1024 SizeMB FROM sys.M_files WHERE DB_NAME(database_id) = 'M2'")
		$M2.Name = (Invoke-Sqlcmd -ServerInstance $Computer -query "SELECT DB_NAME(database_id) AS DatabaseName, Name AS Logical_Name, Physical_Name, (growth * 8) / 1024 SizeMB FROM sys.M_files WHERE DB_NAME(database_id) = 'M2'")
		return $M2
	}
		$M2Size = ($M2Value.Size).SizeMB
		$M2Growth = ($M2Value.Growth).SizeMB
		$M2Name = ($M2Value.Name).Logical_Name	
		$i = 0
		While($i -le $M2Size.Length)   {
			$Name = $M2Name[$i]
			$Size = $M2Size[$i]
			$Growth = $M2Growth[$i]
			If($Name -eq "M2dev")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].M2dev_Size
				If($Size -ne $inputVM[0].M2dev_Size)   {
					$column = 26
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect M2dev size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   { 
					$column = 26
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tM2dev size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].M2dev_Growth
				If($Growth -ne $inputVM[0].M2dev_Growth)   {
					$column = 27					
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect M2dev growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 27
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tM2dev growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}
			}
			ElseIf($Name -eq "M2log")   {
				$errorComment = "Expected Value: `n" + $inputVM[0].M2log_Size
				If($Size -ne $inputVM[0].M2log_Size)   {
					$column = 28
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect M2log size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Red
				}
				Else   {    
					$column = 28
					$Sheet1.Cells.Item($intRow, $column) = $Size
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tM2log size: " -ForegroundColor White -NoNewline
					Write-Host $Size -ForegroundColor Green
				}
				$errorComment = "Expected Value: `n" + $inputVM[0].M2log_Growth
				If($Growth -ne $inputVM[0].M2log_Growth)   {
					$column = 29
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=3
					$Sheet1.Cells.item($intRow, $column).font.bold=$true
					$Sheet1.Cells.Item($intRow, $column).AddComment($errorComment) | Out-Null
					Write-Host "`t`t`tIncorrect M2log growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Red
				}
				Else   {    
					$column = 29
					$Sheet1.Cells.Item($intRow, $column) = $Growth
					$Sheet1.Cells.item($intRow, $column).font.colorIndex=10
					Write-Host "`t`tM2log growth: " -ForegroundColor White -NoNewline
					Write-Host $Growth -ForegroundColor Green
				}			
			}
		$i++	
		}
}

<#
Function Get-QuoteService   {
	param ([String] $Computer, [String] $FQDN, [int]$IntRow)
	
	[int]$totalProc = (Get-WmiObject -class "Win32_computersystem" -ComputerName $Computer).numberoflogicalprocessors	
	$errorComment = "Expected Value: `n" + $inputVM[0].Processors

	If($totalProc -ne $inputVM[0].Processors)   {
		$Sheet1.Cells.Item($intRow, 3) = $totalProc
		$Sheet1.Cells.item($intRow, 3).font.colorIndex=3
		$Sheet1.Cells.item($intRow, 3).font.bold=$true
		$Sheet1.Cells.Item($intRow, 3).AddComment($errorComment) | Out-Null
		$Message = "`tNumber of Processors INCORRECT"
	}
	Else   {
		$Sheet1.Cells.Item($intRow, 3) = $totalProc
		$Sheet1.Cells.item($intRow, 3).font.colorIndex=10
		Write-Host "`t`tNumber of Processors: " -ForegroundColor White -NoNewline
		Write-Host $totalProc -ForegroundColor Green
	}
}
#>

#Endregion Functions

#Region Main Processing

Test-Administrator

$Password = Get_Password

#build server array
$i = 0
$count = ($inputCSV.Azure_Computer).Count
While($i -lt $count)   {
	$servers = $servers + $inputCSV[$i].Azure_Computer
	$i++
}

#validate servers to process
If($servers -eq $null)   {
	$Message = "No Servers to Process. Exiting Script"
	Write-Warning $Message
	Break
}

Write-Header $count
Create-ExcelQC
Sleep 5

#set excel row below headers
$intCurrentRow = 2

#process servers
Foreach($strComputerName in $servers)   {
	#get inputCSV server specification
	$inputVM = $inputcsv | ?{$_.Azure_Computer -eq $strComputerName}
	
	#verIfy server is reachable bofore processing
	Write-Host "Processing: " -NoNewline
	Write-Host $strComputerName -ForegroundColor Green
	$FQDN = ("{0}.domain" -f $strComputerName)

	If(Test-Connection -ComputerName $strComputerName -Quiet -Count 1 -ErrorAction SilentlyContinue)   {
		$Sheet1.Cells.Item($intCurrentRow,1) = $strComputerName
		$Sheet1.Cells.Item($intCurrentRow,1).font.bold=$true

		$Sheet1.Cells.Item($intCurrentRow,2) = $inputVM.On_Premise_Computer
		$Sheet1.Cells.Item($intCurrentRow,2).font.bold=$true		
		
#		Get-MDOP $strComputerName $FQDN $intCurrentRow				
#EXIT
		#Query servers for QC properties
		Get-Proccessors $strComputerName $FQDN $intCurrentRow
		Get-TimeZone $strComputerName $FQDN $intCurrentRow
		Get-Memory $strComputerName $FQDN $intCurrentRow
		Get-Disk $strComputerName $FQDN $intCurrentRow
		Get-SQLVersion $strComputerName $FQDN $intCurrentRow
		Get-MDOP $strComputerName $FQDN $intCurrentRow		
		Get-BackupCompression $strComputerName $FQDN $intCurrentRow
		Get-Tempdb $strComputerName $FQDN $intCurrentRow		
		Get-Databases $strComputerName $FQDN $intCurrentRow

		Write-Host
		Write-Host "Finished Processing: " -NoNewline
		Write-Host $strComputerName -ForegroundColor Green
		$Message = "============================================="
	}
	Else   {
		#write to log computer unreachable
		$Sheet1.Cells.Item($intCurrentRow,1) = $strComputerName
		$Sheet1.Cells.Item($intCurrentRow,1).font.colorIndex=3
		$Sheet1.Cells.Item($intCurrentRow,1).font.bold=$true
		$Sheet1.Cells.Item($intCurrentRow,2) = "Server UNREACHABLE!"
		$Sheet1.Cells.Item($intCurrentRow,2).font.colorIndex=3
		$Sheet1.Cells.Item($intCurrentRow,2).font.bold=$true
		Write-Host
		$Message = "Computer $strComputerName unavailable"
		Write-Warning $Message
	}

	#Autofit the excel columns
	$script:WorkBook.EntireColumn.AutoFit() | Out-Null
	$intCurrentRow++
}

#close excel report add date generated information
$Sheet1.Cells.Item($intCurrentRow + 2,1) = "Report Generated: "
$Sheet1.Cells.Item($intCurrentRow + 2,2) = $Date

$Reportfile = $location + $User + "-QC-Report-" + $logdate + ".xls"
$Sheet1.SaveAs($Reportfile,1)

Write-Footer $count

#Endregion Main Processing