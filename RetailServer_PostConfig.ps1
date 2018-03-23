<# 
	.SYNOPSIS
	    Installs the latest IPAK required to patch a Retail Server as well as installs SQL via SQL IPAK along with
			default values used for setting up a new Retail Server.
	.Description
		This script will install IPAK and SQL along with the SQL prerequisites on a given host.
			SQL will be setup via the SQL IPAK as well as the default values associated to SQL.
			The script then configures the prerequisites required for future SQL configuration.
	.PARAMETER Computer
		The functional definition of which device being used to install post configurations.
	.PARAMETER Environment
		The functional definition of which environment to run the script as (PPE or PROD).
	.PARAMETER Platform
		The functional definition of where the device is located (On-Premise or Azure).
	.PARAMETER Skip
		The functional definition of which item to skip within the script (IPAK, SQL,or Both).
	.INPUTS
		None. You cannot pipe objects to SQL-PostConfig.ps1.
	.OUTPUTS
		Console and log file
    .NOTES
		Author: Marcellus Seamster Jr
	.EXAMPLE
		.\RetailServer_PostConfig.ps1 -Computer computerName -Environment PPE -Platform On-Premise
	.EXAMPLE
		.\RetailServer_PostConfig.ps1 -Computer computerName -Environment PPE -Platform On-Premise -Skip IPAK	
	.EXAMPLE
		.\RetailServer_PostConfig.ps1 -Computer computerName -Environment PPE -Platform On-Premise -Skip SQL
	.EXAMPLE
		.\RetailServer_PostConfig.ps1 -Computer computerName -Environment PPE -Platform On-Premise -Skip Both
#>

Param( $Computer ,[ValidateSet('Prod','PPE')]$Environment = 'PPE', [ValidateSet('On-Premise','Azure')]$Platform,[ValidateSet('IPAK','SQL','Both')]$Skip)


#Region Variable declartion
	#logging variables
	$logdate = "{0:MMddyyyy}" -f (Get-Date)
	$logpath = "fileName"
	$logfile = $logpath + "RetailServer_PostConfig_" + $Computer + "_" + $logdate + ".log"
	
	$Result = @()
	$FinalReboot = @()
	$Complete = @()
	$Reboot = @()
	$count = @()
	$Password = @()
	
	#Memory variables
	$totalMem = @()
	$cleanMemMB = @()
	$maxMem = @()
	$minMem = @()
				
	#MDOP variables
	$totalProc = @()
	$halfProcCount = @()

	#TempDB variables
	$tempname = @()
	$tempfile = @()
	$tempfileName = @()
	
	#Sleep variables
    $Time = @()
    $Minute = @()
	
	#Administrator variables
	$AdminUser = @()
    $Admin = @()
	
	#VerIfy Admin variables
	$VerIfy = @()
	
	#RegKey Variables
	$TempSession = @()
	
	$RegValue = @()	
	$Reg = @()
	$RegKey = @()
		
	$WarningColor = (Get-Host).PrivateData
	$WarningColor.WarningBackgroundColor = "yellow"
	$WarningColor.WarningForegroundColor = "blue"
	
#EndRegion Variable declaration

#Region Get control of powershell UI
    $a = (Get-Host).UI.RawUI
    $a.WindowTitle = "RetailServer PostConfig"
    $a.BackgroundColor = "Black"
    $a.ForegroundColor = "White"
    cls
#EndRegion

#Region Set startup variables
	$version = "1.0"
	$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()
	$Date = Get-Date
	$ErrorActionPreference = "SilentlyContinue"
	
	$displayRows = 10
	$i = 0
	$RegTracker = 0
	$InstallationStatus = 0
	$Environment = $Environment.toUpper()

	$Reboot = $false
	$Complete = $false
	$Check = $false
	$SQLCheck = $false
	$VersionCheck = $false

	$User = $env:username
	$Domain = $env:userdomain
	$UserName = $Domain + "\" + $User
	$cred = Get-Credential -Credential $UserName

    $InstallationStatusRegPath = 'RegLocation'
    $InstallationStatusRegKey = 'Status'	
	
	$StartPath = "fileLocation"
	$logApplication = "MSNIPAK"
	$Folder = "\\" + $Computer + "folderName"
	$FQDN = $Computer + "domain"
	
	$Line = "================================================================================"
		
	If ($Platform -eq "Azure")   {
		$FinalPath = "\\" + $Computer + "\g$\"
	}
	ElseIf ($Platform -eq "On-Premise")   {
		$FinalPath = "\\" + $Computer + "\d$\"
	}
	
	If ([string]::IsNullOrWhiteSpace($Computer))   {
		Write-Host "`t`tInvalid parameter(s) entered. See: " -ForegroundColor White -NoNewline
		Write-Host "Get-Help .\RetailServer_PostConfig.ps1 -Full" -ForegroundColor Cyan
		Write-Host
		Exit
	}
	Else   {
		If (Test-Connection -ComputerName $Computer -Quiet -Count 1)   {
			$Computer = $Computer.toUpper()
		}
		Else   {
			$Computer = $Computer.toUpper()
			Write-Host "`t`tVerIfy " -ForegroundColor White -NoNewline
			Write-Host $Computer -ForegroundColor Cyan -NoNewline
			Write-Host " is online and rerun script!" -ForegroundColor White
			Write-Host
			Exit
		}
	}

#EndRegion Set startup variables

#Region WorkFlows

WorkFlow Restart_Host{
 param ([string[]]$Computer)
	Restart-Computer -PSComputerName $Computer -Force -Wait
}

#EndRegion WorkFlows

#Region Functions

Function Write-Log {
	#region Parameters
		[cmdletbinding()]
		Param(
			[Parameter(ValueFromPipeline=$true,Mandatory=$true)] [ValidateNotNullOrEmpty()]
			[string] $Message,
			[Parameter()] [ValidateSet("Error", "Warn", "Info")]
			[string] $Level = "Info",
			[Parameter()]
			[Switch] $NoConsoleOut,
			[Parameter()]
			[String] $ConsoleForeground = 'Black',
			[Parameter()] [ValidateRange(1,30)]
			[Int16] $Indent = 0,
			[Parameter()]
			[IO.FileInfo] $Path = "$env:temp\PowerShellLog.txt",
			[Parameter()]
			[Switch] $Clobber,
			[Parameter()]
			[String] $EventLogName,
			[Parameter()]
			[String] $EventSource,
			[Parameter()]
			[Int32] $EventID = 1,
			[Parameter()]
			[String] $LogEncoding = "ASCII"
		) #end Parameters
	#endregion

	Begin {}
	Process {
		Try   {			
			$msg = '{0}{1} : {2} : {3}' -f (" " * $Indent), (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level.ToUpper(), $Message
			If ($NoConsoleOut -eq $false)   {
				switch ($Level)   {
					'Error' { Write-Error $Message }
					'Warn' { Write-Warning $Message }
					'Info' { Write-Host ('{0}{1}' -f (" " * $Indent), $Message) -ForegroundColor $ConsoleForeground}
				}
			}
			If (-not $Path.Exists)   {
				New-Item -Path $Path.FullName -ItemType File -Force | Out-Null
			}
			If ($Clobber)   {
				$msg | Out-File -FilePath $Path -Encoding $LogEncoding -Force
			} Else   {
				$msg | Out-File -FilePath $Path -Encoding $LogEncoding -Append
			}
			If ($EventLogName)   {
				If (-not $EventSource)   {
					$EventSource = ([IO.FileInfo] $MyInvocation.ScriptName).Name
				}
				If (-not [Diagnostics.EventLog]::SourceExists($EventSource))   { 
					[Diagnostics.EventLog]::CreateEventSource($EventSource, $EventLogName) 
		        } 
				$log = New-Object System.Diagnostics.EventLog  
			    $log.set_log($EventLogName)  
			    $log.set_source($EventSource) 
				switch ($Level)   {
					"Error" { $log.WriteEnTry($Message, 'Error', $EventID) }
					"Warn"  { $log.WriteEnTry($Message, 'Warning', $EventID) }
					"Info"  { $log.WriteEnTry($Message, 'Information', $EventID) }
				}
			}
		} Catch   {
			throw "Failed to create log enTry in: ‘$Path’. The error was: ‘$_’."
		}
	} #End Process

	End {}
	<#
		.SYNOPSIS
			Writes logging information to screen and log file simultaneously.
		.DESCRIPTION
			Writes logging information to screen and log file simultaneously. Supports multiple log levels.
		.PARAMETER Message
			The message to be logged.
		.PARAMETER Level
			The type of message to be logged.
		.PARAMETER NoConsoleOut
			SpecIfies to not display the message to the console.
		.PARAMETER ConsoleForeground
			SpecIfies what color the text should be be displayed on the console. Ignored when switch 'NoConsoleOut' is specIfied.
		.PARAMETER Indent
			The number of spaces to indent the line in the log file.
		.PARAMETER Path
			The log file path.	
		.PARAMETER Clobber
			Existing log file is deleted when this is specIfied.
		.PARAMETER EventLogName
			The name of the system event log, e.g. 'Application'.
		.PARAMETER EventSource
			The name to appear as the source attribute for the system event log enTry. This is ignored unless 'EventLogName' is specIfied.
		.PARAMETER EventID
			The ID to appear as the event ID attribute for the system event log enTry. This is ignored unless 'EventLogName' is specIfied.

		.PARAMETER LogEncoding
			The text encoding for the log file. Default is ASCII.
		.EXAMPLE
			PS C:\> Write-Log -Message "It's all good!" -Path C:\MyLog.log -Clobber -EventLogName 'Application'
		.EXAMPLE
			PS C:\> Write-Log -Message "Oops, not so good!" -Level Error -EventID 3 -Indent 2 -EventLogName 'Application' -EventSource "My Script"
		.INPUTS
			System.String
		.OUTPUTS
			No output.
		.NOTES
			Revision History:
				2011-03-10 : Andy Arismendi - Created.
				2011-07-23 : Will Steele - Updated.
				2011-07-23 : Andy Arismendi 
					- Added missing comma in param block. 
					- Added support for creating missing directories in log file path.
	#>
}

Function Write-Header   {
	Write-Host $Line -ForegroundColor Green
	Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "`tRetail Server PostConfig Version $version script Started at " -NoNewline
	Write-Host $Date -ForegroundColor Cyan
	$Message = "Retail Server PostConfig Version $version script Started at $(get-date)"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host
	Write-Host "`t`t`t`tScript Parameters:"
	$Message = "Script Parameters:"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host
	Write-Host "`t`t`tComputer:" -NoNewline
	Write-Host "`t$Computer" -ForegroundColor Cyan
	$Message = "`tComputer: $Computer"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tUser:" -NoNewline
	Write-Host "`t`t$UserName" -ForegroundColor Cyan
	$Message = "`tUser: $UserName"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tEnvironment:" -NoNewline
	Write-Host "`t$Environment" -ForegroundColor Cyan
	$Message = "`tEnvironment: $Environment"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tPlatform:" -NoNewline
	Write-Host "`t$Platform" -ForegroundColor Cyan
	$Message = "`tPlatform: $Platform"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tStartPath:" -NoNewline
	Write-Host "`t$StartPath" -ForegroundColor Cyan
	$Message = "`tStartPath: $StartPath"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tFinalPath:" -NoNewline
	Write-Host "`t$FinalPath" -ForegroundColor Cyan
	$Message = "`tFinalPath: $FinalPath"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Log -Message " "  -Path $logfile -NoConsoleOut
	Write-Host
}

Function Sleep_Host   {
    $Time = $args[0]
    $Minute = $Time / "60"
    Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut    
	Write-Host "                            Script will resume in " -NoNewline
	Write-Host $Minute -ForegroundColor Cyan -NoNewline
	Write-Host " minutes"
    $Message = "Script will resume in $Minute minutes"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut    	
    Write-Host
	Write-log -Message ""  -Path $logfile -NoConsoleOut	
    While ($Time -ne "0")   {
        Write-Host "." -NoNewline
        Start-Sleep -s 5
        $Time = $Time - "5"
    }
	Write-Host "  RESUMING" -ForegroundColor Yellow
    $Message = "  RESUMING"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host
	Write-Host
	Write-log -Message ""  -Path $logfile -NoConsoleOut
	Write-log -Message ""  -Path $logfile -NoConsoleOut
}

Function Check_IPAKInstallation   {
	$IPAKEvents = Get-WinEvent -ComputerName $Computer -LogName $logApplication -MaxEvents $displayRows | where ProviderName -ne SQL
	$i = 0
	$count = $IPAKEvents.Count
	While ($i -le $count)   {
		$Result = $IPAKEvents[$i].Message -like "*Please reboot the server and re-run the Ipak*"
		If ($Result -eq $true)   {
			$Reboot = $Result
			Write-Host "`t`t`tPreparing to reboot " -NoNewline
			Write-Host $Computer -ForegroundColor Cyan
			$Message = "`t`tPreparing to reboot $Computer"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			Restart_Host -Computer $Computer	
			Write-Host "`t`t`t`t$Computer" -ForegroundColor Cyan -NoNewline
			Write-Host " rebooted"
			$Message = "`t`t`t$Computer rebooted"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			$i = $count
			$Check = $false
		}
		Else   {
			$Result = $IPAKEvents[$i].Message -like "*WL IPAK successfully installed*"
			If ($Result -eq $true)   {
				$Complete = $Result
				Write-Host "`t`t IPAK successfully installed on " -ForegroundColor Green -NoNewline
				Write-Host $Computer -ForegroundColor Cyan
				$Message = "IPAK successfully installed on $Computer"
				Write-Log -Message $Message -Path $logfile -NoConsoleOut
				$i = $count
				$Check = $true
			}
			Else   {
				$Result = $IPAKEvents[$i].Message -like "*WL IPAK successfully installed, but a reboot is required to complete the installation.*"
				If ($Result -eq $true)   {
					$FinalReboot = $Result
					Write-Host "`t`t`tPreparing to reboot " -NoNewline
					Write-Host $Computer -ForegroundColor Cyan -NoNewline
					Write-Host " to complete IPAK installation"
					$Message = "`t`tPreparing to reboot $Computer to complete IPAK installation"
					Write-Log -Message $Message -Path $logfile -NoConsoleOut
					Restart_Host -Computer $Computer
					Write-Host "`t`t`t`t$Computer" -ForegroundColor Cyan -NoNewline
					Write-Host " rebooted"
					$Message = "`t`t`t$Computer rebooted"
					Write-Log -Message $Message -Path $logfile -NoConsoleOut
					$i = $count
					$Check = $true					
				}
			}
		}
		$i++
	}
	return $Check
}

Function Check_SQLInstallation   {
	$SQLEvents = Get-WinEvent -ComputerName $Computer -LogName $logApplication -MaxEvents $displayRows | where ProviderName -eq SQL
	$i = 0
	$count = $SQLEvents.Count
	While ($i -le $count)   {
		$Result = $SQLEvents[$i].Message -like "*Please reboot the server and re-run the Ipak*"
		If ($Result -eq $true)   {
			$Reboot = $Result
			Write-Host "`t`t`tPreparing to reboot " -NoNewline
			Write-Host $Computer -ForegroundColor Cyan
			$Message = "`t`tPreparing to reboot $Computer"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			Restart_Host -Computer $Computer	
			Write-Host "`t`t`t`t$Computer" -ForegroundColor Cyan -NoNewline
			Write-Host " rebooted"
			$Message = "`t`t`t$Computer rebooted"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			$i = $count
			$SQLCheck = $false
		}
		Else   {
			$Result = $SQLEvents[$i].Message -like "*WL IPAK successfully installed*"
			If ($Result -eq $true)   {
				$Complete = $Result
				Write-Host "`t`t SQL IPAK successfully installed on " -ForegroundColor Green -NoNewline
				Write-Host $Computer -ForegroundColor Cyan
				$Message = "SQL IPAK successfully installed on $Computer"
				Write-Log -Message $Message -Path $logfile -NoConsoleOut
				$i = $count
				$SQLCheck = $true
			}
			Else   {
				$Result = $SQLEvents[$i].Message -like "*WL IPAK successfully installed, but a reboot is required to complete the installation.*"
				If ($Result -eq $true)   {
					$FinalReboot = $Result
					Write-Host "`t`t`tPreparing to reboot " -NoNewline
					Write-Host $Computer -ForegroundColor Cyan -NoNewline
					Write-Host " to complete SQL IPAK installation"
					$Message = "`t`tPreparing to reboot $Computer to complete SQL IPAK installation"
					Write-Log -Message $Message -Path $logfile -NoConsoleOut
					Restart_Host -Computer $Computer	
					Write-Host "`t`t`t`t$Computer" -ForegroundColor Cyan -NoNewline
					Write-Host " rebooted"
					$Message = "`t`t`t$Computer rebooted"
					Write-Log -Message $Message -Path $logfile -NoConsoleOut
					$i = $count
					$SQLCheck = $true					
				}
			}
		}
		$i++
	}	
	return $SQLCheck
}

Function Check_VersionInstallation   {
	$VersionEvents = Get-WinEvent -ComputerName $Computer -LogName $logApplication -MaxEvents $displayRows | where ProviderName -eq PATCH
	$i = 0
	$count = $VersionEvents.Count
	While ($i -le $count)   {
		$Result = $VersionEvents[$i].Message -like "*Please reboot the server and re-run the Ipak*"
		If ($Result -eq $true)   {
			$Reboot = $Result
			Write-Host "`t`t`tPreparing to reboot " -NoNewline
			Write-Host $Computer -ForegroundColor Cyan
			$Message = "`t`tPreparing to reboot $Computer"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			Restart_Host -Computer $Computer	
			Write-Host "`t`t`t`t$Computer" -ForegroundColor Cyan -NoNewline
			Write-Host " rebooted"
			$Message = "`t`t`t$Computer rebooted"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			$i = $count
			$VersionCheck = $false
		}
		Else   {
			$Result = $VersionEvents[$i].Message -like "*WL IPAK successfully installed*"
			If ($Result -eq $true)   {
				$Complete = $Result
				Write-Host "`t`t SQL version update successfully installed on " -ForegroundColor Green -NoNewline
				Write-Host $Computer -ForegroundColor Cyan
				$Message = "SQL version update successfully installed on $Computer"
				Write-Log -Message $Message -Path $logfile -NoConsoleOut
				$i = $count
				$VersionCheck = $true
			}
			Else   {
				$Result = $VersionEvents[$i].Message -like "*WL IPAK successfully installed, but a reboot is required to complete the installation.*"
				If ($Result -eq $true)   {
					$FinalReboot = $Result
					Write-Host "`t`t`tPreparing to reboot " -NoNewline
					Write-Host $Computer -ForegroundColor Cyan -NoNewline
					Write-Host " to complete SQL version update installation"
					$Message = "`t`tPreparing to reboot $Computer to complete SQL version update installation"
					Write-Log -Message $Message -Path $logfile -NoConsoleOut
					Restart_Host -Computer $Computer	
					Write-Host "`t`t`t`t$Computer" -ForegroundColor Cyan -NoNewline
					Write-Host " rebooted"
					$Message = "`t`t`t$Computer rebooted"
					Write-Log -Message $Message -Path $logfile -NoConsoleOut
					$i = $count
					$VersionCheck = $true					
				}
			}
		}
		$i++
	}
	return $VersionCheck
}

Function Start_IPAK   {
	Write-Host
	Write-log -Message ""  -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
	Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tStarting IPAK" -ForegroundColor Cyan
	$Message = "Starting IPAK"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	psexec \\$Computer -accepteula -h -u "$UserName" -p "$Password" cmd.exe /c "IPAK location"
	Write-Host $Line -ForegroundColor Green
	Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host
	Write-log -Message ""  -Path $logfile -NoConsoleOut
}

Function Start_SQLIPAK   {
	Write-Host
	Write-log -Message ""  -Path $logfile -NoConsoleOut	
	Write-Host $Line -ForegroundColor Green
	Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tStarting SQL IPAK" -ForegroundColor Cyan
	$Message = "Starting SQL IPAK"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	If ($Platform -eq "Azure")   {
		If ($Environment -eq "PROD")   {
			psexec \\$Computer -accepteula -h -u "$UserName" -p "$Password" cmd.exe /c "IPAK Location"
		}
		Else   {
			psexec \\$Computer -accepteula -h -u "$UserName" -p "$Password" cmd.exe /c "IPAK Location"
		}
	}
	If ($Platform -eq "On-Premise")   {
		If ($Environment -eq "PROD")   {
			psexec \\$Computer -accepteula -h -u "$UserName" -p "$Password" cmd.exe /c "IPAK Location"
		}
		Else   {
			psexec \\$Computer -accepteula -h -u "$UserName" -p "$Password" cmd.exe /c "IPAK Location"
		}
	}
	Write-Host $Line -ForegroundColor Green
	Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host
	Write-log -Message ""  -Path $logfile -NoConsoleOut
}

Function Install_VersionUpdate   {
	Write-Host
	Write-log -Message ""  -Path $logfile -NoConsoleOut	
	Write-Host $Line -ForegroundColor Green
	Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tInstalling SQL Version Update" -ForegroundColor Cyan
	$Message = "Installing SQL Version Update"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	psexec \\$Computer -accepteula -h -u "$UserName" -p "$Password" cmd.exe /c "version Location"
	Write-Host $Line -ForegroundColor Green
	Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host
	Write-log -Message ""  -Path $logfile -NoConsoleOut
}

Function Start_Copy   {
	Write-Host
	Write-log -Message ""  -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
	Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "    Starting file copy from " -NoNewline
	Write-Host "`t$StartPath" -ForegroundColor Cyan
	$Message = "    Starting file copy from $StartPath"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Copy-Item $StartPath -Destination $FinalPath -Recurse -Force
	Write-Host "    File copy completed to " -NoNewline
	Write-Host "`t$FinalPath" -ForegroundColor Cyan
	$Message = "    File copy completed to $FinalPath"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
	Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
}

Function Write-Footer   {
	Write-Host
	Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
	$Message = "Script Completed at " + $(get-date)
	Write-log -Message $Message  -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tScript Completed at " -NoNewline
	Write-Host "`t$(get-date)" -ForegroundColor Cyan	
	$Message = "Total Elapsed Time: {0:00}:{1:00}:{2:00}" -F $ElapsedTime.Elapsed.Hours, $ElapsedTime.Elapsed.Minutes, $ElapsedTime.Elapsed.Seconds
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = "Total Elapsed Time:{0:00}:{1:00}:{2:00}" -F $ElapsedTime.Elapsed.Hours, $ElapsedTime.Elapsed.Minutes, $ElapsedTime.Elapsed.Seconds	
	$Message = $Message.Replace("Total Elapsed Time:","")
	Write-Host "`t`t`tTotal Elapsed Time:" -NoNewline
	Write-Host "`t$Message" -ForegroundColor Cyan
	Write-Host "`t`t`tComputer:" -NoNewline
	Write-Host "`t`t$Computer" -ForegroundColor Cyan
	$Message = "`tComputer: $Computer"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tUser:" -NoNewline
	Write-Host "`t`t`t$UserName" -ForegroundColor Cyan
	$Message = "`tUser: $UserName"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tEnvironment:" -NoNewline
	Write-Host "`t`t$Environment" -ForegroundColor Cyan
	$Message = "`tEnvironment: $Environment"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tPlatform:" -NoNewline
	Write-Host "`t`t$Platform" -ForegroundColor Cyan
	$Message = "`tPlatform: $Platform"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tSkipped:" -NoNewline
	Write-Host "`t`t$Skip" -ForegroundColor Cyan
	$Message = "`tSkipped: $Skip"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tStartPath:" -NoNewline
	Write-Host "`t`t$StartPath" -ForegroundColor Cyan
	$Message = "`tStartPath: $StartPath"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host "`t`t`tFinalPath:" -NoNewline
	Write-Host "`t`t$FinalPath" -ForegroundColor Cyan
	$Message = "`tFinalPath: $FinalPath"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
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

Function Configure_Memory   {
    Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "`tConfiguring Min and Max memory settings"
	$Message = "Configuring Min and Max memory settings"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		Try   {
			$totalMem = ((Get-WmiObject -class "Win32_OperatingSystem").TotalVisibleMemorySize)
			Write-Host "`t`tTotal memory detected: " -NoNewline
			Write-Host $totalMem -ForegroundColor Cyan
			$cleanMemMB = ([math]::round($totalMem / 1024, 0))
			Write-Host "`t`tRounded memory in MB: " -NoNewline
			Write-Host $cleanMemMB -ForegroundColor Cyan
			If ($cleanMemMB -gt 24400)   {
				$maxMem = ([math]::Round($cleanMemMB - 8192, 0))
				$minMem = ([math]::Round($maxMem * .8, 0))
				Write-Host "`t`tMemory is sufficient for standard sizing considerations!"
				Write-Host "`t`t`tMax memory(MB) is: " -NoNewline
				Write-Host $maxMem -ForegroundColor Cyan
				Write-Host "`t`t`tMin memory(MB) is: " -NoNewline
				Write-Host $minMem -ForegroundColor Cyan
			}
			Else   {
				$maxMem = ([math]::Round($cleanMemMB * .8, 0))
				$minMem = ([math]::Round($maxMem * .8, 0))
				Write-Host "`t`tMemory is insufficient for standard sizing considerations!" -ForegroundColor Red
				Write-Host "`t`t`tMax memory(MB) is: " -NoNewline
				Write-Host $maxMem -ForegroundColor Red
				Write-Host "`t`t`tMin memory(MB) is: " -NoNewline
				Write-Host $minMem -ForegroundColor Red
			}
			$sqlMem = @"
				Exec sp_configure 'min server memory', $minMem ;
				Exec sp_configure 'max server memory', $maxMem
				GO
				RECONFIGURE
"@
				Invoke-Sqlcmd -query $sqlMem | Write-Output
				Write-Host "`tConfiguring max and minimum memory settings - Complete"
		}
		Catch   {
			Throw "SQL Memory failed to configure properly $error!"
		}
	}
	$totalMem = ((Get-WmiObject -class "Win32_OperatingSystem").TotalVisibleMemorySize)	
	$cleanMemMB = ([math]::round($totalMem / 1024, 0))	
	$Message = "`tTotal memory detected: $totalMem"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = "`tRounded memory in MB: $cleanMemMB"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	
	If ($cleanMemMB -gt 24400)   {
		$maxMem = ([math]::Round($cleanMemMB - 8192, 0))
		$minMem = ([math]::Round($maxMem * .8, 0))	
		$Message = "`tMemory is sufficient for standard sizing considerations!"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut
		$Message = "`tMax memory(MB) is: $maxMem"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut
		$Message = "`tMin memory(MB) is: $minMem"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut
	}
	Else   {
		$maxMem = ([math]::Round($cleanMemMB * .8, 0))
		$minMem = ([math]::Round($maxMem * .8, 0))
		$Message = "`tMemory is insufficient for standard sizing considerations!"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut
		$Message = "`tMax memory(MB) is: $maxMem"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut
		$Message = "`tMin memory(MB) is: $minMem"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut
	}
	$Message = "Configuring max and minimum memory settings - Complete"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host
}

Function Configure_MDOP   {
    Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "`tConfiguring Max Degree of Parallelism"
	$Message = "Configuring Max Degree of Parallelism"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		Try   {
			[int]$totalProc = (Get-WmiObject -class "Win32_computersystem").numberoflogicalprocessors
			Write-Host "`t`tLogical Processors Detected: " -NoNewline
			Write-Host $totalProc -ForegroundColor Cyan
			If ($totalProc -le 3)   {
				$halfProcCount = 1
			}
			# hard coded to 1 due to device not being reporting service
			Else   {
				#$halfProcCount = $totalProc/2
				$halfProcCount = 1
				If ($halfProcCount -gt 8)   {
					# $halfProcCount = 8
					$halfProcCount = 1
				}
			}
			$MDOPRPT = @"
			sp_configure 'max degree of parallelism', $halfProcCount ;   
			GO
			RECONFIGURE 
"@
			Write-Host "`t`tRunning MDOP reporting workload query"
			Invoke-Sqlcmd -query $MDOPRPT
			Write-Host "`tConfiguring Max Degree of Parallelism - Complete"
		}
		Catch   {
			Throw "MDOP Reporting failed to configure properly!"
		}
	}
	[int]$totalProc = (Get-WmiObject -class "Win32_computersystem").numberoflogicalprocessors	
	$Message = "`tLogical Processors Detected: $totalProc"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = "`tRunning MDOP reporting workload query"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = "Configuring Max Degree of Parallelism - Complete"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host
}

Function Configure_ServerAgent   {
    Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "`tConfiguring SQLServerAgent Properties"
	$Message = "Configuring SQLServerAgent Properties"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		Try   {
			$SQLAgent = @"
			exec msdb.dbo.sp_set_SQLagent_properties
			@SQLserver_restart           	= 1, -- Auto restart SQL Server If it stops unexpectedly.
			@jobhistory_max_rows         	= 100000, -- Max job history logsize in no. of rows(stored in msdb.dbo.SysJobHistory).
			@jobhistory_max_rows_per_job = 1000,		-- Max job history rows per job.
			@monitor_autostart           	= 1 -- Auto restart SQLServerAgent If it stops unexpectedly.
"@
			Write-Host "`t`tRunning SQL Agent query"
			Invoke-Sqlcmd -query $SQLAgent
			Write-Host "`tConfiguring SQLServerAgent Properties - Complete"
		}
		Catch   {
			Throw "SQL Agent failed to configure properly!"
		}
	}
	$Message = "`tRunning SQL Agent query"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut	
	$Message = "Configuring SQLServerAgent Properties - Complete"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut	
	Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host
}

Function Config_BackupCompression   {
    Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "`tConfiguring Backup Compression"
	$Message = "Configuring Backup Compression"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		Try   {
			$backUpComp =@"
			sp_configure 'backup compression default', 1
			GO
			RECONFIGURE
"@
			Write-Host "`t`tRunning backup compression query"
			Invoke-Sqlcmd -query $backUpComp
			Write-Host "`tConfiguring Backup Compression - Complete"
		}
		Catch   {
			Throw "Backup compression failed to configure properly!"
		}
	}
	$Message = "`tRunning backup compression query"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = "Configuring Backup Compression - Complete"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host
}

Function Configure_Databases   {
    Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "`tConfiguring Databases (master, msdb, and model)"
	$Message = "Configuring Databases (master, msdb, and model)"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		Try   {
			$SQLSystem = @"
			Alter database [model] set recovery simple with no_wait

			Alter database [master]
			modIfy file (Name = [master], Size = 50 MB)
			Go
			Alter database [master]
			modIfy file (Name = [master], Filegrowth = 5 MB)
			Go
			Alter database [master]
			modIfy file (Name = [mastLog], Size = 20 MB)
			Go
			Alter database [master]
			modIfy file (Name = [mastLog], Filegrowth = 5 MB)
			Go

			Alter database [msdb]
			modIfy file (Name = [msdbData], Size = 50 MB)
			Go
			Alter database [msdb]
			modIfy file (Name = [msdbData], Filegrowth = 5 MB)
			Go
			Alter database [msdb]
			modIfy file (Name = [msdbLog], Size = 30 MB)
			Go
			Alter database [msdb]
			modIfy file (Name = [msdbLog], Filegrowth = 5 MB)
			Go

			Alter database [model]
			modIfy file (Name = [modeldev], Size = 50 MB)
			Go
			Alter database [model]
			modIfy file (Name = [modeldev], Filegrowth = 5 MB)
			Go
			Alter database [model]
			modIfy file (Name = [modellog], Size = 20 MB)
			Go
			Alter database [model]
			modIfy file (Name = [modellog], Filegrowth = 5 MB)
			Go
"@
			Write-Host "`t`tRunning system DB query"
			Invoke-Sqlcmd -query $SQLSystem -ErrorAction SilentlyContinue
			Write-Host "`tConfiguring Databases (master, msdb, and model) - Complete"
		}
		Catch   {
			Throw "System DBs failed to configure properly!"
		}
	}
	$Message = "`tRunning system DB query"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = "Configuring Databases (master, msdb, and model) - Complete"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host
}

Function Configure_Tempdb   {
    Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "`tConfiguring Tempdb"
	$Message = "Configuring Tempdb"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		Try   {
			[int]$totalProc = (Get-WmiObject -class "Win32_computersystem").numberoflogicalprocessors
			Write-Host "`t`tTotal Processors: " -NoNewline
			Write-Host $totalProc -ForegroundColor Cyan
			Write-Host "`t`tCreating " -NoNewline
			Write-Host "fileName" -ForegroundColor Cyan
			new-item "directoryName" -type directory | Out-Null
			new-item "directoryName" -type directory | Out-Null
			new-item "directoryName" -type directory | Out-Null
			Write-Host "`t`tMoving " -NoNewline
			Write-Host "fileName" -ForegroundColor Cyan -NoNewline
			Write-Host "and " -NoNewline
			Write-Host "fileName" -ForegroundColor Cyan -NoNewline
			Write-Host "to " -NoNewline
			Write-Host "fileName" -ForegroundColor Cyan

			$SQLTemp = @"
			USE [master]
			GO
			ALTER DATABASE [tempdb] MODIFY FILE ( NAME = tempdev, FILENAME = 'fileName')
			ALTER DATABASE [tempdb] MODIFY FILE ( NAME = templog, FILENAME = 'fileName');
			GO
"@
			Write-Host "`t`tRunning Tempdb query"
			Invoke-Sqlcmd -query $SQLTemp -ErrorAction SilentlyContinue

			Get-Service -Name MSSQLSERVER -ComputerName $Computer | Restart-Service -Force
			Start-Sleep 10
			Get-Service -Name SQLServerAgent -ComputerName $Computer | Restart-Service -Force
			Start-Sleep 10
		}
		Catch   {
			Throw "Temp DB failed to configure properly!"
		}
	}

<#	
	[int]$totalProc = (Get-WmiObject -class "Win32_computersystem").numberoflogicalprocessors	
	$Message = "`tTotal Processors: $totalProc"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = "`tCreating fileName"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = "`tMoving tempdb.mdf and tempdb.ldf to fileName"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = "`tRunning Tempdb query"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$i = 1
	While ($i -lt $totalProc)   {
		$tempname = "tempdev$i"
		$tempfile = "tempdb$i.ndf"
		$tempfileName = "tempdb$i.ndf"		
		$Message = "`tMoving $tempname to fileName"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut
		$i++
	}
	$Message = "Configuring Tempdb - Complete"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host
#>	
	
}

Function Configure_Tempfiles   {
    Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "`tConfiguring Tempdb files"
	$Message = "Configuring Tempdb files"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		Try   {
			$i = 1
			[int]$totalProc = (Get-WmiObject -class "Win32_computersystem").numberoflogicalprocessors			
			While ($i -lt $totalProc)   {
				$tempname = "tempdev$i"
				$tempfile = "tempdb$i"
				$tempfileName = "tempdb$i.ndf"
				Write-Host "`t`tMoving " -NoNewline
				Write-Host $tempname -ForegroundColor Cyan -NoNewline
				Write-Host " to " -NoNewline
				Write-Host "fileName" -ForegroundColor Cyan

				$SQLTempDev = @"
				USE [master]
				ALTER DATABASE [tempdb] MODIFY FILE ( NAME = N'$tempname', FILENAME = 'fileName')
"@
				$i++
				Invoke-Sqlcmd -query $SQLTempDev -ErrorAction SilentlyContinue
			}

			Write-Host "`tConfiguring Tempdb - Complete"
		}
		Catch   {
			Throw "Temp DB failed to configure properly!"
		}
	}
	[int]$totalProc = (Get-WmiObject -class "Win32_computersystem").numberoflogicalprocessors	
	$Message = "`tTotal Processors: $totalProc"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = "`tCreating fileName"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = "`tMoving fileName and fileName to fileLocation"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = "`tRunning Tempdb query"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$i = 1
	While ($i -lt $totalProc)   {
		$tempname = "tempdev$i"
		$tempfile = "fileName"
		$tempfileName = "fileName"		
		$Message = "`tMoving $tempname to fileLocation"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut
		$i++
	}
	$Message = "Configuring Tempdb - Complete"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host
}

Function Remove_Files   {
    Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "`t`tRemoving folder: " -NoNewline
	Write-Host $Folder -ForegroundColor Cyan
	$Message = "`tRemoving folder: $Folder"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut	
	Get-Service -Name "serviceName" -ComputerName $Computer | Stop-Service -Force
	Get-Service -Name "serviceName" -ComputerName $Computer | Stop-Service -Force
	Remove-Item $Folder -Recurse -Force
	Get-Service -Name "serviceName" -ComputerName $Computer | Start-Service
	Get-Service -Name "serviceName" -ComputerName $Computer | Start-Service
	Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host
}

Function Setup_RegKey   {
    Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "`tSetting up RegKey"
	$Message = "Setting up RegKey"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$TempSession = New-PSSession -ComputerName $Computer -Credential $cred
	Invoke-Command -Session $TempSession -ScriptBlock {New-ItemProperty -Path "RegLocation" -Name NumErrorLogs -PropertyType DWord -Value 30}
	Write-Host "`tRegKey setup completed"
	$Message = "RegKey setup completed"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host
}

Function Restart_Service   {
    Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host "`tRestarting Services"
	$Message = "Restarting Services"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host "`t`tRestarting SQL Services: " -NoNewline
	Write-Host "serviceName" -ForegroundColor Cyan -NoNewline
	Write-Host " and " -ForegroundColor White -NoNewline
	Write-Host "serviceName" -ForegroundColor Cyan
	$Message = "`tRestarting SQL Services: serviceName and serviceName"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Get-Service -Name "serviceName" -ComputerName $Computer | Restart-Service -Force
	Sleep_Host 30
	Get-Service -Name "serviceName" -ComputerName $Computer | Restart-Service -Force
	Sleep_Host 30
	Write-Host "`tServices Restarted"
	$Message = "Services Restarted"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host
}

Function Create_RegKey   {
	$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)
	$RegKey = $Reg.OpenSubKey($InstallationStatusRegPath)
	If (!($RegKey))   {
	    Write-Host $Line -ForegroundColor Green
		Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
		Write-Host "`tCreating new RegKey tracker"
		$Message = "Creating new RegKey tracker"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut	
		$TempSession = New-PSSession -ComputerName $Computer -Credential $cred
		Invoke-Command -Session $TempSession -ScriptBlock {New-Item -Path "RegLocation"}
		Invoke-Command -Session $TempSession -ScriptBlock {New-Item -Path "RegLocation"}
		Invoke-Command -Session $TempSession -ScriptBlock {New-ItemProperty -Path "RegLocation" -Name Status -PropertyType String -Value 1}
		Set-RegKeyStatus 1
		$InstallationStatus = Get-RegKeyStatus
		Write-Host "`tRegKey tracker creation completed"
		$Message = "RegKey tracker creation completed"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut
		Write-Host $Line -ForegroundColor Green
		Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
		Write-Host
	}
}

Function Get-RegKeyStatus   {
	$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)
	$RegKey = $Reg.OpenSubKey($InstallationStatusRegPath)
	If ($RegKey)   {
		$RegTracker = $RegKey.GetValue($InstallationStatusRegKey)
		return $RegTracker
	}
	Else   {
		Write-Host "RegKey doesn't exist" -ForegroundColor Red
		Create_RegKey
	}
}

Function Set-RegKeyStatus   {
	$RegValue = $args
	$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)
	$RegKey = $Reg.OpenSubKey($InstallationStatusRegPath)
	If ($RegKey)   {
		$TempSession = New-PSSession -ComputerName $Computer -Credential $cred
		Invoke-Command -Session $TempSession -ScriptBlock {Set-ItemProperty -Path "RegLocation" -Name Status -Value $Using:RegValue}	
	}
	Else   {
		Write-Host "RegKey doesn't exist" -ForegroundColor Red
		Create_RegKey
		Set-RegKeyStatus $RegValue
	}
}

Function Remove_LogFiles   {
    Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut

	Invoke-command -computername $FQDN -credential $cred -scriptblock   {
		[int]$totalProc = (Get-WmiObject -class "Win32_computersystem").numberoflogicalprocessors	
		$i = 1
		While ($i -lt $totalProc)   {
			$tempname = "fileName"
			Write-Host "`t`tRemoving " -NoNewline
			Write-Host $tempname -ForegroundColor Cyan
			$SQLRemove = @"
			ALTER DATABASE [tempdb] REMOVE FILE $tempname
"@
			$i++
			Invoke-Sqlcmd -query $SQLRemove -ErrorAction SilentlyContinue
		}
	}
	[int]$totalProc = (Get-WmiObject -class "Win32_computersystem").numberoflogicalprocessors
	$i = 1
	While ($i -lt $totalProc)   {
		$tempname = "fileName"
		$Message = "`tRemoving $tempname"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut
		$i++
	}
	Write-Host $Line -ForegroundColor Green
	Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
	Write-Host	
}

Function Get-SQLVersion   {
	$TempSession = New-PSSession -ComputerName $Computer -Credential $cred
	$PatchLevel = Invoke-Command -Session $TempSession -ScriptBlock {(Get-ItemProperty "RegLocation").Patchlevel}
	If ( $PatchLevel -eq "11.2.5343.0")   {
		$SQLVersion = $true
		return $SQLVersion
	}
	Else   {
		$SQLVersion = $false
		Write-Host "`t`tSQL Version is: " -NoNewline
		Write-Host $PatchLevel -ForegroundColor Red
		Write-Host "`t`tSQL Version should be: " -NoNewline
		Write-Host "11.2.5343.0" -ForegroundColor Cyan
		$Message = "`tSQL Version is: $PatchLevel"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut
		$Message = "`tSQL Version should be: 11.2.5343.0"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut		
		return $SQLVersion
	}
}

Function Run_Switch   {
	$InstallationStatus = $args

	Switch($InstallationStatus)   {

		{$_ -eq 1}   {
			Write-Host
			Write-log -Message ""  -Path $logfile -NoConsoleOut
			Write-Host $Line -ForegroundColor Yellow
			Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
			Write-Host "`t`t`tBeginning IPAK"
			$Message = "`t`tBeginning IPAK"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			Write-Host $Line -ForegroundColor Yellow
			Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
			Start_IPAK

			While (!($Check))   {
				Write-Host "`t`t`tChecking IPAK Status"
				$Message = "Checking IPAK Status"
				Write-Log -Message $Message -Path $logfile -NoConsoleOut
				$Check = Check_IPAKInstallation

				If ($Check -eq $false)   {
					Sleep_Host 15
					Start_IPAK
				}
			}
			Write-Host $Line -ForegroundColor Yellow
			Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
			Write-Host "`t`t`tIPAK Complete"
			$Message = "`t`tIPAK Complete"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			Write-Host $Line -ForegroundColor Yellow
			Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
			Write-Host
			Write-log -Message ""  -Path $logfile -NoConsoleOut
			Set-RegKeyStatus 2
			$InstallationStatus = Get-RegKeyStatus
			Run_Switch $InstallationStatus
		}

		{$_ -eq 2}   {
			Write-log -Message ""  -Path $logfile -NoConsoleOut
			Write-Host $Line -ForegroundColor Yellow
			Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
			Write-Host "`t`t`tBeginning SQL IPAK"
			$Message = "`t`tBeginning SQL IPAK"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			Write-Host $Line -ForegroundColor Yellow
			Write-log -Message "============================================"  -Path $logfile -NoConsoleOut	
			Start_SQLIPAK

			While (!($SQLCheck))   {
				Write-Host "`t`t`tChecking SQL IPAK Status"
				$Message = "Checking SQL IPAK Status"
				Write-Log -Message $Message -Path $logfile -NoConsoleOut
				$SQLCheck = Check_SQLInstallation
				
				If ($SQLCheck -eq $false)   {
					Sleep_Host 15
					Start_SQLIPAK
				}
			}
			Write-Host $Line -ForegroundColor Yellow
			Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
			Write-Host "`t`t`tSQL IPAK Complete"
			$Message = "`t`tSQL IPAK Complete"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			Write-Host $Line -ForegroundColor Yellow
			Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
			Write-Host
			Write-log -Message ""  -Path $logfile -NoConsoleOut
			Set-RegKeyStatus 3
			$InstallationStatus = Get-RegKeyStatus
			Run_Switch $InstallationStatus
		}

		{$_ -eq 3}   {
			If (Get-SQLVersion)   {
				Set-RegKeyStatus 5
			$InstallationStatus = Get-RegKeyStatus			
			Run_Switch $InstallationStatus
			}
			Else   {
				Set-RegKeyStatus 4
			$InstallationStatus = Get-RegKeyStatus			
			Run_Switch $InstallationStatus
			}
		}

		{$_ -eq 4}   {
			Write-log -Message ""  -Path $logfile -NoConsoleOut
			Write-Host $Line -ForegroundColor Yellow
			Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
			Write-Host "`t`t`tBeginning SQL Version Update"
			$Message = "`t`tBeginning SQL Version Update"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			Write-Host $Line -ForegroundColor Yellow
			Write-log -Message "============================================"  -Path $logfile -NoConsoleOut		
			Install_VersionUpdate

			While (!($VersionCheck))   {
				Write-Host "`t`t`tChecking SQL Version Update Status"
				$Message = "Checking SQL Version Update Status"
				Write-Log -Message $Message -Path $logfile -NoConsoleOut
				$VersionCheck = Check_VersionInstallation
				
				If ($VersionCheck -eq $false)   {
					Sleep_Host 15
					Install_VersionUpdate
				}
			}
			Write-Host $Line -ForegroundColor Yellow
			Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
			Write-Host "`t`t`tSQL Version Update Complete"
			$Message = "`t`tSQL Version Update Complete"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			Write-Host $Line -ForegroundColor Yellow
			Write-Log -Message "============================================"  -Path $logfile -NoConsoleOut
			Write-Host
			Write-log -Message ""  -Path $logfile -NoConsoleOut		
			Set-RegKeyStatus 5
			$InstallationStatus = Get-RegKeyStatus			
			Run_Switch $InstallationStatus
		}
		
		{$_ -eq 5}   {	
			Start_Copy

			Sleep_Host 15

			Write-Host "`t`t`tPreparing to reboot " -NoNewline
			Write-Host $Computer -ForegroundColor Cyan
			$Message = "`t`tPreparing to reboot $Computer"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			Restart_Host -Computer $Computer
			Write-Host "`t`t`t`t$Computer" -ForegroundColor Cyan -NoNewline
			Write-Host " rebooted"
			$Message = "`t`t`t$Computer rebooted"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			
			Set-RegKeyStatus 6
			$InstallationStatus = Get-RegKeyStatus			
			Run_Switch $InstallationStatus
		}
		
		{$_ -eq 6}   {
			Configure_Memory
			Set-RegKeyStatus 7
			$InstallationStatus = Get-RegKeyStatus			
			Run_Switch $InstallationStatus
		}
		
		{$_ -eq 7}   {
			Configure_MDOP
			Set-RegKeyStatus 8
			$InstallationStatus = Get-RegKeyStatus			
			Run_Switch $InstallationStatus
		}
		
		{$_ -eq 8}   {		
			Config_BackupCompression
			Set-RegKeyStatus 9
			$InstallationStatus = Get-RegKeyStatus			
			Run_Switch $InstallationStatus
		}

		{$_ -eq 9}   {
			Configure_Databases
			Set-RegKeyStatus 10		
			$InstallationStatus = Get-RegKeyStatus			
			Run_Switch $InstallationStatus
		}

		{$_ -eq 10}   {
			Configure_Tempdb
			Set-RegKeyStatus 11
			$InstallationStatus = Get-RegKeyStatus			
			Run_Switch $InstallationStatus
		}
		
		{$_ -eq 11}   {
			Configure_Tempfiles
			Set-RegKeyStatus 12
			$InstallationStatus = Get-RegKeyStatus			
			Run_Switch $InstallationStatus
		}
		
		{$_ -eq 12}   {
			Restart_Service
			Set-RegKeyStatus 13
			$InstallationStatus = Get-RegKeyStatus			
			Run_Switch $InstallationStatus
		}

		{$_ -eq 13}   {
			Remove_Files
			Set-RegKeyStatus 14
			$InstallationStatus = Get-RegKeyStatus			
			Run_Switch $InstallationStatus
		}
		
		{$_ -eq 14}   {
			Remove_LogFiles
			Restart_Service
			Set-RegKeyStatus 15
			$InstallationStatus = Get-RegKeyStatus			
			Run_Switch $InstallationStatus
		}
		
		{$_ -eq 15}   {
			Setup_RegKey
			Set-RegKeyStatus 16
			$InstallationStatus = Get-RegKeyStatus			
			Run_Switch $InstallationStatus
		}

		{$_ -eq 16}   {
			Configure_ServerAgent
			Set-RegKeyStatus 17
			$InstallationStatus = Get-RegKeyStatus			
		}
			
		{$_ -eq 17}   {
			Restart_Service
			Set-RegKeyStatus 18
			$InstallationStatus = Get-RegKeyStatus			
		}
	}
}

#EndRegion Functions

#Region Main Processing

Test-Administrator

$Password = Get_Password

Write-Header

Create_RegKey

If (!([string]::IsNullOrWhiteSpace($Skip)))   {
	If ($Skip -eq "IPAK")   {
		Set-RegKeyStatus 2
		$InstallationStatus = Get-RegKeyStatus
		Run_Switch $InstallationStatus
	}
	If ($Skip -eq "SQL")   {
		Start_IPAK

		While (!($Check))   {
			Write-Host "`t`t`tChecking IPAK Status"
			$Message = "Checking IPAK Status"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			$Check = Check_IPAKInstallation

			If ($Check -eq $false)   {
				Sleep_Host 15
				Start_IPAK
			}
		}
		Set-RegKeyStatus 3
		$InstallationStatus = Get-RegKeyStatus
		Run_Switch $InstallationStatus
	}
	If ($Skip -eq "Both")   {
		Set-RegKeyStatus 3
		$InstallationStatus = Get-RegKeyStatus
		Run_Switch $InstallationStatus
	}
}

$InstallationStatus = Get-RegKeyStatus

Run_Switch $InstallationStatus

Write-Footer
Remove-PSSession *

#EndRegion Main Processing