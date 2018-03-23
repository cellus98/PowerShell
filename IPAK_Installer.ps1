<# 
	.SYNOPSIS
    	Installs the latest IPAK required to patch a Retail Server.  And also installs SQL via SQL IPAK along with
			default values used for setting up a new Retail Server.
	.Description
		This script will install IPAK and SQL on the host given.  The script installs SQL via SQL IPAK along with 
			the default values assocaited to SQL.
	.PARAMETER Computer
		The funcitonal definition of which device to IPAK.
	.PARAMETER Environment
		The funcitonal definition of which environment to run the script as (PPE or PROD).
	.PARAMETER Platform
		The functional definition of where the device is located (On-Premise or Azure).
			On-Premise: Device is in Microsoft Store (FL or SS)
			Azure: Device is in Azure
	.INPUTS
		None. You cannot pipe objects to IPAK_Installer.ps1.
	.OUTPUTS
		Console and log file
	.NOTES
		Author: Marcellus Seamster Jr
	.EXAMPLE
		IPAK_Installer.ps1 -Computer computerName
    .EXAMPLE
		IPAK_Installer.ps1 -Computer computerName -Environment PPE -Platform On-Premise
#>

Param( $Computer ,[ValidateSet('Prod','PPE')]$Environment = 'PPE', [ValidateSet('On-Premise','Azure')]$Platform)

#Region Variable declartion
	#logging variables
	$logdate = "{0:MMddyyyy}" -f (Get-Date)
	$logpath = "fileName"
	$logfile = $logpath + "IPAK_Installer_" + $Computer + "_" + $logdate + ".log"
	
	$Result = @()
	$FinalReboot = @()
	$Complete = @()
	$Reboot = @()
	$count = @()
	$Password = @()
	
	$WarningColor = (Get-Host).PrivateData
	$WarningColor.WarningBackgroundColor = "yellow"
	$WarningColor.WarningForegroundColor = "blue"
	
#EndRegion Variable declaration

#Region Get control of powershell UI
    $a = (Get-Host).UI.RawUI
    $a.WindowTitle = "IPAK Installer"
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
	$Environment = $Environment.toUpper()
	If([string]::IsNullOrWhiteSpace($Computer))   {
		Write-Host "`t`tInvalid parameter(s) entered. See: " -ForegroundColor White -NoNewline
		Write-Host "Get-Help .\restart_script.ps1 -Full" -ForegroundColor Cyan
		Write-Host
		Exit
	}
	Else   {
		If(Test-Connection -ComputerName $computer -Quiet -Count 1)   {
			$Computer = $Computer.toUpper()
		}
		Else   {
			$Computer = $Computer.toUpper()
			Write-Host "`t`tVerify " -ForegroundColor White -NoNewline
			Write-Host $Computer -ForegroundColor Cyan -NoNewline
			Write-Host " is online and rerun script!" -ForegroundColor White
			Write-Host
			Exit			
		}
	}
	$User = $env:username
	$Domain = $env:userdomain
	$UserName = $Domain + "\" + $User
	if($Platform -eq "Azure")   {
		$FinalPath = "\\" + $Computer + "driveName"
	}
	else   {
		$FinalPath = "\\" + $Computer + "driveName"
	}
	$StartPath = "fileName"
	$logApplication = "MSNIPAK"
	$Reboot = $false
	$Complete = $false
	$Check = $false
	$SQLCheck = $false
	$Line = "================================================================================"
		

#EndRegion Set startup variables

#Region WorkFlows

WorkFlow Restart_Host {
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
		try {			
			$msg = '{0}{1} : {2} : {3}' -f (" " * $Indent), (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level.ToUpper(), $Message
			if ($NoConsoleOut -eq $false) {
				switch ($Level) {
					'Error' { Write-Error $Message }
					'Warn' { Write-Warning $Message }
					'Info' { Write-Host ('{0}{1}' -f (" " * $Indent), $Message) -ForegroundColor $ConsoleForeground}
				}
			}
			if (-not $Path.Exists) {
				New-Item -Path $Path.FullName -ItemType File -Force | Out-Null
			}
			if ($Clobber) {
				$msg | Out-File -FilePath $Path -Encoding $LogEncoding -Force
			} else {
				$msg | Out-File -FilePath $Path -Encoding $LogEncoding -Append
			}
			if ($EventLogName) {
				if (-not $EventSource) {
					$EventSource = ([IO.FileInfo] $MyInvocation.ScriptName).Name
				}
				If (-not [Diagnostics.EventLog]::SourceExists($EventSource)) { 
					[Diagnostics.EventLog]::CreateEventSource($EventSource, $EventLogName) 
		        } 
				$log = New-Object System.Diagnostics.EventLog  
			    $log.set_log($EventLogName)  
			    $log.set_source($EventSource) 
				switch ($Level) {
					"Error" { $log.WriteEntry($Message, 'Error', $EventID) }
					"Warn"  { $log.WriteEntry($Message, 'Warning', $EventID) }
					"Info"  { $log.WriteEntry($Message, 'Information', $EventID) }
				}
			}
		} catch {
			throw "Failed to create log entry in: ‘$Path’. The error was: ‘$_’."
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
			Specifies to not display the message to the console.
		.PARAMETER ConsoleForeground
			Specifies what color the text should be be displayed on the console. Ignored when switch 'NoConsoleOut' is specified.
		.PARAMETER Indent
			The number of spaces to indent the line in the log file.
		.PARAMETER Path
			The log file path.	
		.PARAMETER Clobber
			Existing log file is deleted when this is specified.
		.PARAMETER EventLogName
			The name of the system event log, e.g. 'Application'.
		.PARAMETER EventSource
			The name to appear as the source attribute for the system event log entry. This is ignored unless 'EventLogName' is specified.
		.PARAMETER EventID
			The ID to appear as the event ID attribute for the system event log entry. This is ignored unless 'EventLogName' is specified.

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
	Write-Host "`tIPAK Installer Version $version script Started at " -NoNewline
	Write-Host $Date -ForegroundColor Cyan
	$Message = "IPAK Installer Version $version script Started at $(get-date)"
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
	While($i -le $count)   {
		$Result = $IPAKEvents[$i].Message -like "*Please reboot the server and re-run the Ipak*"
		If ($Result -eq $true)   {
			$Reboot = $Result
			Write-Host "`t`t`tPreparing to reboot " -NoNewline
			Write-Host $Computer -ForegroundColor Cyan
			$Message = "Preparing to reboot $Computer"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			Restart_Host -Computer $Computer	
			Write-Host "`t`t`t$Computer" -ForegroundColor Cyan -NoNewline
			Write-Host " rebooted"
			$Message = "$Computer rebooted"
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
					$Message = "Preparing to reboot $Computer"
					Write-Log -Message $Message -Path $logfile -NoConsoleOut
					Restart_Host -Computer $Computer
					Write-Host "`t`t`t$Computer" -ForegroundColor Cyan -NoNewline
					Write-Host " rebooted"
					$Message = "$Computer rebooted"
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
	While($i -le $count)   {
		$Result = $SQLEvents[$i].Message -like "*Please reboot the server and re-run the Ipak*"
		If ($Result -eq $true)   {
			$Reboot = $Result
			Write-Host "`t`t`tPreparing to reboot " -NoNewline
			Write-Host $Computer -ForegroundColor Cyan
			$Message = "Preparing to reboot $Computer"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			Restart_Host -Computer $Computer	
			Write-Host "`t`t`t$Computer" -ForegroundColor Cyan -NoNewline
			Write-Host " rebooted"
			$Message = "$Computer rebooted"
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
					Restart_Host -Computer $Computer	
					Write-Host "`t`t`t$Computer" -ForegroundColor Cyan -NoNewline
					Write-Host " rebooted"
					$Message = "$Computer rebooted"
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
			psexec \\$Computer -accepteula -h -u "$UserName" -p "$Password" cmd.exe "IPAK location"
		}
		Else   {
			psexec \\$Computer -accepteula -h -u "$UserName" -p "$Password" cmd.exe /c "IPAK location"
		}
	}
	If ($Platform -eq "On-Premise")   {
		If ($Environment -eq "PROD")   {
			psexec \\$Computer -accepteula -h -u "$UserName" -p "$Password" cmd.exe /c "IPAK location"
		}
		Else   {
			psexec \\$Computer -accepteula -h -u "$UserName" -p "$Password" cmd.exe /c "IPAK location"
		}
	}

<#	
	If ($Platform -eq "Specialty")   {
		If ($Environment -eq "PROD")   {
			psexec \\$Computer -accepteula -h -u "$UserName" -p "$Password" cmd.exe /c "IPAK location"
		}
		Else   {
			psexec \\$Computer -accepteula -h -u "$UserName" -p "$Password" cmd.exe /c "IPAK location"
		}
	}	
#>	
	
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
    Verify_Admin $Admin
}

Function Verify_Admin   {
    $Verify = $args[0]
    If (!$Verify)   {
        Write-Host $Line -ForegroundColor Yellow
		Write-Host "`tClose down this window and run script from Administrator prompt" -ForegroundColor Magenta
        Write-Host $Line -ForegroundColor Yellow
		Write-Host
	 	Exit
    }
}

Function Get_Password   {
	$cred = Get-Credential -Credential $UserName
	$bstr = "default password"
	$Password =  "password"
	return $Password
}

#EndRegion Functions

#Region Main Processing

Test-Administrator

$Password = Get_Password

Write-Header

Write-Host
Write-log -Message ""  -Path $logfile -NoConsoleOut
Write-Host $Line -ForegroundColor Green
Write-log -Message "============================================"  -Path $logfile -NoConsoleOut

Start_IPAK

While (!($Check))   {
	Write-Host "`t`t`tChecking IPAK Status"
	$Message = "Checking IPAK Status"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Check = Check_IPAKInstallation

	If ($Check -eq $false)   {
		Sleep_Host 10
		Start_IPAK
	}
}

Write-Host $Line -ForegroundColor Green
Write-log -Message "============================================"  -Path $logfile -NoConsoleOut
Write-Host
Write-log -Message ""  -Path $logfile -NoConsoleOut
Write-Host $Line -ForegroundColor Green
Write-log -Message "============================================"  -Path $logfile -NoConsoleOut

Start_SQLIPAK

While (!($SQLCheck))   {
	Write-Host "`t`t`tChecking SQL IPAK Status"
	$Message = "Checking SQL IPAK Status"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$SQLCheck = Check_SQLInstallation
	
	If ($SQLCheck -eq $false)   {
		Sleep_Host 10
		Start_SQLIPAK
	}
}

Write-Host $Line -ForegroundColor Green
Write-log -Message "============================================"  -Path $logfile -NoConsoleOut

Start_Copy
Write-Footer

#EndRegion Main Processing