<# 
	.SYNOPSIS
    Update XML file required for use by NSO script.  And also updates input file required to setup DHCP scopes.
	
	.Description
	This script will update the XML used within the NSO script, and update CSV to update the DHCP scopes.
	Given a StoreCode (IE: ARL) the script will update data in the XML as well as CSV.
	The script will automatically pull information from an IP spreadsheet and update the XML and CSV,
		the only information required during the script will be location and time zone.
	
	.PARAMETER StoreCode
	The funcitonal definition of which store to update.

	.PARAMETER StoreLocation
	The funcitonal definition of where the store is located.		
		
	.PARAMETER StoreTimeZone
	The funcitonal definition of the stores time zone.
	
    .NOTES
		Author: Marcellus Seamster Jr
		
	.EXAMPLE
	Add_Security_Group.ps1
#>


#Region Variable declartion
	$Device = @()
	$SecurityGroup = @()
	$SG = @()
	$Domain = @()
	$group = @()
	$tracker = 0
	
	$Date = Get-Date -Format 'MM_dd_yyyy'

	#logging variables
	$logdate = "{0:MMddyyyy}" -f (Get-Date)
	$logpath = " "
	$logfile = $logpath + "Add_Security_Group_" + $logdate + ".log"	
	
	$Line = "=========================================================================="	
	
#EndRegion Variable declaration


#Region Get control of powershell UI
	$a = (Get-Host).UI.RawUI
	$a.WindowTitle = "Add Security Group"
	$a.BackgroundColor = "Black"
	$a.ForegroundColor = "White"
	$WarningColor = (Get-Host).PrivateData
	$WarningColor.WarningBackgroundColor = "yellow"
	$WarningColor.WarningForegroundColor = "blue"	
	cls

	$version = "1.0"
	$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()
	
#EndRegion Get control of powershell UI

 
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
				if(-not [Diagnostics.EventLog]::SourceExists($EventSource)) { 
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

Function Check_Admin   {
	$group = Get-Localadmin
	Write-Host $group -ForegroundColor Magenta
	$Message = " "
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = " " 
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
    $Message = "Checking if $SG exists on $Device"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut	
	Write-Host "Checking if " -ForegroundColor White -NoNewline
	Write-Host $SG -ForegroundColor Cyan -NoNewline
	Write-Host " exists on " -ForegroundColor White -NoNewline
	Write-Host $Device -ForegroundColor Cyan
	
	Foreach ($g in $group)   {
        If ($g -eq $SG)   {
            Write-Host $Line -ForegroundColor Green
			Write-Host $SG -ForegroundColor Green -NoNewline
			Write-Host " exists on " -ForegroundColor White -NoNewline
			Write-Host $Device -ForegroundColor Green
			Write-Host $Line -ForegroundColor Green
        	$Message = "`t$SG exists on $Device"
			Write-Log -Message $Message -Path $logfile -NoConsoleOut
			$tracker = 1
        }
    }
    If ($tracker -eq 0)   {
        Write-Host $Line -ForegroundColor Yellow
		Write-Host $SG -ForegroundColor Yellow -NoNewline
		Write-Host " does not exist on " -ForegroundColor White -NoNewline
		Write-Host $Device -ForegroundColor Yellow
        Write-Host $Line -ForegroundColor Yellow
		$Message = "`t$SG does not exist on $Device"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut
		Add_Admin $tracker
    }
}

Function Get-Localadmin {  
	$admins = Gwmi win32_groupuser –computer $Device
	$admins = $admins |? {$_.groupcomponent –like '*"Administrators"'}

	$admins |% {  
	$_.partcomponent –match “.+Domain\=(.+)\,Name\=(.+)$” > $nul  
	$matches[1].trim('"') + “\” + $matches[2].trim('"')  
	} | sort
}

Function Add_Admin   {
    $t = $args[0]
	$Message = " " 
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = " " 
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	$Message = "Adding $SecurityGroup to $Device"
	Write-Log -Message $Message -Path $logfile -NoConsoleOut
	Write-Host "Adding " -ForegroundColor White -NoNewline
	Write-Host $SecurityGroup -ForegroundColor Cyan -NoNewline
	Write-Host " to " -ForegroundColor White -NoNewline
	Write-Host $Device -ForegroundColor Cyan
	
    If ($t -eq 0)   {
		([ADSI]"WinNT://$Device/Administrators,group").Add("WinNT://$Domain/$SecurityGroup")
		Write-Host $Line -ForegroundColor Magenta
		Write-Host "`t`t$SecurityGroup has been added to $Device" -ForegroundColor Magenta
		Write-Host $Line -ForegroundColor Magenta
		$Message = "`t$SecurityGroup has been added to $Device"
		Write-Log -Message $Message -Path $logfile -NoConsoleOut
		Check_Admin
    }
}

#EndRegion Functions

#Region Main Processing


Test-Administrator

#$Device = Read-Host "Enter Device Name (Example: hostName)"	

$Computers = " "


$SecurityGroup = Read-Host "Enter Security Group to be added (Example: Admin)"
$Domain = Read-Host "Enter Domain (Example: Domain)"
$SG = $Domain + "\" + $SecurityGroup

Foreach($Computer in $Computers)   {
	$Device = $Computer
	Check_Admin
}

#EndRegion Main Processing