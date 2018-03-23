<# 
	.SYNOPSIS
    Update XML file required for use by NSO script.  And also updates input file required to setup DHCP scopes.
	
	.Description
	This script will update the XML used within the NSO script, and update CSV to update the DHCP scopes.
	Given a StoreCode (IE: ML) the script will update data in the XML as well as CSV.
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
	Update_NSO_XML.ps1 -StoreCode ML
	
	.EXAMPLE
	Update_NSO_XML_XL.ps1 -StoreCode ML -StoreLocation "Moses Lake, WA"
	
	.EXAMPLE
	Update_NSO_XML_XL.ps1 -StoreCode ML -StoreTimeZone 255
	
    .EXAMPLE
	Update_NSO_XML_XL.ps1 -StoreCode ML -StoreLocation "Moses Lake, WA" -StoreTimeZone 255
#>

[CmdletBinding()]
Param(
	[Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
		  [String]$StoreCode,
		  [String]$StoreLocation,
		  [String]$StoreTimeZone
	)

Import-Module ActiveDirectory

#Region Variable declartion
	$InfoMessage = @()
	$MachineList = @()
	$Count = 0
	$Date = Get-Date -Format 'MM_dd_yyyy'
	$Path = "fileName"
	$TimeZone = "fileName"
		
	$Unresponsive = @()
	$Line = "=========================================================================="
	$WarningColor = (Get-Host).PrivateData
	$WarningColor.WarningBackgroundColor = "yellow"
	$WarningColor.WarningForegroundColor = "blue"
	
#EndRegion Variable declaration

#Region Get control of powershell UI
    $a = (Get-Host).UI.RawUI
    $a.WindowTitle = "NSO XML Updater"
    $a.BackgroundColor = "Black"
    $a.ForegroundColor = "White"
    cls
#EndRegion

#Region Set startup variables
	$FilePath = "fileName"
	$StoreXML = "fileName"
	$SheetName = "Device"
	$NetworkName = "VLAN"
	$script:Excel = New-Object -Com Excel.Application
	$script:Excel.visible = $false
	$WorkBook = $script:Excel.Workbooks.Open($FilePath)
	$WorkSheet = $WorkBook.sheets.item($SheetName)
	$NetworkSheet = $WorkBook.sheets.item($NetworkName)
	$SearchString = $StoreCode.toUpper()
	$PoolName = $SearchString + "-VLANIPPOOL"
	$SubnetName = $SearchString + "-VLANSubnet"
	$RName = $SearchString + "MSTRDC01"
	$TName = $SearchString + "serverName"
	$UName = $SearchString + "serverName"
	$NName = $SearchString + "serverName"
	$vName = $SearchString + "serverName"
	$CName = $SearchString + "serverName"
	$ASName = $SearchString + "serverName"
	$AWName = $SearchString + "serverName"
	$A01Name = $SearchString + "serverName"
	$A02Name = $SearchString + "serverName"
	$P1Name = $SearchString + "serverName"
	$P2Name = $SearchString + "serverName"
	$Range = $WorkSheet.UsedRange
	$ScopeRange = $NetworkSheet.UsedRange
	$SearchRange = $Range.find($SearchString)
	$SearchScope = $ScopeRange.find($SearchString)
	$SearchPool =  $Range.find($PoolName)
	$SearchSubnet = $Range.find($SubnetName)
	$SearchR = $Range.find($RName)
	$SearchT = $Range.find($TName)
	$SearchvU = $Range.find($UName)
	$SearchvN = $Range.find($NName)
	$SearchvS = $Range.find($vName)
	$SearchX = $Range.find($CName)
	$SearchAS = $Range.find($ASName)
	$SearchAW = $Range.find($AWName)
	$SearchA01 = $Range.find($A01Name)
	$SearchA02 = $Range.find($A02Name)
	$SearchP1 = $Range.find($P1Name)
	$SearchP2 = $Range.find($P2Name)
	$A01IP = $SearchA01.Row
	$A02IP = $SearchA02.Row
	$P1IP = $SearchP1.Row
	$P2IP = $SearchP2.Row	
	$GatewayIP = $SearchRange.Row + 2	
	$RIP = $SearchR.Row
	$TIP = $SearchT.Row
	$vUIP = $SearchvU.Row
	$vNIP = $SearchvN.Row
	$vSIP = $SearchvS.Row
	$XIP = $SearchX.Row
	$ASIP = $SearchAS.Row
	$AWIP = $SearchAW.Row
	$Column = $SearchScope.Column
	$StoreIDIP = $SearchRange.Row + 1
	$StoreID = "0" + $worksheet.Rows.Item($StoreIDIP).Columns.Item(1).Text
	If($StoreID.length -ge 4)   {
		$StoreID = $StoreID.TrimStart("0")
	}
	$StorageName = $SearchString + $StoreID + "storageName"
	$SearchStorage = $Range.find($StorageName)

#EndRegion Set startup variables

#Region Functions

Function Update-Store   {
	Write-Host
	Write-Host "Updating NSO XML and DHCP Scope file" -ForegroundColor Cyan
	Write-Host $Line -ForegroundColor Green
	$StoreInfo = $XMLfile.ServerBuildList.Store

	$StoreInfo.StoreCode = $StoreCode		
	$StoreIDIP = $SearchRange.Row + 1
	$StoreID = "0" + $worksheet.Rows.Item($StoreIDIP).Columns.Item(1).Text	
	If($StoreID.length -gt 4)   {
		$StoreID = $StoreID.TrimStart("0")
		$StoreInfo.ID = $StoreID
	}
	Else   {
		$StoreInfo.ID = $StoreID
	}
	If($StoreLocation -eq "")   {
		$InfoMessage = "`t`t Enter Store Location (EX: Moses Lake, WA)"
		[String]$StoreLocation = Read-Host $InfoMessage
		$StoreInfo.StoreLocation = $StoreLocation
	}
	Else   {
		$StoreInfo.StoreLocation = $StoreLocation
	}

	If($StoreTimeZone -eq "" )   {
		Write-Host
		Write-Host "`t`tUse " -ForegroundColor White -NoNewline
		Write-Host "Time Zone " -ForegroundColor Cyan -NoNewline
		Write-Host "spreadsheet to obtain Store Time Zone" -ForegroundColor White
		Write-Host
		&$TimeZone
		$InfoMessage = "`t`t Enter Store Time Zone (EX: 020)"
		[String]$StoreTimeZone = Read-Host $InfoMessage
		$StoreInfo.StoreTimeZone = $StoreTimeZone
	}
	Else   {
		$StoreInfo.StoreTimeZone = $StoreTimeZone
	}
	
	Write-Host
	Write-Host "`tStore ID: " -ForegroundColor White -NoNewline
	Write-Host "`t`t$StoreID" -ForegroundColor Green
	Write-Host "`tStore Code: " -ForegroundColor White -NoNewline
	Write-Host "`t`t$SearchString" -ForegroundColor Green
	Write-Host "`tStore Location: " -ForegroundColor White -NoNewline
	Write-Host "`t$StoreLocation" -ForegroundColor Green
	Write-Host "`tTime Zone: " -ForegroundColor White -NoNewline
	Write-Host "`t`t$StoreTimeZone" -ForegroundColor Green	
	Write-Host $Line -ForegroundColor Green	
}

Function Update-Cluster   {
	Write-Host
	$ClusterIP = $XMLfile.ServerBuildList.Store.Cluster
	$ClusterName = $SearchString + "clusterName"
	$SearchCluster = $Range.find($ClusterName)
	$ClusterInfo = $SearchCluster.Row
	$Cluster = $worksheet.Rows.Item($ClusterInfo).Columns.Item(4).Text
	$ClusterIP.IP = $Cluster

	Write-Host "`tCluster IP: " -ForegroundColor White -NoNewline
	Write-Host "`t`t$Cluster" -ForegroundColor Green	
}

Function Update-VM   {
	Write-Host
	$Servers = @()
	$Servers += $XMLfile.ServerBuildList.Servers
	$Serverstype = $Servers.Server.Type
	$ServersIP = $Servers.Server.IP

	$RIP = $SearchR.Row
	$TIP = $SearchT.Row
	$vUIP = $SearchvU.Row
	$vNIP = $SearchvN.Row
	$vSIP = $SearchvS.Row
	$XIP = $SearchX.Row
	$ASIP = $SearchAS.Row
	$AWIP = $SearchAW.Row
		
	For ($i = 0; $i -lt $Serverstype.Count; $i++) {
		$R = $worksheet.Rows.Item($RIP).Columns.Item(4).Text
		$Servers.Server[$i].IP = $R
		$i++

		$T = $worksheet.Rows.Item($TIP).Columns.Item(4).Text
		$Servers.Server[$i].IP = $T
		$i++		
		$vU = $worksheet.Rows.Item($vUIP).Columns.Item(4).Text
		$Servers.Server[$i].IP = $vU
		$i++
		$vN = $worksheet.Rows.Item($vNIP).Columns.Item(4).Text
		$Servers.Server[$i].IP = $vN
		$i++
		$vS = $worksheet.Rows.Item($vSIP).Columns.Item(4).Text
		$Servers.Server[$i].IP = $vS
		$i++
		$X = $worksheet.Rows.Item($XIP).Columns.Item(4).Text
		$Servers.Server[$i].IP = $X
		$i++
		$AW = $worksheet.Rows.Item($AWIP).Columns.Item(4).Text	
		$Servers.Server[$i].IP = $AW
		$i++
		$AS = $worksheet.Rows.Item($ASIP).Columns.Item(4).Text		
		$Servers.Server[$i].IP = $AS
		$i++		
	}

	$InfoMessage = $SearchString + "R"
	Write-Host "`t$InfoMessage : " -ForegroundColor White -NoNewline
	Write-Host "`t`t$R" -ForegroundColor Green
	$InfoMessage = $SearchString + "-T-01"
	Write-Host "`t$InfoMessage : " -ForegroundColor White -NoNewline
	Write-Host "`t`t$T" -ForegroundColor Green
	$InfoMessage = $SearchString + "-vU-01"
	Write-Host "`t$InfoMessage : " -ForegroundColor White -NoNewline
	Write-Host "`t`t$vU" -ForegroundColor Green
	$InfoMessage = $SearchString + "-vN-01"
	Write-Host "`t$InfoMessage : " -ForegroundColor White -NoNewline
	Write-Host "`t`t$vN" -ForegroundColor Green
	$InfoMessage = $SearchString + "-vS-01"
	Write-Host "`t$InfoMessage : " -ForegroundColor White -NoNewline
	Write-Host "`t`t$vS" -ForegroundColor Green
	$InfoMessage = $SearchString + "-X-01"
	Write-Host "`t$InfoMessage : " -ForegroundColor White -NoNewline
	Write-Host "`t`t$X" -ForegroundColor Green
	$InfoMessage = $SearchString + "-AS-01"
	Write-Host "`t$InfoMessage : " -ForegroundColor White -NoNewline
	Write-Host "`t`t$AS" -ForegroundColor Green
	$InfoMessage = $SearchString + "-AW-01"
	Write-Host "`t$InfoMessage : " -ForegroundColor White -NoNewline
	Write-Host "`t`t$AW" -ForegroundColor Green
}

Function Update-ClusterHost   {
	Write-Host
	$ClusterHost = @()
	$ClusterHost += $XMLfile.ServerBuildList.BMHosts.BMHost
		
	For ($i = 0; $i -lt $ClusterHost.Count; $i++) {
		If ( $i -eq "0")   {
			$Code = "A01"
			$InfoMessage = $SearchString + "-R-" + $Code
			$HostIp = $worksheet.Rows.Item($A01IP).Columns.Item(4).Text
			$iLOIp = $worksheet.Rows.Item($A01IP).Columns.Item(8).Text
			Write-Host "`t$InfoMessage IP: " -ForegroundColor White -NoNewline
			Write-Host "`t$HostIp" -ForegroundColor Green
			Write-Host "`t$InfoMessage iLO: " -ForegroundColor White -NoNewline
			Write-Host "`t$iLOIp" -ForegroundColor Green			
		}
		ElseIf ( $i -eq "1")   {
			$Code = "A02"
			$InfoMessage = $SearchString + "-R-" + $Code
			$HostIp = $worksheet.Rows.Item($A02IP).Columns.Item(4).Text
			$iLOIp = $worksheet.Rows.Item($A02IP).Columns.Item(8).Text
			Write-Host "`t$InfoMessage IP: " -ForegroundColor White -NoNewline
			Write-Host "`t$HostIp" -ForegroundColor Green
			Write-Host "`t$InfoMessage iLO: " -ForegroundColor White -NoNewline
			Write-Host "`t$iLOIp" -ForegroundColor Green
		}
		ElseIf ($i -eq "2")   {
			$Code = "P1"
			$InfoMessage = $SearchString + "-" + $Code
			$HostIp = $worksheet.Rows.Item($P1IP).Columns.Item(4).Text
			$iLOIp = $worksheet.Rows.Item($P1IP).Columns.Item(8).Text
			Write-Host "`t$InfoMessage IP: " -ForegroundColor White -NoNewline
			Write-Host "`t$HostIp" -ForegroundColor Green
			Write-Host "`t$InfoMessage iLO: " -ForegroundColor White -NoNewline
			Write-Host "`t$iLOIp" -ForegroundColor Green

		}
		ElseIf ($i -eq "3")   {
			$Code = "P2"
			$InfoMessage = $SearchString + "-" + $Code
			$HostIp = $worksheet.Rows.Item($P2IP).Columns.Item(4).Text
			$iLOIp = $worksheet.Rows.Item($P2IP).Columns.Item(8).Text
			Write-Host "`t$InfoMessage IP: " -ForegroundColor White -NoNewline
			Write-Host "`t$HostIp" -ForegroundColor Green
			Write-Host "`t$InfoMessage iLO: " -ForegroundColor White -NoNewline
			Write-Host "`t$iLOIp" -ForegroundColor Green
		}
		$ClusterHost[$i].BMCIP = $iLOIp
		$ClusterHost[$i].RetailIP = $HostIp
	}	
}

Function Update-Storage   {
	Write-Host
	$StorageIP = $XMLfile.ServerBuildList.Store.Storage
	$InServIP = $SearchStorage.Row
	$InServ = $worksheet.Rows.Item($InServIP).Columns.Item(8).Text 
	$StorageIP.IP = [String]$InServ
	Write-Host "`tInServ IP: " -ForegroundColor White -NoNewline
	Write-Host "`t$InServ" -ForegroundColor Green	
}

Function Update-Network   {
	Write-Host
	$NetworkVlan = $XMLfile.ServerBuildList.Store.LogicalNetworks.VLAN	
	$NetworkDNS = $XMLfile.ServerBuildList.Store.LogicalNetworks.DNSS
	$GatewayIP = $SearchRange.Row + 2
	$DNS1IP = $SearchSubnet.Row + 3
	$SubnetIP = $SearchSubnet.Row
	$IPPoolIP = $SearchPool.Row
	$Gateway = $worksheet.Rows.Item($GatewayIP).Columns.Item(4).Text
	$DNS1 = $worksheet.Rows.Item($DNS1IP).Columns.Item(4).Text
	$Subnet = $worksheet.Rows.Item($SubnetIP).Columns.Item(4).Text
	$Pool = $worksheet.Rows.Item($IPPoolIP).Columns.Item(4).Text 
	$IPPool = (($Pool.split("-")[0]) + "-" + ($Pool.split("-")[1]).split(".")[3])
	$NetworkVlan.Gateway = $Gateway
	$NetworkDNS.DNS1 = $DNS1
	$NetworkVlan.IPSubnet = $Subnet
	$NetworkVlan.IPPool = $IPPool

	Write-Host "`tGateway IP: " -ForegroundColor White -NoNewline
	Write-Host "`t$Gateway" -ForegroundColor Green	
	Write-Host "`tIP Pool: " -ForegroundColor White -NoNewline
	Write-Host "`t$IPPool" -ForegroundColor Green	
	Write-Host "`tSubnet: " -ForegroundColor White -NoNewline
	Write-Host "`t$Subnet" -ForegroundColor Green
	Write-Host "`tDNS1: " -ForegroundColor White -NoNewline
	Write-Host "`t`t$DNS1" -ForegroundColor Green
}

Function Update-Scopes   {
	Write-Host	
	$StoreID = "0" + $Networksheet.Rows.Item(3).Columns.Item($Column).Text
	If($StoreID.length -ge 4)   {
		$StoreID = $StoreID.TrimStart("0")
	}
	$ScopePath = "fileName"	
	$FinalPath = "fileName" + $StoreID + "-input.csv"
	[io.file]::readalltext($ScopePath).replace("xxxx",$StoreID) | Out-File $FinalPath -Encoding ascii –Force

	$Data = Import-CSV $FinalPath
	$AName = "A  With ERP"
	$RIP = $SearchRange.Row + 6
	$R = $worksheet.Rows.Item($RIP).Columns.Item(4).Text
	$2Row = ($ScopeRange.find("B")).Row
	$2Scope = $Networksheet.Rows.Item($2Row).Columns.Item($Column).Text
	$2Mask = $Networksheet.Rows.Item($2Row).Columns.Item(4).Text
	$8Row = ($ScopeRange.find("B")).Row
	$8Scope = $Networksheet.Rows.Item($8Row).Columns.Item($Column).Text
	$8Mask = $Networksheet.Rows.Item($8Row).Columns.Item(4).Text
	$22Row = ($ScopeRange.find("net")).Row
	$22Scope = $Networksheet.Rows.Item($22Row).Columns.Item($Column).Text
	$22Mask = $Networksheet.Rows.Item($22Row).Columns.Item(4).Text
	$24Row = ($ScopeRange.find($AName)).Row
	$24Scope = $Networksheet.Rows.Item($24Row).Columns.Item($Column).Text
	$24Mask = $Networksheet.Rows.Item($24Row).Columns.Item(4).Text
	$25Row = ($ScopeRange.find("V")).Row
	$25Scope = $Networksheet.Rows.Item($25Row).Columns.Item($Column).Text
	$25Mask = $Networksheet.Rows.Item($25Row).Columns.Item(4).Text
	$85Row = ($ScopeRange.find("S")).Row
	$85Scope = $Networksheet.Rows.Item($85Row).Columns.Item($Column).Text
	$85Mask = $Networksheet.Rows.Item($85Row).Columns.Item(4).Text

	$count = $Data.Scope.length
	$i = 0
	while ($i -ne $count)   {
		If ( $i -eq 0)   {
			[io.file]::readalltext($FinalPath).replace($Data.Scope[$i],$2Scope) | Out-File $FinalPath -Encoding ascii –Force
			[io.file]::readalltext($FinalPath).replace($Data.CIDR[$i],$2Mask) | Out-File $FinalPath -Encoding ascii –Force
			[io.file]::readalltext($FinalPath).replace($Data.R[$i],$R) | Out-File $FinalPath -Encoding ascii –Force
			$i++
		}
		ElseIf ($i -eq 1) {
			[io.file]::readalltext($FinalPath).replace($Data.Scope[$i],$8Scope) | Out-File $FinalPath -Encoding ascii –Force
			[io.file]::readalltext($FinalPath).replace($Data.CIDR[$i],$8Mask) | Out-File $FinalPath -Encoding ascii –Force
			[io.file]::readalltext($FinalPath).replace($Data.R[$i],$R) | Out-File $FinalPath -Encoding ascii –Force
			$i++
		}
		ElseIf ($i -eq 2) {
			[io.file]::readalltext($FinalPath).replace($Data.Scope[$i],$22Scope) | Out-File $FinalPath -Encoding ascii –Force
			[io.file]::readalltext($FinalPath).replace($Data.CIDR[$i],$22Mask) | Out-File $FinalPath -Encoding ascii –Force
			[io.file]::readalltext($FinalPath).replace($Data.R[$i],$R) | Out-File $FinalPath -Encoding ascii –Force
			$i++
		}
		ElseIf ($i -eq 3) {
			[io.file]::readalltext($FinalPath).replace($Data.Scope[$i],$24Scope) | Out-File $FinalPath -Encoding ascii –Force
			[io.file]::readalltext($FinalPath).replace($Data.CIDR[$i],$24Mask) | Out-File $FinalPath -Encoding ascii –Force
			[io.file]::readalltext($FinalPath).replace($Data.R[$i],$R) | Out-File $FinalPath -Encoding ascii –Force
			$i++
		}
		ElseIf ($i -eq 4) {
			[io.file]::readalltext($FinalPath).replace($Data.Scope[$i],$25Scope) | Out-File $FinalPath -Encoding ascii –Force
			[io.file]::readalltext($FinalPath).replace($Data.CIDR[$i],$25Mask) | Out-File $FinalPath -Encoding ascii –Force
			[io.file]::readalltext($FinalPath).replace($Data.R[$i],$R) | Out-File $FinalPath -Encoding ascii –Force
			$i++
		}
		ElseIf ($i -eq 5) {
			[io.file]::readalltext($FinalPath).replace($Data.Scope[$i],$85Scope) | Out-File $FinalPath -Encoding ascii –Force
			[io.file]::readalltext($FinalPath).replace($Data.CIDR[$i],$85Mask) | Out-File $FinalPath -Encoding ascii –Force
			[io.file]::readalltext($FinalPath).replace($Data.R[$i],$R) | Out-File $FinalPath -Encoding ascii –Force
			$i++
		}
	}
	
	Write-Host "`tB Scope: " -ForegroundColor White -NoNewline
	Write-Host "`t`t$2Scope" -ForegroundColor Green
	Write-Host "`tB CIDR: " -ForegroundColor White -NoNewline
	Write-Host "`t`t$2Mask" -ForegroundColor Green
	Write-Host "`tB Scope: " -ForegroundColor White -NoNewline
	Write-Host "`t`t$8Scope" -ForegroundColor Green
	Write-Host "`tB CIDR: " -ForegroundColor White -NoNewline
	Write-Host "`t`t$8Mask" -ForegroundColor Green
	Write-Host "`tnet Scope: " -ForegroundColor White -NoNewline
	Write-Host "`t$22Scope" -ForegroundColor Green
	Write-Host "`tnet CIDR: " -ForegroundColor White -NoNewline
	Write-Host "`t$22Mask" -ForegroundColor Green
	Write-Host "`tA Scope: " -ForegroundColor White -NoNewline
	Write-Host "`t`t$24Scope" -ForegroundColor Green
	Write-Host "`tA CIDR: " -ForegroundColor White -NoNewline
	Write-Host "`t`t$24Mask" -ForegroundColor Green
	Write-Host "`tV Scope: " -ForegroundColor White -NoNewline
	Write-Host "`t$25Scope" -ForegroundColor Green
	Write-Host "`tV CIDR: " -ForegroundColor White -NoNewline
	Write-Host "`t`t$25Mask" -ForegroundColor Green
	Write-Host "`tS Scope: " -ForegroundColor White -NoNewline
	Write-Host "`t`t$85Scope" -ForegroundColor Green
	Write-Host "`tS CIDR: " -ForegroundColor White -NoNewline
	Write-Host "`t`t$85Mask" -ForegroundColor Green	
}

Function Update-QCFile   {
	Write-Host
	$QCFile = "fileName"
	Write-Host "Loading QC File: " -ForegroundColor Yellow -NoNewline
	Write-Host "`t"$QCFile -ForegroundColor Green
	
	$R = $worksheet.Rows.Item($RIP).Columns.Item(4).Text
	$T = $worksheet.Rows.Item($TIP).Columns.Item(4).Text
	$vU = $worksheet.Rows.Item($vUIP).Columns.Item(4).Text
	$vN = $worksheet.Rows.Item($vNIP).Columns.Item(4).Text
	$vS = $worksheet.Rows.Item($vSIP).Columns.Item(4).Text
	$X = $worksheet.Rows.Item($XIP).Columns.Item(4).Text
	$AW = $worksheet.Rows.Item($AWIP).Columns.Item(4).Text
	$AS = $worksheet.Rows.Item($ASIP).Columns.Item(4).Text
	$Gateway = $worksheet.Rows.Item($GatewayIP).Columns.Item(4).Text
	
	$A01HostIp = $worksheet.Rows.Item($A01IP).Columns.Item(4).Text
	$A02HostIp = $worksheet.Rows.Item($A02IP).Columns.Item(4).Text
	$P1HostIp = $worksheet.Rows.Item($P1IP).Columns.Item(4).Text
	$P2HostIp = $worksheet.Rows.Item($P2IP).Columns.Item(4).Text

	$QCLocation = $Path + $StoreCode.toUpper() + "\" + $StoreCode.toUpper() + "-QC.csv"
	Copy-Item -Path $QCFile -Destination $QCLocation
	(Get-Content $QCFile) | 
	Foreach-Object {$_ -replace "XXX", $StoreCode.toUpper() -replace "0.0.0.0", $R -replace "1.1.1.1", $Gateway -replace "2.2.2.2", $T -replace "3.3.3.3", $vU -replace "4.4.4.4", $vN -replace "5.5.5.5", $vS -replace "6.6.6.6", $X -replace "7.7.7.7", $AW -replace "8.8.8.8", $AS -replace "9.9.9.9", $A01HostIp -replace "w.w.w.w", $A02HostIp -replace "y.y.y.y", $P1HostIp -replace "z.z.z.z", $P2HostIp } | 
	Set-Content $QCLocation
	Write-Host "Updated QC File is: " -ForegroundColor Yellow -NoNewline
	Write-host "`t"$QCLocation -ForegroundColor Green
}

#EndRegion Functions

#Region Main Processing

#set excel row below headers
$intCurrentRow = 2

#Load XML File
Write-Host $Line -ForegroundColor Green
Write-Host "Loading XML File: " -ForegroundColor Yellow -NoNewline
Write-host "`t"$StoreXML -ForegroundColor Green
If(Test-Path $StoreXML -PathType Leaf)   {
	Write-Host "`t`t.......... XML File Loaded .........."
	[XML]$XMLfile = Get-Content "fileName"
	$XMLPath = $Path + $SearchString + "\"
	If (!(Test-Path $XMLPath))  {
		New-Item -ItemType Directory -Force -Path $XMLPath | Out-Null
        Write-Host "`t`t$XMLPath " -ForegroundColor Green -NoNewline
		Write-Host "has been created" -ForegroundColor White
		$File = $Path + $SearchString + "\" + $SearchString + "fileName"
    }
    Else   {
        Write-Host "`t`t$XMLPath " -ForegroundColor Green -NoNewline
		Write-Host "already exists" -ForegroundColor White
		$Message = "`t$XMLPath already exists"
		$File = $Path + $SearchString + "\" + $SearchString + "fileName"
	}
}
Else   {
	Write-Warning "Unable to open input file.  Exiting Script!"
	Write-Footer
	Exit
}
Write-Host "Updated XML File is: " -ForegroundColor Yellow -NoNewline
Write-host "`t"$File -ForegroundColor Green

$XMLfile.Save($File)
[XML]$XMLfile = Get-Content $File
Write-Host $Line -ForegroundColor Green

Update-Store
Update-Cluster
Update-ClusterHost
Update-VM
Update-Storage
Update-Network
Update-Scopes
Update-QCFile

$XMLfile.Save($File)
(Get-Content $File ) | Foreach-Object {$_ -replace "DNS1", "DNS"} | Set-Content $File
(Get-Content $File ) | Foreach-Object {$_ -replace "ï»¿", ""} | Set-Content $File
Copy-Item -Path $File -Destination "fileName"

$script:Excel.Quit()

#EndRegion Main Processing