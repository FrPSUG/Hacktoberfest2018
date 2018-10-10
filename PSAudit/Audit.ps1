<#
	.SYNOPSIS
		System Audit Script

	.DESCRIPTION
	    Perfom several system audit & reporting
	       
	.PARAMETER auditlist
		The file is optional and needs to be a plain text list of computers to be audited
		one on each line, if no list is specified the local machine will be audited.		

	.PARAMETER ComputerName
	   List of Computers

	.PARAMETER ExportCSV
	    If set, export all info to CSV files

	.EXAMPLE	  
		.\audit.ps1 -auditlist C:\temp\list.txt -ExportCSV
	
		Will make an audit on all computer in the txt file and then export all info to HTML as well as CSV file
		
	.EXAMPLE	
		.\Audit.ps1 -ExportCSV
	
		Will make an audit on the local computer and then export all info to HTML as well as CSV file

	.OUTPUTS
		HTML and CSV file if set in the same directory as the ps1 file.

	.NOTES
		Author: Audit script V3 by Alan Renouf - Virtu-Al
		Blog  : http://virtu-al.net/	

		Modif : Update 1 by Mathieu Allegret
			TerminalService
			Local Users & Groups 
			PageFile
		Modif : Update 2 by Mathieu Allegret
			Windows Licence
			EventLog boot errors & oldest event
			Schedule Tasks

    	Update 3 by Mathieu Allegret
			Export-CSV des infos

		Update 4 by MAL on 05/26/2017
			New parameters
			Improve software info gathering
			Improve License info gathering
			Exporting function to separate files

		Update 5 by MAL on 25/05/2018
			Change Schedule Tasks collecting and reporting info
			Bugs fixing

	.LINK
		http://virtu-al.net/

#>

[CmdletBinding()]
param(
	[string]$auditlist,
	[string[]]$ComputerName,
	[switch]$ExportCSV
)

#Date for execution time calculation (function CalculDuree)
$script:StartTime = Get-Date

. .\Audit_functions.ps1
. .\Audit_HTML_functions.ps1


if (!(Test-Path .\PSRemoteRegistry)) {
	Write-Host "Le module PSRemoteRegistry est manquant, abort" -f red
}
else {
	Import-Module .\PSRemoteRegistry\PSRemoteRegistry.psm1 -Verbose:$false

	CalculDuree -message "Import Module PSRemoteRegistry"

	if ($auditlist -eq "" -or $auditlist -eq $null){
		if ($ComputerName -eq "" -or $ComputerName -eq $null) {
			Write-Host "No list specified, using $env:ComputerName"
			$targets = $env:ComputerName
		} else {
			$targets = $ComputerName
		}
	}
	else
	{
		if ((Test-Path $auditlist) -eq $false)
		{
			Write-Host "Invalid audit path specified: $auditlist"
			exit
		}
		else
		{
			Write-Host "Using Audit list: $auditlist"
			$Targets = Get-Content $auditlist
		}
	}
	
	CalculDuree -message "Generating Server list"

	$RegionalOptions = @()
	$SystemInfo = @()
	$WindowsLicenseInfo = @()
	$HotfixInfo = @()
	$LogicalDisksInfo = @()
	$NetworkInfo = @()
	$SoftwareInfo = @()
	$SharesInfo = @()
	$PrintersInfo = @()
	$LocalUsersInfo = @()
	$LocalGroupsInfo = @()
	$ServicesInfo = @()
	$RegionalSettingsInfo = @()
	$EventLogInfo = @()
	$EventLogBootInfo = @()
	$EventLogOldestEvt = @()
	$ScheduleTasksInfo = @()

	Foreach ($Target in $Targets){
		Clear-Variable pingResult -ErrorAction SilentlyContinue

		$pingResult = RPC-Ping -ComputerName $Target 
		Write-Debug "RPC-Ping return $($pingResult)" 

		if ($pingResult -eq $True) {
			
			CalculDuree -message "RPC Ping"

			#region ServerInfo
			Write-Output "$Target is alive"
			Write-Output "Collating Detail for $Target"
			$ComputerSystem = Get-WmiObject -ComputerName $Target Win32_ComputerSystem
			
			CalculDuree -message "Getting ComputerSystem"

			switch ($ComputerSystem.DomainRole){
				0 { $ComputerRole = "Standalone Workstation" }
				1 { $ComputerRole = "Member Workstation" }
				2 { $ComputerRole = "Standalone Server" }
				3 { $ComputerRole = "Member Server" }
				4 { $ComputerRole = "Domain Controller" }
				5 { $ComputerRole = "Domain Controller" }
				default { $ComputerRole = "Information not available" }
			}
			$LicenseWin = Get-WmiObject -Class SoftwareLicensingProduct -ComputerName $Target | Where-Object {$_.ProductKeyID -and $_.name -like "Windows*"} 
			
			CalculDuree -message "Getting License info"

			switch ($LicenseWin.LicenseStatus) {
				"0" { $LicenseStatus = "Unlicensed" ; break}
				"1" { $LicenseStatus = "Licensed" ; break}
				"2" { $LicenseStatus = "Out-Of-Box Grace Period" ; break}
				"3" { $LicenseStatus = "Out-Of-Tolerance Grace Period" ; break}
				"4" { $LicenseStatus = "Non-Genuine Grace Period" ; break}
				default { $LicenseStatus = "Unkown" ; break}
			}
			$OperatingSystem = Get-WmiObject -ComputerName $Target Win32_OperatingSystem
			
			CalculDuree -message "Getting OperatingSystem"

			$TimeZone = Get-WmiObject -ComputerName $Target Win32_Timezone
						
			CalculDuree -message "Getting Timezone"

			$ObjKeyboards = Get-WmiObject -ComputerName $Target Win32_Keyboard
			
			CalculDuree -message "Getting Keyboard"

			$ScheduleTasks = Get-MALScheduledTasks -ComputerName $Target -Subfolders | Where-Object {$_.taskname -notlike "\Microsoft*" -and $_.taskname -notlike "\Optimize Start*"}
			
			CalculDuree -message "Getting ScheduledTasks"
			
			$PageFileInfo = (Get-WmiObject -ComputerName $Target Win32_PageFileusage)[0]
			
			CalculDuree -message "Getting PageFileusage"

			if ($OperatingSystem.OSArchitecture -eq '64-bit') {            
				$OSArchi = "64-Bit"            
			   } else  {            
				$OSArchi = "32-Bit"            
			   }   

			$OSVer = $OperatingSystem.Version
			if ($Osver -ge "6.0") {
				$TSInfo = Get-WmiObject -Namespace "root\CIMV2\TerminalServices" -ComputerName $Target -Class "Win32_TerminalServiceSetting" -ea 0
			}
			else {
				$TSInfo = Get-WmiObject -ComputerName $Target -Class "Win32_TerminalServiceSetting" -ea 0
			}

			if ($TSInfo.TerminalServerMode -eq 0) {
				$TSMode = "Administration"
			}
			elseif (!$TSInfo) {
				$TSMode = "Administration"
			}
			else {
				$TSMode = "Application"
			}
		
			if ($TSMode -eq "Application") {
				$TSKey = "SYSTEM\CurrentControlSet\services\TermService\Parameters\LicenseServers"
				if ($Osver -ge "6.0") {
					$TSLicenseServers = Get-RegValue -ComputerName $Target -Key $TSKey -Value SpecifiedLicenseServers | ForEach-Object {$_.Data}
				}
				else {
					$TSLicenseServersKeys =	Get-RegKey -ComputerName $Target -Key $TSKey | ForEach-Object {$_.Key}
					$TSLicenseServers = @()
					foreach ($key in $TSLicenseServersKeys) {
						$varTempo = ($key.split("\")).count
						$TSLicenseServers += ($key.split("\"))[$varTempo-1]
					}
				}
				$TSLicenseServers = $TSLicenseServers -join ","
			}

			CalculDuree -message "Getting TSE mode & license"
			
			switch ($ComputerRole){
				"Member Workstation" { $CompType = "Computer Domain"; break }
				"Domain Controller" { $CompType = "Computer Domain"; break }
				"Member Server" { $CompType = "Computer Domain"; break }
				default { $CompType = "Computer Workgroup"; break }
			}

			$LBTime = $OperatingSystem.ConvertToDateTime($OperatingSystem.Lastbootuptime)
			$keyboardmap = @{
			"00000402" = "BG" 
			"00000404" = "CH" 
			"00000405" = "CZ" 
			"00000406" = "DK" 
			"00000407" = "GR" 
			"00000408" = "GK" 
			"00000409" = "US" 
			"0000040A" = "SP" 
			"0000040B" = "SU" 
			"0000040C" = "FR" 
			"0000040E" = "HU" 
			"0000040F" = "IS" 
			"00000410" = "IT" 
			"00000411" = "JP" 
			"00000412" = "KO" 
			"00000413" = "NL" 
			"00000414" = "NO" 
			"00000415" = "PL" 
			"00000416" = "BR" 
			"00000418" = "RO" 
			"00000419" = "RU" 
			"0000041A" = "YU" 
			"0000041B" = "SL" 
			"0000041C" = "US" 
			"0000041D" = "SV" 
			"0000041F" = "TR" 
			"00000422" = "US" 
			"00000423" = "US" 
			"00000424" = "YU" 
			"00000425" = "ET" 
			"00000426" = "US" 
			"00000427" = "US" 
			"00000804" = "CH" 
			"00000809" = "UK" 
			"0000080A" = "LA" 
			"0000080C" = "BE" 
			"00000813" = "BE" 
			"00000816" = "PO" 
			"00000C0C" = "CF" 
			"00000C1A" = "US" 
			"00001009" = "US" 
			"0000100C" = "SF" 
			"00001809" = "US" 
			"00010402" = "US" 
			"00010405" = "CZ" 
			"00010407" = "GR" 
			"00010408" = "GK" 
			"00010409" = "DV" 
			"0001040A" = "SP" 
			"0001040E" = "HU" 
			"00010410" = "IT" 
			"00010415" = "PL" 
			"00010419" = "RU" 
			"0001041B" = "SL" 
			"0001041F" = "TR" 
			"00010426" = "US" 
			"00010C0C" = "CF" 
			"00010C1A" = "US" 
			"00020408" = "GK" 
			"00020409" = "US" 
			"00030409" = "USL" 
			"00040409" = "USR" 
			"00050408" = "GK"
			}
			
			if (@($ObjKeyboards).count -gt 1)
			{
				$keyboard = $keyboardmap.$($ObjKeyboards[0].Layout)
			}
			else
			{
				$keyboard = $keyboardmap.$($ObjKeyboards.Layout)
			}
			if (!$keyboard){ 
				$keyboard = "Unknown"
			}

			CalculDuree -message "Getting Keyboard map"

			switch ($OperatingSystem.oslanguage){
				"1031" { $OSLang = "1031 - German"; break }
				"1033" { $OSLang = "1033 - English - United States"; break }
				"1034" { $OSLang = "1034 - Spanish"; break }
				"1036" { $OSLang = "1036 - French"; break }
				"1040" { $OSLang = "1040 - Italian"; break }
				"1041" { $OSLang = "1041 - Japanese"; break }
				"1049" { $OSLang = "1049 - Russian"; break }
				default { $OSLang = "Other language"; break }
			}

			CalculDuree -message "Getting oslang code"

			
			switch ($OperatingSystem.Locale){
				"0407" { $OSLocal = "0407 - German"; break }
				"0409" { $OSLocal = "0409 - English - United States"; break }
				"040A" { $OSLocal = "040A - Spanish"; break }
				"040C" { $OSLocal = "040C - French"; break }
				"0410" { $OSLocal = "0410 - Italian"; break }
				"0411" { $OSLocal = "0411 - Japanese"; break }
				"0419" { $OSLocal = "0419 - Russian"; break }
				default { $OSLocal = "Other language"; break }
			}
			
			CalculDuree -message "Getting locale"

			switch ($OperatingSystem.Countrycode){
				"49" { $OSCountryCode = "49 - Germany"; break }
				"1" { $OSCountryCode = "1 - United States"; break }
				"34" { $OSCountryCode = "34 - Spain"; break }
				"33" { $OSCountryCode = "33 - France"; break }
				"39" { $OSCountryCode = "39 - Italy"; break }
				"81" { $OSCountryCode = "81 - Japan"; break }
				"7" { $OSCountryCode = "7 - Russia"; break }
				default { $OSCountryCode = "Unknown"; break }
			}
			
			CalculDuree -message "Getting country code"

			switch ($OperatingSystem.DataExecutionPrevention_SupportPolicy) {
				0 {$policyCode = "AlwaysOff"}
				1 {$policyCode = "AlwaysOn"}
				2 {$policyCode = "OptIn (defaultConf)"}
				3 {$policyCode = "OptOut"}
			}

			CalculDuree -message "Getting DataExecution"

			#Determine PendingFileRenameOperations exists or not 
			$AutoUpdate = $False
			if (Get-RegKey -ComputerName $Target -Key "Software\Microsoft\Windows\CurrentVersion\Component Based Servicing" -Name RebootPending -EA 0)
			{
				$Pending = $true
				if(Get-RegKey -ComputerName $Target -Key "SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" -Name RebootRequired -EA 0)
				{
					$AutoUpdate = $true
				}
				else
				{
					$AutoUpdate = $false
				}
			}
			else
			{
				$Pending = $false
			}
			if ($AutoUpdate -eq $true)
			{
				$PendingReboot = "Yes (WindowsUpdate)"
			}
			elseif ($Pending -eq $true) {
				$PendingReboot = "Yes"
			}
			else
			{
				$PendingReboot = "No"
			}

			CalculDuree -message "Getting Pending Reboot"

			#Determine PowerShell Version
			$keyPS = "SOFTWARE\Microsoft\PowerShell"
			if (Get-RegKey -Key $keyPS -ComputerName $Target -EA 0)
			{
				$keyPS3 = $keyPS + "\3"
				if (Get-RegKey -Key $keyPS3 -ComputerName $Target -EA 0)
				{
					$PSVersion = (Get-RegValue -Key $keyPS3"\PowerShellEngine" -Value PowerShellVersion -ComputerName $Target -EA 0).data
				}
				else
				{
					$keyPS1 = $keyPS + "\1"
					$PSVersion = (Get-RegValue -Key $keyPS1"\PowerShellEngine" -Value PowerShellVersion -ComputerName $Target -EA 0).data
				}
			}
			else
			{
				$PSVersion = "Not Installed"
			}
			#endregion ServerInfo

			CalculDuree -message "Getting PSH version"
			
			#region RegionalOptions
			Write-Output "..Regional Options"
			$MyReport = Get-CustomHTML "$Target Audit"
			$MyReport += Get-CustomHeader0  "$Target Details"
			$MyReport += Get-CustomHeader "2" "General"
				$MyReport += Get-HTMLDetail "Computer Name" ($ComputerSystem.Name)
				$MyReport += Get-HTMLDetail "Computer Role" ($ComputerRole)
				$MyReport += Get-HTMLDetail $CompType ($ComputerSystem.Domain)
				$MyReport += Get-HTMLDetail "Operating System" ($OperatingSystem.Caption)
				$MyReport += Get-HTMLDetail "Service Pack" ($OperatingSystem.CSDVersion)
				$MyReport += Get-HTMLDetail "Manufacturer" ($ComputerSystem.Manufacturer)
				$MyReport += Get-HTMLDetail "Model" ($ComputerSystem.Model)
				$MyReport += Get-HTMLDetail "Number of Processors (socket)" ($ComputerSystem.NumberOfProcessors)
				$MyReport += Get-HTMLDetail "Number of Processors (logical)" ($ComputerSystem.NumberOfLogicalProcessors)
				$MyReport += Get-HTMLDetail "Memory" ($ComputerSystem.TotalPhysicalMemory | ConvertTo-KMG)
				$MyReport += Get-HTMLDetail "Registered User" ($ComputerSystem.PrimaryOwnerName)
				$MyReport += Get-HTMLDetail "Registered Organisation" ($OperatingSystem.Organization)
				$MyReport += Get-CustomHeaderClose
				
				$Details = "" | Select-Object "Computer Name" , "Computer Role" , $CompType, "Operating System", "Service Pack" , Manufacturer, Model, NumberOfSocketProcessors, NumberOfLogicalProcessors, Memory, "Registered User", "Registered Organisation"
				$Details."Computer Name" = $ComputerSystem.Name
				$Details."Computer Role" = $ComputerRole
				$Details.$CompType = $ComputerSystem.Domain
				$Details."Operating System" = $OperatingSystem.Caption
				$Details."Service Pack" = $OperatingSystem.CSDVersion
				$Details."Manufacturer" = $ComputerSystem.Manufacturer
				$Details."Model" = $ComputerSystem.Model
				$Details."NumberOfSocketProcessors" = $ComputerSystem.NumberOfProcessors
				$Details."NumberOfLogicalProcessors" = $ComputerSystem.NumberOfLogicalProcessors
				$Details."Memory" = $ComputerSystem.TotalPhysicalMemory | ConvertTo-KMG
				$Details."Registered User" = $ComputerSystem.PrimaryOwnerName
				$Details."Registered Organisation" = $OperatingSystem.Organization
				$RegionalOptions += $Details
			#endregion RegionalOptions
			
			CalculDuree -message "Regional options"
				
			#region SystemInfo
				Write-Output "..System Info"
				$MyReport += Get-CustomHeader "2" "System Info"
				$MyReport += Get-HTMLDetail "OS Version" ($OperatingSystem.version)
				$MyReport += Get-HTMLDetail "OS Architecture" ($OSArchi)
				$MyReport += Get-HTMLDetail "System Root" ($OperatingSystem.SystemDrive)
				$MyReport += Get-HTMLDetail "PowerShell Version" ($PSVersion)
				$MyReport += Get-HTMLDetail "PageFile Location" ($PageFileInfo.Name)
				$MyReport += Get-HTMLDetail "PageFile Size" ($PageFileInfo.AllocatedBaseSize * 1MB | ConvertTo-KMG)
				$MyReport += Get-HTMLDetail "Data Execution Prevention" ($policyCode)
				$MyReport += Get-HTMLDetail "TerminalService Mode" $TSMode
				if ($TSInfo.TerminalServerMode -ne 0) {
					if ($TSInfo) {
						$MyReport += Get-HTMLDetail "TerminalService License Server" $TSLicenseServers
					}
				}
				$MyReport += Get-HTMLDetail "Pending reboot" ($PendingReboot)
				$MyReport += Get-HTMLDetail "Last System Boot" ($LBTime)
				$MyReport += Get-CustomHeaderClose
				
				$Details = "" | Select-Object "Computer Name","OS Version","OS Architecture","System Root","PowerShell Version","PageFile Location","PageFile Size","Data Execution Prevention","TerminalService Mode","TerminalService License Server","Pending reboot","Last System Boot"
				$Details."Computer Name" = $ComputerSystem.Name
				$Details."OS Version" = $OperatingSystem.version
				$Details."OS Architecture" = $OperatingSystem.OSarchitecture
				$Details."System Root" = $OperatingSystem.SystemDrive
				$Details."PowerShell Version" = $PSVersion
				$Details."PageFile Location" = $PageFileInfo.Name
				$Details."PageFile Size" = $PageFileInfo.AllocatedBaseSize * 1MB | ConvertTo-KMG
				$Details."Data Execution Prevention" = $policyCode
				$Details."TerminalService Mode" = $TSMode
				if ($TSInfo.TerminalServerMode -ne 0) {
					$Details."TerminalService License Server" = $TSLicenseServers
				} else {
					$Details."TerminalService License Server" = 'N/A'
				}
				$Details."Pending reboot" = ($PendingReboot)
				$Details."Last System Boot" = ($LBTime)
				$SystemInfo += $Details
			#endregion SystemInfo

			CalculDuree -message "System info"
			
			#region WindowsLicenseInfo
			Write-Output "..Windows License Information"
			
				$MyReport += Get-CustomHeader "2" "Windows License Info"
				$Details = "" | Select-Object ComputerName,Description,KMSServer,KMSPort,ProductKeyChannel,PartialProductKey,GracePeriodRemaining,LicenseFamily,LicenseStatus
				$Details.ComputerName = $Target
				$Details.Description = $LicenseWin.Description
				$MyReport += Get-HTMLDetail "License Description" ($LicenseWin.Description)
				if ($LicenseWin.Description -like "*KMS*") {
					$KMSServer = ($LicenseWin.DiscoveredKeyManagementServiceMachineName)
					$KMSPort = ($LicenseWin.DiscoveredKeyManagementServiceMachinePort)
					$MyReport += Get-HTMLDetail "KMS Server" $KMSServer
					$MyReport += Get-HTMLDetail "KMS Port" $KMSPort
					$Details.KMSServer = $KMSServer
					$Details.KMSPort = $KMSPort
					$GracePeriod = "$($([math]::round($LicenseWin.GracePeriodRemaining / 1440)) | Out-String -Stream) jours"
				} else {
					$MyReport += Get-HTMLDetail "KMS Server" "N/A"
					$MyReport += Get-HTMLDetail "KMS Port" "N/A"
					$Details.KMSServer = "N/A"
					$Details.KMSPort = "N/A"
					if ($LicenseWin.GracePeriodRemaining -eq 0) {
						$GracePeriod = "N/A (not KMS)"
					} else {
						$GracePeriod = "$($([math]::round($LicenseWin.GracePeriodRemaining / 1440)) | Out-String -Stream) jours"
					}
				}
				$Details.PartialProductKey = $LicenseWin.ProductKeyChannel
				$MyReport += Get-HTMLDetail "Product Key Channel" ($LicenseWin.ProductKeyChannel)
				$MyReport += Get-HTMLDetail "Partial Product Key" ($LicenseWin.PartialProductKey)
				$MyReport += Get-HTMLDetail "Grace Period Remaining" $GracePeriod
				$MyReport += Get-HTMLDetail "License Family" ($LicenseWin.LicenseFamily)
				$MyReport += Get-HTMLDetail "License Status" ($LicenseStatus)
				$MyReport += Get-CustomHeaderClose
				
				$Details.PartialProductKey = $LicenseWin.PartialProductKey
				$Details.GracePeriodRemaining = $GracePeriod
				$Details.LicenseFamily = $LicenseWin.LicenseFamily
				$Details.LicenseStatus = $LicenseStatus
				$WindowsLicenseInfo += $Details
			#endregion WindowsLicenseInfo

			CalculDuree -message "License info"
			
			#region HotfixInfo
				Write-Output "..Hotfix Information"
				#$colQuickFixes = Get-WmiObject Win32_QuickFixEngineering -ComputerName $Target
				$colQuickFixes = Get-Hotfix -ComputerName $Target
				$MyReport += Get-CustomHeader "2" "HotFixes"
					$MyReport += Get-HTMLTable ($colQuickFixes | Where-Object {$_.HotFixID -ne "File 1" } | Select-Object HotFixID, Description,@{l="InstalledOn";e={[DateTime]$_.psbase.properties["installedon"].value}})
				$MyReport += Get-CustomHeaderClose
				$HotfixInfo += $colQuickFixes | Where-Object {$_.HotFixID -ne "File 1" } | Select-Object @{Name="Computer Name";Expression={$Target}}, HotFixID, Description,@{l="InstalledOn";e={[DateTime]$_.psbase.properties["installedon"].value}}
			#endregion HotfixInfo
				
			CalculDuree -message "Hotfix info"

			#region LogicalDisksInfo
				Write-Output "..Logical Disks"
				$Disks = Get-WmiObject -ComputerName $Target Win32_LogicalDisk
				$MyReport += Get-CustomHeader "2" "Logical Disk Configuration"
					$LogicalDrives = @()
					Foreach ($LDrive in ($Disks | Where-Object {$_.DriveType -eq 3})){
						$Details = "" | Select-Object "Drive Letter", Label, "File System", "Disk Size", "Disk Free Space", "% Free Space"
						$Details."Drive Letter" = $LDrive.DeviceID
						$Details.Label = $LDrive.VolumeName
						$Details."File System" = $LDrive.FileSystem
						$Details."Disk Size" = [math]::round(($LDrive.size)) | ConvertTo-KMG
						$Details."Disk Free Space" = [math]::round(($LDrive.FreeSpace)) | ConvertTo-KMG
						$Details."% Free Space" = [Math]::Round(($LDrive.FreeSpace /1MB) / ($LDrive.Size / 1MB) * 100)
						$LogicalDrives += $Details
					}
					$MyReport += Get-HTMLTable ($LogicalDrives)
				$MyReport += Get-CustomHeaderClose
				$LogicalDrives | Add-Member -MemberType NoteProperty -Name "Computer Name" -Value $Target
				$LogicalDisksInfo += $LogicalDrives
			#endregion LogicalDisksInfo
				
			CalculDuree -message "Logical Disk info"

			#region NetworkInfo
				Write-Output "..Network Configuration"
				$Adapters = Get-WmiObject -ComputerName $Target Win32_NetworkAdapterConfiguration
				$MyReport += Get-CustomHeader "2" "NIC Configuration"
					$IPInfo = @()
					Foreach ($Adapter in ($Adapters | Where-Object {$_.IPEnabled -eq $True})) {
						$Details = "" | Select-Object Description, "Physical address", "IP Address / Subnet Mask", "Default Gateway", "DHCP Enabled", DNS, WINS, "LMHosts Enabled", "Netbios overTCP-IP Enabled"
						$Details.Description = "$($Adapter.Description)"
						$Details."Physical address" = "$($Adapter.MACaddress)"
						if ($Adapter.IPAddress -ne $Null) {
						$Details."IP Address / Subnet Mask" = "$($Adapter.IPAddress)/$($Adapter.IPSubnet)"
							$Details."Default Gateway" = "$($Adapter.DefaultIPGateway)"
						}
						if ($Adapter.DHCPEnabled -eq "True")	{
							$Details."DHCP Enabled" = "Yes"
						}
						else {
							$Details."DHCP Enabled" = "No"
						}
						if ($Adapter.DNSServerSearchOrder -ne $Null)	{
							$DNS = $Adapter.DNSServerSearchOrder -join " / "
							$Details.DNS =  "$DNS"
						}
						$Details.WINS = "$($Adapter.WINSPrimaryServer) $($Adapter.WINSSecondaryServer)"
						if ($Adapter.WINSEnableLMHostsLookup){
							$Details."LMHosts Enabled" = "Yes"
						}
						else {
							$Details."LMHosts Enabled" = "No"
						}
						if ($Adapter.TcpipNetbiosOptions) {
							$Details."Netbios overTCP-IP Enabled" = "Yes"
						}
						else {
							$Details."Netbios overTCP-IP Enabled" = "No"
						}
						$IPInfo += $Details
					}
					$MyReport += Get-HTMLTable ($IPInfo)
				$MyReport += Get-CustomHeaderClose
				$IPInfo | Add-Member -MemberType NoteProperty -Name "Computer Name" -Value $Target
				$NetworkInfo += $IPInfo
			#endregion NetworkInfo

			CalculDuree -message "Network info"
				
			#region SoftwareInfo
				Write-Output "..Software"
				$MyReport += Get-CustomHeader "2" "Software"
					$AllSoft = Get-RemoteProgram -ComputerName $Target -Property Publisher,InstallDate,DisplayVersion | Sort-Object ProgramName
					$MyReport += Get-HTMLTable $AllSoft
				$MyReport += Get-CustomHeaderClose
				$SoftwareInfo += $AllSoft
			#endregion SoftwareInfo

			CalculDuree -message "Software info"
				
			#region SharesInfo
				Write-Output "..Local Shares"
				$Shares = Get-wmiobject -ComputerName $Target Win32_Share
				$MyReport += Get-CustomHeader "2" "Local Shares"
					$MyReport += Get-HTMLTable ($Shares | Select-Object Name, Path, Caption)
				$MyReport += Get-CustomHeaderClose
				$SharesInfo += $Shares | Select-Object Name, Path, Caption,@{Name="Computer Name";Expression={$Target}}
			#endregion SharesInfo

			CalculDuree -message "Shares info"
				
			#region PrintersInfo
				Write-Output "..Printers"
				$InstalledPrinters =  Get-WmiObject -ComputerName $Target Win32_Printer
				$MyReport += Get-CustomHeader "2" "Printers"
					$MyReport += Get-HTMLTable ($InstalledPrinters | Select-Object Name, Location)
				$MyReport += Get-CustomHeaderClose
				$PrintersInfo += $InstalledPrinters | Select-Object Name, Location,@{Name="Computer Name";Expression={$Target}}
			#endregion PrintersInfo

			CalculDuree -message "Printer info"
				
			#region LocalUsersInfo
				Write-Output "..Local Users"
				#Verif si DC, si oui skip !
				if ($ComputerRole -ne "Domain Controller") {
					$adsi = [ADSI]"WinNT://$Target"
					$Users = $adsi.Children | Where-Object {$_.SchemaClassName -eq 'user'} | Foreach-Object {
						$groups = $_.Groups() | Foreach-Object {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)}
						$_ | Select-Object @{n='UserName';e={$_.Name}},@{n='Groups';e={$groups -join ';'}}
					}
				} else {
					$Users = "No User listing, Computer is Domain Controller"
				}
				$MyReport += Get-CustomHeader "1" "Local Users"
					$MyReport += Get-HTMLTable ($Users)
				$MyReport += Get-CustomHeaderClose
				$Users | Add-Member -MemberType NoteProperty -Name "Computer Name" -Value $Target
				$LocalUsersInfo += $Users
			#endregion LocalUsersInfo

			CalculDuree -message "Local User info"
				
			#region LocalGroupsInfo
				Write-Output "..Local Groups"
				$Groups = Get-WmiObject -ComputerName $Target Win32_Group -filter "LocalAccount=True" | Select-Object Name
				$MyReport += Get-CustomHeader "1" "Local Groups"
					$MyReport += Get-HTMLTable ($Groups)
				$MyReport += Get-CustomHeaderClose
				$Groups | Add-Member -MemberType NoteProperty -Name "Computer Name" -Value $Target
				$LocalGroupsInfo += $Groups
			#endregion LocalGroupsInfo

			CalculDuree -message "Local Group info"
				
			#region ServicesInfo
				Write-Output "..Services"
				$ListOfServices = Get-WmiObject -ComputerName $Target Win32_Service #| Sort-Object Name
				$MyReport += Get-CustomHeader "2" "Services"
					$Services = @()
					Foreach ($Service in $ListOfServices) {
						$Details = "" | Select-Object Name,Account,"Start Mode",State,"Expected State"
						$Details.Name = $Service.Caption
						$Details.Account = $Service.Startname
						$Details."Start Mode" = $Service.StartMode
						$Details.State = $Service.State
						If ($Service.StartMode -eq "Auto" -and $Service.State -eq "Stopped") {
							$Details."Expected State" = "Unexpected"
						}
						elseIf ($Service.StartMode -eq "Disabled" -and $Service.State -eq "Running") {
								$Details."Expected State" = "Unexpected"
						}
						else {
							$Details."Expected State" = "OK"
						}
						$Services += $Details
					}
					$MyReport += Get-HTMLTable ($Services | Sort-Object Name)
				$MyReport += Get-CustomHeaderClose
				$Services | Add-Member -MemberType NoteProperty -Name "Computer Name" -Value $Target
				$ServicesInfo += $Services | Sort-Object Name
			#endregion ServicesInfo

			CalculDuree -message "Service info"
			
			#region RegionalSettingsInfo
				$MyReport += Get-CustomHeader "2" "Regional Settings"
					$MyReport += Get-HTMLDetail "Time Zone" ($TimeZone.Description)
					$MyReport += Get-HTMLDetail "Country Code" ($OSCountryCode)
					$MyReport += Get-HTMLDetail "Locale" ($OSLocal)
					$MyReport += Get-HTMLDetail "Operating System Language" ($OSLang)
					$MyReport += Get-HTMLDetail "Keyboard Layout" ($keyboard)
				$MyReport += Get-CustomHeaderClose
				
				$Details = "" | Select-Object "Computer Name","Time Zone","Country Code",Locale,"Operating System Language","Keyboard Layout"
				$Details."Computer Name" = $ComputerSystem.Name
				$Details."Time Zone" = $TimeZone.Description
				$Details."Country Code" = $OSCountryCode
				$Details.Locale = $OSLocal
				$Details."Operating System Language" = $OSLang
				$Details."Keyboard Layout" = $keyboard
				$RegionalSettingsInfo += $Details
			#endregion RegionalSettingsInfo

			CalculDuree -message "Regional Settings info"
				
			#region EventLogInfo
				Write-Output "..Event Log Settings"
				$LogFiles = Get-WmiObject -ComputerName $Target Win32_NTEventLogFile
				$MyReport += Get-CustomHeader "2" "Event Logs"
					$MyReport += Get-CustomHeader "2" "Event Log Settings"
					$LogSettings = @()
					$OldestEvent = @()
					Foreach ($Log in $LogFiles){
						$Details = "" | Select-Object "Log Name", "Overwrite Outdated Records", "Maximum Size", "Current Size"
						$Details."Log Name" = $Log.LogFileName
						If ($Log.OverWriteOutdated -lt 0) {
							$Details."Overwrite Outdated Records" = "Never"
						}
						if ($Log.OverWriteOutdated -eq 0) {
							$Details."Overwrite Outdated Records" = "As needed"
						} Else {
							$Details."Overwrite Outdated Records" = "After $($Log.OverWriteOutdated) days"
						}
						$MaxFileSize = ($Log.MaxFileSize) | ConvertTo-KMG
						$FileSize = ($Log.FileSize) | ConvertTo-KMG
						$Details."Maximum Size" = $MaxFileSize
						$Details."Current Size" = $FileSize
						$LogSettings += $Details
						$OldestEvent += Get-WinEvent -Oldest -FilterHashtable @{logname=$Log.LogFileName} -ComputerName $Target -EA 0 -Verbose:$false | Select-Object -First 1 | Select-Object @{Name="Computer Name";Expression={$Target}},@{Name="LogName";Expression={$($Log.LogFileName)}},LevelDisplayName,ID,TimeCreated,ProviderName,MachineName
					}
					$MyReport += Get-HTMLTable ($LogSettings)
					$MyReport += Get-CustomHeaderClose
					$LogSettings | Add-Member -MemberType NoteProperty -Name "Computer Name" -Value $Target
					$EventLogInfo += $LogSettings
					
					Write-Output "....Event Log Oldest Event"
					$MyReport += Get-CustomHeader "2" "Oldest Event"
					$MyReport += Get-HTMLTable ($OldestEvent)
					$MyReport += Get-CustomHeaderClose
					
					Write-Output "....Event Log Last Boot Errors"
					$BootIndex = (Get-WinEvent -FilterHashtable @{logname='system';ID=6005;StartTime=$LBTime} -ComputerName $Target -EA 0 -Verbose:$false).RecordId
					$BootIndexEnd = $BootIndex + 30
					$LoggedBootErrors = Get-WinEvent -FilterHashtable @{logname='system';ID=6005;StartTime=$LBTime;level=2} -ComputerName $Target -EA 0 -Verbose:$false |  Where-Object {$_.recordid -ge $BootIndex -and $_.recordid -le $BootIndexEnd}
					$MyReport += Get-CustomHeader "2" "LastBOOT ERROR Entries"
						$MyReport += Get-HTMLTable ($LoggedBootErrors | Select-Object InstanceId, Source, TimeWritten, Message)
					$MyReport += Get-CustomHeaderClose
					<#
						Write-Output "....Event Log Errors"
						$WmidtQueryDT = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime([DateTime]::Now.AddDays(-14))
						$LoggedErrors = Get-WmiObject -ComputerName $Target -query ("Select * from Win32_NTLogEvent Where Type='Error' and TimeWritten >='" + $WmidtQueryDT + "'")
						$MyReport += Get-CustomHeader "2" "ERROR Entries"
							$MyReport += Get-HTMLTable ($LoggedErrors | Select EventCode, SourceName, @{N="Time";E={$_.ConvertToDateTime($_.TimeWritten)}}, LogFile, Message)
						$MyReport += Get-CustomHeaderClose
						
						Write-Output "....Event Log Warnings"
						$WmidtQueryDT = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime([DateTime]::Now.AddDays(-14))
						$LoggedWarning = Get-WmiObject -ComputerName $Target -query ("Select * from Win32_NTLogEvent Where Type='Warning' and TimeWritten >='" + $WmidtQueryDT + "'")
						$MyReport += Get-CustomHeader "2" "WARNING Entries"
							$MyReport += Get-HTMLTable ($LoggedWarning | Select EventCode, SourceName, @{N="Time";E={$_.ConvertToDateTime($_.TimeWritten)}}, LogFile, Message)
						$MyReport += Get-CustomHeaderClose
					#>					
				$MyReport += Get-CustomHeaderClose
				if ($LoggedBootErrors) {
					$LoggedBootErrors | Add-Member -MemberType NoteProperty -Name "Computer Name" -Value $Target
				}
				$EventLogBootInfo += $LoggedBootErrors | Select-Object @{Name="Computer Name";Expression={$Target}},InstanceId, Source, TimeWritten, Message
				$EventLogOldestEvt += $OldestEvent
			#endregion EventLogInfo

			CalculDuree -message "Eventlog info"

			#region ScheduleTasksInfo
				Write-Output "..Schedule Tasks"
				$MyReport += Get-CustomHeader "2" "Schedule Tasks"
				foreach ($Task in $ScheduleTasks) {
					$MyReport += Get-CustomHeader "2" "$($Task.TaskName)"
						$MyReport += Get-HTMLDetail "TaskName" ($Task.TaskName)
						$MyReport += Get-HTMLDetail "RunAs" ($Task.RunAs)
						$MyReport += Get-HTMLDetail "ActionNumber" ($Task.ActionNumber)
						$MyReport += Get-HTMLDetail "ActionType" ($Task.ActionType)
						$MyReport += Get-HTMLDetail "Action" ($Task.Action)
						$MyReport += Get-HTMLDetail "LastRunTime" ($Task.LastRunTime)
						$MyReport += Get-HTMLDetail "LastResult" ($Task.LastResult)
						$MyReport += Get-HTMLDetail "NextRunTime" ($Task.NextRunTime)
						$MyReport += Get-HTMLDetail "Enable" ($Task.Enabled)
						$MyReport += Get-HTMLDetail "Author" ($Task.Author)
						$MyReport += Get-HTMLDetail "Created" ($Task.Created)
						$MyReport += Get-HTMLDetail "Elevated" ($Task.Elevated)
					$MyReport += Get-CustomHeaderClose
					
					$Details = "" | Select-Object ComputerName,TaskName,RunAs,ActionNumber,ActionType,Action,LastRunTime,LastResult,NextRunTime,Enable,Author,Created,Elevated
					$Details.ComputerName = $ComputerSystem.Name
					$Details.TaskName = $Task.TaskName
					$Details.RunAs = $Task.RunAs
					$Details.ActionNumber = $Task.ActionNumber
					$Details.ActionType = $Task.ActionType
					$Details.Action = $Task.Action
					$Details.LastRunTime = $Task.LastRunTime
					$Details.LastResult = $Task.LastResult
					$Details.NextRunTime = $Task.NextRunTime
					$Details.Enable = $Task.Enabled
					$Details.Author = $Task.Author
					$Details.Created = $Task.Created
					$Details.Elevated = $Task.Elevated
					$ScheduleTasksInfo += $Details
				}
				$MyReport += Get-CustomHeaderClose
			$MyReport += Get-CustomHeader0Close
			$MyReport += Get-CustomHTMLClose
			$MyReport += Get-CustomHTMLClose
			#endregion ScheduleTasksInfo

			CalculDuree -message "Schedule Tasks info"
			
			#region ExportData
			$Date = Get-Date -f yyyy-MM-dd_HHmm
			$Filename = ".\" + $Target + "_" + $date + ".htm"
			$MyReport | out-file -encoding ASCII -filepath $Filename
			Write-Output "Audit saved as $Filename"
			#endregion ExportData

			CalculDuree -message "Exporting to HTML"


		}
		else {
			Write-Output "Skipping $Target, RPC-ping failed"
		}
	}
	
	if ($ExportCSV) {
		Write-Output "Generating CSV"
		#Export en CSV de l'integralite des machines auditees
		$RegionalOptions 		| Export-Csv -NoTypeInformation ".\RegionalOptions.csv"
		$SystemInfo 			| Export-Csv -NoTypeInformation ".\SystemInfo.csv"
		$WindowsLicenseInfo 	| Export-Csv -NoTypeInformation ".\WindowsLicenseInfo.csv"
		$HotfixInfo 			| Export-Csv -NoTypeInformation ".\HotfixInfo.csv"
		$LogicalDisksInfo 		| Export-Csv -NoTypeInformation ".\LogicalDisksInfo.csv"
		$NetworkInfo 			| Export-Csv -NoTypeInformation ".\NetworkInfo.csv"
		$SoftwareInfo 			| Export-Csv -NoTypeInformation ".\SoftwareInfo.csv"
		$SharesInfo 			| Export-Csv -NoTypeInformation ".\SharesInfo.csv"
		$PrintersInfo 			| Export-Csv -NoTypeInformation ".\PrintersInfo.csv"
		$LocalUsersInfo			| Export-Csv -NoTypeInformation ".\LocalUsersInfo.csv"
	$LocalGroupsInfo 			| Export-Csv -NoTypeInformation ".\LocalGroupsInfo.csv"
		$ServicesInfo 			| Export-Csv -NoTypeInformation ".\ServicesInfo.csv"
		$RegionalSettingsInfo 	| Export-Csv -NoTypeInformation ".\RegionalSettingsInfo.csv"
		$EventLogInfo 			| Export-Csv -NoTypeInformation ".\EventLogInfo.csv"
		if ($EventLogBootInfo) 
		{
			$EventLogBootInfo 	| Export-Csv -NoTypeInformation ".\EventLogBootInfo.csv"
		}
		if ($EventLogOldestEvt) 
		{ 
			$EventLogOldestEvt 	| Export-Csv -NoTypeInformation ".\EventLogOldestEvent.csv"
		}
		$ScheduleTasksInfo		| Export-Csv -NoTypeInformation ".\ScheduleTasksInfo.csv"
	}
}