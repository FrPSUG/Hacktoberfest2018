function CalculDuree {
	<#
	fonction qui sert à afficher une durée d'execution, la variable StartTime est positionnée une première fois dans le scritp principale
	#>
	param ($message)
	$EndTime = Get-Date
    $ExecutionTime = [math]::Round(($EndTime-$script:StartTime).totalseconds,2)
	$script:StartTime = Get-Date
	Write-Verbose "$message : $ExecutionTime sec"
}


function Ping-Host { 
	Param([string]$computername = $(Throw "You must specify a computername.")) 
	Write-Debug "In Ping-Host function" 
	$query = "Select * from Win32_PingStatus where address='$computername'" 
	$wmiping = Get-WmiObject -query $query 
	write $wmiping
}

function RPC-Ping {
	<#
		.SYNOPSIS
			RPC-Ping.ps1 - Test an RPC connection against one or more computer(s)
		.DESCRIPTION
			RPC-Ping - Test an RPC connection (WMI request) against one or more computer(s)
			with test-connection before to see if the computer is reachable or not first
		.PARAMETER ComputerName
			Defines the computer name or IP address to tet the RPC connection. Could be an array of servernames
			Mandatory parameter.
		.NOTES
			File Name   : RPC-Ping.ps1
			Author      : Fabrice ZERROUKI - fabricezerrouki@hotmail.com
		.EXAMPLE
			PS D:\> .\RPC-Ping.ps1 -ComputerName SERVER1
			Open an RPC connection against SERVER1
		.EXAMPLE
			PS D:\> .\RPC-Ping.ps1 -ComputerName SERVER1,192.168.0.23
			Open an RPC connection against SERVER1 and 192.168.0.23
	#>
	Param (
		[Parameter(Mandatory = $true, HelpMessage = "You must provide a computername or an IP address to test")]
		[string[]]$ComputerName
	)
	ForEach ($Computer in $ComputerName)
	{
		if (Test-Connection -ComputerName $Computer -Quiet -Count 1)
		{
			if (Get-WmiObject win32_computersystem -ComputerName $Computer -ErrorAction SilentlyContinue)
			{
				Write-Host "RPC connection on computer $Computer successful." -ForegroundColor cyan;
				return $True
			}
			else
			{
				Write-Host "RPC connection on computer $Computer failed!" -ForegroundColor Magenta;
				return $False
			}
		}
		else
		{
			Write-Host "Computer $Computer doesn't even responds to ping..." -ForegroundColor DarkRed;
			return $False
		}
	}
}

Function Get-MALScheduledTasks {
	[CmdletBinding()]
	param(
	  [parameter(Position=0)] [String[]] $TaskName="*",
	  [parameter(Position=1,ValueFromPipeline=$TRUE)] [String[]] $ComputerName=$ENV:COMPUTERNAME,
	  [switch] $Subfolders,
	  [switch] $Hidden,
	  [System.Management.Automation.PSCredential] $ConnectionCredential
	)
	
	begin {
	  $PIPELINEINPUT = (-not $PSBOUNDPARAMETERS.ContainsKey("ComputerName")) -and (-not $ComputerName)
	  $MIN_SCHEDULER_VERSION = "1.2"
	  $TASK_ENUM_HIDDEN = 1
	  $TASK_STATE = @{0 = "Unknown"; 1 = "Disabled"; 2 = "Queued"; 3 = "Ready"; 4 = "Running"}
	  $ACTION_TYPE = @{0 = "Execute"; 5 = "COMhandler"; 6 = "Email"; 7 = "ShowMessage"}
	
	  # Try to create the TaskService object on the local computer; throw an error on failure
	  try {
		$TaskService = new-object -comobject "Schedule.Service"
	  }
	  catch [System.Management.Automation.PSArgumentException] {
		throw $_
	  }
	
	  # Returns the specified PSCredential object's password as a plain-text string
	  function get-plaintextpwd($credential) {
		$credential.GetNetworkCredential().Password
	  }
	
	  # Returns a version number as a string (x.y); e.g. 65537 (10001 hex) returns "1.1"
	  function convertto-versionstr([Int] $version) {
		$major = [Math]::Truncate($version / [Math]::Pow(2, 0x10)) -band 0xFFFF
		$minor = $version -band 0xFFFF
		"$($major).$($minor)"
	  }
	
	  # Returns a string "x.y" as a version number; e.g., "1.3" returns 65539 (10003 hex)
	  function convertto-versionint([String] $version) {
		$parts = $version.Split(".")
		$major = [Int] $parts[0] * [Math]::Pow(2, 0x10)
		$major -bor [Int] $parts[1]
	  }
	
	  # Returns a list of all tasks starting at the specified task folder
	  function get-task($taskFolder) {
		$tasks = $taskFolder.GetTasks($Hidden.IsPresent -as [Int])
		$tasks | foreach-object { $_ }
		if ($SubFolders) {
		  try {
			$taskFolders = $taskFolder.GetFolders(0)
			$taskFolders | foreach-object { get-task $_ $TRUE }
		  }
		  catch [System.Management.Automation.MethodInvocationException] {
		  }
		}
	  }
	
	  # Returns a date if greater than 12/30/1899 00:00; otherwise, returns nothing
	  function get-OLEdate($date) {
		if ($date -gt [DateTime] "12/30/1899") { $date }
	  }
	
	  function get-scheduledtask2($computerName) {
		# Assume $NULL for the schedule service connection parameters unless -ConnectionCredential used
		$userName = $domainName = $connectPwd = $NULL
		if ($ConnectionCredential) {
		  # Get user name, domain name, and plain-text copy of password from PSCredential object
		  $userName = $ConnectionCredential.UserName.Split("\")[1]
		  $domainName = $ConnectionCredential.UserName.Split("\")[0]
		  $connectPwd = get-plaintextpwd $ConnectionCredential
		}
		try {
		  $TaskService.Connect($ComputerName, $userName, $domainName, $connectPwd)
		}
		catch [System.Management.Automation.MethodInvocationException] {
		  write-warning "$computerName - $_"
		  return
		}
		$serviceVersion = convertto-versionstr $TaskService.HighestVersion
		$vistaOrNewer = (convertto-versionint $serviceVersion) -ge (convertto-versionint $MIN_SCHEDULER_VERSION)
		$rootFolder = $TaskService.GetFolder("\")
		$taskList = get-task $rootFolder
		if (-not $taskList) { return }
		foreach ($task in $taskList) {
		  foreach ($name in $TaskName) {
			# Assume root tasks folder (\) if task folders supported
			if ($vistaOrNewer) {
			  if (-not $name.Contains("\")) { $name = "\$name" }
			}
			if ($task.Path -notlike $name) { continue }
			$taskDefinition = $task.Definition
			$actionCount = 0
			foreach ($action in $taskDefinition.Actions) {
			  $actionCount += 1
			  $output = new-object PSObject
			  # PROPERTY: ComputerName
			  $output | add-member NoteProperty ComputerName $computerName
			  # PROPERTY: ServiceVersion
			  $output | add-member NoteProperty ServiceVersion $serviceVersion
			  # PROPERTY: TaskName
			  if ($vistaOrNewer) {
				$output | add-member NoteProperty TaskName $task.Path
			  } else {
				$output | add-member NoteProperty TaskName $task.Name
			  }
			  #PROPERTY: Enabled
			  $output | add-member NoteProperty Enabled ([Boolean] $task.Enabled)
			  # PROPERTY: ActionNumber
			  $output | add-member NoteProperty ActionNumber $actionCount
			  # PROPERTIES: ActionType and Action
			  # Old platforms return null for the Type property
			  if ((-not $action.Type) -or ($action.Type -eq 0)) {
				$output | add-member NoteProperty ActionType $ACTION_TYPE[0]
				$output | add-member NoteProperty Action "$($action.Path) $($action.Arguments)"
			  } else {
				$output | add-member NoteProperty ActionType $ACTION_TYPE[$action.Type]
				$output | add-member NoteProperty Action $NULL
			  }
			  # PROPERTY: LastRunTime
			  $output | add-member NoteProperty LastRunTime (get-OLEdate $task.LastRunTime)
			  # PROPERTY: LastResult
			  if ($task.LastTaskResult) {
				# If negative, convert to DWORD (UInt32)
				if ($task.LastTaskResult -lt 0) {
				  $lastTaskResult = "0x{0:X}" -f [UInt32] ($task.LastTaskResult + [Math]::Pow(2, 32))
				} else {
				  $lastTaskResult = "0x{0:X}" -f $task.LastTaskResult
				}
			  } else {
				$lastTaskResult = $NULL  # fix bug in v1.0-1.1 (should output $NULL)
			  }
			  $output | add-member NoteProperty LastResult $lastTaskResult
			  # PROPERTY: NextRunTime
			  $output | add-member NoteProperty NextRunTime (get-OLEdate $task.NextRunTime)
			  # PROPERTY: State
			  if ($task.State) {
				$taskState = $TASK_STATE[$task.State]
			  }
			  $output | add-member NoteProperty State $taskState
			  $regInfo = $taskDefinition.RegistrationInfo
			  # PROPERTY: Author
			  $output | add-member NoteProperty Author $regInfo.Author
			  # The RegistrationInfo object's Date property, if set, is a string
			  if ($regInfo.Date) {
				$creationDate = [DateTime]::Parse($regInfo.Date)
			  } else {
				$creationDate = 'N/A'
			  }
			  $output | add-member NoteProperty Created $creationDate
			  # PROPERTY: RunAs
			  $principal = $taskDefinition.Principal
			  $output | add-member NoteProperty RunAs $principal.UserId
			  # PROPERTY: Elevated
			  if ($vistaOrNewer) {
				if ($principal.RunLevel -eq 1) { $elevated = $TRUE } else { $elevated = $FALSE }
			  }
			  $output | add-member NoteProperty Elevated $elevated
			  # Output the object
			  $output
			}
		  }
		}
	  }
	}
	
	process {
	  if ($PIPELINEINPUT) {
		get-scheduledtask2 $_
	  }
	  else {
		$ComputerName | foreach-object {
		  get-scheduledtask2 $_
		}
	  }
	}
	

}

function ts {
	$input | %{
		$_ -as [string]
	}
}

Filter ConvertTo-KMG {
	$bytecount = $_
		switch ([math]::truncate([math]::log($bytecount,1024))) 
		{
			0 {"$bytecount Bytes"}
			1 {"{0:n2} KB" -f ($bytecount / 1kb)}
			2 {"{0:n2} MB" -f ($bytecount / 1mb)}
			3 {"{0:n2} GB" -f ($bytecount / 1gb)}
			4 {"{0:n2} TB" -f ($bytecount / 1tb)}
			Default {"{0:n2} PB" -f ($bytecount / 1pb)}
		}
}

Function Get-RemoteProgram {
	<#
		.Synopsis
		Generates a list of installed programs on a computer

		.DESCRIPTION
		This function generates a list by querying the registry and returning the installed programs of a local or remote computer.

		.NOTES   
		Name       : Get-RemoteProgram
		Author     : Jaap Brasser
		Version    : 1.3
		DateCreated: 2013-08-23
		DateUpdated: 2016-08-26
		Blog       : http://www.jaapbrasser.com

		.LINK
		http://www.jaapbrasser.com

		.PARAMETER ComputerName
		The computer to which connectivity will be checked

		.PARAMETER Property
		Additional values to be loaded from the registry. Can contain a string or an array of string that will be attempted to retrieve from the registry for each program entry

		.PARAMETER ExcludeSimilar
		This will filter out similar programnames, the default value is to filter on the first 3 words in a program name. If a program only consists of less words it is excluded and it will not be filtered. For example if you Visual Studio 2015 installed it will list all the components individually, using -ExcludeSimilar will only display the first entry.

		.PARAMETER SimilarWord
		This parameter only works when ExcludeSimilar is specified, it changes the default of first 3 words to any desired value.

		.EXAMPLE
		Get-RemoteProgram

		Description:
		Will generate a list of installed programs on local machine

		.EXAMPLE
		Get-RemoteProgram -ComputerName server01,server02

		Description:
		Will generate a list of installed programs on server01 and server02

		.EXAMPLE
		Get-RemoteProgram -ComputerName Server01 -Property DisplayVersion,VersionMajor

		Description:
		Will gather the list of programs from Server01 and attempts to retrieve the displayversion and versionmajor subkeys from the registry for each installed program

		.EXAMPLE
		'server01','server02' | Get-RemoteProgram -Property Uninstallstring

		Description
		Will retrieve the installed programs on server01/02 that are passed on to the function through the pipeline and also retrieves the uninstall string for each program

		.EXAMPLE
		'server01','server02' | Get-RemoteProgram -Property Uninstallstring -ExcludeSimilar -SimilarWord 4

		Description
		Will retrieve the installed programs on server01/02 that are passed on to the function through the pipeline and also retrieves the uninstall string for each program. Will only display a single entry of a program of which the first four words are identical.
	#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(ValueFromPipeline              =$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0
        )]
        [string[]]
            $ComputerName = $env:COMPUTERNAME,
        [Parameter(Position=0)]
        [string[]]
            $Property,
        [switch]
            $ExcludeSimilar,
        [int]
            $SimilarWord
    )

    begin {
        $RegistryLocation = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\',
                            'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
        $HashProperty = @{}
        $SelectProperty = @('ProgramName','ComputerName')
        if ($Property) {
            $SelectProperty += $Property
        }
    }

    process {
        foreach ($Computer in $ComputerName) {
            $RegBase = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$Computer)
            $RegistryLocation | ForEach-Object {
                $CurrentReg = $_
                if ($RegBase) {
                    $CurrentRegKey = $RegBase.OpenSubKey($CurrentReg)
                    if ($CurrentRegKey) {
                        $CurrentRegKey.GetSubKeyNames() | ForEach-Object {
                            if ($Property) {
                                foreach ($CurrentProperty in $Property) {
                                    $HashProperty.$CurrentProperty = ($RegBase.OpenSubKey("$CurrentReg$_")).GetValue($CurrentProperty)
                                }
                            }
                            $HashProperty.ComputerName = $Computer
                            $HashProperty.ProgramName = ($DisplayName = ($RegBase.OpenSubKey("$CurrentReg$_")).GetValue('DisplayName'))
                            if ($DisplayName) {
                                New-Object -TypeName PSCustomObject -Property $HashProperty |
                                Select-Object -Property $SelectProperty
                            } 
                        }
                    }
                }
            } | ForEach-Object -Begin {
                if ($SimilarWord) {
                    $Regex = [regex]"(^(.+?\s){$SimilarWord}).*$|(.*)"
                } else {
                    $Regex = [regex]"(^(.+?\s){3}).*$|(.*)"
                }
                [System.Collections.ArrayList]$Array = @()
            } -Process {
                if ($ExcludeSimilar) {
                    $null = $Array.Add($_)
                } else {
                    $_
                }
            } -End {
                if ($ExcludeSimilar) {
                    $Array | Select-Object -Property *,@{
                        name       = 'GroupedName'
                        expression = {
                            ($_.ProgramName -split $Regex)[1]
                        }
                    } |
                    Group-Object -Property 'GroupedName' | ForEach-Object {
                        $_.Group[0] | Select-Object -Property * -ExcludeProperty GroupedName
                    }
                }
            }
        }
    }
}