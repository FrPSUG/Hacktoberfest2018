function Get-RegValue
{

	<#
	.SYNOPSIS
	       Sets the default value (REG_SZ) of the registry key on local or remote computers.

	.DESCRIPTION
	       Use Get-RegValue to set the default value (REG_SZ) of the registry key on local or remote computers.
	       
	.PARAMETER ComputerName
	    	An array of computer names. The default is the local computer.

	.PARAMETER Hive
	   	The HKEY to open, from the RegistryHive enumeration. The default is 'LocalMachine'.
	   	Possible values:
	   	
		- ClassesRoot
		- CurrentUser
		- LocalMachine
		- Users
		- PerformanceData
		- CurrentConfig
		- DynData	   	

	.PARAMETER Key
	       The path of the registry key to open.  

	.PARAMETER Value
	       The name of the registry value, Wildcards are permitted.

	   .PARAMETER Type

	   	A collection of data types of registry values, from the RegistryValueKind enumeration.
	   	Possible values:

		- Binary
		- DWord
		- ExpandString
		- MultiString
		- QWord
		- String
		
		When the parameter is not specified all types are returned, Wildcards are permitted.
		
	   .PARAMETER Recurse
	   	Gets the registry values of the specified registry key and its sub keys.

	.PARAMETER Ping
	       Use ping to test if the machine is available before connecting to it. 
	       If the machine is not responding to the test a warning message is output.
       		
	.EXAMPLE	   
		Get-RegValue -Key SOFTWARE\Microsoft\PowerShell\1 -Recurse

		ComputerName Hive            Key                  Value                     Data                 Type
		------------ ----            ---                  -----                     ----                 ----
		COMPUTER1    LocalMachine    SOFTWARE\Microsof... Install                   1                    DWord
		COMPUTER1    LocalMachine    SOFTWARE\Microsof... PID                       89383-100-0001260... String
		COMPUTER1    LocalMachine    SOFTWARE\Microsof... Install                   1                    DWord
		COMPUTER1    LocalMachine    SOFTWARE\Microsof... ApplicationBase           C:\Windows\System... String
		COMPUTER1    LocalMachine    SOFTWARE\Microsof... PSCompatibleVersion       1.0, 2.0             String
		COMPUTER1    LocalMachine    SOFTWARE\Microsof... RuntimeVersion            v2.0.50727           String
		(...)
		
		
		Description
		-----------
		Gets all values of the PowerShell subkey on the local computer regardless of their type.		

	.EXAMPLE
		"SERVER1" | Get-RegValue -Key SOFTWARE\Microsoft\PowerShell\1 -Type String,DWord -Recurse -Ping

		ComputerName Hive            Key                  Value                     Data                 Type
		------------ ----            ---                  -----                     ----                 ----
		SERVER1      LocalMachine    SOFTWARE\Microsof... Install                   1                    DWord
		SERVER1      LocalMachine    SOFTWARE\Microsof... PID                       89383-100-0001260... String
		SERVER1      LocalMachine    SOFTWARE\Microsof... Install                   1                    DWord
		SERVER1      LocalMachine    SOFTWARE\Microsof... ApplicationBase           C:\Windows\System... String
		SERVER1      LocalMachine    SOFTWARE\Microsof... PSCompatibleVersion       1.0, 2.0             String
		(...)
	
		Description
		-----------
		Gets all String and DWord values of the PowerShell subkey and its subkeys from remote computer SERVER1, ping the remote server first.				

	.EXAMPLE
		Get-RegValue -ComputerName SERVER1 -Key SOFTWARE\Microsoft\PowerShell -Type MultiString -Value t* -Recurse

		ComputerName Hive            Key                  Value  Data                 Type
		------------ ----            ---                  -----  ----                 ----
		SERVER1      LocalMachine    SOFTWARE\Microsof... Types  {virtualmachinema... MultiString
		SERVER1      LocalMachine    SOFTWARE\Microsof... Types  {C:\Program Files... MultiString

		Description
		-----------
		Gets all MultiString value names, from the subkey and its subkeys, that starts with the 't' letter from remote computer SERVER1.				

	.OUTPUTS
		System.Boolean
		PSFanatic.Registry.RegistryValue (PSCustomObject)

	.NOTES
		Author: Shay Levy
		Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/
	
	.LINK
		http://code.msdn.microsoft.com/PSRemoteRegistry

	.LINK
		Set-RegValue
		Test-RegValue
		Remove-RegValue	

	#>	
	

	[OutputType('System.Boolean','PSFanatic.Registry.RegistryValue')]
	[CmdletBinding(DefaultParameterSetName="__AllParameterSets")]
	
	param( 
		[Parameter(
			Position=0,
			ValueFromPipeline=$true,
			ValueFromPipelineByPropertyName=$true
		)]		
		[Alias("CN","__SERVER","IPAddress")]
		[string[]]$ComputerName="",		

		[Parameter(
			Position=1,
			ValueFromPipelineByPropertyName=$true,
			HelpMessage="The HKEY to open, from the RegistryHive enumeration. The default is 'LocalMachine'."
		)]
		[ValidateSet("ClassesRoot","CurrentUser","LocalMachine","Users","PerformanceData","CurrentConfig","DynData")]
		[string]$Hive="LocalMachine",

		[Parameter(
			Mandatory=$true,
			Position=2,
			ValueFromPipelineByPropertyName=$true,
			HelpMessage="The path of the subkey to open."
		)]
		[string]$Key,
		
		[Parameter(
			Mandatory=$false,
			Position=3,
			ValueFromPipelineByPropertyName=$true,
			HelpMessage="The name of the value to set."
		)]	
		[string]$Value="*",		

		[Parameter(
			Mandatory=$false,
			Position=4,
			ValueFromPipelineByPropertyName=$true,
			HelpMessage="The data type of the registry value."
		)]
		[ValidateSet("String","ExpandString","Binary","DWord","MultiString","QWord")]
		[string[]]$Type="*",
		
		[switch]$Ping,
		
		[switch]$Recurse
	) 

	begin
	{
		Write-Verbose "Enter begin block..."
	
		function Recurse($Key){
		
			Write-Verbose "Start recursing, key is [$Key]"

			try
			{
			
				$subKey = $reg.OpenSubKey($key)
				
				if(!$subKey)
				{
					Throw "Key '$Key' doesn't exist."
				}
				

				foreach ($v in $subKey.GetValueNames())
				{
					$vk = $subKey.GetValueKind($v)
					
					foreach($t in $Type)
					{	
						if($v -like $Value -AND $vk -like $t)
						{						
							$pso = New-Object PSObject -Property @{
								ComputerName=$c
								Hive=$Hive
								Value=if(!$v) {"(Default)"} else {$v}
								Key=$Key
								Data=$subKey.GetValue($v)
								Type=$vk
							}

							Write-Verbose "Recurse: Adding format type name to custom object."
							$pso.PSTypeNames.Clear()
							$pso.PSTypeNames.Add('PSFanatic.Registry.RegistryValue')
							$pso
						}
					}
				}
				
				foreach ($k in $subKey.GetSubKeyNames())
				{
					Recurse "$Key\$k"		
				}
				
			}
			catch
			{
				Write-Error $_
			}
			
			Write-Verbose "Ending recurse, key is [$Key]"
		}
		
		Write-Verbose "Exit begin block..."
	}
	

	process
	{


	    	Write-Verbose "Enter process block..."
		
		foreach($c in $ComputerName)
		{	
			try
			{				
				if($c -eq "")
				{
					$c=$env:COMPUTERNAME
					Write-Verbose "Parameter [ComputerName] is not present, setting its value to local computer name: [$c]."
					
				}
				
				if($Ping)
				{
					Write-Verbose "Parameter [Ping] is present, initiating Ping test"
					
					if( !(Test-Connection -ComputerName $c -Count 1 -Quiet))
					{
						Write-Warning "[$c] doesn't respond to ping."
						return
					}
				}
				
				
				Write-Verbose "Starting remote registry connection against: [$c]."
				Write-Verbose "Registry Hive is: [$Hive]."
				$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]$Hive,$c)		
	
								
				if($Recurse)
				{
					Write-Verbose "Parameter [Recurse] is present, calling Recurse function."
					Recurse $Key
				}
				else
				{					
				
					Write-Verbose "Open remote subkey: [$Key]."			
					$subKey = $reg.OpenSubKey($Key)
					
					if(!$subKey)
					{
						Throw "Key '$Key' doesn't exist."
					}
					
					Write-Verbose "Start get remote subkey: [$Key] values."
					foreach ($v in $subKey.GetValueNames())
					{						
						$vk = $subKey.GetValueKind($v)
						
						foreach($t in $Type)
						{					
							if($v -like $Value -AND $vk -like $t)
							{														
								$pso = New-Object PSObject -Property @{
									ComputerName=$c
									Hive=$Hive
									Value= if(!$v) {"(Default)"} else {$v}
									Key=$Key
									Data=$subKey.GetValue($v)
									Type=$vk
								}

								Write-Verbose "Adding format type name to custom object."
								$pso.PSTypeNames.Clear()
								$pso.PSTypeNames.Add('PSFanatic.Registry.RegistryValue')
								$pso
							}
						}
					}				
				}
				
				Write-Verbose "Closing remote registry connection on: [$c]."
				$reg.close()
			}
			catch
			{
				Write-Error $_
			}
		} 
	
		Write-Verbose "Exit process block..."
	}
}