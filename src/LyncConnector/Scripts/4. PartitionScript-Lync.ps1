<#
<copyright file="PartitionScript-Lync.ps1" company="Microsoft">
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
</copyright>
<summary>
	The Partition script for the Skype 2015 / Lync 2010 / 2013 Connector.
</summary>
#>

[CmdletBinding()]
param(
	[parameter(Mandatory = $true)]
	[System.Collections.ObjectModel.KeyedCollection[string,Microsoft.MetadirectoryServices.ConfigParameter]]
	$ConfigParameters,
	[parameter(Mandatory = $true)]
	[Alias('PSCredential')] # To fix mess-up of the parameter name in the RTM version of the PowerShell connector.
	[System.Management.Automation.PSCredential]
	$Credential,
	[parameter(Mandatory = $false)]
	[ValidateScript({ Test-Path $_ -PathType "Container" })]
	[string]
	$ScriptDir = (Join-Path -Path $env:windir -ChildPath "TEMP") # Optional parameter for manipulation by the TestHarness script.
)

Set-StrictMode -Version "2.0"

$Global:DebugPreference = "Continue"
$Global:VerbosePreference = "Continue"

$commonModule = (Join-Path -Path $ScriptDir -ChildPath $ConfigParameters["Common Module Script Name (with extension)"].Value)

if (!(Get-Module -Name (Get-Item $commonModule).BaseName)) { Import-Module -Name $commonModule }

Enter-Script -ScriptType "Partition" -ErrorObject $Error

function Get-Partitions
{
	<#
	.Synopsis
		Gets the partitions in the forest to which the connector service account belongs.
	.Description
		Gets the partitions in the forest to which the connector service account belongs.
	#>

	[CmdletBinding()]
    [OutputType([System.Collections.Generic.List[Microsoft.MetadirectoryServices.Partition]])]
	param(
	)
	
	$partitions = New-GenericObject System.Collections.Generic.List Microsoft.MetadirectoryServices.Partition

	$userName = "{0}\{1}" -f $Credential.GetNetworkCredential().Domain, $Credential.GetNetworkCredential().UserName
	$password = $Credential.GetNetworkCredential().Password

	$rootDSEQuery = "LDAP://{0}/rootDSE" -f $preferredDomainController

	Write-Debug "Enumerating Directory Partitions. RootDSE: '$rootDSEQuery'."

	$rootDSE = New-Object System.DirectoryServices.DirectoryEntry -ArgumentList $rootDSEQuery, $userName, $password

	if ([string]::IsNullOrEmpty($password)) # Check for a bug in ECMA 2.0 on refreshing partion, apparently fixed now.
	{
		# Try with SyncService credentials - work around for the bug in ECMA 2.0, apparently not needed now.
		Write-Debug "Hitting ECMA 2.0 bug. Trying with SyncSerice credentials to connect to RootDSE: '$rootDSEQuery'."

		$rootDSE = New-Object System.DirectoryServices.DirectoryEntry -ArgumentList $rootDSEQuery
	}

	if (!$rootDSE)
	{
		throw "Unable to get RootDSE."
	}

	$configurationContainer = $rootDSE.Properties["configurationnamingcontext"].Value.ToString();

	$searchRootQuery = "LDAP://{0}/CN=Partitions,{1}" -f $preferredDomainController, $configurationContainer

	Write-Debug "Enumerating Directory Partitions. Configuration Container: '$searchRootQuery'."

	$searchRoot = New-Object System.DirectoryServices.DirectoryEntry -ArgumentList $searchRootQuery, $userName, $password

	if ([string]::IsNullOrEmpty($password)) # Check for a bug in ECMA 2.0 on refreshing partion, apparently fixed now.
	{
		# Try with SyncService credentials - work around for the bug in ECMA 2.0, apparently not needed now.
		Write-Debug "Hitting ECMA 2.0 bug. Trying with SyncSerice credentials to connect to Configuration Container: '$searchRootQuery'."

		$searchRoot = New-Object System.DirectoryServices.DirectoryEntry -ArgumentList $searchRootQuery
	}

	$ds = [adsisearcher]"(NETBIOSName=*)"
	$ds.searchroot = $searchRoot
	$ds.PageSize = 1000 
	$ds.SearchScope = "OneLevel"
	$props= "ncname", "netbiosName", "objectguid"
	$ds.PropertiesToLoad.AddRange($props)

	$directoryPartitions = $ds.FindAll()

	foreach ($directoryPartition in $directoryPartitions)
	{
		$props = $directoryPartition.Properties

		# Property names case sensitive, beware!! Must be spelt as defined in $props variable
		if (![string]::IsNullOrEmpty($props.netbiosname))
		{
			$objectGuid = $props.objectguid | foreach { "{0:X2}" -f $_}
			$identifier = [Guid] $objectGuid.Replace(" ", "")
			$dn = $props.ncname
			$name = $props.ncname
			$partition = [Microsoft.MetadirectoryServices.Partition]::Create($identifier, $dn, $name)
			$partition.HiddenByDefault = $false

			[void] $partitions.Add($partition)
		}
	}

	return ,$partitions # Prevent unwinding
}

$preferredDomainController = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "PreferredDomainControllerFQDN"

if (![string]::IsNullOrEmpty($preferredDomainController))
{
	$preferredDomainController = Select-PreferredDomainController -DomainControllerList $preferredDomainController
}

if ([string]::IsNullOrEmpty($preferredDomainController))
{
	$preferredDomainController = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "Domain" # Do this only in the partition script. This is not needed when object DN is available.
}

Get-Partitions

Exit-Script -ScriptType "Partition" -ErrorObject $Error
