<#
<copyright file="HierarchyScript-Lync.ps1" company="Microsoft">
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
</copyright>
<summary>
	The Hierarchy script for the Skype 2015 / Lync 2010 / 2013 Connector.
</summary>
#>

[CmdletBinding()]
param(
	[parameter(Mandatory = $true)]
	[Microsoft.MetadirectoryServices.HierarchyNode]
	$HierarchyNode,
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

Enter-Script -ScriptType "Hierarchy" -ErrorObject $Error

function Get-Hierarchy
{
	<#
	.Synopsis
		Gets the OU hierarchy under the selected node.
	.Description
		Gets the OU hierarchy under the selected node.
	#>

	[CmdletBinding()]
    [OutputType([System.Collections.Generic.List[Microsoft.MetadirectoryServices.HierarchyNode]])]
	param(
	)

	$children = New-GenericObject System.Collections.Generic.List Microsoft.MetadirectoryServices.HierarchyNode

	$rootDN = $HierarchyNode.DN

	if ([string]::IsNullOrEmpty($preferredDomainController))
	{
		$searchPath = "LDAP://{0}" -f $rootDN
	}
	else
	{
		$searchPath = "LDAP://{0}/{1}" -f $preferredDomainController, $rootDN
	}

	Write-Debug "Enumerating Inclusion OrganizationalUnit Hierarchy $searchPath"

	$userName = "{0}\{1}" -f $Credential.GetNetworkCredential().Domain, $Credential.GetNetworkCredential().UserName
	$password = $Credential.GetNetworkCredential().Password
	$searchRoot = New-Object System.DirectoryServices.DirectoryEntry -ArgumentList $searchPath, $userName, $password

	if ([string]::IsNullOrEmpty($password)) # Check for a bug in ECMA 2.0 on refreshing partion, apparently fixed now.
	{
		# Try with SyncService credentials - work around for the bug in ECMA 2.0, apparently not needed now.
		Write-Debug "Hitting ECMA 2.0 bug. Trying with SyncSerice credentials to connect to OrganizationalUnit Hierarchy: '$searchPath'."

		$searchRoot = New-Object System.DirectoryServices.DirectoryEntry -ArgumentList $searchPath
	}

	$ds = [adsisearcher]"(|(objectClass=organizationalUnit)(objectClass=Container))"
	$ds.searchroot = $searchRoot
	$ds.PageSize = 1000 
	$ds.SearchScope = "OneLevel"
	$props= "name", "distinguishedname"
	$ds.PropertiesToLoad.AddRange($props)

	$results = $ds.FindAll()

	foreach ($result in $results)
	{
		$props = $result.Properties

		# Property names case sensitive, beware!! Must be spelt as defined in $props variable
		$dn = $props.distinguishedname 
		$name = $props.name
		[void] $children.Add([Microsoft.MetadirectoryServices.HierarchyNode]::Create($dn, $name))
	}

	return ,$children # Prevent unwinding
}

$preferredDomainController = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "PreferredDomainControllerFQDN"

if (![string]::IsNullOrEmpty($preferredDomainController))
{
	$preferredDomainController = Select-PreferredDomainController -DomainControllerList $preferredDomainController
}

Get-Hierarchy

Exit-Script -ScriptType "Hierarchy" -ErrorObject $Error
