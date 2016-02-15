<#
<copyright file="SchemaScript-Lync.ps1" company="Microsoft">
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
</copyright>
<summary>
	The Schema script for the Skype 2015 / Lync 2010 / 2013 Connector.
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

Enter-Script -ScriptType "Schema"

function Get-ConnectorSchema
{
	<#
	.Synopsis
		Gets the connector space schema.
	.Description
		Gets the connector space schema defined in the "Schema-Lync.xml" file.
	#>

	[CmdletBinding()]
    [OutputType([Microsoft.MetadirectoryServices.Schema])]
	param(
	)

	$extensionsDir = Get-ExtensionsDirectory
	$schemaXml = Join-Path -Path $extensionsDir -ChildPath "Schema-Lync.xml"

	$schema = ConvertFrom-SchemaXml -SchemaXml $schemaXml

	return $schema
}

Get-ConnectorSchema

Exit-Script -ScriptType "Schema"
