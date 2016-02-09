<#
<copyright file="End-ExportScript-Lync.ps1" company="Microsoft">
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
</copyright>
<summary>
	The End-Export script for the Skype 2015 / Lync 2010 / 2013 Connector.
	Closes the RPS session.
</summary>
#>

[CmdletBinding()]
param (
	[parameter(Mandatory = $true)]
	[System.Collections.ObjectModel.KeyedCollection[string,Microsoft.MetadirectoryServices.ConfigParameter]]
	$ConfigParameters,
	[parameter(Mandatory = $true)]
	[Microsoft.MetadirectoryServices.Schema]
	$Schema,
	[parameter(Mandatory = $true)]
	[Microsoft.MetadirectoryServices.OpenExportConnectionRunStep]
	$OpenExportConnectionRunStep,
	[parameter(Mandatory = $true)]
	[Microsoft.MetadirectoryServices.CloseExportConnectionRunStep]
	$CloseExportConnectionRunStep,
	[parameter(Mandatory = $true)]
	[Alias('PSCredential')] # To fix mess-up of the parameter name in the RTM version of the PowerShell connector.
	[System.Management.Automation.PSCredential]
	$Credential,
	[parameter(Mandatory = $false)]
	[ValidateScript({ Test-Path $_ -PathType "Container" })]
	[string]
	$ScriptDir = [Microsoft.MetadirectoryServices.MAUtils]::MAFolder # Optional parameter for manipulation by the TestHarness script.
)

Set-StrictMode -Version "2.0"

$commonModule = (Join-Path -Path $scriptDir -ChildPath $configParameters["Common Module Script Name (with extension)"].Value)

if (!(Get-Module -Name (Get-Item $commonModule).BaseName)) { Import-Module -Name $commonModule }

Enter-Script -ScriptType "End-Export"

if (Test-Variable -Name "Session" -Scope "Global")
{
	Remove-PSSession $Global:Session
}

Exit-Script -ScriptType "End-Export"
