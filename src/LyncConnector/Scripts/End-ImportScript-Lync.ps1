<#
<copyright file="End-ImportScript-Lync.ps1" company="Microsoft">
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
</copyright>
<summary>
	The End-Import script for the Skype 2015 / Lync 2010 / 2013 Connector.
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
	[Microsoft.MetadirectoryServices.OpenImportConnectionRunStep]
	$OpenImportConnectionRunStep,
	[parameter(Mandatory = $true)]
	[Microsoft.MetadirectoryServices.CloseImportConnectionRunStep]
	$CloseImportConnectionRunStep,
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

$commonModule = (Join-Path -Path $ScriptDir -ChildPath $ConfigParameters["Common Module Script Name (with extension)"].Value)

if (!(Get-Module -Name (Get-Item $commonModule).BaseName)) { Import-Module -Name $commonModule }

Write-Debug "$Global:ConnectorName - End-Import Script: Execution Started..."

function Get-CloseImportConnectionResults
{
	<#
	.Synopsis
		Gets the CloseImportConnectionResults.
	.Description
		Gets the CloseImportConnectionResults.
	#>
	
	[CmdletBinding()]
	[OutputType([Microsoft.MetadirectoryServices.CloseImportConnectionResults])]
	param (
	)
	
	$waterMark = [xml]$CloseImportConnectionRunStep.CustomData

	if ($waterMark -ne $null -and $waterMark.WaterMark -ne $null)
	{
		$waterMark.WaterMark.LastRunDateTime = [DateTime]::UtcNow.ToString("yyyy-MM-dd HH:mm:ss");
	}

	Write-Debug ("Watermark finalized to: {0}" -f $waterMark.InnerXml)

	$results = New-Object Microsoft.MetadirectoryServices.CloseImportConnectionResults($watermark.InnerXml)

	return $results
}

Get-CloseImportConnectionResults

if (Test-Variable -Name "Session" -Scope "Global")
{
	Remove-PSSession $Global:Session
}

Write-Debug "$Global:ConnectorName - End-Import Script: Execution Completed."
