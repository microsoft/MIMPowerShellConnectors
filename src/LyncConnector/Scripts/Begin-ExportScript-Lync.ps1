<#
<copyright file="Begin-ExportScript-Lync.ps1" company="Microsoft">
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
</copyright>
<summary>
	The Begin-Export script for the Skype 2015 / Lync 2010 / 2013 Connector.
	Opens the RPS session and imports a set of Lync cmdlets into it.
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

Write-Debug "$Global:ConnectorName - Begin-Export Script: Execution Started..."

$server = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "Server"
$preferredDomainController = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "PreferredDomainControllerFQDN"

if (![string]::IsNullOrEmpty($preferredDomainController))
{
	$preferredDomainController = Select-PreferredDomainController -DomainControllerList $preferredDomainController
}

$Global:PreferredDomainController = $preferredDomainController

$session = Get-PSSession -Name $Global:RemoteSessionName -ErrorAction "SilentlyContinue"

if (!$session)
{
	Write-Debug "Opening a new RPS Session."

	$skipCertificate = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
	$session = New-PSSession -ConnectionUri $server -Credential $Credential -SessionOption $skipCertificate -Name $Global:RemoteSessionName
	$Global:Session = $session
	$lyncCommands = "Get-CsUser", "Get-CsAdUser", "Enable-CsUser", "Disable-CsUser", "Set-CsUser", "Grant-CsArchivingPolicy", "Grant-CsClientPolicy", `
		"Grant-CsClientVersionPolicy", "Grant-CsConferencingPolicy", "Grant-CsDialPlan", "Grant-CsExternalAccessPolicy", "Grant-CsHostedVoicemailPolicy", `
		"Grant-CsLocationPolicy", "Grant-CsPinPolicy", "Grant-CsPresencePolicy", "Grant-CsVoicePolicy", "Move-CsUser"

	Import-PSSession $Global:Session -CommandName $lyncCommands | Out-Null

	Write-Debug "Opened a new RPS Session."
}

Write-Debug "$Global:ConnectorName - Begin-Export Script: Execution Completed."
