<#
<copyright file="ValidationScript-Lync.ps1" company="Microsoft">
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
</copyright>
<summary>
	The Validation script for the Skype 2015 / Lync 2010 / 2013 Connector.
</summary>
#>

[CmdletBinding()]
param(
	[parameter(Mandatory = $true)]
	[System.Collections.ObjectModel.KeyedCollection[string,Microsoft.MetadirectoryServices.ConfigParameter]]
	$ConfigParameters,
	[parameter(Mandatory = $true)]
	[Microsoft.MetadirectoryServices.ConfigParameterPage]
	$ConfigParameterPage,
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

Enter-Script -ScriptType "Validation"

function Test-ConfigParameterPage
{
	<#
	.Synopsis
		Validates the input entries on the current Config Parameter Page.
	.Description
		Validates the input entries on the current Config Parameter Page.
	#>

	[CmdletBinding()]
    [OutputType([Microsoft.MetadirectoryServices.ParameterValidationResult])]
	param(
	)
	
	$validationResult = New-Object -TypeName "Microsoft.MetadirectoryServices.ParameterValidationResult" -ArgumentList "Success", $null, $null

	Write-Debug "Validating Settings: $ConfigParameterPage"

	switch ($ConfigParameterPage)
	{
		"Connectivity"
		{
			$validationResult = Test-ConnectivityConfigParameterPage $_
			break
		}

		"Global"
		{
			$validationResult = Test-GlobalConfigParameterPage
			break
		}

		"Partition"
		{
			$validationResult = Test-ConnectivityConfigParameterPage $_
			break
		}

		"RunStep"
		{
			break
		}
	}

	return $validationResult
}

function Test-ConnectivityConfigParameterPage
{
	<#
	.Synopsis
		Validates the input entries on the Connectivity Config Parameter Page.
	.Description
		Validates the input entries on the Connectivity Config Parameter Page.
	#>

	[CmdletBinding()]
    [OutputType([Microsoft.MetadirectoryServices.ParameterValidationResult])]
	param(
	[parameter(Mandatory = $false)]
	[ValidateSet("Partition", "Connectivity", "")]
	[string]
	$Scope
	)

	$validationResult = New-Object -TypeName "Microsoft.MetadirectoryServices.ParameterValidationResult" -ArgumentList "Success", $null, $null

	$Scope = $null # this will cause Get-ConfigParameter to return values configured in the "upper" scope if not defined on the current scope.
	$impersonateConnectorAccount = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "Impersonate Connector Account" -Scope $Scope
	$server = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "Server" -Scope $Scope
	$domain = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "Domain" -Scope $Scope
	$user = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "User" -Scope $Scope
	$password = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "Password" -Scope $Scope -Encrypted

	Write-Debug ("Validating $Scope settings. ImpersonateConnectorAccount: '{0}'. Server: '{1}'. Domain: '{2}'. User: '{3}'." -f $impersonateConnectorAccount, $server, $domain, $user)

	if ($impersonateConnectorAccount -eq "1")
	{
		$statusCode = "Failure"
		$errorParameter = "Impersonate Connector Account"
		$errorMessage = "Please uncheck {0} checkbox." -f $errorParameter

		$validationResult = New-Object -TypeName "Microsoft.MetadirectoryServices.ParameterValidationResult" -ArgumentList $statusCode, $errorMessage, $errorParameter
	}

	if ([string]::IsNullOrEmpty($server))
	{
		$statusCode = "Failure"
		$errorParameter = "Server"
		$errorMessage = "Please specify a value for {0} field." -f $errorParameter

		$validationResult = New-Object -TypeName "Microsoft.MetadirectoryServices.ParameterValidationResult" -ArgumentList $statusCode, $errorMessage, $errorParameter
	}
	elseif ([string]::IsNullOrEmpty($user))
	{
		$statusCode = "Failure"
		$errorParameter = "User"
		$errorMessage = "Please specify a value for {0} field." -f $errorParameter

		$validationResult = New-Object -TypeName "Microsoft.MetadirectoryServices.ParameterValidationResult" -ArgumentList $statusCode, $errorMessage, $errorParameter
	}
	elseif ([string]::IsNullOrEmpty($password))
	{
		$statusCode = "Failure"
		$errorParameter = "Password"
		$errorMessage = "Please specify a value for {0} field." -f $errorParameter

		$validationResult = New-Object -TypeName "Microsoft.MetadirectoryServices.ParameterValidationResult" -ArgumentList $statusCode, $errorMessage, $errorParameter
	}
	elseif ([string]::IsNullOrEmpty($domain))
	{
		$statusCode = "Failure"
		$errorParameter = "Domain"
		$errorMessage = "Please specify a value for {0} field." -f $errorParameter

		$validationResult = New-Object -TypeName "Microsoft.MetadirectoryServices.ParameterValidationResult" -ArgumentList $statusCode, $errorMessage, $errorParameter
	}
	else
	{
		if (![string]::IsNullOrEmpty($domain))
		{
			$user = "$domain\$user"
		}

		$Credential = New-Object System.Management.Automation.PSCredential($user, $password)

		$statusCode = "Failure"
		$errorParameter = "Server"
		$errorMessage = $null

		$skipCertificate = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
		$session = New-PSSession -ConnectionUri $server -Credential $Credential -SessionOption $skipCertificate -Name $Global:RemoteSessionName -ErrorVariable errorMessage -ErrorAction SilentlyContinue

		if ($errorMessage.Count -gt 0)
		{
			$errorMessage = [string]$errorMessage[0] + "`r`nConnector User: $user"

			$validationResult = New-Object -TypeName "Microsoft.MetadirectoryServices.ParameterValidationResult" -ArgumentList $statusCode, $errorMessage, $errorParameter

			$Error.Clear()
		}
		else
		{
			Remove-PSSession $session
		}
	}

	return $validationResult
}


function Test-GlobalConfigParameterPage
{
	<#
	.Synopsis
		Validates the input entries on the Global Config Parameter Page.
	.Description
		Validates the input entries on the Global Config Parameter Page.
	#>

	[CmdletBinding()]
    [OutputType([Microsoft.MetadirectoryServices.ParameterValidationResult])]
	param(
	)

	$validationResult = New-Object -TypeName "Microsoft.MetadirectoryServices.ParameterValidationResult" -ArgumentList "Success", $null, $null
	
	$scope = "Global"
	$sipAddressType = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "SipAddressType" -Scope $scope
	$sipDomain = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "SipDomain" -Scope $scope

	Write-Debug ("Validating Global settings. SipAddressType: '{0}'. SipDomain: '{1}'." -f $sipAddressType, $sipDomain)

	if ($sipAddressType -ne "UserPrincipalName" -and $sipAddressType -ne "EmailAddress" -and $sipAddressType -ne "FirstLastName" -and $sipAddressType -ne "SamAccountName")
	{
		$statusCode = "Failure"
		$errorParameter = "SipAddressType_Global"
		$errorMessage = "SipAddressType must be one of the values: UserPrincipalName, EmailAddress, FirstLastName, SamAccountName"

		$validationResult = New-Object -TypeName "Microsoft.MetadirectoryServices.ParameterValidationResult" -ArgumentList $statusCode, $errorMessage, $errorParameter
	}
	elseif ($sipAddressType -eq "FirstLastName" -or $sipAddressType -eq "SamAccountName")
	{
		if ([string]::IsNullOrEmpty($sipDomain))
		{
			$statusCode = "Failure"
			$errorParameter = "SipDomain_Global"
			$errorMessage = "When SipAddressType = $sipAddressType, SipDomain must be configured as well."

			$validationResult = New-Object -TypeName "Microsoft.MetadirectoryServices.ParameterValidationResult" -ArgumentList $statusCode, $errorMessage, $errorParameter
		}
	}

	return $validationResult
}

Test-ConfigParameterPage

Exit-Script -ScriptType "Validation"
