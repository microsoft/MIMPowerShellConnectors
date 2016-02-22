<#
<copyright file="ExportScript-Lync.ps1" company="Microsoft">
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
</copyright>
<summary>
	The Main Export script for the Skype 2015 / Lync 2010 / 2013 Connector.
</summary>
#>

[CmdletBinding()]
param(
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
	[System.Collections.Generic.IList[Microsoft.MetadirectoryServices.CSEntryChange]]
	$CSEntries,
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

Enter-Script -ScriptType "Export" -ErrorObject $Error

function Export-CSEntries
{
	<#
	.Synopsis
		Exports the CSEntry changes.
	.Description
		Exports the CSEntry changes.
	#>
	
	[CmdletBinding()]
	[OutputType([System.Collections.Generic.List[Microsoft.MetadirectoryServices.CSEntryChangeResult]])]
	param(
	)

	$csentryChangeResults = New-GenericObject System.Collections.Generic.List Microsoft.MetadirectoryServices.CSEntryChangeResult

	foreach ($csentryChange in $CSEntries)
	{
		$Error.Clear() = $null
		$newAnchorTable = @{}

		$dn = Get-CSEntryChangeDN $csentryChange
		$objectType = $csentryChange.ObjectType
		$objectModificationType = $csentryChange.ObjectModificationType

		Write-Debug "Exporting $objectModificationType to $objectType : $dn"

		try
		{
			switch ($objectType)
			{
				"User"
				{
					$newAnchorTable = Export-User $csentryChange
					break
				}

				"OrganizationalUnit"
				{
					$newAnchorTable = Export-OrganizationalUnit $csentryChange
					break
				}

				default
				{
					throw "Unknown CSEntry ObjectType: $_"
				}
			}
		}
		catch
		{
			Write-Error "$_"
		}

		if ($Error)
		{
			$csentryChangeResult = New-CSEntryChangeExportError -CSEntryChangeIdentifier $csentryChange.Identifier -ErrorObject $Error
		}
		else
		{
			$exportAdd = $objectModificationType -eq "Add"

			$csentryChangeResult = New-CSEntryChangeResult -CSEntryChangeIdentifier $csentryChange.Identifier -NewAnchorTable $newAnchorTable -ExportAdd:$exportAdd

			Write-Debug "Exported $objectModificationType to $objectType : $dn."
		}

		$csentryChangeResults.Add($csentryChangeResult)
	}

	##$exportEntriesResults = New-Object -TypeName "Microsoft.MetadirectoryServices.PutExportEntriesResults" -ArgumentList $csentryChangeResults
	$closedType = [Type] "Microsoft.MetadirectoryServices.PutExportEntriesResults"
	return [Activator]::CreateInstance($closedType, $csentryChangeResults)
}

function Export-User
{
	<#
	.Synopsis
		Exports the changes for User objects.
	.Description
		Exports the changes for User objects.
		Returns the Hashtable of anchor attribute values for Export-Add csentries.
	#>
	
	[CmdletBinding()]
	[OutputType([Hashtable])]
	param(
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange
	)

	$newAnchorTable = @{}

	switch ($CSEntryChange.ObjectModificationType)
	{
		"Add"
		{
			$dn  = Get-CSEntryChangeDN $CSEntryChange
			$identity = Get-CsIdentity $CSEntryChange

			$cmd = "Get-CsAdUser -Identity '$identity'"
			if (![string]::IsNullOrEmpty($preferredDomainController))
			{
				$cmd += " -DomainController '$preferredDomainController'"
			}

			Write-Debug "Invoking $cmd for user: $dn"

			$x = Invoke-Expression $cmd 

			if (!$Error)
			{
				$newAnchorTable.Add("Guid", $x.Guid.ToByteArray())
			}

			Invoke-EnableCsUserCommand $CSEntryChange 
			Invoke-SetCsUserCommand $CSEntryChange
			Invoke-GrantCsPolicyCommands $CSEntryChange

			break
		}

		"Replace"
		{
			Invoke-MoveCsUserCommand $CSEntryChange
			Invoke-SetCsUserCommand $CSEntryChange
			Invoke-GrantCsPolicyCommands $CSEntryChange

			break
		}

		"Update"
		{
			Invoke-MoveCsUserCommand $CSEntryChange
			Invoke-SetCsUserCommand $CSEntryChange
			Invoke-GrantCsPolicyCommands $CSEntryChange

			break
		}

		"Delete"
		{
			Invoke-DisableCsUserCommand $CSEntryChange

			break
		}

		default
		{
			throw "Unknown CSEntry ObjectModificationType: $_"
		}
	}

	return $newAnchorTable
}

function Export-OrganizationalUnit
{
	<#
	.Synopsis
		Exports the changes for OrganizationalUnit objects.
	.Description
		Exports the changes for OrganizationalUnit objects.
		Returns the Hashtable of anchor attribute values for Export-Add csentries.
	#>
	
	[CmdletBinding()]
	[OutputType([Hashtable])]
	param(
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange
	)

	$newAnchorTable = @{}

	switch ($CSEntryChange.ObjectModificationType)
	{
		"Add"
		{
			$dn  = Get-CSEntryChangeDN $CSEntryChange

			if ([string]::IsNullOrEmpty($preferredDomainController))
			{
				$adsPath = "LDAP://{0}" -f $dn
			}
			else
			{
				$adsPath = "LDAP://{0}/{1}" -f $preferredDomainController, $dn
			}

			Write-Debug ("Export-OrganizationalUnit ADS Path: {0}" -f $adsPath)

			$userName = "{0}\{1}" -f $Credential.GetNetworkCredential().Domain, $Credential.GetNetworkCredential().UserName
			$password = $Credential.GetNetworkCredential().Password
			$directoryEntry = New-Object System.DirectoryServices.DirectoryEntry -ArgumentList $adsPath, $userName, $password
			
			$objectGuid = $directoryEntry.ObjectGUID | foreach { "{0:X2}" -f $_}
			$objectGuid = [Guid] $objectGuid.Replace(" ", "")

			$newAnchorTable.Add("Guid", $objectGuid.ToByteArray())

			break
		}

		"Replace"
		{
			break
		}

		"Update"
		{
			break
		}

		"Delete"
		{
			break
		}

		default
		{
			Write-Error "Unknown CSEntry ObjectModificationType: $_"
		}
	}

	return $newAnchorTable
}

function Invoke-EnableCsUserCommand
{
	<#
	.Synopsis
		Invokes Enable-CsUser cmdlet on the specified user csentry.
	.Description
		Invokes Enable-CsUser cmdlet on the specified user csentry.
	#>
	
	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange
	)

	if (!$Error)
	{
		$dn  = Get-CSEntryChangeDN $CSEntryChange
		$identity = Get-CsIdentity $CSEntryChange
		$registrarPool = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "RegistrarPool"
		$cmd = "Enable-CsUser -Identity '$identity' -RegistrarPool '$registrarPool'"
		if (![string]::IsNullOrEmpty($preferredDomainController))
		{
			$cmd += " -DomainController '$preferredDomainController'"
		}

		$sipAddress = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "SipAddress"

		if (![string]::IsNullOrEmpty($sipAddress))
		{
			$cmd += " -SipAddress '$sipAddress'"
		}
		elseif ($sipAddressType -eq "FirstLastName" -or $sipAddressType -eq "SamAccountName")
		{
			$cmd +=  " -SipAddressType '$sipAddressType' -SipDomain '$sipDomain'"
		}
		else
		{
			$cmd +=  " -SipAddressType '$sipAddressType'"
		}

		Write-Debug "Invoking $cmd for user: $dn"

		Invoke-Expression $cmd | Out-Null
	}
}

function Invoke-SetCsUserCommand
{
	<#
	.Synopsis
		Invokes Set-CsUser cmdlet on the specified user csentry.
	.Description
		Invokes Set-CsUser cmdlet on the specified user csentry.
	#>
	
	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange
	)

	if (!$Error)
	{
		$dn  = Get-CSEntryChangeDN $CSEntryChange
		$identity = Get-CsIdentity $CSEntryChange
		$audioVideoDisabled = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "AudioVideoDisabled"
		$enabled = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "Enabled"
		$enterpriseVoiceEnabled = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "EnterpriseVoiceEnabled"
		$hostedVoiceMail = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "HostedVoiceMail"
		$lineURI = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "LineURI"
		$lineServerURI = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "LineServerURI"
		$privateLine = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "PrivateLine"
		$remoteCallControlTelephonyEnabled = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "RemoteCallControlTelephonyEnabled"
		$sipAddress = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "SipAddress"

		$cmd = "Set-CsUser -Identity '$identity'"
		if (![string]::IsNullOrEmpty($preferredDomainController))
		{
			$cmd += " -DomainController '$preferredDomainController'"
		}

		if ($audioVideoDisabled -ne $null) { $cmd += " -AudioVideoDisabled `$$audioVideoDisabled" }
		if ($enabled -ne $null) { $cmd += " -Enabled `$$enabled" }
		if ($enterpriseVoiceEnabled -ne $null) { $cmd += " -EnterpriseVoiceEnabled `$$enterpriseVoiceEnabled" }
		if ($hostedVoiceMail -ne $null) { $cmd += " -HostedVoiceMail `$$hostedVoiceMail" }
		if ($lineURI -ne $null) { $cmd += " -LineURI '$lineURI'" }
		if ($lineServerURI -ne $null) { $cmd += " -LineServerURI '$lineServerURI'" }
		if ($privateLine -ne $null) { $cmd += " -PrivateLine '$privateLine'" }
		if ($remoteCallControlTelephonyEnabled -ne $null) { $cmd += " -RemoteCallControlTelephonyEnabled `$$remoteCallControlTelephonyEnabled" }
		if ($sipAddress -ne $null) { $cmd += " -SipAddress '$sipAddress'" }

		Write-Debug "Invoking $cmd for user: $dn"

		Invoke-Expression $cmd | Out-Null
	}
}

function Invoke-GrantCsPolicyCommands
{
	<#
	.Synopsis
		Invokes Grant-CsPolicy cmdlets on the specified user csentry.
	.Description
		Invokes Grant-CsPolicy cmdlets on the specified user csentry.
	#>
	
	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange
	)


	if (!$Error)
	{
		$dn  = Get-CSEntryChangeDN $CSEntryChange
		$identity = Get-CsIdentity $CSEntryChange
		$archivingPolicy  = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "ArchivingPolicy"
		$archivingPolicyChanged  = Test-CSEntryChangeValueChanged -CSEntryChange $CSEntryChange -AttributeName "ArchivingPolicy"
		$clientPolicy = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "ClientPolicy"
		$clientPolicyChanged = Test-CSEntryChangeValueChanged -CSEntryChange $CSEntryChange -AttributeName "ClientPolicy"
		$clientVersionPolicy = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "ClientVersionPolicy"
		$clientVersionPolicyChanged = Test-CSEntryChangeValueChanged -CSEntryChange $CSEntryChange -AttributeName "ClientVersionPolicy"
		$conferencingPolicy = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "ConferencingPolicy"
		$conferencingPolicyChanged = Test-CSEntryChangeValueChanged -CSEntryChange $CSEntryChange -AttributeName "ConferencingPolicy"
		$dialPlan = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "DialPlan"
		$dialPlanChanged = Test-CSEntryChangeValueChanged -CSEntryChange $CSEntryChange -AttributeName "DialPlan"
		$externalAccessPolicy = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "ExternalAccessPolicy"
		$externalAccessPolicyChanged = Test-CSEntryChangeValueChanged -CSEntryChange $CSEntryChange -AttributeName "ExternalAccessPolicy"
		$hostedVoicemailPolicy = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "HostedVoicemailPolicy"
		$hostedVoicemailPolicyChanged = Test-CSEntryChangeValueChanged -CSEntryChange $CSEntryChange -AttributeName "HostedVoicemailPolicy"
		$locationPolicy = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "LocationPolicy"
		$locationPolicyChanged = Test-CSEntryChangeValueChanged -CSEntryChange $CSEntryChange -AttributeName "LocationPolicy"
		$pinPolicy = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "PinPolicy"
		$pinPolicyChanged = Test-CSEntryChangeValueChanged -CSEntryChange $CSEntryChange -AttributeName "PinPolicy"
		$presencePolicy = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "PresencePolicy"
		$presencePolicyChanged = Test-CSEntryChangeValueChanged -CSEntryChange $CSEntryChange -AttributeName "PresencePolicy"
		$voicePolicy = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "VoicePolicy"
		$voicePolicyChanged = Test-CSEntryChangeValueChanged -CSEntryChange $CSEntryChange -AttributeName "VoicePolicy"

		$cmd = "Get-CsUser -Identity '$identity'"
		if (![string]::IsNullOrEmpty($preferredDomainController))
		{
			$cmd += " -DomainController '$preferredDomainController'"
		}

		if ($archivingPolicyChanged) { $cmd += " | Grant-CsArchivingPolicy -PolicyName '$archivingPolicy' -PassThru" }
		if ($clientPolicyChanged) { $cmd += " | Grant-CsClientPolicy -PolicyName '$clientPolicy' -PassThru" }
		if ($clientVersionPolicyChanged) { $cmd += " | Grant-CsClientVersionPolicy -PolicyName '$clientVersionPolicy' -PassThru" }
		if ($conferencingPolicyChanged) { $cmd += " | Grant-CsConferencingPolicy -PolicyName '$conferencingPolicy' -PassThru" }
		if ($dialPlanChanged) { $cmd += " | Grant-CsDialPlan -PolicyName '$dialPlan' -PassThru" }
		if ($externalAccessPolicyChanged) { $cmd += " | Grant-CsExternalAccessPolicy -PolicyName '$externalAccessPolicy' -PassThru" }
		if ($hostedVoicemailPolicyChanged) { $cmd += " | Grant-CsHostedVoicemailPolicy -PolicyName '$hostedVoicemailPolicy' -PassThru" }
		if ($locationPolicyChanged) { $cmd += " | Grant-CsLocationPolicy -PolicyName '$locationPolicy' -PassThru" }
		if ($pinPolicyChanged) { $cmd += " | Grant-CsPinPolicy -PolicyName '$pinPolicy' -PassThru" }
		if ($presencePolicyChanged) { $cmd += " | Grant-CsPresencePolicy -PolicyName '$presencePolicy' -PassThru" }
		if ($voicePolicyChanged) { $cmd += " | Grant-CsVoicePolicy -PolicyName '$voicePolicy' -PassThru" }

		Write-Debug "Invoking $cmd for user: $dn"

		Invoke-Expression $cmd | Out-Null
	}
}

function Invoke-MoveCsUserCommand
{
	<#
	.Synopsis
		Invokes Move-CsUser cmdlet on the specified user csentry.
	.Description
		Invokes Move-CsUser cmdlet on the specified user csentry.
	#>
	
	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange
	)

	if (!$Error)
	{
		$dn  = Get-CSEntryChangeDN $CSEntryChange
		$identity = Get-CsIdentity $CSEntryChange
		$registrarPool  = Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "RegistrarPool"
		if ($registrarPool)
		{
			$cmd = "Move-CsUser -Identity '$identity' -Target $registrarPool -Force:`$$forceMove -Confirm:`$$false'"
			if (![string]::IsNullOrEmpty($preferredDomainController))
			{
				$cmd += " -DomainController '$preferredDomainController'"
			}

			Write-Debug "Invoking $cmd for user: $dn"

			Invoke-Expression $cmd | Out-Null
		}
	}
}

function Invoke-DisableCsUserCommand
{
	<#
	.Synopsis
		Invokes Disable-CsUser cmdlet on the specified user csentry.
	.Description
		Invokes Disable-CsUser cmdlet on the specified user csentry.
	#>
	
	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange
	)

	if (!$Error)
	{
		$dn  = Get-CSEntryChangeDN $CSEntryChange
		$identity = Get-CsIdentity $CSEntryChange

		$cmd = "Disable-CsUser -Identity '$identity'"
		if (![string]::IsNullOrEmpty($preferredDomainController))
		{
			$cmd += " -DomainController '$preferredDomainController'"
		}

		Write-Debug "Invoking $cmd for user: $dn"

		Invoke-Expression $cmd | Out-Null
	}
}

function Get-CsIdentity
{
	<#
	.Synopsis
		Gets the identifier for specified user csentry.
	.Description
		Gets the identifier for specified user csentry.
		It is the Guid if the Anchor is populated. Otherwise DN.
	#>
	
	[CmdletBinding()]
	[OutputType([string])]
	param(
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange
	)

	if (!$Error)
	{
		$dn  = Get-CSEntryChangeDN $CSEntryChange

		if ($CSEntryChange.AnchorAttributes.Contains("Guid") -and $CSEntryChange.AnchorAttributes["Guid"].Value -ne $null)
		{
			return ([Guid]$CSEntryChange.AnchorAttributes["Guid"].Value).ToString()
		}
		else # should only be here when ObjectModificationType = "Add"
		{
			return $dn
		}
	}
}

$sipAddressType = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "SipAddressType"
$sipDomain = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "SipDomain"

if ($sipAddressType -eq "FirstLastName" -or $sipAddressType -eq "SamAccountName")
{
	if ([string]::IsNullOrEmpty($sipDomain))
	{
		throw "MA configuration error. When SipAddressType = $sipAddressType, SipDomain must be configured as well."
	}
}

$preferredDomainController = $Global:PreferredDomainController

$forceMove = (Get-ConfigParameter -ConfigParameters $configParameters -ParameterName "ForceMove") -eq "Yes"

Export-CSEntries

Exit-Script -ScriptType "Export" -SuppressErrorCheck -ErrorObject $Error

