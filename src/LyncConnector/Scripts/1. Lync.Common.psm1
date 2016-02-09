<#
<copyright file="Lync.Common.psm1" company="Microsoft">
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
</copyright>
<summary>
	Common Utility Functions shared by all Skype 2015 / Lync 2010 / 2013 Connector scripts.
</summary>
#>

Set-StrictMode -Version "2.0"

#region "Global Variables"

$Global:ConnectorName = "LyncPowerShellConnector"
$Global:RemoteSessionName = "LyncPowerShellConnector"
$Error.Clear()

#endregion "Global Variables"

#region "Import Dependent Modules"

# None

#endregion "Import Dependent Modules"

function Enter-Script
{
	<#
	.Synopsis
		Writes the Versbose message saying specified script execution started.
	.Description
		Writes the Versbose message saying specified script execution started.
		Also clear the $Error variable.
	#>
	[CmdletBinding()]
	param (
		[parameter(Mandatory = $true)]
		[string]
		$ScriptType
	)

	Write-Verbose "$Global:ConnectorName - $ScriptType Script: Execution Started..."
	$Error.Clear()
}

function Exit-Script
{
	<#
	.Synopsis
		Checks $Error variable for any Errors. Writes the Versbose message saying specified script execution sucessfully completed.
	.Description
		Checks $Error variable for any Errors. Writes the Versbose message saying specified script execution sucessfully completed.
		Throws an exception if $Error is present
	#>
	[CmdletBinding()]
	param (
		[parameter(Mandatory = $true)]
		[string]
		$ScriptType,
		[parameter(Mandatory = $false)]
		[switch]
		$SuppressErrorCheck,
		[parameter(Mandatory = $false)]
		[Type]
		$ExceptionRaisedOnErrorCheck
	)

	if ($Error.Count -ne 0 -and !$SuppressErrorCheck)
	{
		$errorMessage = [string]$Error[0]

		if ($ExceptionRaisedOnErrorCheck -eq $null)
		{
			$ExceptionRaisedOnErrorCheck = [Microsoft.MetadirectoryServices.ExtensibleExtensionException]
		}

		throw  $errorMessage -as $ExceptionRaisedOnErrorCheck
	}

	Write-Verbose "$Global:ConnectorName - $ScriptType Script: Execution Completed."
}

function Get-ExtensionsDirectory
{
	<#
	.Synopsis
		Gets the path of the "Extensions" folder.
	.Description
		Gets the path of the "Extensions" folder on a FIM Sync server.
		If FIM Sync is not installed on the DEV machine, it returns the present working directory.
	#>
	[CmdletBinding()]
	[OutputType([string])]
	param (
	)

	process
	{
		$scriptDir = $PWD

		$syncDir = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\FIMSynchronizationService\Parameters" -Name Path -ErrorAction SilentlyContinue

		if ($syncDir)
		{
			$scriptDir = Join-Path -Path $syncDir.Path -ChildPath "Extensions"
		}

		return $scriptDir
	}
}

function New-GenericObject
{
	<#
	.Synopsis
		Create a new generic object.
	.Description
		Create a new generic object.
	.Example
		New-GenericObject -TypeName System.Collections.Generic.List  -TypeParameters Microsoft.MetadirectoryServices.CSEntryChange
	#>
	
	[CmdletBinding()]
	[OutputType([object])]
	param (
		[parameter(Mandatory = $true)]
		[string]
		$TypeName,
		[parameter(Mandatory = $true)]
		[string[]]
		$TypeParameters,
		[parameter(Mandatory = $false)]
		[object[]] 
		$ConstructorParameters
	)

	process
	{
		$genericTypeName = $typeName + '`' + $typeParameters.Count
		$genericType = [Type]$genericTypeName

		if (!$genericType)
		{
			throw "Could not find generic type $genericTypeName"
		}

		# Bind the type arguments to it
		$typedParameters = [type[]] $TypeParameters
		$closedType = $genericType.MakeGenericType($typedParameters)
	
		if (!$closedType)
		{
			throw "Could not make closed type $genericType"
		}

		# Create the closed version of the generic type, don't forget comma prefix
		,[Activator]::CreateInstance($closedType, $constructorParameters)
	}
}

function Test-Variable
{
	<#
	.Synopsis
		Tests if a variable is not null.
	.Description
		Tests if a variable is not null. Returns $true if the variable is declared and not null
	.Example
		Test-Variable -Name "session" -Scope "global"
	#>

	[CmdletBinding()]
	[OutputType([bool])]
	param (
		[parameter(Mandatory = $true)]
		[string]
		$Name,
		[parameter(Mandatory = $true)]
		[string]
		$Scope
	)

	process
	{
		# Return $true if the variable is declared and not null
	
		if ($Scope -eq "local")
		{
			$Scope = "1" # Parent Scope
		}

		$declaredAndNotNull = (Test-Path "variable:\${Scope}:$Name") -and (Get-Variable -Name $Name -Scope $Scope -ValueOnly) -ne $null

		return $declaredAndNotNull
	}
}

function ConvertFrom-SchemaXml
{
	<#
	.Synopsis
		Converts a connector schema defined in a xml file into a "Microsoft.MetadirectoryServices.Schema" object.
	.Description
		Converts a connector schema defined in a xml file into a "Microsoft.MetadirectoryServices.Schema" object.
	.Example
		ConvertFrom-SchemaXml -SchemaXml "Schema-Lync.xml"
	#>

	[CmdletBinding()]
	[OutputType([Microsoft.MetadirectoryServices.Schema])]
	param (
		[parameter(Mandatory = $true)]
		[ValidateScript({ Test-Path $_ -PathType "Leaf" })]
		[string]
		$SchemaXml
	)

	process
	{
		$x = [xml](Get-Content $SchemaXml)
	
		$schema = [Microsoft.MetadirectoryServices.Schema]::Create()

		foreach ($t in $x.Schema.Types.SchemaType)
		{
			$lockAnchorDefinition = $true
		
			if ($t.LockAnchorDefinition -eq "0")
			{
				$lockAnchorDefinition = $false
			}
		
			$schemaType = [Microsoft.MetadirectoryServices.SchemaType]::Create($t.Name, $lockAnchorDefinition)

			if ($t.GetElementsByTagName("PossibleDNComponentsForProvisioning").Count -gt 0)
			{
				foreach ($c in $t.PossibleDNComponentsForProvisioning)
				{
					$schemaType.PossibleDNComponentsForProvisioning.Add($c)
				}
			}

			foreach ($a in $t.Attributes.SchemaAttribute)
			{
				if ($a.IsAnchor -eq 1)
				{
					$schemaType.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateAnchorAttribute($a.Name, $a.DataType, $a.AllowedAttributeOperation))
				}
				elseif ($a.IsMultiValued -eq 1)
				{
					$schemaType.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateMultiValuedAttribute($a.Name, $a.DataType, $a.AllowedAttributeOperation))
				}
				else
				{
					$schemaType.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute($a.Name, $a.DataType, $a.AllowedAttributeOperation))
				}
			}

			$schema.Types.Add($schemaType)
		}

		return $schema
	}
}

function Get-CSEntryChangeValue
{
	<#
	.Synopsis
		Gets the value of the specified attribute of the CSEntryChange object.
	.Description
		Gets the value of the specified attribute of the CSEntryChange object.
	.Example
		Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "RegistrarPool"
	.Example
		Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "RegistrarPool" -DefaultValue "pool01.contoso.com"
	.Example
		Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "RegistrarPool" -DefaultValue "pool01.contoso.com" -OldValue
	#>

	[CmdletBinding()]
	[OutputType([object])]
	param (
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange,
		[parameter(Mandatory = $true)]
		[string]
		$AttributeName,
		[parameter(Mandatory = $false)]
		[object]
		$DefaultValue = $null,
		[parameter(Mandatory = $false)]
		[switch]
		$OldValue
	)

	process
	{
		if ($CSEntryChange.AttributeChanges.Contains($AttributeName))
		{
			$returnDefault = $true
		
			$attributeChange = $CSEntryChange.AttributeChanges[$AttributeName]
		
			foreach ($valueChange in $attributeChange.ValueChanges)
			{
				if ($OldValue)
				{
					if ($valueChange.ModificationType -eq "Delete")
					{
						$valueChange.Value # add to return pipeline
						$returnDefault = $false
					}
				}
				else
				{
					if ($valueChange.ModificationType -eq "Add")
					{
						$valueChange.Value # add to return pipeline
						$returnDefault = $false
					}
				}
			}

			if ($returnDefault)
			{
				$DefaultValue # return
			}
		}
		else
		{
			$DefaultValue # return
		}
	}
}

function Get-CSEntryChangeValueIfChanged
{
	<#
	.Synopsis
		Gets the new value of the specified attribute of the CSEntryChange object if it was changed.
	.Description
		Gets the new value of the specified attribute of the CSEntryChange object if it was changed.
	.Example
		Get-CSEntryChangeValue -CSEntryChangeValue $csentryChange -AttributeName "RegistrarPool"
	#>
	
	[CmdletBinding()]
	[OutputType([object])]
	param (
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange,
		[parameter(Mandatory = $true)]
		[string]
		$AttributeName
	)

	process
	{
		$returnValue = $null
	
		if ($CSEntryChange.AttributeChanges.Contains($AttributeName))
		{
			$oldAttributeValue = $null
			$newAttributeValue = $null
		
			$attributeChange = $CSEntryChange.AttributeChanges[$AttributeName]
		
			if ($attributeChange.IsMultiValued)
			{
				throw "Support for multivalued attribute is not implemented in this function. Attribute Name: $AttributeName."
			}
		
			foreach ($valueChange in $attributeChange.ValueChanges)
			{
				if ($valueChange.ModificationType -eq "Delete")
				{
					$oldAttributeValue = $valueChange.Value
				}
				elseif ($valueChange.ModificationType -eq "Add")
				{
					$newAttributeValue = $valueChange.Value
				}
			}

			if ($oldAttributeValue -cne $newAttributeValue)
			{
				$returnValue = $newAttributeValue
			}
		}
	
		return $returnValue
	}
}

function Test-CSEntryChangeValueChanged
{
	<#
	.Synopsis
		Tests if the value of the specified attribute of the CSEntryChange object was changed.
	.Description
		Tests if the value of the specified attribute of the CSEntryChange object was changed.
	.Example
		Get-CSEntryChangeValue -CSEntryChangeValue $csentryChange -AttributeName "RegistrarPool"
	.Example
		Get-CSEntryChangeValue -CSEntryChangeValue $csentryChange -AttributeName "RegistrarPool" -IgnoreCase
	#>
	
	[CmdletBinding()]
	[OutputType([bool])]
	param (
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange,
		[parameter(Mandatory = $true)]
		[string]
		$AttributeName,
		[parameter(Mandatory = $false)]
		[switch]
		$IgnoreCase
	)
	
	process
	{
		$valueChanged = $false
	
		if ($CSEntryChange.AttributeChanges.Contains($AttributeName))
		{
			$oldAttributeValue = $null
			$newAttributeValue = $null
		
			$attributeChange = $CSEntryChange.AttributeChanges[$AttributeName]
		
			if ($attributeChange.IsMultiValued)
			{
					throw "Support for multivalued attribute is not implemented in this function. Attribute Name: $AttributeName."
			}
		
			foreach ($valueChange in $attributeChange.ValueChanges)
			{
				if ($valueChange.ModificationType -eq "Delete")
				{
					$oldAttributeValue = $valueChange.Value
				}
				elseif ($valueChange.ModificationType -eq "Add")
				{
					$newAttributeValue = $valueChange.Value
				}
			}

			if ($IgnoreCase)
			{
				if ($oldAttributeValue -ne $newAttributeValue)
				{
					$valueChanged = $true
				}
			}
			else
			{
				if ($oldAttributeValue -cne $newAttributeValue)
				{
					$valueChanged = $true
				}
			}
		}
	
		return $valueChanged
	}
}

function Test-CSEntryChangeAttributeDeleted
{
	<#
	.Synopsis
		Tests if the value of the specified attribute of the CSEntryChange object was deleted.
	.Description
		Tests if the value of the specified attribute of the CSEntryChange object was deleted.
	.Example
		Get-CSEntryChangeValue -CSEntryChangeValue $csentryChange -AttributeName "RegistrarPool"
	#>
	
	[CmdletBinding()]
	[OutputType([bool])]
	param (
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange,
		[parameter(Mandatory = $true)]
		[string]
		$AttributeName
	)

	process
	{
		$attributeDeleted = $false
	
		if ($CSEntryChange.AttributeChanges.Contains($AttributeName))
		{
			$attributeChange = $CSEntryChange.AttributeChanges[$AttributeName]
		
			if ($attributeChange.ModificationType -eq "Delete")
			{
				$attributeDeleted = $true
			}
		}
	
		return $attributeDeleted
	}
}

function Get-CSEntryChangeDN
{
	<#
	.Synopsis
		Gets the DN of the CSEntryChange object.
	.Description
		Gets the DN of the CSEntryChange object.
	.Example
		Get-CSEntryChangeValue -CSEntryChangeValue $csentryChange
	#>

	[CmdletBinding()]
	[OutputType([string])]
	param (
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange
	)
	process
	{
		return Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "DN" -DefaultValue $csentryChange.DN
	}
}

function Get-ConfigParameter
{
	<#
	.Synopsis
		Gets the value of a configuration parameter.
	.Description
		Gets the value of a configuration parameter.
		If the "Scope" parameter is not specified, the value is looked up in "RunStep", "Partition", "Global", "Connectivity" in that order.
	.Example
		Get-ConfigParameter -ConfigParameters $configParameters -ParameterName $parameterName
	.Example
		Get-ConfigParameter -ConfigParameters $configParameters -ParameterName $parameterName -Scope "Connectivity"
	.Example
		Get-ConfigParameter -ConfigParameters $configParameters -ParameterName $parameterName -Scope "Connectivity" -Encrypted
	#>

	[CmdletBinding()]
	[OutputType([string])]
	param (
		[parameter(Mandatory = $true)]
		[System.Collections.ObjectModel.KeyedCollection[string,Microsoft.MetadirectoryServices.ConfigParameter]]
		$ConfigParameters,
		[parameter(Mandatory = $true)]
		[string]
		$ParameterName,
		[parameter(Mandatory = $false)]
		[ValidateSet("RunStep", "Partition", "Global", "Connectivity", "")]
		[string]
		$Scope,
		[parameter(Mandatory = $false)]
		[switch]
		$Encrypted
	)

	process
	{
		$configParameterValue = $null

		if ([string]::IsNullOrEmpty($Scope))
		{
			$configParameterValue = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName $ParameterName -Scope "RunStep" -Encrypted:$Encrypted

			if ([string]::IsNullOrEmpty($configParameterValue))
			{
				$configParameterValue = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName $ParameterName -Scope "Partition" -Encrypted:$Encrypted

				if ([string]::IsNullOrEmpty($configParameterValue))
				{
					$configParameterValue = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName $ParameterName -Scope "Global" -Encrypted:$Encrypted

					if ([string]::IsNullOrEmpty($configParameterValue))
					{
						$configParameterValue = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName $ParameterName -Scope "Connectivity" -Encrypted:$Encrypted
					}
				}
			}
		}
		elseif ($Scope -eq "RunStep" -or $Scope -eq "Partition" -or $Scope -eq "Global" -or "Connectivity")
		{
			if ($Scope -eq "RunStep" -or $Scope -eq "Partition" -or $Scope -eq "Global")
			{
				$configParameterName = "{0}_{1}" -f $ParameterName, $Scope
			}
			else
			{
				$configParameterName = $ParameterName
			}

			if ($ConfigParameters.Contains($configParameterName))
			{
				if ($Encrypted -ne $true)
				{
					$configParameterValue =  $ConfigParameters[$configParameterName].Value

					if (![string]::IsNullOrEmpty($configParameterValue))
					{
					   $configParameterValue = $configParameterValue.Trim() 
					}

					Write-Verbose ("ConfigParameter: Scope={0}, Name={1}, Value={2}" -f $Scope, $ParameterName, $configParameterValue)
				}
				else
				{
					$configParameterValue =  $ConfigParameters[$configParameterName].SecureValue

					Write-Verbose ("ConfigParameter: Scope={0}, Name={1}, Value={2}" -f $Scope, $ParameterName, "***Encrypted***")
				}
			}
		}
		else
		{
			throw "Invalid ConfigurationParameter scope: $Scope"
		}

		return $configParameterValue
	}
}

#region Import Helpers

function New-CSEntryChange
{
	<#
	.Synopsis
		Creates a new CSEntryChange object from the specified InputObject
	.Description
		Creates a new CSEntryChange object from the specified InputObject.
		The function expects the DN to be $InputObject.DistinguishedName and
		the property names to be same as schema attribute names.
	#>
	
	[CmdletBinding()]
	[OutputType([Microsoft.MetadirectoryServices.CSEntryChange])]
	param(
		[parameter(Mandatory = $true, ValueFromPipeline = $true)]
		[ValidateNotNull()]
		[object]
		$InputObject,
		[parameter(Mandatory = $true)]
		[string]
		$ObjectType,
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.Schema]
		$Schema
	)

	$csentry = [Microsoft.MetadirectoryServices.CSEntryChange]::Create()
	$csentry.ObjectModificationType = "Add"
	$csentry.ObjectType = $ObjectType
	$csentry.DN = $InputObject.DistinguishedName

	foreach ($attribute in $schema.Types[$csentry.ObjectType].Attributes)
	{
		Write-Debug ("Processing CSEntry: '{0}'. Attribute: '{1}'." -f $csentry.DN, $attribute.Name)

		$attributeVal = $InputObject.($attribute.Name)

		Write-Debug ("Processing CSEntry: '{0}'. Attribute: '{1}'. Attribute Value: '{2}'." -f $csentry.DN, $attribute.Name, $attributeVal)

		if ($attributeVal -ne $null)
		{
			if ($attributeVal -is [System.Collections.ArrayList]) # do not compare array with [string]::Empty as 0 is treated as empty string in PowerShell.
			{
				if ($attributeVal.Count -eq 0)
				{
					$attributeVal = $null
				}
			}
			elseif (($attribute.DataType -eq "String" -or  $attribute.DataType -eq "Reference") -and $attributeVal -eq [string]::Empty)
			{
				$attributeVal = $null
			}
			elseif ($attribute.DataType -eq "Integer" -and $attributeVal -eq [string]::Empty)
			{
				$attributeVal = $null
			}
			elseif ($attribute.DataType -eq "Boolean" -and $attributeVal -eq [string]::Empty)
			{
				$attributeVal = $false
			}
		}

		if ($attributeVal -ne $null)
		{
			if ($attributeVal -is [System.Collections.ArrayList])
			{	
				$attributeVal = [string[]] $attributeVal # TODO: Support other multi-valued datatypes
			}
			else
			{
				if ($attribute.DataType -eq "Binary")
				{
					if ($attributeVal -is [Guid])
					{
						$attributeVal = $attributeVal.ToByteArray()
					}
				}
			}

			try
			{
				if ($attribute.IsAnchor)
				{
					$csentry | Add-CSEntryAnchorAttribute -AttributeName $attribute.Name -AttributeValue $attributeVal
				}
				else
				{
					$csentry | Add-CSEntryAttributeChange -AttributeModificationType "Add" -AttributeName $attribute.Name -AttributeValue $attributeVal
				}
			}
			catch
			{
				$e = "Error with property: " + $attribute.Name + ". Error: " + $_.Exception.ToString()
				Write-Debug $e
			}
		}
		else
		{
			Write-Verbose ("Skipped populating null or empty attribute: {0}" -f $attribute.Name)
		}
	}

	return $csentry
}

function Add-CSEntryAnchorAttribute
{
	<#
	.Synopsis
		Adds the specified anchor attribute to the specified CSEntryChange object.
	.Description
		Adds the specified anchor attribute to the specified CSEntryChange object.
	.Example
	#>

	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true, ValueFromPipeline = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange,
		[parameter(Mandatory = $true)]
		[string]
		$AttributeName,
		[parameter(Mandatory = $true)]
		[object]
		$AttributeValue
	)

	process
	{
		[void] $CSEntryChange.AnchorAttributes.Add([Microsoft.MetadirectoryServices.AnchorAttribute]::Create($AttributeName, $AttributeValue))
	}
}

function Add-CSEntryAttributeChange
{
	<#
	.Synopsis
		Adds the specified attribute change to the specified CSEntryChange object.
	.Description
		Adds the specified attribute change to the specified CSEntryChange object.
	.Example
		$csentry | Add-CSEntryAttributeChange -AttributeModificationType "Add" -AttributeName $attribute.Name -AttributeValue $attributeVal
	#>

	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true, ValueFromPipeline = $true)]
		[Microsoft.MetadirectoryServices.CSEntryChange]
		$CSEntryChange,
		[parameter(Mandatory = $true)]
		[ValidateSet("Add", "Update", "Delete", "Replace", "Rename")]
		[string]
		$AttributeModificationType,
		[parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[string]
		$AttributeName,
		[parameter(Mandatory = $false)]
		[object]
		$AttributeValue
	)

	process
	{
		if ($AttributeModificationType -ne 'Rename' -and $AttributeName -eq $null)
		{
			throw ("AttributeName parameter is required. AttributeModificationType: {0}. CSEntry: {1}." -f $AttributeModificationType, $CSEntryChange)
		}

		if ($AttributeModificationType -ne 'Delete' -and $AttributeValue -eq $null)
		{
			throw ("AttributeValue parameter is required. AttributeModificationType: {0}. CSEntry: {1}." -f $AttributeModificationType, $CSEntryChange)
		}

		switch ($AttributeModificationType)
		{
			'Add'
			{
				[void] $CSEntryChange.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeAdd($AttributeName, $AttributeValue))
			}
			'Update'
			{
				[void] $CSEntryChange.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeUpdate($AttributeName, $AttributeValue))
			}
			'Delete'
			{
				[void] $CSEntryChange.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeDelete($AttributeName))
			}
			'Replace'
			{
				[void] $CSEntryChange.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeReplace($AttributeName, $AttributeValue))
			}
			'Rename'
			{
				[void] $CSEntryChange.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateNewDN($AttributeValue))
			}
			default
			{
				throw ("AttributeModificationType: {0} is not handled. AttributeName: {1}. CSEntry: {2}." -f $AttributeModificationType, $AttributeName, $CSEntryChange)
			}
		}
	}
}

#endregion

#region Export Helpers

function New-CSEntryChangeExportError
{
	<#
	.Synopsis
		Creates a new CSEntryChangeResult object for the specified CSEntryChange and specified error.
	.Description
		Creates a new CSEntryChangeResult object for the specified CSEntryChange and specified error.
	#>
	
	[CmdletBinding()]
	[OutputType([Microsoft.MetadirectoryServices.CSEntryChangeResult])]
	param(
		[parameter(Mandatory = $true)]
		[Guid]
		$CSEntryChangeIdentifier,
		[parameter(Mandatory = $true)]
		[object]
		$ErrorObject
	)

	$csentryChangeResult = $null

	try
	{
		foreach ($cmdStatus in $ErrorObject)
		{
			$exception = $cmdStatus.GetBaseException()
			$exceptionType = $exception.GetType().Name
			$exceptionMessage = $exception.Message

			Write-Warning ("CSEntry Identifier: {0}. ErrorName: {1}. ErrorDetail: {2}" -f $CSEntryChangeIdentifier, $exceptionType, $exceptionMessage)
			$csentryChangeResult = [Microsoft.MetadirectoryServices.CSEntryChangeResult]::Create($CSEntryChangeIdentifier, $null, "ExportErrorCustomContinueRun", $exceptionType, $exceptionMessage)
			Write-Warning ("CSEntryChangeResult Identifier: {0}. ErrorCode: {1}. ErrorName: {2}. ErrorDetail: {3}" -f $csentryChangeResult.Identifier, $csentryChangeResult.ErrorCode, $csentryChangeResult.ErrorName, $csentryChangeResult.ErrorDetail)
					
			break # report the first error and stop
		}
	}
	catch
	{
		foreach ($cmdStatus in $ErrorObject)
		{
			$exceptionType = "RUNTIME_EXCEPTION"
			$exceptionMessage = $cmdStatus.ToString()
			Write-Warning ("CSEntry Identifier: {0}. Error: {1}" -f $CSEntryChangeIdentifier, $exceptionMessage)
			$csentryChangeResult = [Microsoft.MetadirectoryServices.CSEntryChangeResult]::Create($CSEntryChangeIdentifier, $null, "ExportErrorCustomContinueRun", $exceptionType, $exceptionMessage)
			Write-Warning ("CSEntryChangeResult Identifier: {0}. ErrorCode: {1}. ErrorName: {2}. ErrorDetail: {3}" -f $csentryChangeResult.Identifier, $csentryChangeResult.ErrorCode, $csentryChangeResult.ErrorName, $csentryChangeResult.ErrorDetail)

			break # report the first error and stop
		}
	}

	return $csentryChangeResult
}

function New-CSEntryChangeResult
{
	<#
	.Synopsis
		Creates a new CSEntryChangeResult object for the specified CSEntryChange.
	.Description
		Creates a new CSEntryChangeResult object for the specified CSEntryChange.
	#>
	
	[CmdletBinding()]
	[OutputType([Microsoft.MetadirectoryServices.CSEntryChangeResult])]
	param(
		[parameter(Mandatory = $true)]
		[Guid]
		$CSEntryChangeIdentifier,
		[parameter(Mandatory = $false)]
		[Hashtable]
		$NewAnchorTable,
		[Switch]
		$ExportAdd
	)
	
	$attributeChanges = $null

	if ($ExportAdd)
	{
		if ($NewArchorTable -eq $null -or !$NewAnchorTable.Keys.Count -eq 0)
		{
			throw "The NewAnchorTable parameter must not be null."
		}

		foreach ($anchor in $NewAnchorTable.Keys)
		{
			$anchorValue = $NewAnchorTable[$anchor]

			$attributeChanges = New-GenericObject System.Collections.Generic.List Microsoft.MetadirectoryServices.AttributeChange
			$anchorAttribute = [Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeAdd($anchor, $anchorValue)
			$attributeChanges.Add($anchorAttribute)
		}
	}

	$csentryChangeResult = [Microsoft.MetadirectoryServices.CSEntryChangeResult]::Create($CSEntryChangeIdentifier, $attributeChanges, "Success")

	return $csentryChangeResult
}

#endregion Export Helpers

function Select-PreferredDomainController
{
	<#
	.Synopsis
		Selects the preferred domain controller that is online from the specified comma-separated list of preferred domain controllers.
	.Description
		Selects the preferred domain controller that is online from the specified comma-separated list of preferred domain controllers.
	.Example
		Select-PreferredDomainController -$DomainControllerList $PreferredDomainController
	#>

	[CmdletBinding()]
	[OutputType([string])]
	param (
		[parameter(Mandatory = $false)]
		[string]
		$DomainControllerList
	)

	process
	{
		if (![string]::IsNullOrEmpty($DomainControllerList))
		{
			$selected = $false

			foreach ($preferredDC in $DomainControllerList.Split(","))
			{
				if (![string]::IsNullOrEmpty($preferredDC))
				{
					try
					{
						$conn = New-Object "Net.Sockets.TcpClient"
						$conn.Connect($preferredDC, 389)
						$selected = $true

						Write-Debug ("Preferred Domain Controller is: {0}." -f $preferredDC)

						break
					}
					catch
					{
						Write-Warning ("Domain Controller {0} is unavailable." -f $preferredDC)
					}
				}
			}

			if ($selected)
			{
				return $preferredDC
			}
			else
			{
				throw [Microsoft.MetadirectoryServices.ServerDownException] ("None of the servers from the Preferred Domain Controller List '{0}' is online." -f $DomainControllerList)
			}
		}
	}
}

Export-ModuleMember -Function * -Variable *
