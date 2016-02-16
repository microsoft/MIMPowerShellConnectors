<#
<copyright file="TestHarness.Common.psm1" company="Microsoft">
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
</copyright>
<summary>
	Common Functions shared by all PowerShell Connectors Test Harness scripts.
</summary>
#>

Set-StrictMode -Version "2.0"

function ConvertTo-RunProfileAuditFile
{
	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.GetImportEntriesResults]
		$GetImportEntriesResults,
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.Schema]
		$Schema,
		[parameter(Mandatory = $true)]
		[string]
		$AuditFileName,
		[parameter(Mandatory = $true)]
		[ValidateSet("full-import", "delta-import", "export")]
		[string]
		$RunProfileStepType
	)

	process
	{
		$RunProfileStepType = $RunProfileStepType.ToLowerInvariant()
		$namespaceUri = "http://www.microsoft.com/mms/mmsml/v2"
		$document = [xml] "<mmsml xmlns=`"$namespaceUri`" step-type=`"$RunProfileStepType`"><directory-entries/></mmsml>"
		[System.Xml.XmlNamespaceManager] $nsm = $document.NameTable
		$nsm.AddNamespace("x", $namespaceUri)
		$directoryEntries = $document.SelectSingleNode("//x:directory-entries", $nsm)

		foreach ($csentry in $GetImportEntriesResults.CSEntries)
		{
			$type = $Schema.Types[$csentry.ObjectType]

			$directoryEntry = $document.CreateElement("delta", $namespaceUri)
			[void]$directoryEntries.AppendChild($directoryEntry)

			#region set operation and DN

			$directoryEntryAttribute = $document.CreateAttribute("operation")
			$directoryEntryAttribute.Value = ([string]$csentry.ObjectModificationType).ToLowerInvariant()
			[void]$directoryEntry.Attributes.Append($directoryEntryAttribute)

			$directoryEntryAttribute = $document.CreateAttribute("dn")
			$directoryEntryAttribute.Value = $csentry.DN
			[void]$directoryEntry.Attributes.Append($directoryEntryAttribute)

			#endregion set operation and DN

			#region set anchor

			$directoryEntryAnchor = $document.CreateElement("anchor", $namespaceUri)
			[void]$directoryEntry.AppendChild($directoryEntryAnchor)

			$anchorAttributeValue = ""
			$binaryAnchor = $false
			foreach ($attribute in $csentry.AnchorAttributes)
			{
				$anchorAttributeValue += 
					switch ($type.Attributes[$attribute.Name].DataType)
					{
							"Binary"
							{
								$anchor = @(16,0,0,0) + $attribute.Value # seems to be the case how anchor is encoded
								[Convert]::ToBase64String($anchor)
								$binaryAnchor = $true
							}

							default
							{
								$attribute.Value + "+"
							}
					}
			}
			
			$directoryEntryAnchor.InnerText = $anchorAttributeValue.TrimEnd("+")

			if ($binaryAnchor)
			{
				$directoryEntryAnchorEncoding = $document.CreateAttribute("encoding")
				$directoryEntryAnchorEncoding.Value = "base64"
				[void]$directoryEntryAnchor.Attributes.Append($directoryEntryAnchorEncoding)
			}

			#endregion set anchor

			#region set objectclass

			$directoryEntryObjectClass = $document.CreateElement("primary-objectclass", $namespaceUri)
			[void]$directoryEntry.AppendChild($directoryEntryObjectClass)
			$directoryEntryObjectClass.InnerText = $csentry.ObjectType

			$directoryEntryObjectClass = $document.CreateElement("objectclass", $namespaceUri)
			[void]$directoryEntry.AppendChild($directoryEntryObjectClass)
			$directoryEntryObjectClassValue = $document.CreateElement("oc-value", $namespaceUri)
			[void]$directoryEntryObjectClass.AppendChild($directoryEntryObjectClassValue)
			$directoryEntryObjectClassValue.InnerText = $csentry.ObjectType

			#endregion set objectclass

			#region set attributes

			foreach ($attribute in $csentry.AttributeChanges)
			{
				$directoryEntryAttribute = $document.CreateElement("attr", $namespaceUri)
				[void]$directoryEntry.AppendChild($directoryEntryAttribute)

				$directoryEntryAttributeName = $document.CreateAttribute("name")
				[void]$directoryEntryAttribute.Attributes.Append($directoryEntryAttributeName)
				$directoryEntryAttributeName.Value = $attribute.Name

				if ($RunProfileStepType -ne "full-import" -and $csentry.ObjectModificationType -ne "Add")
				{
					$directoryEntryAttributeOperation = $document.CreateAttribute("operation")
					[void]$directoryEntryAttribute.Attributes.Append($directoryEntryAttributeOperation)
					$directoryEntryAttributeOperation.Value = ([string]$attribute.ModificationType).ToLowerInvariant()
				}

				$directoryEntryAttributeType = $document.CreateAttribute("type")
				[void]$directoryEntryAttribute.Attributes.Append($directoryEntryAttributeType)
				$directoryEntryAttributeType.Value = ([string]$type.Attributes[$attribute.Name].DataType).ToLowerInvariant()

				$directoryEntryAttributeMultivalued = $document.CreateAttribute("multivalued")
				[void]$directoryEntryAttribute.Attributes.Append($directoryEntryAttributeMultivalued)
				$directoryEntryAttributeMultivalued.Value = $type.Attributes[$attribute.Name].IsMultiValued -eq 1

				foreach ($value in $attribute.ValueChanges)
				{
					$directoryEntryAttributeValue = $document.CreateElement("value", $namespaceUri)
					[void]$directoryEntryAttribute.AppendChild($directoryEntryAttributeValue)
					$directoryEntryAttributeValue.InnerText = 
						switch ($type.Attributes[$attribute.Name].DataType)
						{
								"Binary"
								{
									[Convert]::ToBase64String($value.Value)
								}

								default
								{
									$value.Value
								}
						}
				}
			}

			#end region set attributes
		}

		$document.Save($AuditFileName)
	}
}

function ConvertFrom-RunProfileAuditFile
{
	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true)]
		[string]
		$AuditFileName,
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.Schema]
		$Schema
	)

	process
	{
		$document = [xml] (Get-Content $AuditFileName)

		$namespaceUri = "http://www.microsoft.com/mms/mmsml/v2"
		[System.Xml.XmlNamespaceManager] $nsm = $document.NameTable
		$nsm.AddNamespace("x", $namespaceUri)
		$directoryEntries = $document.SelectNodes("//x:directory-entries/x:delta", $nsm)

		$csentryChanges = New-GenericObject System.Collections.Generic.List Microsoft.MetadirectoryServices.CSEntryChange

		foreach ($entry in $directoryEntries)
		{
			$objectModificationType = $entry.operation

			$csentryChange = [Microsoft.MetadirectoryServices.CSEntryChange]::Create()
			$csentryChange.ObjectModificationType = $entry.operation
			$csentryChange.ObjectType = $entry.'primary-objectclass' # Make sure the export drop file is edited to added a "primary-objectclass" element for export-update or export-delete
			$csentryChange.DN = $entry.dn

			if ($entry.SelectSingleNode("x:anchor", $nsm) -ne $null)
			{
				$anchorAttributes = Get-CSEntryAnchorAttribute -AnchorValue $entry.anchor -SchemaType $Schema.Types[$csentryChange.ObjectType]
				
				foreach ($anchorAttribute in $anchorAttributes)
				{
					[void]$csentryChange.AnchorAttributes.Add($anchorAttribute)
				}
			}

			switch -Regex ($objectModificationType)
			{
				"^(add|replace|update)$"
				{
					foreach ($attr in $entry.attr)
					{
						$attributeChange = Get-CSEntryAttributeChange $attr $Schema.Types[$csentryChange.ObjectType]

						if ($attributeChange -ne $null)
						{
							[void] $csentryChange.AttributeChanges.Add($attributeChange)
						}
					}

					break
				}

				"^(Delete)$"
				{
					break
				}

				default
				{
					Write-Warning ("Object {0} is skipped." -f $csentryChange.DN)
				}

			}

			[void] $csentryChanges.Add($csentryChange)
		}

		return ,$csentryChanges # prevents unwinding
	}
}

function Get-CSEntryAnchorAttribute
{
	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true)]
		[object]
		$AnchorValue,
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.SchemaType]
		$SchemaType
	)

	process
	{
		foreach ($schemaAttribute in $SchemaType.Attributes)
		{
			if ($schemaAttribute.IsAnchor)
			{
				$attributeValue = ConvertTo-DataType -Data $AnchorValue -DataType $schemaAttribute.DataType
				if ($schemaAttribute.DataType -eq "Binary")
				{
					$attributeValue = [byte[]]$attributeValue[4..($attributeValue.Length-1)] # seems to be the case how anchor is encoded
				}

				$anchorAttribute = New-Object -TypeName "Microsoft.MetadirectoryServices.DetachedObjectModel.AnchorAttributeDetached" -ArgumentList $schemaAttribute.Name, $schemaAttribute.DataType, $attributeValue
	
				$anchorAttribute #return on pipe-line
				break #TODO - Handle more than one anchor attributes
			}
		}
	}
}

function Get-CSEntryAttributeChange
{
	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true)]
		[System.Xml.XmlElement]
		$Attr,
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.SchemaType]
		$SchemaType
	)

	$attributeChange = $null
	$attributeName = $Attr.name
	$attributeModificationType = "Add"

	if ($Attr.Attributes["operation"] -ne $null)
	{
		$attributeModificationType = $Attr.Attributes["operation"].Value
	}
	
	if (!$SchemaType.Attributes.Contains($attributeName))
	{
		throw ("The connector schema does not contain attribute '{0}'" -f $attributeName)
	}

	$schemaAttribute = $SchemaType.Attributes[$attributeName]

	switch -Regex ($attributeModificationType)
	{
		"^(add|replace|update)$"
		{
			$attributeValue = Get-CSEntryAttributeChangeValueChanges -Attr $Attr -SchemaAttribute $schemaAttribute
			##$attributeChange = [Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeAdd($attributeName, $attributeValue) # commented as it does not attach attribute schema to the AttributeChange
			$attributeChange = New-Object -TypeName "Microsoft.MetadirectoryServices.DetachedObjectModel.AttributeChangeDetached" -ArgumentList $schemaAttribute, $attributeModificationType, $attributeValue
		}

		"^(delete)$"
		{
			$attributeChange = [Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeDelete($attributeName)
		}

		default
		{
			Write-Warning ("Attribute {0} is skipped" -f $attributeName)
		}
	}
	
	return $attributeChange
}

function Get-CSEntryAttributeChangeValueChanges
{
	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true)]
		[System.Xml.XmlElement]
		$Attr,
		[parameter(Mandatory = $true)]
		[Microsoft.MetadirectoryServices.SchemaAttribute]
		$SchemaAttribute
	)

	process
	{
		$attributeValue = $null
		$attributeName = $Attr.Name

		$attributeValue = New-GenericObject "System.Collections.Generic.List" "Microsoft.MetadirectoryServices.ValueChange"

		foreach ($value in $Attr.Value)
		{
			$valueModificationType = "Add"
			if ($value -is [System.Xml.XmlElement] -and $value.Attributes["operation"] -ne $null)
			{
				$valueModificationType = $value.Attributes["operation"].Value
			}

			$value = ConvertTo-DataType -Data $value -DataType $schemaAttribute.DataType

			if ($valueModificationType -eq "Add")
			{
				$attributeValue.Add([Microsoft.MetadirectoryServices.ValueChange]::CreateValueAdd($value))
			}
			else
			{
				$attributeValue.Add([Microsoft.MetadirectoryServices.ValueChange]::CreateValueDelete($value))
			}
		}

		return ,$attributeValue # prevent unwinding
	}
}

function ConvertTo-DataType
{
	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true)]
		[object]
		$Data,
		[parameter(Mandatory = $true)]
		[ValidateSet("String", "Integer", "Boolean", "Binary", "Reference")]
		[string]
		$DataType
	)

	process
	{
		$value = $null

		if ($Data -ne $null)
		{
			if ($Data.GetType().Name -eq "XmlElement")
			{
				$value = $Data.InnerText
			}
			else
			{
				$value = $Data
			}
		}

		switch ($DataType)
		{
			"String"
			{
				break
			}

			"Integer"
			{
				$value = [Convert]::ToInt64($value)
				break
			}

			"Boolean"
			{
				$value = [Convert]::ToBoolean($value)
				break
			}

			"Binary"
			{
				$value = [Convert]::FromBase64String($value)
				break
			}

			"Reference"
			{
				break
			}
		}

		if ($value -is [array]) # handle byte[]
		{
			return ,$value
		}
		else
		{
			return $value
		}
	}
}
