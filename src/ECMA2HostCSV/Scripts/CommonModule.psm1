Set-PSDebug -Strict

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
  param(
    [Parameter(Mandatory = $true)]
    [string]
    $ScriptType,
    [Parameter(Mandatory = $false)]
    [ValidateNotNull()]
    [System.Collections.ArrayList]
    $ErrorObject
  )

  process
  {
    Write-Verbose "$Global:ConnectorName - $ScriptType Script: Execution Started..."
    if ($ErrorObject)
    {
      $ErrorObject.Clear()
    }
  }
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
  param(
    [Parameter(Mandatory = $true)]
    [string]
    $ScriptType,
    [Parameter(Mandatory = $false)]
    [ValidateNotNull()]
    [System.Collections.ArrayList]
    $ErrorObject,
    [Parameter(Mandatory = $false)]
    [switch]
    $SuppressErrorCheck,
    [Parameter(Mandatory = $false)]
    [type]
    $ExceptionRaisedOnErrorCheck
  )

  process
  {
    if (!$SuppressErrorCheck -and $ErrorObject -and $ErrorObject.Count -ne 0)
    {
      # Take the first one otherwise you get "An error occurred while enumerating through a collection: Collection was modified; enumeration operation may not execute.."
      # Seems like a bug in Remote PSH
      $errorMessage = $ErrorObject[0] # | Out-String -ErrorAction SilentlyContinue

      if ($ExceptionRaisedOnErrorCheck -eq $null)
      {
        $ExceptionRaisedOnErrorCheck = [Microsoft.MetadirectoryServices.ExtensibleExtensionException]
      }

      $ErrorObject.Clear()

      throw $errorMessage -as $ExceptionRaisedOnErrorCheck
    }

    Write-Verbose "$Global:ConnectorName - $ScriptType Script: Execution Completed."
  }
}

function Get-ExtensionsDirectory
{
  <#
    .Synopsis
    Gets the path of the "Extensions" folder.
    .Description
    Gets the path of the "Extensions" folder.
  #>
  [CmdletBinding()]
  [OutputType([string])]
  param(
  )

  process
  {
    $scriptDir = "C:\\Program Files\\Microsoft ECMA2Host\\Service\\ECMA"

    return $scriptDir
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
    ConvertFrom-SchemaXml -SchemaXml "Schema.xml"
    #>  

  [CmdletBinding()]
  [OutputType([Microsoft.MetadirectoryServices.Schema])]
  param(
    [Parameter(Mandatory = $true)]
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

      $schemaType = [Microsoft.MetadirectoryServices.SchemaType]::Create($t.Name,$lockAnchorDefinition)

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
          $schemaType.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateAnchorAttribute($a.Name,$a.DataType,$a.AllowedAttributeOperation))
        }
        elseif ($a.IsMultiValued -eq 1)
        {
          $schemaType.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateMultiValuedAttribute($a.Name,$a.DataType,$a.AllowedAttributeOperation))
        }
        else
        {
          $schemaType.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute($a.Name,$a.DataType,$a.AllowedAttributeOperation))
        }
      }

      $schema.Types.Add($schemaType)
    }

    return $schema
  }
}


function Get-xADSyncPSConnectorSetting
{
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [Alias('InputObject')]
    [System.Collections.ObjectModel.KeyedCollection[string, Microsoft.MetadirectoryServices.ConfigParameter]]
    $ConfigurationParameters,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Name,
    [Parameter(Mandatory = $true)]
    [ValidateSet('Global','Partition','RunStep')]
    [string]
    $Scope,
    $DefaultValue
  )
  process
  {
    try
    {
      $scopedName = '{0}_{1}' -f $Name,$Scope

      if ($ConfigurationParameters[$scopedName].Value)
      {
        return $ConfigurationParameters[$scopedName].Value
      }
      elseif ($PSBoundParameters.ContainsKey('DefaultValue'))
      {
        return $DefaultValue
      }
      else
      {
        return $null
      }
    }
    catch [System.Collections.Generic.KeyNotFoundException]
    {
      # if they gave us a default, go ahead and return it
      if ($PSBoundParameters.ContainsKey('DefaultValue'))
      {
        return $DefaultValue
      }
      else
      {
        throw
      }
    }
  }
}

function Get-xADSyncPSConnectorFolder
{
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter(Mandatory = $true,Position = 0)]
    [ValidateSet('ManagementAgent','Extensions')]
    [string]
    $Folder
  )

  switch ($Folder)
  {
    'ManagementAgent'
    {
      return [Microsoft.MetadirectoryServices.MAUtils]::MAFolder
    }
    'Extensions'
    {
      return [Microsoft.MetadirectoryServices.Utils]::ExtensionsDirectory
    }
    default
    {
      throw "Folder '$Folder' is not supported"
    }
  }
}

#region Schema Helpers
function New-xADSyncPSConnectorSchema
{
  [CmdletBinding()]
  [OutputType([Microsoft.MetadirectoryServices.Schema])]
  param()

  return [Microsoft.MetadirectoryServices.Schema]::Create()
}

function New-xADSyncPSConnectorSchemaType
{
  [CmdletBinding()]
  [OutputType([Microsoft.MetadirectoryServices.SchemaType])]
  param(
    [ValidateNotNullOrEmpty()]
    [string]
    $Name,
    [switch]
    $LockAnchorAttributeDefinition
  )

  return [Microsoft.MetadirectoryServices.SchemaType]::Create($Name,$LockAnchorAttributeDefinition.ToBool())
}

function Add-xADSyncPSConnectorSchemaAttribute
{
  [CmdletBinding(DefaultParameterSetName = 'Singlevalued')]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $true)]
    [Microsoft.MetadirectoryServices.SchemaType]
    [ValidateNotNull()]
    $InputObject,
    [ValidateNotNullOrEmpty()]
    [string]
    [Parameter(Mandatory = $true,ParameterSetName = 'Anchor')]
    [Parameter(Mandatory = $true,ParameterSetName = 'Multivalued')]
    [Parameter(Mandatory = $true,ParameterSetName = 'Singlevalued')]
    $Name,
    [Parameter(ParameterSetName = 'Anchor')]
    [switch]
    $Anchor,
    [Parameter(ParameterSetName = 'Multivalued')]
    [switch]
    $Multivalued,
    [Parameter(Mandatory = $true,ParameterSetName = 'Anchor')]
    [Parameter(Mandatory = $true,ParameterSetName = 'Multivalued')]
    [Parameter(Mandatory = $true,ParameterSetName = 'Singlevalued')]
    [ValidateSet('Binary','Boolean','Integer','Reference','String')]
    [string]
    $DataType,
    [Parameter(Mandatory = $true,ParameterSetName = 'Anchor')]
    [Parameter(Mandatory = $true,ParameterSetName = 'Multivalued')]
    [Parameter(Mandatory = $true,ParameterSetName = 'Singlevalued')]
    [ValidateSet('ImportOnly','ExportOnly','ImportExport')]
    [string]
    $SupportedOperation
  )

  process
  {
    switch ($PSCmdlet.ParameterSetName)
    {
      'Singlevalued'
      {
        $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute($Name,$DataType,$SupportedOperation))
      }
      'Multivalued'
      {
        if ($Multivalued.ToBool() -eq $true)
        {
          $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateMultiValuedAttribute($Name,$DataType,$SupportedOperation))
        }
        else
        {
          $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute($Name,$DataType,$SupportedOperation))
        }
      }
      'Anchor'
      {
        if ($Anchor.ToBool() -eq $true)
        {
          $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateAnchorAttribute($Name,$DataType,$SupportedOperation))
        }
        else
        {
          $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute($Name,$DataType,$SupportedOperation))
        }
      }
      default
      {
        throw "Parameter set '$($PSCmdlet.ParameterSetName)' is not supported"
      }
    }
  }
}
#endregion

#region Partition Helpers
function New-FIMPSConnectorPartition
{
  [CmdletBinding()]
  [OutputType([Microsoft.MetaDirectoryServices.Partition])]
  param(
    [Parameter(Mandatory = $true)]
    [guid]
    $Identifier,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DistinguishedName,
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DisplayName
  )

  if ($PSBoundParameters.ContainsKey('DisplayName'))
  {
    return [Microsoft.MetadirectoryServices.Partition]::Create($Identifier,$DistinguishedName,$DisplayName)
  }
  else
  {
    return [Microsoft.MetadirectoryServices.Partition]::Create($Identifier,$DistinguishedName)
  }
}
#endregion
function New-xADSyncPSConnectorHierarchyNode
{
  [CmdletBinding()]
  [OutputType([Microsoft.MetadirectoryServices.HierarchyNode])]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DistinguishedName,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DisplayName
  )

  return [Microsoft.MetadirectoryServices.HierarchyNode]::Create($DistinguishedName,$DisplayName)
}
#region Hierarchy Helpers

#endregion

#region Import Helpers
function New-xADSyncPSConnectorCSEntryChange
{
  [CmdletBinding()]
  [OutputType([Microsoft.MetadirectoryServices.CSEntryChange])]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $ObjectType,
    [Parameter(Mandatory = $true)]
    [ValidateSet('Add','Delete','Update','Replace','None')]
    [string]
    $ModificationType,
    [ValidateNotNullOrEmpty()]
    [Alias('DistinguishedName')]
    [string]
    $DN,
    [ValidateNotNullOrEmpty()]
    [Alias('RelativeDistinguishedName')]
    [string]
    $RDN
  )

  $csEntry = [Microsoft.MetadirectoryServices.CSEntryChange]::Create()
  $csEntry.ObjectModificationType = $ModificationType
  $csEntry.ObjectType = $ObjectType

  if ($PSBoundParameters.ContainsKey('DN'))
  {
    $csEntry.DN = $DN
  }

  if ($PSBoundParameters.ContainsKey('RDN'))
  {
    $csEntry.RDN = $RDN
  }

  Write-Output $csEntry
}

function Add-xADSyncPSConnectorCSAttribute
{
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $true)]
    [Microsoft.MetadirectoryServices.CSEntryChange]
    [ValidateNotNull()]
    $InputObject,
    [Parameter(Mandatory = $true)]
    [ValidateSet('Add','Update','Delete','Replace','Rename')]
    [string]
    $ModificationType,
    [ValidateNotNullOrEmpty()]
    [string]
    $Name,
    $Value
  )

  process
  {
    if ($ModificationType -ne 'Rename' -and $Name -eq $null)
    {
      throw 'Name parameter is required'
    }

    if ($ModificationType -ne 'Delete' -and $Value -eq $null)
    {
      throw 'Value parameter is required'
    }

    switch ($ModificationType)
    {
      'Add'
      {
        $InputObject.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeAdd($Name,$Value))
      }
      'Update'
      {
        $InputObject.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeUpdate($Name,$Value))
      }
      'Delete'
      {
        $InputObject.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeDelete($Name))
      }
      'Replace'
      {
        $InputObject.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeReplace($Name,$Value))
      }
      'Rename'
      {
        $InputObject.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateNewDN($Value))
      }
      default
      {
        throw "Modification type $ModificationType is not supported"
      }
    }
  }
}
#endregion

function New-GenericObject
{
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $TypeName,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $TypeParameters,
    [Parameter(Mandatory = $false)]
    [object[]]
    $ConstructorParameters
  )

  $genericTypeName = $typeName + '
              `r`n' + $typeParameters.Count
  $genericType = [type]$genericTypeName

  if (!$genericType)
  {
    throw "Could not find generic type $genericTypeName"
  }

  ## Bind the type arguments to it
  $typedParameters = [Type[]]$typeParameters
  $closedType = $genericType.MakeGenericType($typedParameters)

  if (!$closedType)
  {
    throw "Could not make closed type $genericType"
  }

  ## Create the closed version of the generic type. Don't forget comma prefix
  ,[Activator]::CreateInstance($closedType,$constructorParameters)
}

Export-ModuleMember -Function * -Verbose:$false -Debug:$false