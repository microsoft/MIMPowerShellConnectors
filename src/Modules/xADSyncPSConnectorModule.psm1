Set-PSDebug -Strict

function Get-xADSyncPSConnectorSetting
{
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [Alias('InputObject')]
        [System.Collections.ObjectModel.KeyedCollection[[string], [Microsoft.MetadirectoryServices.ConfigParameter]]]
        $ConfigurationParameters,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,
        [parameter(Mandatory = $true)]
        [ValidateSet('Global', 'Partition', 'RunStep')]
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
        [parameter(Mandatory = $true, Position = 0)]
        [ValidateSet('ManagementAgent', 'Extensions')]
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
        [Switch]
        $LockAnchorAttributeDefinition
    )

    return [Microsoft.MetadirectoryServices.SchemaType]::Create($Name, $LockAnchorAttributeDefinition.ToBool())
}

function Add-xADSyncPSConnectorSchemaAttribute
{
    [CmdletBinding(DefaultParameterSetName = 'Singlevalued')]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [Microsoft.MetadirectoryServices.SchemaType]
        [ValidateNotNull()]
        $InputObject,
        [ValidateNotNullOrEmpty()]
        [string]
        [parameter(Mandatory = $true, ParameterSetName='Anchor')]
        [parameter(Mandatory = $true, ParameterSetName = 'Multivalued')]
        [parameter(Mandatory = $true, ParameterSetName = 'Singlevalued')]
        $Name,
        [parameter(ParameterSetName='Anchor')]
        [Switch]
        $Anchor,
        [parameter(ParameterSetName = 'Multivalued')]
        [Switch]
        $Multivalued,
        [parameter(Mandatory = $true, ParameterSetName='Anchor')]
        [parameter(Mandatory = $true, ParameterSetName = 'Multivalued')]
        [parameter(Mandatory = $true, ParameterSetName = 'Singlevalued')]
        [ValidateSet('Binary', 'Boolean', 'Integer', 'Reference', 'String')]
        [string]
        $DataType,
        [parameter(Mandatory = $true, ParameterSetName='Anchor')]
        [parameter(Mandatory = $true, ParameterSetName = 'Multivalued')]
        [parameter(Mandatory = $true, ParameterSetName = 'Singlevalued')]
        [ValidateSet('ImportOnly', 'ExportOnly', 'ImportExport')]
        [string]
        $SupportedOperation
    )
    
    process
    {
        switch ($PSCmdlet.ParameterSetName)
        {
            'Singlevalued'
            {
               $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute($Name, $DataType, $SupportedOperation))
            }
            'Multivalued'
            {
                if ($Multivalued.ToBool() -eq $true)
                {
                    $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateMultiValuedAttribute($Name, $DataType, $SupportedOperation))
                }
                else
                {
                    $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute($Name, $DataType, $SupportedOperation))
                }
            }
            'Anchor'
            {
                if ($Anchor.ToBool() -eq $true)
                {
                    $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateAnchorAttribute($Name, $DataType, $SupportedOperation))
                }
                else
                {
                    $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute($Name, $DataType, $SupportedOperation))
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
        [parameter(Mandatory = $true)]
        [Guid]
        $Identifier,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $DistinguishedName,
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]
        $DisplayName
    )

    if ($PSBoundParameters.ContainsKey('DisplayName'))
    {
        return [Microsoft.MetadirectoryServices.Partition]::Create($Identifier, $DistinguishedName, $DisplayName)
    }
    else
    {
        return [Microsoft.MetadirectoryServices.Partition]::Create($Identifier, $DistinguishedName)
    }
}
#endregion
function New-xADSyncPSConnectorHierarchyNode
{
    [CmdletBinding()]
    [OutputType([Microsoft.MetadirectoryServices.HierarchyNode])]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $DistinguishedName,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $DisplayName
    )

    return [Microsoft.MetadirectoryServices.HierarchyNode]::Create($DistinguishedName, $DisplayName)
}
#region Hierarchy Helpers

#endregion

#region Import Helpers
function New-xADSyncPSConnectorCSEntryChange
{
    [CmdletBinding()]
    [OutputType([Microsoft.MetadirectoryServices.CSEntryChange])]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $ObjectType,
        [parameter(Mandatory = $true)]
        [ValidateSet('Add', 'Delete', 'Update', 'Replace', 'None')]
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
        [parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [Microsoft.MetadirectoryServices.CSEntryChange]
        [ValidateNotNull()]
        $InputObject,
        [parameter(Mandatory = $true)]
        [ValidateSet('Add', 'Update', 'Delete', 'Replace', 'Rename')]
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
                $InputObject.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeAdd($Name, $Value))
            }
            'Update'
            {
                $InputObject.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeUpdate($Name, $Value))
            }
            'Delete'
            {
                $InputObject.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeDelete($Name))
            }
            'Replace'
            {
                $InputObject.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeReplace($Name, $Value))
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
    param (
      [parameter(Mandatory = $true)]
      [ValidateNotNullOrEmpty()]
      [string]
      $TypeName,
      [parameter(Mandatory = $true)]
      [ValidateNotNullOrEmpty()]
      [string[]]
      $TypeParameters,
      [parameter(Mandatory = $false)]
      [object[]] 
      $ConstructorParameters
    )

    $genericTypeName = $typeName + '`' + $typeParameters.Count
    $genericType = [Type]$genericTypeName

    if (!$genericType)
    {
      throw "Could not find generic type $genericTypeName"
    }

    ## Bind the type arguments to it
    $typedParameters = [Type[]] $typeParameters
    $closedType = $genericType.MakeGenericType($typedParameters)
    
    if (!$closedType)
    {
     throw "Could not make closed type $genericType"
    }

    ## Create the closed version of the generic type. Don't forget comma prefix
    ,[Activator]::CreateInstance($closedType, $constructorParameters)
}

Export-ModuleMember -Function * -Verbose:$false -Debug:$false 
