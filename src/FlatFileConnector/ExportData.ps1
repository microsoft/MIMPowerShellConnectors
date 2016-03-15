param(    
    [System.Collections.ObjectModel.KeyedCollection[[string], [Microsoft.MetadirectoryServices.ConfigParameter]]]
    $ConfigParameters,
    [Microsoft.MetadirectoryServices.Schema]
    $Schema,
    [Microsoft.MetadirectoryServices.OpenExportConnectionRunStep]
    $OpenExportConnectionRunStep,
    [System.Collections.Generic.IList[Microsoft.MetaDirectoryServices.CSEntryChange]]
    $CSEntries,
    [PSCredential]
    $PSCredential
)
Set-PSDebug -Strict

Import-Module (Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath 'xADSyncPSConnectorModule.psm1') -Verbose:$false

function CreateCustomPSObject
{
    param
    (
        $PropertyNames = @()
    )
    $template = New-Object -TypeName System.Object

    foreach ($property in $PropertyNames)
    {
        $template | Add-Member -MemberType NoteProperty -Name $property -Value $null
    }

    return $template
}

$exportCsvParameters = @{
    Path = (Join-Path -Path (Get-xADSyncPSConnectorFolder -Folder ManagementAgent) -ChildPath (Get-xADSyncPSConnectorSetting -Name 'FileName' -Scope Global -ConfigurationParameters $ConfigParameters))
}

$csentryChangeResults = New-GenericObject System.Collections.Generic.List Microsoft.MetadirectoryServices.CSEntryChangeResult

if ((Test-Path ([IO.Path]::GetDirectoryName($exportCsvParameters['Path'])) -PathType Container) -eq $false)
{
    ##TODO: ECMA exception?
    throw "Could not find $($exportCsvParameters['Path'])"   
}

Write-Verbose "Export path: $($exportCsvParameters['Path'])"

$delimiter = Get-xADSyncPSConnectorSetting -Name 'Delimiter' -Scope Global -ConfigurationParameters $ConfigParameters
if ($delimiter)
{
    $exportCsvParameters.Add('Delimiter', $delimiter)
    Write-Verbose "Setting delimiter to $delimiter)"
}

$encoding = Get-xADSyncPSConnectorSetting -Name 'Encoding' -Scope Global -ConfigurationParameters $ConfigParameters
if ($encoding)
{
    ##TODO: Validation
    $exportCsvParameters.Add('Encoding', $encoding)
    Write-Verbose "Setting encoding to $encoding)"
}

$columnsToExport = @()
foreach ($attribute in $Schema.Types[0].Attributes)
{
    $columnsToExport += $attribute.Name
    Write-Verbose "Added attribute $($attribute.Name) to export list"
}

Write-Verbose "Loaded $($columnsToExport.Count) attributes to export" 

$csvSource = @()
foreach ($entry in $CSEntries)
{
    Write-Verbose "Processing object $($entry.Identifier)"

    [bool]$objectHasAttributes = $false
    $baseObject = CreateCustomPSObject -PropertyNames $columnsToExport

    if ($entry.ModificationType -ne 'Delete')
    {
        foreach ($attribute in $columnsToExport)
        {                              
            if (($entry.AttributeChanges.Contains($attribute)) -eq $false -and ($entry.AnchorAttributes.Contains($attribute) -eq $false))
            {
                continue
            }
            
            if ($entry.AnchorAttributes[$attribute].Value)
            {
                $baseObject.$attribute = $entry.AnchorAttributes[$attribute].Value
                $objectHasAttributes = $true
            }
            elseif ($entry.AttributeChanges[$attribute].ValueChanges[0].Value)
            {
                $baseObject.$attribute = $entry.AttributeChanges[$attribute].ValueChanges[0].Value
                $objectHasAttributes = $true
            }            
        }

        if ($objectHasAttributes)
        {
            $csvSource += $baseObject
        }
$csentryChangeResult = [Microsoft.MetadirectoryServices.CSEntryChangeResult]::Create($entry.Identifier, $null, "Success")
    } 

    $csentryChangeResults.Add($csentryChangeResult) 
    Write-Verbose "Completed processing object $($entry.Identifier)"   
}

$csvSource | Export-Csv @exportCsvParameters -NoTypeInformation

$result = New-Object -TypeName Microsoft.MetadirectoryServices.PutExportEntriesResults

$closedType = [Type] "Microsoft.MetadirectoryServices.PutExportEntriesResults"
return [Activator]::CreateInstance($closedType, $csentryChangeResults) 