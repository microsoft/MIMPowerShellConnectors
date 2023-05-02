param(
  [System.Collections.ObjectModel.KeyedCollection[string, Microsoft.MetadirectoryServices.ConfigParameter]]
  [ValidateNotNull()]
  $ConfigParameters,
  [Microsoft.MetadirectoryServices.Schema]
  [ValidateNotNull()]
  $Schema,
  [Microsoft.MetadirectoryServices.OpenImportConnectionRunStep]
  $OpenImportConnectionRunStep,
  [Microsoft.MetadirectoryServices.ImportRunStep]
  $GetImportEntriesRunStep,
  [pscredential]
  $PSCredential
)

Set-PSDebug -Strict

$commonModule = (Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath $ConfigParameters['Common Module Script Name (with extension)'].Value)
Import-Module -Name $commonModule -Verbose:$false -ErrorAction Stop

$importResults = New-Object -TypeName 'Microsoft.MetadirectoryServices.GetImportEntriesResults'

$csEntries = New-Object -TypeName 'System.Collections.Generic.List[Microsoft.MetadirectoryServices.CSEntryChange]'

$columnsToImport = $Schema.Types[0].Attributes

Write-Verbose "Loaded $($columnsToImport.Count) attributes to import"

$importCsvParameters = @{

  Path = (Join-Path -Path (Get-xADSyncPSConnectorFolder -Folder ManagementAgent) -ChildPath (Get-xADSyncPSConnectorSetting -Name 'FileName' -Scope Global -ConfigurationParameters $ConfigParameters))

}

if ((Test-Path $importCsvParameters['Path'] -PathType Leaf) -eq $false)

{

  ##TODO: ECMA exception?

  throw "Could not find $($importCsvParameters['Path'])"

}

Write-Verbose "Import path: $($importCsvParameters['Path'])"

$delimiter = Get-xADSyncPSConnectorSetting -Name 'Delimiter' -Scope Global -ConfigurationParameters $ConfigParameters

if ($delimiter)

{

  $importCsvParameters.Add('Delimiter',$delimiter)

  Write-Verbose "Setting delimiter to $delimiter)"

}

$encoding = Get-xADSyncPSConnectorSetting -Name 'Encoding' -Scope Global -ConfigurationParameters $ConfigParameters

if ($encoding)

{

  ##TODO: Validation

  $importCsvParameters.Add('Encoding',$encoding)

  Write-Verbose "Setting encoding to $encoding)"

}

$recordsToImport = Import-Csv @importCsvParameters

Write-Verbose "Imported $($recordsToImport.Count) records"

foreach ($record in $recordsToImport)

{

  Write-Verbose 'Starting new record'

  ##TODO: Handle a missing anchor (what exception to throw?)

  $foundValidColumns = $false

  $entrySchema = $Schema.Types[0];
  $csEntry = New-xADSyncPSConnectorCSEntryChange -ObjectType $entrySchema.Name -ModificationType Add

  foreach ($column in $columnsToImport)

  {

    $columnName = $column.Name

    Write-Verbose "Processing column $columnName"

    if ($record.$columnName)

    {

      Write-Verbose 'Found column'

      $foundValidColumns = $true

      ##TODO: Support multivalue?

      $anchorAttrName = $entrySchema.AnchorAttributes[0].Name
      $value = [string]$record.$columnName

      Write-Verbose "$columnName with value equal $value"


      if ($columnName -eq $anchorAttrName)
      {


        $csEntry.AnchorAttributes.Add([Microsoft.MetadirectoryServices.AnchorAttribute]::Create($columnName,$value))
      }


      $csEntry | Add-xADSyncPSConnectorCSAttribute -ModificationType Add -Name $columnName -Value ([Collections.IList]($record.$columnName.Split(";")))

    }

  }

  if ($foundValidColumns)

  {

    Write-Verbose 'Publishing CSEntryChange'

    $csEntries.Add($csEntry)

  }

  Write-Verbose 'Record completed'

}

##TODO: Support paging

$importResults.CSEntries = $csEntries

$importResults.MoreToImport = $false

Write-Output $importResults
