param(
  [System.Collections.ObjectModel.KeyedCollection[string, Microsoft.MetadirectoryServices.ConfigParameter]]
  $ConfigParameters,
  [Microsoft.MetadirectoryServices.Schema]
  $Schema,
  [Microsoft.MetadirectoryServices.OpenExportConnectionRunStep]
  $OpenExportConnectionRunStep,
  [System.Collections.Generic.IList[Microsoft.MetaDirectoryServices.CSEntryChange]]
  $CSEntries,
  [pscredential]
  $PSCredential
)

Set-PSDebug -Strict


$commonModule = (Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath $ConfigParameters['Common Module Script Name (with extension)'].Value)
Import-Module -Name $commonModule -Verbose:$false -ErrorAction Stop

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


function DeleteFromCsv
{
  param($CsvParameters,[string]$ColumnName,[string]$ColumnValue)

  try
  {
    Write-Verbose "Delete from CSV. File: $($CsvParameters.Path)"

    $csv = Import-Csv -Path $CsvParameters.Path -Delimiter $CsvParameters.Delimiter | Where-Object $ColumnName -NE $ColumnValue

    $csv | Export-Csv -Path $CsvParameters.Path -Delimiter $CsvParameters.Delimiter -NoTypeInformation
  }
  catch
  {
    Write-Error $_.ErrorDetails.Message
  }
}

$exportCsvParameters = @{

  Path = (Join-Path -Path (Get-xADSyncPSConnectorFolder -Folder ManagementAgent) -ChildPath (Get-xADSyncPSConnectorSetting -Name 'FileName' -Scope Global -ConfigurationParameters $ConfigParameters))

}



$csentryChangeResults = New-Object "System.Collections.Generic.List[Microsoft.MetadirectoryServices.CSEntryChangeResult]"

if ((Test-Path ([IO.Path]::GetDirectoryName($exportCsvParameters['Path'])) -PathType Container) -eq $false)

{

  ##TODO: ECMA exception?

  throw "Could not find $($exportCsvParameters['Path'])"

}

Write-Verbose "Export path: $($exportCsvParameters['Path'])"

$delimiter = Get-xADSyncPSConnectorSetting -Name 'Delimiter' -Scope Global -ConfigurationParameters $ConfigParameters

if ($delimiter)

{

  $exportCsvParameters.Add('Delimiter',$delimiter)

  Write-Verbose "Setting delimiter to $delimiter)"

}

$encoding = Get-xADSyncPSConnectorSetting -Name 'Encoding' -Scope Global -ConfigurationParameters $ConfigParameters

if ($encoding)

{

  ##TODO: Validation

  $exportCsvParameters.Add('Encoding',$encoding)

  Write-Verbose "Setting encoding to $encoding)"

}

$columnsToExport = @()

foreach ($attribute in $Schema.Types[0].Attributes)

{

  $columnsToExport += $attribute.Name

  Write-Verbose "Added attribute $($attribute.Name) to export list"

}



$csvSource = @()

Write-Verbose "Processing object $($entry.Identifier)"


foreach ($entry in $CSEntries)

{

  Write-Verbose "Processing object $($entry.Identifier). ObjectModificationType $($entry.ObjectModificationType)"

  [bool]$objectHasAttributes = $false

  $baseObject = CreateCustomPSObject -PropertyNames $columnsToExport

  if ($entry.ObjectModificationType -eq 'Replace')
  {
    $anchorAttributeName = $entry.AnchorAttributes[0].Name;
    $anchorAttributeValue = $entry.AnchorAttributes[0].Value.ToString();
    Write-Verbose "Remove the object with attribute '$($anchorAttributeName)' equals '$($anchorAttributeValue)' before replacing it with new object"
    DeleteFromCsv -CsvParameters $exportCsvParameters -ColumnName $anchorAttributeName -ColumnValue $anchorAttributeValue
  }


  if ($entry.ObjectModificationType -ne 'Delete')

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

        $baseObject.$attribute = ($entry.AttributeChanges[$attribute].ValueChanges | Select-Object -Expand Value) -join ";"

        $objectHasAttributes = $true

      }

    }

    if ($objectHasAttributes)

    {

      foreach ($property in $baseObject.PSObject.Properties)
      { 
          if ($property.Value -eq $null)
          {
              $baseObject.($property.Name) = ""
          }
      }
      
      $csvSource += $baseObject

    }

  }
  else
  {
    $anchorAttributeName = $entry.AnchorAttributes[0].Name;
    $anchorAttributeValue = $entry.AnchorAttributes[0].Value.ToString();
    Write-Verbose "Delete the object with attribute '$($anchorAttributeName)' equals '$($anchorAttributeValue)'"
    DeleteFromCsv -CsvParameters $exportCsvParameters -ColumnName $anchorAttributeName -ColumnValue $anchorAttributeValue
  }

  $csentryChangeResult = [Microsoft.MetadirectoryServices.CSEntryChangeResult]::Create($entry.Identifier,$null,"Success")
  $csentryChangeResults.Add($csentryChangeResult)

  Write-Verbose "Completed processing object $($entry.Identifier)"

}

$csvSource | Export-Csv @exportCsvParameters -NoTypeInformation -Append -Force

$closedType = [type]"Microsoft.MetadirectoryServices.PutExportEntriesResults"

return [Activator]::CreateInstance($closedType,$csentryChangeResults)
