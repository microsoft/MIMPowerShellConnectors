param( 
    [System.Collections.ObjectModel.KeyedCollection[[string], [Microsoft.MetadirectoryServices.ConfigParameter]]]  
    $ConfigParameters,
    [PSCredential]
    $PSCredential
)

Set-PSDebug -Strict

Import-Module (Join-Path -Path ([Environment]::GetEnvironmentVariable('TEMP', [EnvironmentVariableTarget]::Machine)) -ChildPath 'xADSyncPSConnectorModule.psm1') -Verbose:$false

$TemplateFile = (Join-Path -Path (Get-xADSyncPSConnectorFolder -Folder Extensions) -ChildPath 'SampleInputFile.txt')
$Delimiter = ';'
$Encoding = 'Default'

if ((Test-Path -Path $TemplateFile -PathType Leaf) -eq $false)
{
    throw "Cannot find template file $TemplateFile"
}

$TemplateCSV = Import-Csv -Path $TemplateFile -Delimiter $Delimiter -Encoding $Encoding
if ($TemplateCSV -eq $null)
{
    throw 'Imported CSV is null'
}

$Schema = New-xADSyncPSConnectorSchema

$SchemaType = New-xADSyncPSConnectorSchemaType -Name 'Row'

$Columns = $TemplateCSV[0] | Get-Member -MemberType NoteProperty

foreach ($c in $Columns)
{
    $SchemaType | Add-xADSyncPSConnectorSchemaAttribute -Name $c.Name -DataType String -SupportedOperation ImportExport
}

$Schema.Types.Add($SchemaType)

Write-Output $Schema
