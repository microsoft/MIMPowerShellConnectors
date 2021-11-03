param(
    # Table of all configuration parameters set at instance of connector
    [System.Collections.ObjectModel.KeyedCollection[[string], [Microsoft.MetadirectoryServices.ConfigParameter]]]
    $ConfigParameters,

    # Contains any credentials entered by the administrator on the Connectivity tab
    [PSCredential]
    $PSCredential,

    # Filename of Template Input file, please use unique filename for each connector.
    [String]
    $TemplateInputFile = 'MAName_SampleInputFile.txt',

    # Delimiter used in Template Input File
    [String]
    $Delimiter = ';',

    # Encoding used in Template Input File
    [String]
    $Encoding = 'UTF8',

    # Used as objecttype when no complex schema is used.
    [String]
    $DefaultObjectTypeName = 'Object',

    # Path to debug log
    [String]
    $LogFilePath = "$([System.Environment]::GetEnvironmentVariable('Temp', 'Machine'))\MIMPS_SchemaScript.log"
)

Set-PSDebug -Strict
function Write-Log {
    <#
    .SYNOPSIS
    Function for logging, modify to suit your needs.
    #>
    [CmdletBinding()]
    param([string]$Message, [String]$Path)
    # Uncomment this line to enable debug logging
    # Out-File -InputObject $Message -FilePath $Path -Append
}

try {
    # Remove log if exists:
    Remove-Item -Path $LogFilePath -Force -ErrorAction 'Stop'
} catch {
    # We don't care about errors here
}

$PSDefaultParameterValues['Write-Log:Path'] = $LogFilePath

try {
    $commonModule = Join-Path -Path ([System.Environment]::GetEnvironmentVariable('Temp', 'Machine')) -ChildPath $ConfigParameters['Common Module Script Name (with extension)'].Value
    Import-Module -Name $commonModule -Verbose:$false -ErrorAction Stop
    Write-Log -Message 'CommonModule imported'
} catch {
    throw "Failed to import common module with error [$_]"
}

$FolderPath = Get-xADSyncPSConnectorFolder -Folder 'Extensions'
Write-Log -Message 'ExtensionFolder found'
$TemplateFile = Join-Path -Path $FolderPath -ChildPath $TemplateInputFile

if ((Test-Path -Path $TemplateFile -PathType 'Leaf') -eq $false) {
    $Message =  "Cannot find template file $TemplateFile"
    Write-Log -Message  "Cannot find template file $TemplateFile"
    throw $Message
}

Write-Log -Message "Importing template input file: $TemplateFile"
$CSVParams = @{
    Path = $TemplateFile
    Delimiter = $Delimiter
    Encoding = $Encoding
}
Write-Log -Message ($CSVParams | ConvertTo-Json)
$TemplateCSV = Import-Csv @CSVParams
Write-Log -Message 'File imported'
if ($null -eq $TemplateCSV) {
    throw 'Imported CSV is null'
}

Write-Log -Message 'Creating schema'
$Schema = New-xADSyncPSConnectorSchema
Write-Log -Message 'Schema created'

$SchemaType = $TemplateCSV | Get-Member | Select-Object -ExpandProperty 'TypeName' -First 1
Write-Log -Message "SchemaType: $SchemaType"

if($SchemaType -eq 'CSV:MIMPSConnector.SchemaConfig') {
    Write-Log -Message 'Complex Schema'
    foreach($SchemaType in $TemplateCSV) {
        $ObjectType = $SchemaType | Select-Object -ExpandProperty 'ObjectType'
        Write-Log -Message "CreatingType: $ObjectType"
        $FileName = $SchemaType | Select-Object -ExpandProperty 'FileName'
        $FilePath = Join-Path -Path $FolderPath -ChildPath $FileName
        $ReferenceAttributes = @{}
        ($SchemaType | Select-Object -ExpandProperty 'ReferenceAttributes') -split '\|' | Foreach-Object {
            $Key,$Value = $_ -split '#'
            $ReferenceAttributes.Add($Key,$Value)
        }
        $SchemaObjectCSV = Import-Csv -Path $FilePath -Delimiter $Delimiter -Encoding $Encoding
        $Columns = $SchemaObjectCSV[0] | Get-Member -MemberType 'NoteProperty'
        $SchemaTypeObject = New-xADSyncPSConnectorSchemaType -Name $ObjectType
        foreach ($c in $Columns) {
            if($c.Name -in $ReferenceAttributes.Keys) {
                $DataType = 'Reference'
            } else {
                $DataType = 'String'
            }
            Add-xADSyncPSConnectorSchemaAttribute -Name $c.Name -SupportedOperation 'ImportExport' -DataType $DataType -InputObject $SchemaTypeObject
        }
        Write-Log -Message 'Adding Schematypeobject:'
        Write-Log -Message ($SchemaTypeObject|ConvertTo-Json)
        $Schema.Types.Add($SchemaTypeObject)
        Write-Log -Message 'Object Added'
    }
} else {
    Write-Log -Message 'Simple Schema'
    $Columns = $TemplateCSV[0] | Get-Member -MemberType 'NoteProperty'
    Write-Log -Message ($Columns | ConvertTo-Json)
    Write-Log -Message $DefaultObjectTypeName
    $SchemaTypeObject = New-xADSyncPSConnectorSchemaType -Name $DefaultObjectTypeName
    foreach ($c in $Columns) {
        Add-xADSyncPSConnectorSchemaAttribute -Name $c.Name -SupportedOperation 'ImportExport' -DataType 'String' -InputObject $SchemaTypeObject
    }
    Write-Log -Message 'Adding Schematypeobject'
    Write-Log -Message ($SchemaTypeObject|ConvertTo-Json)
    $Schema.Types.Add($SchemaTypeObject)
    Write-Log -Message 'Object Added'
}
Write-Log -Message ($Schema | ConvertTo-Json)
Write-Output $Schema
