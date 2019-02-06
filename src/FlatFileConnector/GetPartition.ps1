param(
	# Table of all configuration parameters set at instance of connector
    [System.Collections.ObjectModel.KeyedCollection[[string], [Microsoft.MetadirectoryServices.ConfigParameter]]]
    $ConfigParameters,

    # Contains any credentials entered by the administrator on the Connectivity tab
    [parameter()]
    [Alias('PSCredential')]
    [System.Management.Automation.PSCredential]
    $Credential,

    # Path to debug log
    [String]
    $LogFilePath = "$([System.Environment]::GetEnvironmentVariable('Temp', 'Machine'))\MIMPS_PartitionScript.log"
)

Set-PSDebug -Strict
try {
    # Remove log if exists:
    Remove-Item -Path $LogFilePath -Force -ErrorAction 'Stop'
} catch {
    # We don't care about errors here
}
$PSDefaultParameterValues['Write-Log:Path'] = $LogFilePath
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
    $commonModule = Join-Path -Path ([System.Environment]::GetEnvironmentVariable('Temp', 'Machine')) -ChildPath $ConfigParameters['Common Module Script Name (with extension)'].Value
    Import-Module -Name $commonModule -Verbose:$false -ErrorAction 'Stop'
    Write-Log -Message 'CommonModule imported'
} catch {
    throw "Failed to import common module with error [$_]"
}

try {
    # Import parameter default values from Schema script
    $SchemaScriptBlock = [Scriptblock]::Create($ConfigParameters['Schema Script'].Value)
    if (-not($TemplateInputFile = $SchemaScriptBlock.Ast.FindAll({ $args[0] -is [System.Management.Automation.Language.ParameterAst] -and $args[0].Name -like '$TemplateInputFile' },$false).DefaultValue.Value)) {
        $Message = 'TemplateInputFile parameter not found in Schema script'
        Write-Log -Message $Message
        throw $Message
    }
    if (-not($Delimiter = $SchemaScriptBlock.Ast.FindAll({ $args[0] -is [System.Management.Automation.Language.ParameterAst] -and $args[0].Name -like '$Delimiter' },$false).DefaultValue.Value)) {
        $Message = 'Delimiter parameter not found in Schema script'
        Write-Log -Message $Message
        throw $Message
    }
    if (-not($Encoding = $SchemaScriptBlock.Ast.FindAll({ $args[0] -is [System.Management.Automation.Language.ParameterAst] -and $args[0].Name -like '$Encoding' },$false).DefaultValue.Value)) {
        $Message = 'Encoding parameter not found in Schema script'
        Write-Log -Message $Message
        throw $Message
    }
} catch {
    throw $_
}

$FolderPath = Get-xADSyncPSConnectorFolder -Folder 'Extensions'
$TemplateFile = Join-Path -Path $FolderPath -ChildPath $TemplateInputFile

if ((Test-Path -Path $TemplateFile -PathType 'Leaf') -eq $false) {
    throw "Cannot find template file $TemplateFile"
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
Write-Log -Message 'Schema created'

$SchemaType = $TemplateCSV | Get-Member | Select-Object -ExpandProperty 'TypeName' -First 1
Write-Log -Message "SchemaType: $SchemaType"

if($SchemaType -eq 'CSV:MIMPSConnector.SchemaConfig') {
    Write-Log -Message 'Complex Schema'
    $Partitions = New-GenericObject System.Collections.Generic.List Microsoft.MetadirectoryServices.Partition

    foreach($SchemaType in $TemplateCSV) {
        $ObjectType = $SchemaType | Select-Object -ExpandProperty 'ObjectType'
        $Identifier = [System.Guid]::NewGuid()
        $dn = "OBJECT=$ObjectType"
        $Partition = [Microsoft.MetadirectoryServices.Partition]::Create($Identifier, $dn, $ObjectType)
        $Partition.HiddenByDefault = $false
        $null = $Partitions.Add($Partition)
    }
    return ,$Partitions
}
