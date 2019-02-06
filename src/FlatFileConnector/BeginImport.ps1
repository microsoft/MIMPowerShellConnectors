[CmdletBinding()]
param(
    # Table of all configuration parameters set at instance of connector
    [parameter(Mandatory = $true)]
    [System.Collections.ObjectModel.KeyedCollection[string, Microsoft.MetadirectoryServices.ConfigParameter]]
    $ConfigParameters,

    # Representation of Schema
    [parameter(Mandatory = $true)]
    [Microsoft.MetadirectoryServices.Schema]
    $Schema,

    # Informs the script about the type of import run (delta or full), partition, hierarchy, watermark, and expected page size.
    [parameter(Mandatory = $true)]
    [Microsoft.MetadirectoryServices.OpenImportConnectionRunStep]
    $OpenImportConnectionRunStep,

    # Contains any credentials entered by the administrator on the Connectivity tab
    [parameter()]
    [Alias('PSCredential')]
    [System.Management.Automation.PSCredential]
    $Credential,

    # Path to debug log
    [String]
    $LogFilePath = "$([System.Environment]::GetEnvironmentVariable('Temp', 'Machine'))\MIMPS_BeginImportScript.log"
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
$PSDefaultParameterValues['Write-Log:Path'] = $LogFilePath
Write-Log -Message "BeginImport Started at: $(Get-Date -Format 'o')"

try {
    $commonModule = (Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath $ConfigParameters['Common Module Script Name (with extension)'].Value)
    Import-Module -Name $commonModule -Verbose:$false -ErrorAction Stop
    Write-Log -Message "CommonModule imported"
} catch {
    $Message = "Failed to import common module with error [$_]"
    Write-Log -Message $Message
    throw $Message
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
}
catch {
    throw $_
}

$FolderPath = Get-xADSyncPSConnectorFolder -Folder 'Extensions'
$TemplateFile = Join-Path -Path $FolderPath -ChildPath $TemplateInputFile
Write-Log -Message "Importing template input file: $TemplateFile"
$CSVParams = @{
    Path = $TemplateFile
    Delimiter = $Delimiter
    Encoding = $Encoding
}
Write-Log -Message ($CSVParams | ConvertTo-Json)
$TemplateCSV = Import-Csv @CSVParams
Write-Log -Message 'File imported'
#########################################
if ($null -eq $TemplateCSV) {
    $Message = 'Imported CSV is empty'
    Write-Log -Message $Message
    throw $Message
}

$ReferenceAttributeData = @{}
foreach($Entry in $TemplateCSV) {

    $ObjectType = $Entry | Select-Object -ExpandProperty 'ObjectType'
    Write-Log -Message "CreatingType: $ObjectType"

    $RefAttList = ($Entry | Select-Object -ExpandProperty 'ReferenceAttributes') -split '\|'

    $ReferenceAttributes = @{}
    foreach($RefAtt in $RefAttList) {
        if(-not [string]::IsNullOrWhiteSpace($RefAtt)) {
            $Key,$Value = $RefAtt -split '#'
            if([string]::IsNullOrWhiteSpace($Value)) {
                $Value = $ObjectType
            }
            $ReferenceAttributes.Add($Key,$Value)
        }
    }
    $ReferenceAttributeData.Add($ObjectType,$ReferenceAttributes)
}

$ConnectionResultString = $ReferenceAttributeData | ConvertTo-Json
$OpenImportConnectionResults = New-Object Microsoft.MetadirectoryServices.OpenImportConnectionResults($ConnectionResultString)
Write-Output -InputObject $OpenImportConnectionResults