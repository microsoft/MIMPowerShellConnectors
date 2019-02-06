param(
    # Table of all configuration parameters set at instance of connector
    [System.Collections.ObjectModel.KeyedCollection[string,Microsoft.MetadirectoryServices.ConfigParameter]]
    $ConfigParameters,

    # Representation of Schema
    [Microsoft.MetadirectoryServices.Schema]
    $Schema,

    # Informs the script about the partition
    [Microsoft.MetadirectoryServices.OpenExportConnectionRunStep]
    $OpenExportConnectionRunStep,

    # Contains any credentials entered by the administrator on the Connectivity tab
    [Alias('PSCredential')]
    [System.Management.Automation.PSCredential]
    $Credential,

    # Path to debug log
    [String]
    $LogFilePath = "$([System.Environment]::GetEnvironmentVariable('Temp', 'Machine'))\MIMPS_BeginExportScript.log"
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

try {
    $commonModule = (Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath $ConfigParameters['Common Module Script Name (with extension)'].Value)
    Import-Module -Name $commonModule -Verbose:$false -ErrorAction Stop
    Write-Log -Message "CommonModule imported"
} catch {
    throw "Failed to import common module with error [$_]"
}

# Uncomment below lines to log details about input data
# Write-Log -Message 'ConfigParameters:'
# Write-Log -Message ($ConfigParameters | ConvertTo-Json)
# Write-Log -Message 'Schema:'
# Write-Log -Message ($Schema | ConvertTo-Json -Depth 6)
# Write-Log -Message 'OpenExportConnectionRunStep:'
# Write-Log -Message ($OpenExportConnectionRunStep | ConvertTo-Json)

$Global:Delimiter = Get-xADSyncPSConnectorSettingNearest -Name 'Delimiter' -ConfigurationParameters $ConfigParameters -DefaultValue ';'
$MAFolderPath = Get-xADSyncPSConnectorFolder -Folder ManagementAgent
$FileName = Get-xADSyncPSConnectorSettingNearest -Name 'FileName' -ConfigurationParameters $ConfigParameters
$Global:FilePath = Join-Path -Path $MAFolderPath -ChildPath $FileName
$Global:Encoding = Get-xADSyncPSConnectorSettingNearest -Name 'Encoding' -ConfigurationParameters $ConfigParameters -DefaultValue 'UTF8'
try {
    Remove-Item -Path $Global:FilePath -Force -ErrorAction Stop
} catch {
    # We don't care about errors here
}

Write-Log -Message "Export path: $($Global:FilePath)"
if ((Test-Path -Path $MAFolderPath -PathType Container) -eq $false) {
    New-xADEntryChangeResult -Identifier $null -AttributeChanges $null -ErrorCode ExportErrorCustomStopRun -ErrorName 'script-error' -ErrorDetail 'MA Folder not found'
    throw "Could not find MA Folder: $MAFolderPath"
}

# Load schema-script and parse default values of TemplateInputfile, Delimiter and Encoding
# This will be used read TemplateInputfile and determine order of headers in export-file
try {
    # Load SchemaScript for parsing
    $SchemaScriptBlock = [Scriptblock]::Create($ConfigParameters['Schema Script'].Value)

    # Get default value from parameter TemplateInputFile in SchemaScript
    if (-not(
        $TemplateInputFile = $SchemaScriptBlock.Ast.FindAll({
            $args[0] -is [System.Management.Automation.Language.ParameterAst] -and $args[0].Name -like '$TemplateInputFile'
        },$false).DefaultValue.Value
    )) {
        throw 'TemplateInputFile parameter not found in Schema script'
    }
    # Get default value from parameter Delimiter in SchemaScript
    if (-not(
        $TemplateDelimiter = $SchemaScriptBlock.Ast.FindAll({
            $args[0] -is [System.Management.Automation.Language.ParameterAst] -and $args[0].Name -like '$Delimiter'
        },$false).DefaultValue.Value
    )) {
        throw 'Delimiter parameter not found in Schema script'
    }
    # Get default value from parameter Encoding in SchemaScript
    if (-not(
        $TemplateEncoding = $SchemaScriptBlock.Ast.FindAll({
            $args[0] -is [System.Management.Automation.Language.ParameterAst] -and $args[0].Name -like '$Encoding'
        },$false).DefaultValue.Value
    )) {
        throw 'Encoding parameter not found in Schema script'
    }

    $FolderPath = Get-xADSyncPSConnectorFolder -Folder 'Extensions'
    $TemplateFilePath = Join-Path -Path $FolderPath -ChildPath $TemplateInputFile
    if(-not(Test-Path -Path $TemplateFilePath)) {
        throw 'TemplateFile not found'
    }
    $TemplateFileData = Import-Csv -Path $TemplateFilePath -Delimiter $TemplateDelimiter -Encoding $TemplateEncoding -ErrorAction Stop
    $TemplateFileDataType = $TemplateFileData | Get-Member | Select-Object -ExpandProperty 'TypeName' -First 1

    if ($OpenExportConnectionRunStep.StepPartition.DN -like 'OBJECT=*') {
        # Complex, select right type
        $Type = $Schema.Types | Where-Object -FilterScript {"OBJECT=$($_.Name)" -eq $OpenExportConnectionRunStep.StepPartition.DN}
    } else {
        # Not complex, only one type supported
        $Type = $Schema.Types
    }

    if ($TemplateFileDataType -eq 'CSV:MIMPSConnector.SchemaConfig') {
        # Dealing with partitions, need to find right TemplateFile
        Write-Log -Message 'Partitions!'
        if (-not($TemplateInputFile = $TemplateFileData | Where-Object -FilterScript {$_.ObjectType -eq $Type.Name} | Select-Object -ExpandProperty FileName)) {
            $Message = "Failed to locate templatefile for object type: $($Type.Name)"
            Write-Log -Message $Message
            throw $Message
        }
        $TemplateFilePath = Join-Path -Path $FolderPath -ChildPath $TemplateInputFile
        $TemplateFileData = Import-Csv -Path $TemplateFilePath -Delimiter $TemplateDelimiter -Encoding $TemplateEncoding -ErrorAction Stop
    }
    $global:columnsToExport = $TemplateFileData[0].psobject.Properties.Name | Where-Object -FilterScript {$Type.Attributes.Name -eq $_}
} catch {
    # Something failed along the way, just take attributes from schema.
    Write-Log -Message "Failed to get attributes from template file with error: $_"
    $global:columnsToExport = $Type.Attributes.Name
}

Write-Log -Message "Loaded $($columnsToExport.Count) attributes to export"