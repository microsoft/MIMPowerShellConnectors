param(
    # Table of all configuration parameters set at instance of connector
    [System.Collections.ObjectModel.KeyedCollection[[string], [Microsoft.MetadirectoryServices.ConfigParameter]]]
    $ConfigParameters,

    # Representation of Schema
    [Microsoft.MetadirectoryServices.Schema]
    [ValidateNotNull()]
    $Schema,

    # Informs the script about the type of import run (delta or full), partition, hierarchy, watermark, and expected page size.
    [Microsoft.MetadirectoryServices.OpenImportConnectionRunStep]
    $OpenImportConnectionRunStep,

    # Holds the watermark (CustomData) that can be used during paged imports and delta imports.
    [Microsoft.MetadirectoryServices.ImportRunStep]
    $GetImportEntriesRunStep,

    # Contains any credentials entered by the administrator on the Connectivity tab
    [Alias('PSCredential')]
    [System.Management.Automation.PSCredential]
    $Credential,

    # Path to debug log
    [String]
    $LogFilePath = "$([System.Environment]::GetEnvironmentVariable('Temp', 'Machine'))\MIMPS_ImportScript.log"
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
Write-Log -Message "Import Started at: $(Get-Date -Format 'o')"
# Uncomment below lines to log details about input data
# Write-Log -Message 'ConfigParameters:'
# Write-Log -Message ($ConfigParameters | ConvertTo-Json)
# Write-Log -Message 'OpenImportConnectionRunStep:'
# Write-Log -Message ($OpenImportConnectionRunStep | ConvertTo-Json)
# Write-Log -Message 'GetImportEntriesRunStep:'
# Write-Log -Message ($GetImportEntriesRunStep | ConvertTo-Json)
# Write-Log -Message 'Schema:'
# Write-Log -Message ($Schema | ConvertTo-Json -Depth 6)

try {
    $commonModule = (Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath $ConfigParameters['Common Module Script Name (with extension)'].Value)
    Import-Module -Name $commonModule -Verbose:$false -ErrorAction Stop
    Write-Log -Message "CommonModule imported"
} catch {
    $Message = "Failed to import common module with error [$_]"
    Write-Log -Message $Message
    throw $Message
}

$MoreToImport = 1
$PageSize = $OpenImportConnectionRunStep.PageSize

$MAFolderPath = Get-xADSyncPSConnectorFolder -Folder ManagementAgent
$importResults = New-Object -TypeName 'Microsoft.MetadirectoryServices.GetImportEntriesResults'
$csEntries = New-GenericObject -TypeName 'System.Collections.Generic.List' -TypeParameters 'Microsoft.MetadirectoryServices.CSEntryChange'

$ReferenceAttributeData = $getImportEntriesRunStep.CustomData | ConvertFrom-Json

if($OpenImportConnectionRunStep.StepPartition.DN -like 'OBJECT=*') {
    $Type = $Schema.Types | Where-Object -FilterScript {"OBJECT=$($_.Name)" -eq $OpenImportConnectionRunStep.StepPartition.DN}
    $AnchorAttributeNames = $Type.AnchorAttributes.Name
    $IsPartitionImport = $true
} else {
    $Type = $Schema.Types
    $IsPartitionImport = $false
}
$columnsToImport = $Type.Attributes
$ModificationType = 'Add'

Write-Log -Message "Processing $($Type.Name)"

try {
    # If this is first page, $Global:StreamReader does not exist and this will throw
    $null = Get-Variable -Scope Global -Name StreamReader -ErrorAction Stop
    Write-Log -Message 'Not first page, using existing stream reader'
} catch {
    # This is first page, create global stream reader
    Write-Log -Message "This is first page, creating stream reader for $($Type.Name)"

    $FileName = Get-xADSyncPSConnectorSettingNearest -Name 'FileName' -ConfigurationParameters $ConfigParameters
    $FilePath = Join-Path -Path $MAFolderPath -ChildPath $FileName
    $Encoding = Get-xADSyncPSConnectorSettingNearest -Name 'Encoding' -ConfigurationParameters $ConfigParameters -DefaultValue 'UTF8'
    try {
        $Global:StreamReader = New-Object -TypeName System.IO.StreamReader -ArgumentList $FilePath, $Encoding
        $Global:Delimiter = Get-xADSyncPSConnectorSettingNearest -Name 'Delimiter' -ConfigurationParameters $ConfigParameters -DefaultValue ';'
        # Pattern for splitting on any delimiter that does not have an even number of quotes following them
        # This should solve any problem with text qualifiers, i.e. "this,is",just,"three headers"
        $Pattern = "$Global:Delimiter(?=(?:[^`"]*`"[^`"]*`")*[^`"]*$)"
        $Global:Headers = $Global:StreamReader.ReadLine() -split $Pattern -replace '^"|"$'
    } catch {
        $Message = "Failed to open file: $FilePath with encoding $Encoding with error: $_"
        Write-Log -Message $Message
        throw $Message
    }
}

# Read one page of lines
for($i = 1; $i -le $PageSize; $i++) {
    if($row = $Global:StreamReader.ReadLine()) {
        $CSEntryParam = @{}
        $Entry = $row | ConvertFrom-Csv -Delimiter $Global:Delimiter -Header $Global:Headers
        if($IsPartitionImport) {
            $AnchorValues = foreach($AnchorAttributeName in $AnchorAttributeNames) {
                $Entry.$AnchorAttributeName
            }
            $CN = $AnchorValues -join '+' | Set-EscapeLdapDN

            $CSEntryParam['dn'] = "CN=$CN,$($OpenImportConnectionRunStep.StepPartition.DN)"
        }
        $csEntry = New-xADSyncPSConnectorCSEntryChange -ObjectType $Type.Name -ModificationType $ModificationType @CSEntryParam

        foreach ($Attribute in $Entry.PSObject.Properties.Name) {
            if (($columnsToImport | Where-Object -Property 'Name' -eq $Attribute).DataType -eq 2) {
                # This is a reference attribute, calculate dn
                if($Entry.$Attribute -match '\S') {
                    $ReferenceType = $ReferenceAttributeData.($Type.Name).$Attribute
                    $PartitionDN = "OBJECT=$ReferenceType"
                    $Entry.$Attribute = "CN=$($Entry.$Attribute),$PartitionDN"
                }
            }
            if($Entry.$Attribute -match '\S') {
                Get-CSAttributeChange -InputObject $csEntry -ModificationType $ModificationType -ColumnsToImport $columnsToImport -Attribute $Attribute -Value $Entry.$Attribute
            }
        }
        if ($csEntry.AttributeChanges) {
            $csEntries.Add($csEntry)
        }
    } else {
        Write-Log -Message 'End of file'
        $i = $PageSize+1
        $MoreToImport = 0
        $Global:StreamReader.Close()
        $Global:StreamReader.Dispose()
    }
}

$importResults.CSEntries = $csEntries
$importResults.MoreToImport = $MoreToImport -gt 0
Write-Log -Message "Returning $($csEntries.Count) objects with more to import: $($importResults.MoreToImport)"
Write-Output $importResults
