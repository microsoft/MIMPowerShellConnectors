param(
    # Table of all configuration parameters set at instance of connector
    [System.Collections.ObjectModel.KeyedCollection[[string], [Microsoft.MetadirectoryServices.ConfigParameter]]]
    $ConfigParameters,

    # Representation of Schema
    [Microsoft.MetadirectoryServices.Schema]
    $Schema,

    # Informs the script about the partition
    [Microsoft.MetadirectoryServices.OpenExportConnectionRunStep]
    $OpenExportConnectionRunStep,

    # List of objects to export
    [System.Collections.Generic.IList[Microsoft.MetaDirectoryServices.CSEntryChange]]
    $CSEntries,

    # Contains any credentials entered by the administrator on the Connectivity tab
    [Alias('PSCredential')]
    [System.Management.Automation.PSCredential]
    $Credential,

    # Path to debug log
    [String]
    $LogFilePath = "$([System.Environment]::GetEnvironmentVariable('Temp', 'Machine'))\MIMPS_ExportScript.log"
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
Write-Log -Message "Export Started at: $(Get-Date -Format 'o')"

try {
    $commonModule = (Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath $ConfigParameters['Common Module Script Name (with extension)'].Value)
    Import-Module -Name $commonModule -Verbose:$false -ErrorAction Stop
    Write-Log -Message "CommonModule imported"
} catch {
    throw "Failed to import common module with error [$_]"
}

function New-CustomPSObject {
    param(
        [Parameter(Mandatory)]
        [string[]]
        $PropertyNames
    )
    $template = New-Object -TypeName System.Object
    foreach ($property in $PropertyNames) {
        Add-Member -InputObject $template -MemberType NoteProperty -Name $property -Value $null
    }

    return $template
}

if ($OpenExportConnectionRunStep.StepPartition.DN -like 'OBJECT=*') {
    # Complex
    $Type = $Schema.Types | Where-Object -FilterScript {"OBJECT=$($_.Name)" -eq $OpenExportConnectionRunStep.StepPartition.DN}
} else {
    $Type = $Schema.Types
}

$exportCsvParameters = @{
    Path = $Global:FilePath
    Encoding = $Global:Encoding
    Delimiter = $Global:Delimiter
    NoTypeInformation = $true
    Append = $true
    ErrorAction = 'Stop'
}

$csentryChangeResults = New-GenericObject System.Collections.Generic.List Microsoft.MetadirectoryServices.CSEntryChangeResult

$ExportObjects = foreach ($entry in $CSEntries) {
    $baseObject = New-CustomPSObject -PropertyNames $global:columnsToExport
    if ($entry.ModificationType -ne 'Delete') {
        foreach ($attribute in $global:columnsToExport) {
            if (($entry.AttributeChanges.Contains($attribute)) -eq $false -and ($entry.AnchorAttributes.Contains($attribute) -eq $false)) {
                continue
            }
            if ($entry.AnchorAttributes[$attribute].Value) {
                $baseObject.$attribute = $entry.AnchorAttributes[$attribute].Value
            } elseif ($entry.AttributeChanges[$attribute].ValueChanges[0].Value) {
                $baseObject.$attribute = $entry.AttributeChanges[$attribute].ValueChanges[0].Value
            }
        }
        if ($baseObject.psobject.Properties.Value) {
            Write-Output -InputObject $baseObject
        }
    }
    $csentryChangeResult = [Microsoft.MetadirectoryServices.CSEntryChangeResult]::Create($entry.Identifier, $null, "Success")
    $csentryChangeResults.Add($csentryChangeResult)
}

$ExportObjects | Export-Csv @exportCsvParameters

$closedType = [Type] "Microsoft.MetadirectoryServices.PutExportEntriesResults"
return [Activator]::CreateInstance($closedType, $csentryChangeResults)