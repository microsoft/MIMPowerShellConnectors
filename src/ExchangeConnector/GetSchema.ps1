param(
    # Table of all configuration parameters set at instance of connector
    [System.Collections.ObjectModel.KeyedCollection[[string], [Microsoft.MetadirectoryServices.ConfigParameter]]]
    $ConfigParameters,

    # Contains any credentials entered by the administrator on the Connectivity tab
    [parameter(Mandatory = $true)]
	[Alias('PSCredential')]
	[System.Management.Automation.PSCredential]
	$Credential,

    # Path to debug log
	[String]
    $LogFilePath = "$([System.Environment]::GetEnvironmentVariable('Temp', 'Machine'))\MIMPS_SchemaScript.log",

    # Optional parameter for manipulation by the TestHarness script.
	[parameter(Mandatory = $false)]
	[ValidateScript({ Test-Path $_ -PathType "Container" })]
	[string]
	$ScriptDir = [System.Environment]::GetEnvironmentVariable('Temp','Machine')
)
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

#region Import modules
try {
    $commonModule = Join-Path -Path ([System.Environment]::GetEnvironmentVariable('Temp', 'Machine')) -ChildPath $ConfigParameters['Common Module Script Name (with extension)'].Value
    Import-Module -Name $commonModule -Verbose:$false -ErrorAction Stop
    Write-Log -Message "CommonModule imported"
} catch {
    throw "Failed to import modules"
}
Write-Log -Message "Modules imported OK"
#endregion

#region Initiate Script

Write-Log -Message "Creating new schema object"
$Schema = New-xADSyncPSConnectorSchema

#endregion

#region Building Schema
$UserBaseAttributes = @{
    'dn' = @{Multivalued=$False;DataType='String';SupportedOperation='ImportExport'}
    'mailNickname' = @{Multivalued=$False;DataType='String';SupportedOperation='ImportExport'}
    '_MailboxType' = @{Multivalued=$False;DataType='String';SupportedOperation='ImportExport'}
    '_isMailboxEnabled' = @{Multivalued=$False;DataType='Boolean';SupportedOperation='ImportExport'}
}
foreach($TypeName in 'User') {
    Write-Log -Message "Processing $Type Schema"

    $TypeObject = New-xADSyncPSConnectorSchemaType -Name $TypeName

    Write-Log -Message "Adding Anchor attribute"
    Add-xADSyncPSConnectorSchemaAttribute -Name 'objectGuid' -DataType 'Binary' -SupportedOperation 'ImportExport' -InputObject $TypeObject -Anchor

    Write-Log -Message 'Adding Base attributes'
    foreach($Attribute in $UserBaseAttributes.GetEnumerator()) {
            $Param = $Attribute.Value.Clone()
            if(-Not$Param.Multivalued){$Param.Remove('MultiValued')}
            Add-xADSyncPSConnectorSchemaAttribute -Name $Attribute.Name @Param -InputObject $TypeObject
    }

    $Schema.Types.Add($TypeObject)
}
#endregion

Write-Output $Schema