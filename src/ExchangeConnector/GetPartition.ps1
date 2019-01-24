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
    Import-Module -Name $commonModule -Verbose:$false -ErrorAction 'Stop'
    Write-Log -Message 'CommonModule imported'
} catch {
    throw "Failed to import common module with error [$_]"
}

$ForestFQDN = Get-xADSyncPSConnectorSettingNearest -Name 'ForestFQDN' -ConfigurationParameters $ConfigParameters -DefaultValue $null
$UserName = $Credential.UserName
$Password = $Credential.GetNetworkCredential().Password

if($null -eq $ForestFQDN) {
    $RootDSE = [ADSI]"LDAP://RootDSE"
    $PartitionList = $RootDSE.namingContexts
} else {
    $ArgumentList = "LDAP://$ForestFQDN/RootDSE", $UserName, $Password
    $DirectoryEntry = New-Object -TypeName 'System.DirectoryServices.DirectoryEntry' -ArgumentList $ArgumentList
    $PartitionList = $DirectoryEntry.namingContexts
}

$Partitions = New-GenericObject System.Collections.Generic.List Microsoft.MetadirectoryServices.Partition

foreach($PartitionEntry in $PartitionList) {
    $Identifier = [System.Guid]::NewGuid()
    $dn = $PartitionEntry
    $Partition = [Microsoft.MetadirectoryServices.Partition]::Create($Identifier, $dn)
    if($dn -match '^CN=Configuration,|^CN=Schema,|^DC=DomainDnsZones,|^DC=ForestDnsZones,') {
        $Partition.HiddenByDefault = $true
    } else {
        $Partition.HiddenByDefault = $false
    }
    $null = $Partitions.Add($Partition)
}

return ,$Partitions
