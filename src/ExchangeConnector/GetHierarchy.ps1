[CmdletBinding()]
param(
	# DistinguishedName
	[parameter(Mandatory = $true)]
	[Microsoft.MetadirectoryServices.HierarchyNode]
	$HierarchyNode,

	# Table of all configuration parameters set at instance of connector
	[parameter(Mandatory = $true)]
	[System.Collections.ObjectModel.KeyedCollection[string,Microsoft.MetadirectoryServices.ConfigParameter]]
	$ConfigParameters,

	# Contains any credentials entered by the administrator on the Partitions and Hierachies tab
	[parameter(Mandatory = $true)]
	[Alias('PSCredential')]
	[System.Management.Automation.PSCredential]
	$Credential,

	[parameter(Mandatory = $false)]
	[ValidateScript({ Test-Path $_ -PathType "Container" })]
	[string]
	$ScriptDir = [System.Environment]::GetEnvironmentVariable('Temp','Machine'),

	# Path to debug log
	[String]
	$LogFilePath = "$([System.Environment]::GetEnvironmentVariable('Temp', 'Machine'))\MIMPS_HierarchyScript.log"
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

$Containers = New-GenericObject System.Collections.Generic.List Microsoft.MetadirectoryServices.HierarchyNode

$FindParams = @{
	SearchScope = 'OneLevel'
	LDAPFilter  = '(|(objectClass=container)(objectClass=organizationalUnit))'
	SearchBase  = $HierarchyNode.DN
	Property    = 'Name', 'DistinguishedName'
}

if($null -ne $Credential) {
	$FindParams['Credential'] = $Credential
}

if($null -ne $ConfigParameters.Server) {
	$FindParams['Server'] = $ConfigParameters.Server
}

$SearchResult = Find-LDAPObject @FindParams

foreach($Entry in $SearchResult) {
    $null = $Containers.Add(
        [Microsoft.MetadirectoryServices.HierarchyNode]::Create(
            $Entry.DistinguishedName,
            $Entry.Name
        )
    )
}

return ,$Containers
