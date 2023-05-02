[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [System.Collections.ObjectModel.KeyedCollection[string,Microsoft.MetadirectoryServices.ConfigParameter]]
  $ConfigParameters,
  [Parameter(Mandatory = $false)]
  [Alias('PSCredential')] # To fix mess-up of the parameter name in the RTM version of the PowerShell connector.
  [System.Management.Automation.PSCredential]
  $Credential,
  [Parameter(Mandatory = $false)]
  [ValidateScript({ Test-Path $_ -PathType "Container" })]
  [string]
  $ScriptDir = (Join-Path -Path $env:windir -ChildPath "TEMP") # Optional parameter for manipulation by the TestHarness script.
)

Set-StrictMode -Version "2.0"

$Global:DebugPreference = "Continue"
$Global:VerbosePreference = "Continue"

$commonModule = (Join-Path -Path ([System.Environment]::GetEnvironmentVariable('Temp', 'Machine')) -ChildPath $ConfigParameters["Common Module Script Name (with extension)"].Value)

if (!(Get-Module -Name (Get-Item $commonModule).BaseName)) { Import-Module -Name $commonModule }

Enter-Script -ScriptType "Schema" -ErrorObject $Error

function Get-ConnectorSchema
{
<#
    .Synopsis
    Gets the connector space schema.
    .Description
    Gets the connector space schema defined in the "Schema.xml" file.
#>

  [CmdletBinding()]
  [OutputType([Microsoft.MetadirectoryServices.Schema])]
  param(
  )

  $extensionsDir = Get-ExtensionsDirectory
  $schemaXml = Join-Path -Path $extensionsDir -ChildPath "Schema.xml"

  $schema = ConvertFrom-SchemaXml -SchemaXml $schemaXml

  return $schema
}

Get-ConnectorSchema

Exit-Script -ScriptType "Schema" -ErrorObject $Error