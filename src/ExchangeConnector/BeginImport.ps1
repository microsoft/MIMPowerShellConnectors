[CmdletBinding()]
param(
    # Table of all configuration parameters set at instance of connector
	[parameter(Mandatory = $true)]
	[System.Collections.ObjectModel.KeyedCollection[string,Microsoft.MetadirectoryServices.ConfigParameter]]
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
	[parameter(Mandatory = $true)]
	[Alias('PSCredential')]
	[System.Management.Automation.PSCredential]
    $Credential,

    # Optional parameter for manipulation by the TestHarness script
	[parameter(Mandatory = $false)]
	[ValidateScript({ Test-Path $_ -PathType 'Container' })]
	[string]
	$ScriptDir = [Microsoft.MetadirectoryServices.MAUtils]::MAFolder
)
if($OpenImportConnectionRunStep.ImportType -eq 'Full') {
    $waterMarkXml = "<WaterMark>"
    $waterMarkXml += "<StartIndex>0</StartIndex><DirSyncCookie></DirSyncCookie>"

    foreach ($type in $Schema.Types) {
        $waterMarkXml += "<{0}><MoreToImport>1</MoreToImport></{0}>" -f $type.Name
    }
    $waterMarkXml += "</WaterMark>"
    $waterMark = [xml]$waterMarkXml
} else {
    $waterMark = [xml]$OpenImportConnectionRunStep.CustomData
    if ($null -eq $waterMark -or $null -eq $waterMark.WaterMark) {
        throw ("Invalid Watermark. Please run Full Import first.")
    }
    $waterMark.WaterMark.StartIndex = "0"
    foreach ($type in $Schema.Types) {
        $waterMark.WaterMark.($type.Name).MoreToImport = "1"
    }
}
$OpenImportConnectionResults = New-Object Microsoft.MetadirectoryServices.OpenImportConnectionResults($Watermark.InnerXml)
Write-Output -InputObject $OpenImportConnectionResults