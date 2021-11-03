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
    [parameter(Mandatory = $true)]
    [Alias('PSCredential')]
    [System.Management.Automation.PSCredential]
    $Credential,

    # Optional parameter for manipulation by the TestHarness script.
    [parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ -PathType "Container" })]
    [string]
    $ScriptDir = [Microsoft.MetadirectoryServices.MAUtils]::MAFolder,

    # Path to debug log
    [String]
    $LogFilePath = "$([System.Environment]::GetEnvironmentVariable('Temp', 'Machine'))\MIMPS_ImportScript.log"
)

try {
    # Initiate script
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

    Write-Log -Message "ImportScript started at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

    $commonModule = (Join-Path -Path $ScriptDir -ChildPath $ConfigParameters["Common Module Script Name (with extension)"].Value)
    Import-Module -Name $commonModule -Verbose:$false -ErrorAction Stop

    $customData = [xml]$getImportEntriesRunStep.CustomData
    $PageSize = $OpenImportConnectionRunStep.PageSize

    if([int]$CustomData.watermark.StartIndex -gt 0) {
        $StartIndex = [int]$CustomData.watermark.StartIndex
    } else {
        $StartIndex = 0
    }

    # Import modules
    try {
        Import-Module -Name $commonModule -Verbose:$false -ErrorAction Stop
    } catch {
        throw 'failed to import modules'
    }

    # Create return objects
    $importResults = New-Object -TypeName 'Microsoft.MetadirectoryServices.GetImportEntriesResults'
    $csEntries = New-GenericObject -TypeName 'System.Collections.Generic.List' -TypeParameters 'Microsoft.MetadirectoryServices.CSEntryChange'

    # Connect to connected data source

    # Loop over object types
    foreach($Type in $Schema.Types) {
        $ObjectType = $Type.Name
        if($customData.WaterMark.$ObjectType.MoreToImport -eq '1') {
            $columnsToImport = $Type.Attributes

            switch($Type.Name) {
                'User' {
                    Write-Log -Message 'Processing Users'
                    try {
                        # If this is first run, $Global:AllUsers does not exist and this will throw
                        $null = Get-Variable -Scope Global -Name AllUsers -ErrorAction Stop -ValueOnly
                        Write-Log -Message 'Not first run, using existing user array'
                    } catch {
                        # Get list of all users
                        Write-Log -Message 'Getting users'

                        $SearchScope = [System.DirectoryServices.SearchScope]::Subtree
                        $UserName = $Credential.UserName
                        $Password = $Credential.GetNetworkCredential().Password
                        $PartitionDN = $OpenImportConnectionRunStep.StepPartition.DN
                        $Server = $ConfigParameters.Server
                        $DeltaPropertiesToLoad = $null
                        if($null -ne $Server) {
                            $SearchRoot = "LDAP://$Server/$PartitionDN"
                        } else {
                            $SearchRoot = "LDAP://$PartitionDN"
                        }

                        $ArgumentList = $SearchRoot, $UserName, $Password
                        $DirectoryEntry = New-Object -TypeName 'System.DirectoryServices.DirectoryEntry' -ArgumentList $ArgumentList

                        $Searcher = New-Object -TypeName 'System.DirectoryServices.DirectorySearcher' -ArgumentList $DirectoryEntry, "(&(objectCategory=person)(objectClass=user))", $DeltaPropertiesToLoad, $SearchScope
                        if($OpenImportConnectionRunStep.ImportType -eq 'Delta') {
                            $Searcher.TombStone = $true
                        }
                        $Searcher.CacheResults = $false

                        if ($null -eq $CustomData.watermark.DirSyncCookie) {
                            $Searcher.directorysynchronization = new-object -TypeName system.directoryservices.directorysynchronization
                        } else {
                            # grab the watermark from last run and pass that to the searcher
                            $DirSyncCookie = ,[System.Convert]::FromBase64String($CustomData.WaterMark.DirSyncCookie)
                            $Searcher.directorysynchronization = new-object -TypeName system.directoryservices.directorysynchronization -ArgumentList $DirSyncCookie
                        }
                        $Global:AllUsers = $Searcher.FindAll()
                        $null = $Global:AllUsers.Count
                        $Global:DirSyncCookieString = [System.Convert]::ToBase64String($Searcher.DirectorySynchronization.GetDirectorySynchronizationCookie())
                    }

                    $EndIndex = [Math]::Min(($Global:AllUsers.Count-1),($StartIndex + $PageSize - 1))

                    Write-Log -Message "Processing users in batch $StartIndex..$EndIndex"

                    # Loop over users in current batch
                    for($i = $StartIndex;$i -le $EndIndex; $i++) {
                        $User = $Global:AllUsers[$i]

                        if($User.Properties.isdeleted) {
                            if($OpenImportConnectionRunStep.ImportType -eq 'Delta') {
                                $csEntry = New-xADSyncPSConnectorCSEntryChange -ObjectType $ObjectType -ModificationType 'Delete' -DN $User.Properties.'distinguishedname'
                                Add-xADSyncPSConnectorCSAnchor -InputObject $csEntry -Name 'objectGuid' -Value $User.Properties.'objectguid'[0]
                                $csEntries.Add($csEntry)
                            }
                        } else {
                            if($OpenImportConnectionRunStep.ImportType -eq 'Delta') {
                                $ModificationType = 'Update'
                            } else {
                                $ModificationType = 'Add'
                            }
                            $csEntry = New-xADSyncPSConnectorCSEntryChange -ObjectType $ObjectType -ModificationType $ModificationType -DN $User.Properties.'distinguishedname'
                            # Process Attributes

                            if($null -ne $User.Properties.'samaccountname') {
                                Get-CSAttributeChange -InputObject $csEntry -ColumnsToImport $columnsToImport -Attribute 'sAMAccountName' -Value $User.Properties.'samaccountname'[0] -ModificationType $ModificationType
                            }

                            if($null -ne $User.Properties.'mailnickname') {
                                Get-CSAttributeChange -InputObject $csEntry -ColumnsToImport $columnsToImport -Attribute 'mailNickname' -Value $User.Properties.'mailnickname'[0] -ModificationType $ModificationType
                            }

                            if($null -ne $User.Properties.'objectguid') {
                                Get-CSAttributeChange -InputObject $csEntry -ColumnsToImport $columnsToImport -Attribute 'objectGuid' -Value $User.Properties.'objectguid'[0] -ModificationType $ModificationType
                            }

                            if($User.Properties.Contains('msexchmailboxguid')) {
                                if($User.Properties.'msexchmailboxguid'.Count -eq 1) {
                                    Get-CSAttributeChange -InputObject $csEntry -ColumnsToImport $columnsToImport -Attribute '_isMailboxEnabled' -Value $true -ModificationType $ModificationType
                                } else {
                                    Get-CSAttributeChange -InputObject $csEntry -ColumnsToImport $columnsToImport -Attribute '_isMailboxEnabled' -Value $false -ModificationType $ModificationType
                                }
                            }

                            if($User.Properties.Contains('msexchrecipienttypedetails')) {
                                if($User.Properties.'msexchrecipienttypedetails'.Count -eq 1) {
                                    $MailboxType = switch($User.Properties.'msexchrecipienttypedetails') {
                                        {$_ -band 0x80000000} {
                                            'RemoteMailbox'
                                            break
                                        }

                                        {$_ -band 1} {
                                            'Mailbox'
                                            break
                                        }
                                    }
                                    Get-CSAttributeChange -InputObject $csEntry -ColumnsToImport $columnsToImport -Attribute '_MailboxType' -Value $MailboxType  -ModificationType $ModificationType
                                } elseif ($OpenImportConnectionRunStep.ImportType -eq 'Delta') {
                                    Get-CSAttributeChange -InputObject $csEntry -ColumnsToImport $columnsToImport -Attribute '_MailboxType' -ModificationType 'Delete'
                                }
                            }

                            $csEntries.Add($csEntry)

                        }

                    }

                    if($EndIndex -lt ($Global:AllUsers.Count - 1)) {
                        $customData.WaterMark.$ObjectType.MoreToImport = '1'
                    } else {
                        $customData.WaterMark.$ObjectType.MoreToImport = '0'
                        Remove-Variable -Scope Global -Name AllUsers -Force -ErrorAction Ignore
                    }
                }
            }
        }
    }

    # Add csEntries to importResults
    $importResults.CSEntries = $csEntries
    $customData.WaterMark.StartIndex = [string]([int]$customData.WaterMark.StartIndex + [int]$PageSize)

    # Check if this is last batch
    $importResults.MoreToImport = $false
    foreach($type in $Schema.Types) {
        $ObjectType = $type.Name
        if($customData.WaterMark.$ObjectType.MoreToImport -eq '1') {
            $importResults.MoreToImport = $true
        }
    }
    if($importResults.MoreToImport -eq $false) {
        $CustomData.WaterMark.DirSyncCookie = $Global:DirSyncCookieString
    }
    $importResults.CustomData = $customData.InnerXml

    Write-Output $importResults
} catch {
    try {
        Write-Log -Message ($_ | Format-Custom | Out-String )
    } catch {
        Write-Log -Message "Some objects are not serializable to JSON, no error details are shown."
    }
}