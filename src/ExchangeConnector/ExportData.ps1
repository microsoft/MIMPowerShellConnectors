param(
    # Table of all configuration parameters set at instance of connector
    [parameter(Mandatory = $true)]
    [System.Collections.ObjectModel.KeyedCollection[string, Microsoft.MetadirectoryServices.ConfigParameter]]
    $ConfigParameters,

    # Representation of Schema
    [parameter(Mandatory = $true)]
    [Microsoft.MetadirectoryServices.Schema]
    $Schema,

    # Export information passed into the OpenExportConnection method
    [parameter(Mandatory = $true)]
    [Microsoft.MetadirectoryServices.OpenExportConnectionRunStep]
    $OpenExportConnectionRunStep,

    # List of ConnectorSpace Entries to be exported
    [parameter(Mandatory = $true)]
    [System.Collections.Generic.IList[Microsoft.MetadirectoryServices.CSEntryChange]]
    $CSEntries,

    # Contains any credentials entered by the administrator on the Connectivity tab
    [parameter(Mandatory = $true)]
    [Alias('PSCredential')]
    [System.Management.Automation.PSCredential]
    $Credential,

     # Optional parameter for manipulation by the TestHarness script.
    [parameter(Mandatory = $false)]
    [ValidateScript( { Test-Path $_ -PathType "Container" })]
    [string]
    $ScriptDir = [Microsoft.MetadirectoryServices.MAUtils]::MAFolder,

    # Path to debug log
	[String]
    $LogFilePath = "$([System.Environment]::GetEnvironmentVariable('Temp', 'Machine'))\MIMPS_ExportScript.log"
)

function Write-Log
{
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

Write-Log -Message "ExportScript started at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

$commonModule = (Join-Path -Path $ScriptDir -ChildPath $ConfigParameters["Common Module Script Name (with extension)"].Value)
Import-Module -Name $commonModule -Verbose:$false -ErrorAction Stop

# Import modules and connect to connected data source
try {
    try {
        # If this is first run, $Global:ExchSession does not exist and this will throw
        $null = Get-Variable -Scope Global -Name 'ExchSession' -ErrorAction Stop -ValueOnly
        Write-Log -Message 'Not first run, using existing user array'
    } catch {
        # Create session to exchange
        $ExchConnectionUri = Get-xADSyncPSConnectorSettingNearest -Name 'ExchConnectionUri' -ConfigurationParameters $ConfigParameters -DefaultValue $null
        $Global:ExchSession = New-PSSession -ConnectionUri $ExchConnectionUri -ConfigurationName 'Microsoft.Exchange' -Credential $Credential -Authentication 'Negotiate' -AllowRedirection -ErrorAction Stop
        $null = Import-PSSession -Session $Global:ExchSession -DisableNameChecking
    }
} catch {
    throw "Failed to connect to exchange with message: $($_.Exception.Message)"
}

$UserName = $Credential.UserName
$Password = $Credential.GetNetworkCredential().Password
$csentryChangeResults = New-GenericObject -TypeName 'System.Collections.Generic.List' -TypeParameters 'Microsoft.MetadirectoryServices.CSEntryChangeResult'

try {
    foreach ($Entry in $CSEntries) {

        $ChangeResult = $null

        switch ($Entry.ObjectModificationType) {
            'Add' {
                switch ($Entry.ObjectType) {
                    'User' {
                        try {
                            $MailNickname = $Entry.AttributeChanges.Where{$_.Name -eq 'mailNickname'}.ValueChanges.Value
                            $MailboxParams = @{
                                Identity = $Entry.DN
                                Alias = $MailNickname
                            }

                            if(-not [string]::IsNullOrWhiteSpace($ConfigParameters.Server)) {
                                $MailboxParams.Add('DomainController',$ConfigParameters.Server)
                                $LdapDN = "LDAP://$($ConfigParameters.Server)/$($Entry.DN)"
                            } else {
                                $LdapDN = "LDAP://$($Entry.DN)"
                            }
                            $ADUser = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $LdapDN,$UserName, $Password
                            Add-xADSyncPSConnectorCSAttribute -InputObject $Entry -ModificationType 'Add' -Name 'objectGuid' -Value $ADUser.'objectguid'[0]

                            if($Entry.AttributeChanges.Where{$_.Name -eq '_MailboxType'}.ValueChanges.Value -eq 'Mailbox') {
                                $null = Enable-Mailbox @MailboxParams
                            } elseif ($Entry.AttributeChanges.Where{$_.Name -eq '_MailboxType'}.ValueChanges.Value -eq 'RemoteMailbox') {
                                $null = Enable-RemoteMailbox @MailboxParams
                            }
                            $ChangeResult = New-xADEntryChangeResult -Identifier $Entry.Identifier -AttributeChanges $Entry.AttributeChanges -ErrorCode 'Success'

                        } catch {
                            $ChangeResult = New-xADEntryChangeResult -Identifier $Entry.Identifier -AttributeChanges $Entry.AttributeChanges -ErrorCode 'ExportErrorCustomContinueRun' -ErrorName 'script-error' -ErrorDetail "$($_.Exception.Message)"
                        }
                    }
                }
                break
            }

            'Delete' {
                break
            }

            'Replace' {
                # Modification of object
                switch ($Entry.ObjectType) {
                    'User' {
                        try {

                            $Guid = [guid][byte[]]$Entry.AnchorAttributes.Where{$_.Name -eq 'ObjectGuid'}.Value | Foreach-Object -MemberName ToString
                            $MailNickname = $Entry.AttributeChanges.Where{$_.Name -eq 'mailNickname'}.ValueChanges.Value

                            $MailboxParams = @{
                                Identity = $Guid
                                Alias = $MailNickname
                            }

                            if (-not [string]::IsNullOrWhiteSpace($ConfigParameters.Server)) {
                                $MailboxParams.Add('DomainController',$ConfigParameters.Server)
                            }

                            if($Entry.AttributeChanges.Where{$_.Name -eq '_MailboxType'}.ValueChanges.Value -eq 'Mailbox') {
                                # Mailbox type set to Mailbox, enable mailbox
                                $null = Enable-Mailbox @MailboxParams -ErrorAction 'Stop'
                            } elseif ($Entry.AttributeChanges.Where{$_.Name -eq '_MailboxType'}.ValueChanges.Value -eq 'RemoteMailbox') {
                                # Mailbox type set to RemoteMailbox, enable in cloud
                                $null = Enable-RemoteMailbox @MailboxParams -ErrorAction 'Stop'
                            } elseif (-not $Entry.AttributeChanges.Where{$_.Name -eq '_MailboxType'}) {
                                # Mailbox type deleted, disable mailbox
                                if($MailboxParams.ContainsKey('Alias')) { $MailboxParams.Remove('Alias') }
                                $null = Disable-Mailbox @MailboxParams -Confirm:$False -ErrorAction 'Stop'
                            }

                            $ChangeResult = New-xADEntryChangeResult -Identifier $Entry.Identifier -AttributeChanges $Entry.AttributeChanges -ErrorCode 'Success'

                        } catch {
                            try {
                                Write-Log -Message ($_ | ConvertTo-Json -Depth 3 -ErrorAction Stop | Out-String )
                            } catch {
                                Write-Log -Message "MailboxError: Some objects are not serializable to JSON, no error details are shown."
                            }
                            $ChangeResult = New-xADEntryChangeResult -Identifier $Entry.Identifier -AttributeChanges $Entry.AttributeChanges -ErrorCode 'ExportErrorCustomContinueRun' -ErrorName 'script-error' -ErrorDetail "$($_.Exception.Message)"
                        }
                        break
                    }

                    default {
                        $ChangeResult = New-xADEntryChangeResult -Identifier $Entry.Identifier -AttributeChanges $Entry.AttributeChanges -ErrorCode 'ExportErrorCustomContinueRun' -ErrorName 'script-error' -ErrorDetail "Unsupported object type: $($Entry.ObjectType)"
                    }
                }
                break
            }

            Default {
                Write-Log -Message "Unsupported ObjectModificationType!"
            }
        }

        if ($null -ne $ChangeResult) {
            $csentryChangeResults.Add($ChangeResult)
        }
    }
} catch {
    try {
        Write-Log -Message ($_ | ConvertTo-Json -Depth 3 | Out-String )
    } catch {
        Write-Log -Message "Some objects are not serializable to JSON, no error details are shown."
    }
}

$closedType = [Type] "Microsoft.MetadirectoryServices.PutExportEntriesResults"
return [Activator]::CreateInstance($closedType, $csentryChangeResults)