<#
<copyright file="Register-EventSource.ps1" company="Microsoft">
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
</copyright>
<summary>
	The script to register event sources for the "ConnectorsLog" Event Log of the FIM PowerShell and all other new connectors.
	The script to needs to run with elevated privileges.
	Please create / edit <system.diagnostics> configuration element as directed in the app.config file.
</summary>
#>

$eventSources = @{
	"ConnectorsLog" = "ConnectorsLog"
}

function TestIsAdministrator 
{ 
	$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent() 
	(New-Object Security.Principal.WindowsPrincipal $currentUser).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator) 
}

function RegisterEventSource()
{
	foreach($source in $eventSources.Keys)
	{
		$logName = $eventSources[$source]
	
		Write-Host "Creating event source $source in event log $logName"
		
		if ([System.Diagnostics.EventLog]::SourceExists($source) -eq $false) {
			New-EventLog -Source $source -LogName $logName
			Write-Host -ForegroundColor green "Event source $source created in event log $logName"
		}
		else
		{
			$eventLog = Get-EventLog -List | Where-Object {$_.Log -eq $logName}

			if ($eventLog -ne $null)
			{
				Write-Host -ForegroundColor yellow "Warning: Event source $source already exists in event log $logName"
			}
			else
			{
				Write-Host -ForegroundColor yellow "Warning: Event source $source already exists, but not in event log $logName. It will be deleted and recreated. You'll need to reboot the machine to see the events in the new event log."
				[System.Diagnostics.EventLog]::DeleteEventSource($source)
				New-EventLog -Source $source -LogName $logName
			}
		}

		Limit-EventLog -LogName $logName -MaximumSize 20480KB

		Write-Host -ForegroundColor green "Writing a test event in the event log '$logName'"

		[System.Diagnostics.EventLog]::WriteEntry($logName, "Test Event")
	}
}

if(!(TestIsAdministrator))  
{
	throw $("Admin rights are required for this script")
}

RegisterEventSource
