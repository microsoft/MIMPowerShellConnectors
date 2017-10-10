<#
<copyright file="ImportScript-Lync.ps1" company="Microsoft">
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
</copyright>
<summary>
	The Main Import script for the Skype 2015 / Lync 2010 / 2013 Connector.
</summary>
#>

[CmdletBinding()]
param(
	[parameter(Mandatory = $true)]
	[System.Collections.ObjectModel.KeyedCollection[string,Microsoft.MetadirectoryServices.ConfigParameter]]
	$ConfigParameters,
	[parameter(Mandatory = $true)]
	[Microsoft.MetadirectoryServices.Schema]
	$Schema,
	[parameter(Mandatory = $true)]
	[Microsoft.MetadirectoryServices.OpenImportConnectionRunStep]
	$OpenImportConnectionRunStep,
	[parameter(Mandatory = $true)]
	[Microsoft.MetadirectoryServices.GetImportEntriesRunStep]
	$GetImportEntriesRunStep,
	[parameter(Mandatory = $true)]
	[Alias('PSCredential')] # To fix mess-up of the parameter name in the RTM version of the PowerShell connector.
	[System.Management.Automation.PSCredential]
	$Credential,
	[parameter(Mandatory = $false)]
	[ValidateScript({ Test-Path $_ -PathType "Container" })]
	[string]
	$ScriptDir = [Microsoft.MetadirectoryServices.MAUtils]::MAFolder # Optional parameter for manipulation by the TestHarness script.
)

Set-StrictMode -Version "2.0"

Function Escape-IllegalCharacters {
    
param(
	[parameter(Mandatory = $true)]
	[string]$DN
)

    $l = $DN.Length
    $i = 0
    $parts = New-Object System.Collections.ArrayList
    $delims = @("OU=","DC=","CN=")
    $illegalChars = ',\/#+<>;"='.ToCharArray()
    $charArr = $DN.ToCharArray()
    $tempString = ""
    $newDN = ""

    do {
        
        if ($charArr[$i] -eq ",") {
            # Check if end of part base on next 3 chars
            $endOfPart=$false
            if ($i+3 -lt $l) {
                #not at end of string
                $nextThree = ($charArr[$i+1] + $charArr[$i+2] + $charArr[$i+3]).ToUpper()
                if ($delims.Contains($nextThree)) {
                    $endOfPart = $true
                }
            }
            if ($endOfPart) {
                $parts.Add($tempString) | out-null
                $tempString = ""
                $i++
            }
            else {
                $tempString += $charArr[$i]
                $i++
            }
        }
        else {
            # Add to tempString
            $tempString += $charArr[$i]
            if ($i -eq $l-1) {
                $parts.Add($tempString) | out-null
            }
            $i++
        }
    

    }
    while ($i -lt $l)

    foreach ($part in $parts) {
        $type = $part.Substring(0,3)
        $value = $part.Substring(3)
        $newValue = ""
        foreach ($char in $value.ToCharArray()) {
            if ($illegalChars.Contains($char)) {
                $newValue += ("\"+$char)
            }
            else {
                $newValue+=$char
            }
        }
        $part = $type + $newValue
        if ($newDN -ne "") {
            $newDN += ","
        }
        $newDN+=$part
    }

    $newDN
}


Function Write-MIMLog ([string]$m) {
	$logFile = Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath "PSMA.txt"
	Write-Verbose $m
	$logOn = $true
	if ($logOn) {
		$d = [string](Get-Date -Format s)
		($d + "   " + $m) | Out-File $logFile -Append
	}
}

function Import-CSEntries
{
	<#
	.Synopsis
		Imports the users and OU's from the connected source.
	.Description
		Imports the users and OU's from the connected source.
	#>
	
	[CmdletBinding()]
    [OutputType([Microsoft.MetadirectoryServices.GetImportEntriesResults])]
	param(
	)
	
	$importEntriesResults = New-Object -TypeName "Microsoft.MetadirectoryServices.GetImportEntriesResults"
	$importEntriesResults.CSEntries = New-GenericObject System.Collections.Generic.List Microsoft.MetadirectoryServices.CSEntryChange

	foreach ($type in $schema.Types)
	{
		$objectType = $type.Name

		if ($customData.WaterMark.$objectType.MoreToImport -eq "1")
		{
			$lastRunDateTime = $null
			if ($deltaImport)
			{
				$lastRunDateTime = $customData.WaterMark.LastRunDateTime
			}

			$filterData = Get-PagingFilter -ObjectType $objectType -LastRunDateTime $lastRunDateTime

			Write-Debug ("Importing {0}. LdapFilter: {1}" -f $objectType, $filterData["LdapFilter"])

			switch ($objectType)
			{
				"User"
				{
					$importdata = Import-Users -LdapFilter $filterData["LdapFilter"]
					break
				}

				"OrganizationalUnit"
				{
					$importdata = Import-OrganizationalUnits -LdapFilter $filterData["LdapFilter"]
					break
				}

				default
				{
					throw "Unknown ObjectType: $_"
				}
			}
			
			$importEntriesResults.CSEntries.AddRange($importdata.CSEntries)

			if (!$filterData["MoreToImport"])
			{ 
				$customData.WaterMark.CurrentPageIndex = "0"
				$customData.WaterMark.$objectType.MoreToImport = "0"
			}
			else
			{
				$customData.WaterMark.CurrentPageIndex = [string]$filterData["NextPageIndex"]
			}

			Write-Debug ("Imported {0}. LdapFilter: {1} " -f $objectType, $filterData["LdapFilter"])

			break
		}
	}

	$importEntriesResults.MoreToImport = $false

	foreach ($type in $schema.Types)
	{
		$objectType = $type.Name

		if ($customData.WaterMark.$objectType.MoreToImport -eq "1")
		{
			$importEntriesResults.MoreToImport = $true
			break
		}
	}
	
	$importEntriesResults.CustomData = $customData.InnerXml

	Write-Debug ("WaterMark saved is: {0}" -f $importEntriesResults.CustomData)
	Write-Debug ("ImportEntriesResults.MoreToImport is: {0}" -f $importEntriesResults.MoreToImport)

	return $importEntriesResults
}

function Import-Users
{
	<#
	.Synopsis
		Imports the users from the connected source.
	.Description
		Imports the users from the connected source.
	#>
	
	[CmdletBinding()]
    [OutputType([Microsoft.MetadirectoryServices.GetImportEntriesResults])]
	param(
		[parameter(Mandatory = $true)]
		[string]
		$LdapFilter
	)

	$importReturnInfo = New-Object -TypeName "Microsoft.MetadirectoryServices.GetImportEntriesResults"
	$importReturnInfo.MoreToImport = $false
	$importReturnInfo.CSEntries = New-GenericObject System.Collections.Generic.List Microsoft.MetadirectoryServices.CSEntryChange

	$cmd = "Get-CsUser -LdapFilter '$LdapFilter'"
	if (![string]::IsNullOrEmpty($preferredDomainController))
	{
		$cmd += " -DomainController '$preferredDomainController'"
	}

	$statusMsg = "Invoking $cmd"
	$activityName = $MyInvocation.InvocationName

	Write-Progress -Id 1 -Activity $activityName -Status $statusMsg

	$x = Invoke-Expression $cmd

	if ($x)
	{
		if ($x -is [array])
		{
			$importReturnInfo.CSEntries.Capacity = $x.Count
		}

		$currentParition = $openImportConnectionRunStep.StepPartition.DN

		foreach ($i in $x) 
		{
			if ($i.DistinguishedName.EndsWith($currentParition, "OrdinalIgnoreCase") -eq $false)
			{
				Write-Debug ("Identity {0} does not belong to the current partition {1}. Dropping it from import results..." -f $i.DistinguishedName, $currentParition)
				continue
			}
			 
			$csentry = New-CSEntryChange -InputObject $i -ObjectType "User" -Schema $schema
			[void] $importReturnInfo.CSEntries.Add($csentry)
		}
	}

	return $importReturnInfo
}

function Get-PagingFilter
{
	<#
	.Synopsis
		Gets the paging filter for the specified object type.
	.Description
		Gets the paging filter for the specified object type.
	#>
	
	[CmdletBinding()]
    [OutputType([string])]
	param(
		[parameter(Mandatory = $true)]
		[string]
		$ObjectType,
		[parameter(Mandatory = $false)]
		[string]
		$LastRunDateTime
	)

	$filterData = @{}
		
	switch ($ObjectType)
	{
		"User"
		{
			$pageIndex = [int]$customData.WaterMark.CurrentPageIndex

			if ($pageIndex -ge $userPages.Length)
			{
				throw ("Unexpected Page Index $pageIndex to import objectType {0}. Max Index: {1}" -f $ObjectType, $userPages.Length)
			}

			$ldapFilter = "(&(objectCategory=person)(objectClass=user)(msRTCSIP-PrimaryUserAddress=sip:{0}*)" -f $userPages[$pageIndex].Trim() # no spaces in the LDAP query or it will fail.

			if (![string]::IsNullOrEmpty($LastRunDateTime))
			{
				$ldapFilter += "(whenChanged>={0:yyyyMMddHHmmss}.0Z)" -f ([DateTime]$LastRunDateTime).AddMinutes(-1*[int]$lastRunDateTimeOffset) # The LastRunDateTime in watermark is already in UTC.
			}

			$ldapFilter += ")"

			++$pageIndex
			$moreToImport = ($pageIndex -lt $userPages.Length)

			$filterData.Add("LdapFilter", $ldapFilter)
			$filterData.Add("NextPageIndex", $pageIndex)
			$filterData.Add("MoreToImport", $moreToImport)

			break
		}

		"OrganizationalUnit"
		{
			$pageIndex = [int]$customData.WaterMark.CurrentPageIndex

			if ($pageIndex -ge $ouPages.Length)
			{
				throw ("Unexpected Page Index $pageIndex to import objectType {0}. Max Index: {1}" -f $ObjectType, $ouPages.Length)
			}

			$ldapFilter = "(name={0}*)" -f $ouPages[$pageIndex].Trim()
			if (![string]::IsNullOrEmpty($LastRunDateTime))
			{
				$ldapFilter += "(whenChanged>={0:yyyyMMddHHmmss}.0Z)" -f ([DateTime]$LastRunDateTime).AddMinutes(-1*[int]$lastRunDateTimeOffset) # The LastRunDateTime in watermark is already in UTC.
			}

			++$pageIndex
			$moreToImport = ($pageIndex -lt $ouPages.Length)

			$filterData.Add("LdapFilter", $ldapFilter)
			$filterData.Add("NextPageIndex", $pageIndex)
			$filterData.Add("MoreToImport", $moreToImport)

			break
		}

		default
		{
			throw "Unexpected ObjectType $_"
		}
	}

	return $filterData
}

function Import-OrganizationalUnits
{
	<#
	.Synopsis
		Imports the OU's from the connected source.
	.Description
		Imports the OU's from the connected source.
	#>
	
	[CmdletBinding()]
    [OutputType([Microsoft.MetadirectoryServices.GetImportEntriesResults])]
	param(
		[parameter(Mandatory = $true)]
		[string]
		$LdapFilter
	)

	$importReturnInfo = New-Object -TypeName "Microsoft.MetadirectoryServices.GetImportEntriesResults"
	$importReturnInfo.MoreToImport = $false
	$importReturnInfo.CSEntries = New-GenericObject System.Collections.Generic.List Microsoft.MetadirectoryServices.CSEntryChange

	$statusMsg = "Importing OrganizationalUnits -LdapFilter $LdapFilter"
	$activityName = $MyInvocation.InvocationName

	Write-Progress -Id 1 -Activity $activityName -Status $statusMsg

	$x = Get-OrganizationalUnits -LdapFilter $LdapFilter

	if ($x)
	{
		if ($x -is [array])
		{
			$importReturnInfo.CSEntries.Capacity = $x.Count
		}

		foreach ($i in $x) 
		{ 
			$csentry = New-CSEntryChange -InputObject $i -ObjectType "OrganizationalUnit" -Schema $schema
			[void] $importReturnInfo.CSEntries.Add($csentry)
		}
	}

	return $importReturnInfo
}

function Get-OrganizationalUnits
{
	<#
	.Synopsis
		Imports the OU's from the connected source.
	.Description
		Imports the OU's from the connected source.
	#>
	
	[CmdletBinding()]
    [OutputType([string[]])]
	param(
		[parameter(Mandatory = $true)]
		[string]
		$LdapFilter
	)

	$includedNodeOUs = @()

	foreach ($includedNode in $openImportConnectionRunStep.InclusionHierarchyNodes)
	{
		$rootDN = Escape-IllegalCharacters -DN $includedNode.DN

		Write-Debug "Enumerating Inclusion OrganizationalUnit Hierarchy for $rootDN"

		$includedNodeOUs += Get-OrganizationalUnitHierarchy -RootDN $rootDN -LdapFilter $LdapFilter
	}

	$excludedNodeOUs = @()
	foreach ($exludedNode in $openImportConnectionRunStep.ExclusionHierarchyNodes)
	{
		$rootDN = Escape-IllegalCharacters -DN $includedNode.DN

		Write-Debug "Enumerating Exclusion OrganizationalUnit Hierarchy for $rootDN"

		$excludedNodeOUs += Get-OrganizationalUnitHierarchy -RootDN $rootDN -LdapFilter $LdapFilter
	}

	$organizationalUnit = @()
	foreach ($includedNodeOU in $includedNodeOUs)
	{
		$exclude = $false
		# seems excludedNodeOUs can just be ignored. 
		##foreach ($excludedNodeOU in $excludedNodeOUs)
		##{
		##	if ($includedNodeOU.DistinguishedName -eq $excludedNodeOU.DistinguishedName)
		##	{
		##		$exclude = $true
		##		break
		##	}
		##}

		if (!$exclude)
		{
			$organizationalUnit += $includedNodeOU
		}
	}

	return $organizationalUnit
}

function Get-OrganizationalUnitHierarchy
{
	<#
	.Synopsis
		Imports the specifed OU Hierarchy from the connected source.
	.Description
		Imports the specifed OU Hierarchy from the connected source.
	#>
	
	[CmdletBinding()]
    [OutputType([string[]])]
	param(
		[parameter(Mandatory = $true)]
		[string]
		$RootDN,
		[parameter(Mandatory = $true)]
		[string]
		$LdapFilter
	)

	$attributeNameMapping = @{ Name = "name"; DistinguishedName = "distinguishedname"; ObjectClass = "objectclass";  ObjectCategory = "objectcategory"; Guid = "objectguid" }

	if ([string]::IsNullOrEmpty($preferredDomainController))
	{
		$searchPath = "LDAP://{0}" -f $RootDN
	}
	else
	{
		$searchPath = "LDAP://{0}/{1}" -f $preferredDomainController, $RootDN
	}

	Write-Debug ("Get-OrganizationalUnitHierarchy Search Path: {0}" -f $searchPath)

	$userName = "{0}\{1}" -f $Credential.GetNetworkCredential().Domain, $Credential.GetNetworkCredential().UserName
	$password = $Credential.GetNetworkCredential().Password
	$searchRoot = New-Object System.DirectoryServices.DirectoryEntry -ArgumentList $searchPath, $userName, $password
	$ds = [adsisearcher]"(&(|(objectClass=organizationalUnit)(objectClass=Container))$LdapFilter)"
	$ds.searchroot = $searchRoot
	$ds.PageSize = 1000 
	$ds.SearchScope = "Subtree"
	$ds.PropertiesToLoad.AddRange($attributeNameMapping.Values)

	$results = $ds.FindAll()

	$organizationalUnits = @() 

	foreach ($result in $results)
	{
		$organizationalUnit = New-Object PSObject

		$props = $result.Properties

		foreach ($propName in $attributeNameMapping.Keys)
		{
			if (@($props.item($attributeNameMapping.$propName)).count -gt 1)
			{ 
				$values = [string[]]$props.item($attributeNameMapping.$propName)
				$organizationalUnit | Add-Member -MemberType NoteProperty -Name $propName -Value $values
			} 
			else
			{
				$organizationalUnit | Add-Member -MemberType NoteProperty -Name $propName -Value $props.item($attributeNameMapping.$propName)[0]
			} 
		}

		$organizationalUnits += $organizationalUnit 
	}

	return $organizationalUnits
}

try {

	
	
	$commonModule = (Join-Path -Path $ScriptDir -ChildPath $ConfigParameters["Common Module Script Name (with extension)"].Value)

	if (!(Get-Module -Name (Get-Item $commonModule).BaseName)) { Import-Module -Name $commonModule }

	Enter-Script -ScriptType "Import" -ErrorObject $Error


	$userPages = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "UserPages"

	if ([string]::IsNullOrEmpty($userPages))
	{
		$userPages = "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z,0,1,2,3,4,5,6,7,8,9".Split(",") 
	} 
	else
	{
		$userPages = $userPages.Split(",")
	}

	$ouPages = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "OrganizationalUnitPages"
	if ([string]::IsNullOrEmpty($ouPages))
	{
		$ouPages = @("")
	} 
	else
	{
		$ouPages = $ouPages.Split(",")
	}

	$lastRunDateTimeOffset = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "LastRunDateTimeOffsetMinutes" 

	if ([string]::IsNullOrEmpty($lastRunDateTimeOffset))
	{
		$lastRunDateTimeOffset = 30 # in minutes
	}
	else
	{
		$lastRunDateTimeOffset = $lastRunDateTimeOffset.Trim()
	}

	$deltaImport = $openImportConnectionRunStep.ImportType -eq "Delta"
	$customData = [xml]$getImportEntriesRunStep.CustomData

	$preferredDomainController = $customData.WaterMark.PreferredDomainController

	Write-Debug ("GetImportEntriesRunStep.CustomData received is: {0}" -f $customData.InnerXml)

	Import-CSEntries

	Exit-Script -ScriptType "Import" -ErrorObject $Error
}

catch {
	$ErrorMessage = $_.Exception.Message
    Write-MIMLog "ERROR!! Exception caught: $ErrorMessage"
}
