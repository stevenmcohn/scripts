<#
.SYNOPSIS
Export CSV of Jira issues, including number of days state remains In Progress.

.PARAMETER Email
Your Jira email address used for basic authentication, e.g. your Waters email.
Override with $env:JIRA_EMAIL

.PARAMETER File
The path of the output CSV file, default is ./issues.csv

.PARAMETER PAT
Your Jira personal access token, created from your Atlassian account.
Override with $env:JIRA_PAT

.PARAMETER Project
The project code to query, e.g. INFSCP.
Override with $env:JIRA_PROJECT

.PARAMETER Since
Filters issues since the given date. The parameter value must be a valid
date string that can be parsed by the DateTime class.

.DESCRIPTION
Jira Cloud export feature does not include the changelog associated with each
issue and there is no way to query the changelog to determine when issues
transitioned from state to state, other than visually looking at the issue
in the browser.

This script uses the Jira REST API to extract the changelog for each issue in
a specified Jira project. The data is stored as a CSV so it can be pulled into
Excel to calculate average days "In Progres" for each story size in points.

Basic Auth for REST APIs
https://developer.atlassian.com/cloud/jira/platform/basic-auth-for-rest-apis/

REST APIs
https://developer.atlassian.com/cloud/jira/platform/rest/v3/intro/#about
#>

# CmdletBinding adds -Verbose functionality, SupportsShouldProcess adds -WhatIf
[CmdletBinding(SupportsShouldProcess = $true)]

param (
	[string] $Email,
	[string] $PAT,
	[string] $Project,
	[string] $Since,
	[string] $File
)

Begin
{
	# SETTINGS...

	$OutputFile = './issues.csv'
	$URI = 'https://waterscorporation.atlassian.net/rest/api/3'
	$Header = 'Accept: application/json'

	# used /field API to list all custom fields; story points is a custom field.
	$PointsField = 'customfield_10201'
	$StartState = 'In Progress'
	$EndState = 'Verified'


	function AppendToPowerShellProfile
	{
		param($key, $command)
		# PowerShell >= 6.0
		$0 = "$HOME\Documents\PowerShell\Microsoft.PowerShell_profile.ps1"
		if (!(Test-Path $0)) { New-Item $0 -Force -Confirm:$false | Out-Null }
		if ((Get-Content $0 | select-String $key).Count -eq 0) { Add-Content $0 $command }
		# PowerShell <= 5.1
		$0 = "$HOME\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1"
		if (!(Test-Path $0)) { New-Item $0 -Force -Confirm:$false | Out-Null }
		if ((Get-Content $0 | select-String $key).Count -eq 0) { Add-Content $0 $command }
	}

	function InstallChocolatey
	{
		# Modules/Scripts contains a better version but this is a stand-alone copy for the
		# top-level Install scripts so they can remain independent of the Module scripts
		if ((Get-Command choco -ErrorAction:SilentlyContinue) -eq $null)
		{
			# touch $profile prior to install
			if (!(Test-Path $profile))
			{
				$folder = [System.IO.Path]::GetDirectoryName($profile)
				if (!(Test-Path $folder)) { New-Item $folder -Itemtype Directory -Force -Confirm:$false | Out-Null }
				New-Item -Path $profile -ItemType File -Value '' -Confirm:$false | Out-Null
			}
	
			Write-Host '... Installing Chocolatey'
			Set-ExecutionPolicy Bypass -Scope Process -Force
			Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
	
			AppendToPowerShellProfile 'chocolateyProfile.psm1' 'Import-Module $env:ChocolateyInstall\helpers\chocolateyProfile.psm1'
			. $profile
		}
	
		# import profile to ensure we can use refreshenv
		$0 = 'C:\ProgramData\chocolatey\helpers\chocolateyProfile.psm1'
		if (Test-Path $0)
		{
			Import-Module $0
		}
	}

	function InstallCurl
	{
		# default PowerShell curl alias points to Invoke-WebRequest which is... disappointing
		Remove-Item alias:curl -ErrorAction:SilentlyContinue
	
		$cmd = Get-Command curl -ErrorAction:SilentlyContinue
		if ($cmd -ne $null)
		{
			if ($cmd.Source.Contains('curl.exe')) { return }
		}
	
		if ((Get-Command choco -ErrorAction:SilentlyContinue) -eq $null)
		{
			InstallChocolatey
		}
	
		if ((choco list -l 'curl' | Select-string 'curl ').count -gt 0) { return }
	
		Write-Host '... Installing Curl'
		choco install -y curl
	
		AppendToPowerShellProfile 'alias:curl' 'Remove-Item alias:curl -ErrorAction SilentlyContinue'
		. $profile
	}
	
	function PromptForValue
	{
		param($prompt, $value)
		if ([String]::IsNullOrWhiteSpace($value))
		{
			While ([String]::IsNullOrWhiteSpace($value))
			{
				$value = Read-Host $prompt
			}
		}
		else
		{
			$val = Read-Host "$prompt [$value]"
			if (![String]::IsNullOrWhiteSpace($val))
			{
				$value = $val
			}
		}

		return $value
	}

	# ==================================================================================

	function GetIssues
	{
		$startAt = 0

		$Updated = ''
		if ($Since)
		{
			$ms = ([DateTimeOffset]([DateTime]::Parse($Since))).ToUnixTimeMilliseconds()
			$updated = " AND updated>=$ms"
		}

		do
		{
			$jql = [System.Web.HttpUtility]::UrlEncode(
				"project=$Project AND issuetype=Story AND status=$EndState$updated")

			$url = "$URI/search?jql=$jql&startAt=$startAt"
			$page = curl -s --request GET --url $url --user "$email`:$PAT" --header $Header | ConvertFrom-Json

			if ($startAt -eq 0)
			{
				Write-Host "Checking $($page.total) $EndState issues from $Project" -NoNewline
				if ($updated -eq '') { Write-Host } else { Write-Host " since $Since" }
			}

			$page.issues | foreach { MeasureIssue $_ }

			$startAt += $page.maxResults
			$lastPage = $page.total - $page.maxResults

		} while ($startAt -le $lastpage)

		Write-Host
	}

	function MeasureIssue
	{
		param($issue)
		
		$points = $issue.fields.$PointsField
		if ([String]::IsNullOrWhiteSpace($points))
		{
			# indicates that the story points are unspecified for this issue
			Write-Host '-' -NoNewline
			return
		}

		$url = "$uri/issue/$($issue.key)/changelog"
		$changelog = curl -s --request GET -url $url --user "$email`:$PAT" --header $Header | ConvertFrom-Json

		$item = $changelog.values | where {
			$_.items | where { $_.field -eq 'status' -and $_.toString -eq $StartState }
		} | select -first 1

		if ($changelog.total -gt $changelog.maxResults)
		{
			# NOTE this indicates there are more pages of changelog;
			# we may need to enhance this script to query those extra pages
			Write-Host '+' -NoNewline
		}
		else
		{
			Write-Host '.' -NoNewline
		}

		if ($item -and $item.created)
		{
			$started = $item.created
	
			$item = $changelog.values | where {
				$_.items | where { $_.field -eq 'status' -and $_.toString -eq $EndState }
			} | select -first 1

			if ($item -and $item.created)
			{
				$verified = $item.created
				$days = [int][Math]::Ceiling(($verified - $started).TotalDays)
				"$($issue.Key),$points,$started,$verified,$days" | Out-File -FilePath $File -Append
			}
		}
	}
}
Process
{
	if ([String]::IsNullOrWhiteSpace($Email))
 {
		$email = $env:JIRA_EMAIL
		if ([String]::IsNullOrWhiteSpace($Email))
		{
			PromptForValue 'Your email'
		}
	}

	if ([String]::IsNullOrWhiteSpace($PAT))
	{
		$PAT = $env:JIRA_PAT
		if ([String]::IsNullOrWhiteSpace($PAT))
		{
			PromptForValue 'Your PAT'
		}
	}

	if ([String]::IsNullOrWhiteSpace($Project))
	{
		$Project = $env:JIRA_PROJECT
		if ([String]::IsNullOrWhiteSpace($Project))
		{
			PromptForValue 'Project key'
		}
	}

	if (!$File)
	{
		$File = $OutputFile
		if (Test-Path $File)
		{
			Remove-Item $File -Force -Confirm:$false
		}
	}

	InstallCurl

	GetIssues
}
