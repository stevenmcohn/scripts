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
	$StartStatus = 'In Progress'
	$EndStatus = 'Verified'


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

	function CountWeekDays
	{
		param($start, $end)
		$days = ($end - $start).TotalDays
		for ($d = $start;$d -le $end; $d = $d.AddDays(1)) {
			if ($d.DayOfWeek -match "Sunday|Saturday") {
				$days -= 1.0
			}
		}
		return $days
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
		'Key,Points,StartedDt,InTestDt,PassedDt,VerifiedDt,Days,WeekDays,InProgress,InTest,Passed,Reworked,Repassed,Reverified' | Out-File -FilePath $File

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
				"project=$Project AND issuetype=Story AND status=$EndStatus$updated")

			$url = "$URI/search?jql=$jql&startAt=$startAt"
			$page = curl -s --request GET --url $url --user "$email`:$PAT" --header $Header | ConvertFrom-Json

			if ($startAt -eq 0)
			{
				Write-Host "Checking $($page.total) $EndStatus issues from $Project" -NoNewline
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
			Write-Host 'x' -NoNewline
			return
		}

		Write-Host '.' -NoNewline

		$changes = GetChangeLog $issue.key

		# find starting state
		$item = $changes | where { $_.toStatus -eq $StartStatus } | select -first 1
		if (-not $item) { Write-Host '-' -NoNewline; "$($issue.key) no start" | out-file 'issues.log' -append; return }

		$started = $item.created

		# find finished state
		$item = $changes | where { $_.toStatus -eq $EndStatus } | select -last 1
		if (-not $item) { Write-Host '-' -NoNewline; "$($issue.key) no end" | out-file 'issues.log' -append; return }

		$finished = $item.created
		$days = [int][Math]::Ceiling(($finished - $started).TotalDays)
		$weekdays = (CountWeekDays $started $finished).ToString('0.##')

		# total in progress duration, across one or more test>prog>test transitions
		$item = $changes | where { $_.fromStatus -eq $StartStatus } | select -last 1
		$progress = (CountWeekDays $started $item.created).ToString('0.##')

		# last date moved to Passed and calc days held in the Passed status
		# this can also be considered the time it took to verify
		$passed = $finished
		$item = $changes | where { $_.toStatus -eq 'Passed' } | select -last 1
		if ($item) { $passed = $item.created }
		$passedDays = (CountWeekDays $passed $finished).ToString('0.##')

		# last date moved to In Test and calc days held in the In Test status
		# this can also be considered the time it took to test
		$tested = $passed
		$item = $changes | where { $_.toStatus -eq 'In Test' } | select -last 1
		if ($item) { $tested = $item.created }
		$testedDays = (CountWeekDays $tested $passed).ToString('0.##')

		# look for backwards transitions from In Test
		$reworked = ($changes | where {
			$_.fromStatus -eq 'In Test' -and $_.toStatus -notmatch 'Passed|Verified|Rejected'
		} | measure).Count
		if ($reworked -gt 0) { $reworked = 1 }

		# look for backwards transitions from Passed
		$repassed = ($changes | where {
			$_.fromStatus -eq 'Passed' -and $_.toStatus -notmatch 'Verified|Rejected'
		} | measure).Count
		if ($repassed -gt 0) { $repassed = 1 }

		# look for backwards transitions from Verified
		$reverified = ($changes | where { $_.fromStatus -eq 'Verified' } | measure).Count
		if ($reverified -gt 0) { $reverified = 1 }
		
		"$($issue.Key),$points,$started,$tested,$passed,$finished,$days,$weekdays,$progress,$testedDays,$passedDays,$reworked,$repassed,$reverified" | Out-File -FilePath $File -Append
	}

	function GetChangeLog
	{
		param($key)
		$startAt = 0
		$changes = @()

		do {
			$url = "$uri/issue/$key/changelog?startAt=$startAt"
			$changelog = curl -s --request GET -url $url --user "$email`:$PAT" --header $Header | ConvertFrom-Json

			# grab only status changes and flatten into a custom collection
			$changes += $changelog.values | foreach {
				$item = $_.items | where {
					$_.field -eq 'status' -and $_.fromString -ne $null -and $_.toString -ne $null
				} | select -first 1

				if ($item -ne $null) {
					[PSCustomObject]@{
						# key = $key
						# id = $_.id
						# author = $_.author.displayName
						created = $_.created
						fromStatus = $item.fromString
						toStatus = $item.toString
					}
				}
			}

			$startAt += $changelog.maxResults

		} while ($changelog.isLast -eq $false)
		return $changes | sort -property created
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
