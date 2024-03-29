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

.PARAMETER Sprints
Collect the last 6 months of sprints and write those to a separate file.
The filename is sprints.csv. This is mutually exclusive with all other
parameters. If specified, this file is generated and the script exits.

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
	[string] $File,
	[switch] $Sprints
)

Begin
{
	. $PSScriptRoot\common.ps1

	# SETTINGS...

	$OutputFile = './issues.csv'
	$SprintsFile = './sprints.csv'
	$URI = 'https://waterscorporation.atlassian.net/rest/api/3'
	$ARI = 'https://waterscorporation.atlassian.net/rest/agile/1.0'
	$Header = 'Accept: application/json'

	# used /field API to list all custom fields; story points is a custom field.
	$PointsField = 'customfield_10201'
	$SprintField = 'customfield_10020'
	$TeamField = 'customfield_10253'
	$StartStatus = 'In Progress'
	$EndStatus = 'Verified'


	function GetIssues
	{
		if (Test-Path $File)
		{
			Remove-Item $File -Force -Confirm:$false
		}

		'Sprint,Team,User,Epic,Key,Type,Points,StartedDt,InTestDt,PassedDt,VerifiedDt,Days,WeekDays,InProgress,InTest,Passed,Reworked,Repassed,Reverified' | Out-File -FilePath $File

		Write-Host 'Legend: [.] OK, [+] reverified, [-] skip no start or end, [x] skip no points'
		Write-Host

		$startAt = 0

		$Updated = ''
		if ($Since)
		{
			$ms = ([DateTimeOffset]([DateTime]::Parse($Since))).ToString('yyyy-MM-dd')
			$updated = " AND updated>=$ms"
		}

		do
		{
			$jql = [System.Web.HttpUtility]::UrlEncode(
				"project=$Project AND issuetype IN (Story, Defect) AND status=$EndStatus$updated")

			$url = "$URI/search?jql=$jql&startAt=$startAt"
			Write-Verbose $url

			# to see verbose output including http header, change -s to -v
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

		$type = $issue.fields.issuetype.name
		$sprint = $issue.fields.$SprintField.name
		$team = $issue.fields.$TeamField.value
		$user = $issue.fields.assignee.displayName

		# extract sprint name without team, match on SCP 'Sprint n' or Scalars '2024Q1.3'
		$mats = ([regex]'Sprint \d+|\d{4}Q\d+\.\d+').Matches($sprint)
		if ($mats.Count -gt 0)
		{
			# if story was in multiple sprints, choose latest one
			$sprint = ($mats | sort | select -last 1).Value
		}
		else
		{
			Write-Host 't' -NoNewline
			return
		}

		$epic = ''
		if ($issue.fields.parent.fields.issuetype.name -eq 'Epic')
		{
			$epic = $issue.fields.parent.key
		}

		$changes = GetChangeLog $issue.key

		# find starting state
		$item = $changes | where { $_.toStatus -eq $StartStatus } | select -first 1
		if (-not $item) { Write-Host '-' -NoNewline; "$($issue.key) no start" | out-file 'issues.log' -append; return }

		$started = $item.created

		# find first occurrence of finished state; assume multiple occurrences mean just
		# a fix to Components or FixVersion with no actual rework
		$item = $changes | where { $_.toStatus -eq $EndStatus } | select -first 1
		if (-not $item) { Write-Host '-' -NoNewline; "$($issue.key) no end" | out-file 'issues.log' -append; return }

		$finished = $item.created
		$days = [int][Math]::Ceiling(($finished - $started).TotalDays)
		$weekdays = (CountWeekDays $started $finished).ToString('0.##')

		# look for backwards transitions from Verified
		$reverified = ($changes | where { $_.fromStatus -eq 'Verified' } | measure).Count
		if ($reverified -gt 0) { $reverified = 1 }

		$marker = '.'

		# ignore remaining
		$index = $changes.indexOf($item)
		if ($index -lt $changes.Length - 1)
		{
			$marker = '+'
			$changes = $changes[0..$index]
		}

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

		Write-Host $marker -NoNewline

		"$sprint,$team,$user,$epic,$($issue.Key),$type,$points,$started,$tested,$passed,$finished,$days,$weekdays,$progress,$testedDays,$passedDays,$reworked,$repassed,$reverified" | Out-File -FilePath $File -Append
	}

	function GetChangeLog
	{
		param($key)
		$startAt = 0
		$changes = @()

		do
		{
			$url = "$uri/issue/$key/changelog?startAt=$startAt"
			$changelog = curl -s --request GET -url $url --user "$email`:$PAT" --header $Header | ConvertFrom-Json

			# grab only status changes and flatten into a custom collection
			$changes += $changelog.values | foreach {
				$item = $_.items | where {
					$_.field -eq 'status' -and $_.fromString -ne $null -and $_.toString -ne $null
				} | select -first 1

				if ($item -ne $null)
				{
					[PSCustomObject]@{
						# key = $key
						# id = $_.id
						# author = $_.author.displayName
						created    = $_.created
						fromStatus = $item.fromString
						toStatus   = $item.toString
					}
				}
			}

			$startAt += $changelog.maxResults

		} while ($changelog.isLast -eq $false)
		return $changes | sort -property created
	}

	function GetSprints
	{
		Write-Host
		Write-Host "... Reporting six months of sprint to $SprintsFile"

		if (Test-Path $SprintsFile)
		{
			Remove-Item $SprintsFile -Force -Confirm:$false
		}

		if ($since)
		{
			$window = $since
		}
		else
		{
			# last 6 months
			$window = (Get-Date).AddMonths(-6).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
		}

		$url = "$ARI/board?projectKeyOrId=$($Project.toUpper())&type=scrum"
		$boards = curl -s --request GET -url $url --user "$email`:$PAT" --header $Header | ConvertFrom-Json

		'Board,Sprint,StartDt,EndDt,Goal' | Out-File -FilePath $SprintsFile

		$boards.values | foreach {
			$name = $_.name

			$url = "$($_.self)/sprint?state=active,closed"
			$board = curl -s --request GET -url $url --user "$email`:$PAT" --header $Header | ConvertFrom-Json

			$board.values | where { $_.startDate -gt $window } | sort -property startDate | foreach {
				# extract sprint name without team, match on SCP 'Sprint n' or Scalars '2024Q1.3'
				$mats = ([regex]'Sprint \d+|\d{4}Q\d+\.\d+').Matches($_.name)
				if ($mats.Count -gt 0)
				{
					# if story was in multiple sprints, choose latest one
					$sprint = ($mats | sort | select -last 1).Value
				}
				else
				{
					$sprint = 'sprint'
				}
				# clean multi-line, commas, bullets from goal for CSV
				$goal = $_.goal -replace "`n|`r|,|^[^\w]*", ''
				$startDate = [DateTime]::Parse($_.startDate).ToLocalTime()
				$endDate = [DateTime]::Parse($_.endDate).ToLocalTime()
				"$name,$sprint,$startDate,$endDate,$goal" | Out-File -FilePath $SprintsFile -Append
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

	if ($File)
	{
		$dir = Split-Path $File
		if ($dir -eq '')
		{
			$SprintsFile = "$(Split-Path $File -LeafBase)-sprints.csv"
		}
		else
		{
			$SprintsFile = Join-Path (Split-Path $File) "$(Split-Path $File -LeafBase)-sprints.csv"
		}
	}
	else
	{
		$File = $OutputFile
	}

	InstallCurl

	if ($Sprints)
	{
		GetSprints
	}
	else
	{
		GetIssues
	}
}
