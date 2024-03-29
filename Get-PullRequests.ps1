<#
.SYNOPSIS
Export CSV of Github pull requests, including number of comments, commits, etc.

.PARAMETER Append
Specify this switch to append to the output file rather than recreate it.
Makes it easy to append to a single file when using the User parameter.

.PARAMETER File
The path of the output CSV file, default is ./issues.csv

.PARAMETER Org
The Github organization name.
Override with $env:GH_ORG

.PARAMETER Since
Filters PRs since the given date. The parameter value must be a valid
date string that can be parsed by the DateTime class, e.g. "12-JAN-2024"

.PARAMETER Token
Your personal Github access token, created from your Github account.
Override with $env:GH_TOKEN

.PARAMETER User
A single Github user account name to report. For multiple users in a single
file, add the Append parameter for subsequent calls after the first user.

.PARAMETER Username
Your Github account username used for authentication.
Override with $env:GH_USERNAME

.PARAMETER Users
A collection of GitHub user account names to collate into a single file.
This is a faster way of using [-User] followed by [-User -Append]

.DESCRIPTION
blah
#>

# CmdletBinding adds -Verbose functionality, SupportsShouldProcess adds -WhatIf
[CmdletBinding(SupportsShouldProcess = $true)]

param (
	[string] $Org,
	[string] $Username,
	[string] $Token,
	[string] $User,
	[string[]] $Users,
	[string] $Since,
	[string] $File,
	[switch] $Append
)

Begin
{
	. $PSScriptRoot\common.ps1

	# SETTINGS...

	$OutputFile = './pull-requests.csv'
	$URI = 'https://api.github.com'
	$Header = 'Accept: application/json'

	$script:TotalCount = 0
	$script:lastUpdated = '2000-01-01T01:01:01Z'

	function GetPullRequests
	{
		param(
			[string] $usr,
			[bool] $writeHeader
		)

		if ($writeHeader) {
			'User,Repo,PR,State,Commits,Post-Commits,Comments,Created,Closed,Updated,Days,Sonars' | Out-File -FilePath $File
		}

		Write-Host

		$url = "$URI/search/issues?q=author:$usr+org:$Org+type:pr" #+sort:updated+direction:desc+per_page:99"
		Write-Verbose "$url --user ""$Username`:$Token"""

		$raw = curl -s --request GET --url $url --user "$Username`:$Token" --header $Header
		$page = $raw | ConvertFrom-Json

		if ($LASTEXITCODE -ne 0) {
			$page
			return
		}

		$items = $page.items | where { $_.updated_at -ge $Since }

		Write-Host "Checking $($items.Count) out of $($page.total_count) PRs for $usr" -NoNewline
		$items | foreach { ReportItem $_ $usr }

		Write-Host
	}


	function ReportItem
	{
		param($item, $usr)

		$number = $item.number

		$repo = $item.repository_url.substring($item.repository_url.lastIndexOf("/") + 1)

		if ($item.draft) {
			$state = 'Draft'
		}
		else {
			$state = $item.state
		}

		# examine commits
		$u = $item.comments_url.Replace("/issues", "/pulls").Replace("/comments", "/commits")
		$data = curl -s --request GET --url $u --user "$Username`:$Token" --header $Header | ConvertFrom-Json
		$commits = $data.Length
		$postcommits = ($data | where { $_.commit.author.date -gt $item.created_at }).Count

		# adjust comments, subtracting those from automated tools
		$comments = $item.comments
		$data = curl -s --request GET --url $item.comments_url --user "$Username`:$Token" --header $Header | ConvertFrom-Json
		$data | foreach { if ($_.performed_via_github_app -ne $null) { $comments = $comments - 1 } }

		$sonars = 0
		$data | where { $_.user.login -eq 'sonarcloud[bot]' } | foreach {
			if ($_.body -match '\[(\d+) New issues\]') {
				if ($matches -and ($matches.count -gt 0)) {
					$sonars = $sonars + [int]$matches[1]
				}
			}
		}

		$created = $item.created_at
		if (![String]::IsNullOrWhiteSpace($created)) {
			$created = [System.TimeZone]::CurrentTimeZone.ToLocalTime($created)
		}

		$closed = $item.closed_at
		$days = ''

		if ([String]::IsNullOrWhiteSpace($closed)) {
			$closed = ''
		}
		else {
			$closed = [System.TimeZone]::CurrentTimeZone.ToLocalTime($closed)
			$days = (CountWeekDays $created $closed).ToString('0.##')
		}

		$updated = $item.updated_at
		if ($updated -gt $lastUpdated)
		{
			$script:lastUpdated = $updated
		}

		"$usr,$repo,$number,$state,$commits,$postcommits,$comments,$created,$closed,$updated,$days,$sonars" | Out-File -FilePath $File -Append

		Write-Host '.' -NoNewline
		$script:TotalCount = $TotalCount + 1
	}
}
Process
{
	if ([String]::IsNullOrWhiteSpace($Org))
	{
		$Org = $env:GH_ORG
		if ([String]::IsNullOrWhiteSpace($Org))
		{
			$Org = PromptForValue 'Your Github org'
		}
	}

	if ([String]::IsNullOrWhiteSpace($Username))
	{
		$Username = $env:GH_USERNAME
		if ([String]::IsNullOrWhiteSpace($Username))
		{
			$Username = PromptForValue 'Your Github username'
		}
	}

	if ([String]::IsNullOrWhiteSpace($Token))
	{
		$Token = $env:GH_TOKEN
		if ([String]::IsNullOrWhiteSpace($Token))
		{
			$Token = PromptForValue 'Your Github token'
		}
	}

	if ([String]::IsNullOrWhiteSpace($User) -and
		(-not $PSBoundParameters.ContainsKey('Users') -or $Users.Length -lt 1))
	{
		$User = PromptForValue 'Github user to report'
	}

	$Org = [System.Web.HttpUtility]::UrlEncode($Org)

	if ($Since)
	{
		$script:Since = ([DateTimeOffset]([DateTime]::Parse($Since))).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
	}

	if (!$File)
	{
		$File = $OutputFile
		if ((Test-Path $File) -and -not $Append)
		{
			Remove-Item $File -Force -Confirm:$false
		}
	}

	InstallCurl

	if ($Users -and $Users.Length -gt 0) {
		$first = $true
		$Users | foreach {
			$user = [System.Web.HttpUtility]::UrlEncode($_)
			GetPullRequests $user $first
			$first = $false
		}
	}
	else
	{
		$user = [System.Web.HttpUtility]::UrlEncode($User)
		write-host "user=$user"
		write-host "app=$append"
		GetPullRequests $user ((-not $Append) -eq $true)
	}

	Write-Host "`nFound $TotalCount PRs. Most recently updated: $lastUpdated"
}
