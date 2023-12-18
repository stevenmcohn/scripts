<#
.SYNOPSIS
Export CSV of Github pull requests, including number of comments, commits, etc.

.PARAMETER Append
Specify this switch to append to the output file rather than recreate it.
Makes it easy to append multiple users to a single file.

.PARAMETER File
The path of the output CSV file, default is ./issues.csv

.PARAMETER Org
The Github organization name.
Override with $env:GH_ORG

.PARAMETER Token
Your personal Github access token, created from your Github account.
Override with $env:GH_TOKEN

.PARAMETER User
The Github user to report

.PARAMETER Username
Your Github account username used for authentication.
Override with $env:GH_USERNAME

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


	function GetPullRequests
	{
		if (-not $Append) {
			'User,Repo,PR,State,Commits,Post-Commits,Comments,Created,Closed,Days,Sonars' | Out-File -FilePath $File
		}

		Write-Host

		$url = "$URI/search/issues?q=author:$User+org:$Org+type:pr" #+sort:updated+direction:desc+per_page:99"
		Write-Verbose "$url --user ""$Username`:$Token"""

		$raw = curl -s --request GET --url $url --user "$Username`:$Token" --header $Header
		$page = $raw | ConvertFrom-Json

		if ($LASTEXITCODE -ne 0) {
			$page
			return
		}

		Write-Host "Checking $($page.total_count) PRs for $User" -NoNewline
		$page.items | foreach { ReportItem $_ }

		Write-Host
	}


	function ReportItem
	{
		param($item)

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

		"$User,$repo,$number,$state,$commits,$postcommits,$comments,$created,$closed,$days,$sonars" | Out-File -FilePath $File -Append

		Write-Host '.' -NoNewline
	}
}
Process
{
	if ([String]::IsNullOrWhiteSpace($Org))
	{
		$Org = $env:GH_ORG
		if ([String]::IsNullOrWhiteSpace($Org))
		{
			PromptForValue 'Your Github org'
		}
	}

	if ([String]::IsNullOrWhiteSpace($Username))
	{
		$Username = $env:GH_USERNAME
		if ([String]::IsNullOrWhiteSpace($Username))
		{
			PromptForValue 'Your Github username'
		}
	}

	if ([String]::IsNullOrWhiteSpace($Token))
	{
		$Token = $env:GH_TOKEN
		if ([String]::IsNullOrWhiteSpace($Token))
		{
			PromptForValue 'Your Github token'
		}
	}

	if ([String]::IsNullOrWhiteSpace($User))
	{
		PromptForValue 'Github user to report'
	}

	$Org = [System.Web.HttpUtility]::UrlEncode($Org)
	$User = [System.Web.HttpUtility]::UrlEncode($User)

	if (!$File)
	{
		$File = $OutputFile
		if ((Test-Path $File) -and -not $Append)
		{
			Remove-Item $File -Force -Confirm:$false
		}
	}

	InstallCurl

	GetPullRequests
}
