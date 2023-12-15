<#
Library of common functions shared by scripts in this folder.
#>
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
	for ($d = $start; $d -le $end; $d = $d.AddDays(1))
	{
		if ($d.DayOfWeek -match "Sunday|Saturday" -and $days -gt 0)
		{
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
