<#
.SYNOPSIS
    .
.DESCRIPTION
    .
.NOTES
    Author:     Børge Holse
    Date:       23-07-2016
    Version:    3.0
    Comment:    Oprettet
#>
$global:logPath = $null;
$global:messageContinue = $false;
$global:logTitle = "";
try
{
Add-Type -TypeDefinition @"
	// very simple enum type
	public enum LogLevels
	{
		Flow,
		Debug,
		Information,
		ActionInfo,
		Error,
		Warning
	}
"@ -ErrorAction SilentlyContinue
}
catch
{}

function global:ShowMessage([string] $prompt, $logLevel, $continue, $tag) 
{
    <#
    <Summary>
        <Description>Viser en besked</Description>
        <Parameters>
            <Parameter Name="$prompt" Description="Den tekst der skal vises" />
            <Parameter Name="$logLevel" Description="Den måde det skal vises på - Se Loglevels" />
            <Parameter Name="$continue" Description="Om teksten skal fortsætte næste gang den kaldes. " />
            <Parameter Name="$tag" Description="Bruges til logging for at kunne kategorisere" />
        </Parameters>
    </Summary>
    #>
	$starLine = "*******************************************************************************";
	$hyphenLine = "--------------------------------------------------------------------------------";
	if ($prompt -eq "-") {$prompt = $starLine}
	switch ($logLevel)
	{
		[Loglevels]::Flow
		{
			if ($continue)
			{
				if ($global:messageContinue)
				{
					Write-Host $prompt -backgroundcolor white -ForegroundColor Black
				}
				else
				{
					Write-Host $starLine -backgroundcolor white -ForegroundColor Black
					Write-Host $prompt -backgroundcolor white -ForegroundColor Black
				}
				$global:messageContinue = $true;
			}
			else
			{
				if ($global:messageContinue)
				{
					Write-Host $prompt -backgroundcolor white -ForegroundColor Black
					Write-Host $starLine -backgroundcolor white -ForegroundColor Black
				}
				else
				{
					Write-Host $starLine -backgroundcolor white -ForegroundColor Black
					Write-Host $prompt -backgroundcolor white -ForegroundColor Black
					Write-Host $starLine -backgroundcolor white -ForegroundColor Black
				}
				$global:messageContinue = $false;
			}
		}	
		[Loglevels]::Information
		{
			if ($continue)
			{
				Write-Host $prompt -ForegroundColor green -NoNewLine 
			}
			else
			{
				Write-Host $prompt -ForegroundColor green
			}
		}	
		[Loglevels]::Debug
		{
			if ($Global:ShowDebug)
			{
				if ($continue)
				{
					write-host $prompt -ForegroundColor yellow -Backgroundcolor blue -NoNewLine 
				}
				else
				{
					write-host $prompt -ForegroundColor yellow -Backgroundcolor blue					
				}
			}
		}
		[Loglevels]::Warning
		{
			if ($continue)
			{
				Write-Host $prompt -ForegroundColor green -NoNewLine 
			}
			else
			{
				Write-Host $prompt -ForegroundColor green
			}
			write-host $prompt -ForegroundColor yellow 
		}
		[Loglevels]::Error
		{
			if ($continue)
			{
				if ($global:messageContinue)
				{
					write-host $prompt -ForegroundColor red
				}
				else
				{
					write-host $hyphenLine -ForegroundColor red
					write-host $prompt -ForegroundColor red
				}
				$global:messageContinue = $true;
			}
			else
			{
				if ($global:messageContinue)
				{
					write-host $prompt -ForegroundColor red
					write-host $hyphenLine -ForegroundColor red
				}
				else
				{
					write-host $hyphenLine -ForegroundColor red
					write-host $prompt -ForegroundColor red
					write-host $hyphenLine -ForegroundColor red
				}
				$global:messageContinue = $false;
			}
		}
		default
		{
			write-host $logLevel
		}
	}

	if ($global:logPath -ne $null) 
	{
		$logFile = [System.IO.Path]::Combine($global:logPath, "log.csv")
		$mode = [System.IO.FileMode]::Append
		$access = [System.IO.FileAccess]::Write
		$sharing = [IO.FileShare]::Read
		$encoding = [System.Text.Encoding]::UTF8
		$stream = New-Object IO.FileStream $logFile ,$mode, $access, $sharing
		$sw = New-Object System.IO.StreamWriter($stream, $encoding)
		if ($global:logTitle -eq "") 
		{
			$global:logTitle = "Date;Level;Tag;Message"
			$sw.WriteLine($global:logTitle)
		}
		$logLine = (Get-Date -Format g).ToString() + ";" + $logLevel + ";" + $tag + ";" + $prompt
		$sw.WriteLine($logLine)
		$sw.close()
		$sw.dispose()
		$stream.dispose()
	}
}

ShowMessage "Loading Core Version 3.0.0" [LogLevels]::Flow

$thisScript = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
. ($thisScript + '\01_01_ErrorHandling.ps1')
. ($thisScript + '\01_05_General.ps1')
. ($thisScript + '\01_10_xml.ps1')

$global:configurationXML = LoadSetupConfig ($global:root + "\config.xml")
if ($global:configurationXML.Configuration.SolutionPath -ne $null) {
	$global:root = $global:configurationXML.Configuration.SolutionPath;
}