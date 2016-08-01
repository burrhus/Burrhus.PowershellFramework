<#
.SYNOPSIS
    test
.DESCRIPTION
    test2
.PARAMETER Path
    The path to the .
.PARAMETER LiteralPath
    Specifies a path to one or more locations. Unlike Path, the value of 
    LiteralPath is used exactly as it is typed. No characters are interpreted 
    as wildcards. If the path includes escape characters, enclose it in single
    quotation marks. Single quotation marks tell Windows PowerShell not to 
    interpret any characters as escape sequences.
.PARAMETER Interactive
    sdfsdf
.EXAMPLE
    C:\PS> 
    <Description of example>
.NOTES
    Author: Børge Holse
    Date:   23-07-2016    
    Version: 3.0
#>
param(
	[bool]$Interactive=$true, 
	[bool]$UseSetupFile=$true, 
	[string]$URL="http://localhost", 
	[string]$Environment="UDV", 
	[string]$Solution="", 
	[string]$DirectRun="", 
	[string]$UpdateID="",
	[string]$Addon="",
	[bool]$ShowDebug=$true
)

# Initialize
cls
$error.clear()

if ($PSScriptRoot -eq $null) {
	$thisScript = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
} else {
	$thisScript = $PSScriptRoot
}
$global:root = $thisScript
$global:configurationXML = $null;
$Global:ShowDebug = $ShowDebug
#Load scripts
. ($thisScript + '\01_Core\Load.ps1')
ShowMessage "Burrhus.PowerShellFramework" [LogLevels]::Flow $true
ShowMessage "Version 3.0.0" [LogLevels]::Flow $false

if ($Interactive -and ($UseSetupFile -eq $false)) 
{
	$URL = Input "Url for sitecollection?" $URL

	$Environment = Input "Navn på miljøet?" $Environment
}

if ($Solution -eq "" -or $Solution -eq $null) 
{
	$Solution = Input "Navn på løsningen?" $Solution
}

$global:root = $global:root+"\"+$Solution

LoadEnvironment
LoadInstall

if ($PSScriptRoot -eq $null) {
	$thisScript = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
} else {
	$thisScript = $PSScriptRoot
}
LoadScripts $thisScript
. ($thisScript + '\01_Core\Menu.ps1')

#Vis menu
ShowMenu

if ($PSScriptRoot -eq $null) {
	$thisScript = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
} else {
	$thisScript = $PSScriptRoot
}
UnLoadScripts $thisScript
