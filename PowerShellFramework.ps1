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
	[bool]$ShowDebug=$false
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

#Load scripts
. ($thisScript + '\..\ps\01_Core\Load.ps1')

read-host $global:root