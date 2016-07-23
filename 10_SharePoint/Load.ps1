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
ShowMessage "Loading SharePoint Framework Version 3.0.0" [LogLevels]::Flow

function InitializeSharePoint($thisScript){
    ShowMessage "Initializing SharePoint" [LogLevels]::Information

	if ($global:setupXML.Setup.SiteType -eq "SPOnline") {
        . ($thisScript + "/Online/Load.ps1")
	} else {
        . ($thisScript + "/OnPremise/Load.ps1")
	}
}

if ($PSScriptRoot -eq $null) {
	$thisScript = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
} else {
	$thisScript = $PSScriptRoot
}
InitializeSharePoint $thisScript