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

function TerminateSharePoint(){
    ShowMessage "Terminating SharePoint" [LogLevels]::Information
	if ($global:setupXML.Setup.SiteType -eq "SPOnline") {
	} else {
	    stop-spassignment -Global
	}
}

TerminateSharePoint