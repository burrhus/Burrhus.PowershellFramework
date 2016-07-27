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
ShowMessage "Loading SharePoint Online Version 3.0.0" [LogLevels]::Information

<#
Log på 
#>
    $global:invDef = $MyInvocation.MyCommand.Definition 
    $thisScript = Split-Path -Path $global:invDef -Parent
    write-host "Initializing SharePoint online" $thisScript -ForegroundColor Green
    $thisScript = Split-Path -Path $global:invDef -Parent
    add-type -Path "$thisScript\Microsoft.SharePoint.Client.dll"
    add-type -Path "$thisScript\Microsoft.SharePoint.Client.Runtime.dll"
    add-type -Path "$thisScript\Microsoft.SharePoint.Client.Taxonomy.dll"
	$global:clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($global:setupXML.Setup.SiteURL)
	$username = $global:setupXML.Setup.SiteUser;
	if ($username -eq $null)
	{
		$username = Input "Brugernavn" $username
	}
	$password = $global:setupXML.Setup.SitePassword
	if ($password -eq $null)
	{
		$password = Input "Password" $password
	}
	$domain = $global:setupXML.Setup.SiteDomain
	#Konverter kodeord 
	$securePassword = ConvertTo-SecureString $password -AsPlainText -Force 
	if ($domain -ne $null) 
	{
	    $global:credentials = New-Object System.Net.NetworkCredential($username, $securePassword, $domain)
	}
	else
	{
	    $global:credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $securePassword)
	}
		
    $global:clientContext.Credentials = $global:credentials

    #Tjek om login er valid og vi har adgang til SharePoint serverne
    if (!$clientContext.ServerObjectIsNull.Value) 
    {             
        Write-Host "Connected to SharePoint Online site: '"$global:setupXML.Setup.siteURL.ToString()"'" -ForegroundColor Green 
    } 
	$global:site = $global:clientContext.Site;
	$global:clientContext.Load($global:site) 
	$global:clientContext.executeQuery() 

if ($PSScriptRoot -eq $null) {
	$thisScript = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
} else {
	$thisScript = $PSScriptRoot
}


. ($thisScript + "\10_25_Taxonomy.ps1")
. ($thisScript + "\10_30_Fields.ps1")
. ($thisScript + "\10_40_ContentTypes.ps1")
. ($thisScript + "\10_50_Lists.ps1")
. ($thisScript + "\10_55_Files.ps1")