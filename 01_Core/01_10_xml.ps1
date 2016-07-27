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
ShowMessage "Loading XML Version 3.0.0" [LogLevels]::Information

function global:LoadSetupConfig([string] $fileName)
{
    <#
    <Summary>
        <Description>Henter konfigurationsXML</Description>
        <Parameters>
            <Parameter Name="$fileName" Description="Sti og filnavn på xmlfilen" />
        </Parameters>
        <Returns>XML Object</Returns>
    </Summary>    
    #>
    
	ShowMessage ("Reading " + $fileName) [LogLevels]::Information

	[xml]$setupXML = gc $fileName # -ErrorAction:SilentlyContinue
	while ($setupXML -eq $null) { 
		ShowMessage "Kan ikke finde filen" [LogLevels]::Error

		$filename = Input "Indtast sti til XML (skriv x for at afslutte)" $fileName 
		if ($fileName -eq "x") 
		{
			exit
		}
		[xml]$setupXML = gc $fileName -ErrorAction:SilentlyContinue

		ErrorHandling;
	}

	return $setupXML
}

function CreateXML([string] $fileName, [string] $rootElement)
{
    <#
    <Summary>
        <Description>Opretter XML</Description>
        <Parameters>
            <Parameter Name="$fileName" Description="Sti og filnavn på xmlfilen" />
            <Parameter Name="$rootElement" Description="Navn på rodelementet" />
        </Parameters>
        <Returns>XML Object</Returns>
    </Summary>    
    #>
	ShowMessage ("Creating report") [LogLevels]::Information
	# Create a new XML File with config root node
	[System.XML.XMLDocument]$oXMLDocument=New-Object System.XML.XMLDocument
	# New Node
	[System.XML.XMLElement]$oXMLRoot=$oXMLDocument.CreateElement($rootElement)
	# Append as child to an existing node
	$oXMLDocument.appendChild($oXMLRoot)
	# Add a Attribute
	$oXMLRoot.SetAttribute("description","Config file for testing")
	[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("system"))
	$oXMLSystem.SetAttribute("description","Document Management")
	# Save File
	$oXMLDocument.Save($fileName)
}

function LoadEnvironment()
{
	$setupFile = $global:root + "\Setup" + $Environment + ".xml"

	$global:setupXML = LoadSetupConfig $setupFile

	$date = get-date;
	$global:logPath =  $global:root + "\" +  + $date.Year.ToString("0000") + $date.Month.ToString("00") + $date.Day.ToString("00") + "_" +  + $date.Hour.ToString("00") + $date.Minute.ToString("00")  + $date.Second.ToString("00");
	md $global:logPath
}

function LoadInstall()
{
	$setupFile = $global:root + "\Install.xml"

	$global:installXML = LoadSetupConfig $setupFile	
}