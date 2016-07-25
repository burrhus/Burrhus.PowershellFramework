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
ShowMessage "Loading SharePoint Online Contenttypes 3.0.0" [LogLevels]::Information

<#
Menufunktioner
#>
function global:SP-TestContentTypes($command) {
    write-host (SP-GetFields $command) 
    read-host "TEST"
    # SP-GetTaxonomies 
    # $xml = New-Object -TypeName System.Xml.XmlDocument
    # $xml.LoadXML($global:taxonomyTreeXML);

    # write-host $xml.Taxonomies.TaxonomyGroup.count
}

function global:SP-ExportContentTypes($command) {
    $fieldsXML = SP-GetFields $command
	$fieldsFile = [System.IO.Path]::Combine($global:logPath, "Fields.xml")
	$fieldsFile = (Input "Sti til felt fil" $fieldsFile)

	$fieldsXML | out-file  -Encoding "UTF8" $fieldsFile 
}
