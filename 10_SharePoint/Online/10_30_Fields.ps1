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
ShowMessage "Loading SharePoint Online Fields 3.0.0" [LogLevels]::Information

$global:fieldGroups = @{}

<#
Menufunktioner
#>
function global:SP-TestSiteColumns($command) {
    
    write-host (SP-GetFields $command) 
    read-host "TEST"
    # SP-GetTaxonomies 
    # $xml = New-Object -TypeName System.Xml.XmlDocument
    # $xml.LoadXML($global:taxonomyTreeXML);

    # write-host $xml.Taxonomies.TaxonomyGroup.count
}

function global:SP-ExportSiteColumns($command) {
    $fieldsXML = SP-GetFields $command
	$fieldsFile = [System.IO.Path]::Combine($global:logPath, "Fields.xml")
	$fieldsFile = (Input "Sti til felt fil" $fieldsFile)

	$fieldsXML | out-file  -Encoding "UTF8" $fieldsFile 
}

<#
Hjælpefunktioner
#>
function global:SP-GetFields($command)
{
    $groupXML = ""
    $global:fieldGroups = @{}
	$web = $global:clientContext.Web;
	$fields = $web.Fields;
    $global:clientContext.Load($web)
	$global:clientContext.Load($fields);
	$global:clientContext.executeQuery();

    foreach ($field in $fields) {
        ShowMessage (".") [LogLevels]::Information $true;
        if ($global:fieldGroups.Contains($field.Group)) {
            
        } else {
            $global:fieldGroups.Add($field.Group, @{});
        }
        $global:fieldGroups[$field.Group].Add($field.InternalName, $field);
    }

    ShowMessage (".") [LogLevels]::Information;
    ShowMessage ("Genererer XML") [LogLevels]::Information;

    $groupXML += "<Fields><AddFields>`r`n";
    
    foreach ($groupName in $global:fieldGroups.KEYS.GetEnumerator()) {
        $groupXML += "<Group>`r`n";
        $groupXML += "`t<Name>" + $groupName + "</Name>`r`n";
        foreach ($fieldName in $global:fieldGroups[$groupName].KEYS) {
            $groupXML += "`t`t<Field>`r`n";
            $groupXML += "`t`t`t<Name>" + $global:fieldGroups[$groupName][$fieldName].InternalName + "</Name>`r`n";
            $groupXML += "`t`t`t<Description>" + $global:fieldGroups[$groupName][$fieldName].Description + "</Description>`r`n";
            $groupXML += "`t`t`t<FieldType>" + $global:fieldGroups[$groupName][$fieldName].TypeAsString + "</FieldType>`r`n";
            $schema = $global:fieldGroups[$groupName][$fieldName].SchemaXml;
            $xml = New-Object -TypeName System.Xml.XmlDocument
            $xml.LoadXML("<t>" + $schema + "</t>");
            $xml.t.Field.RemoveAttribute("ID")
            $xml.t.Field.RemoveAttribute("SourceID")
            $xml.t.Field.RemoveAttribute("Version")
            $xml.t.Field.RemoveAttribute("Name")
            $xml.t.Field.RemoveAttribute("StaticName")
            $xml.t.Field.RemoveAttribute("Description")
            $xml.t.Field.SetAttribute("DisplayName", "¤Name¤")
            $xml.t.Field.SetAttribute("Group", "¤Group¤")
            $groupXML += "`t`t`t<SchemaXML><![CDATA[" + $xml.t.Field.OuterXML + "]]></SchemaXML>`r`n";
            $groupXML += "`t`t`t<Keep></Keep>`r`n";
            
            
            $groupXML += "`t`t</Field>`r`n";

        }
        $groupXML += "</Group>`r`n";
    }
    $groupXML += "</AddFields></Fields>";
    
    return $groupXML
}
