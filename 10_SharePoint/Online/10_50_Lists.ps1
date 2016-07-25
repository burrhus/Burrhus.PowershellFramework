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
ShowMessage "Loading SharePoint Online Lists 3.0.0" [LogLevels]::Information

<#
Menufunktioner
#>
function global:SP-TestListContents($command) {
    $lists = SP-GetLists $command
    foreach ($listName in $lists.KEYS) {
        $listXML = "<ListItems>`r`n";
        $listItems = SP-GetListItems $lists[$listName];

        ShowMessage "Genererer XML" [LogLevels]::Information $true

        foreach ($listItemID in $listItems.KEYS) {
            ShowMessage "." [LogLevels]::Information $true 
            
            $listXML += "`t<ListItem>`r`n";
                $listXML += "`t`t<ID>" + $listItemID + "</ID>`r`n";
                foreach ($fieldName in $listItems[$listItemID].FieldValues.KEYS) {
                    $listXML += "`t`t<Field>`r`n";
                        $listXML += "`t`t`t<Name>" + $listItems[$listItemID].FieldValues[$fieldName] + "</Name>`r`n";
                        $listXML += "`t`t`t<Value>" + $listItems[$listItemID].FieldValues[$fieldName] + "</Value>`r`n";
                    $listXML += "`t`t</Field>`r`n";            
                }
            $listXML += "`t</ListItem>`r`n";
        }
        $listXML += "</ListItems>`r`n";
        ShowMessage "." [LogLevels]::Information 
        ShowMessage $listXML [LogLevels]::Information 
        
        write-host $listXML;
    }


    # SP-GetTaxonomies 
    # $xml = New-Object -TypeName System.Xml.XmlDocument
    # $xml.LoadXML($global:taxonomyTreeXML);

    # write-host $xml.Taxonomies.TaxonomyGroup.count
}

function global:SP-ExportListContents($command) {
    $lists = SP-GetLists $command
    foreach ($listName in $lists.KEYS) {
        #if ($listName -eq "Configuration" -or $listName -eq "Dokumenter") {
            ShowMessage ("Exporting " + $listName) [LogLevels]::Flow 
            $listItems = SP-GetListItems $lists[$listName];
            
            $listXML = ConvertListContent2XML $listItems
            $listFile = [System.IO.Path]::Combine($global:logPath, "Lists", $listName)
            if (-not (Test-Path $listFile)) {
                md $listFile                 
            }
            foreach ($listItem in $listItems.KEYS) {
                SP-DownloadAttachments $listItems[$listItem] $listFile
            }
            
            $listFile = [System.IO.Path]::Combine($global:logPath, "Lists", $listName, "Export.xml")
            if ($interactive) {
                $listFile = (Input "Sti til felt fil" $listFile)                
            }

            $listXML | out-file  -Encoding "UTF8" $listFile 
        #}
    }
}


<#
Hjælpefunktioner
#>

function global:ConvertListContent2XML($listItems) {
    $listXML = "<ListItems>`r`n";

    ShowMessage "Genererer XML" [LogLevels]::Information $true

    foreach ($listItemID in $listItems.KEYS) {
        ShowMessage "." [LogLevels]::Information $true 
        
        $listXML += "`t<ListItem>`r`n";
            $listXML += "`t`t<ID>" + $listItemID + "</ID>`r`n";
            foreach ($fieldName in $listItems[$listItemID].FieldValues.KEYS) {
                $value = "";
                $listXML += "`t`t<Field>`r`n";
                    $listXML += "`t`t`t<Name>" + $fieldName + "</Name>`r`n";
                    if ($listItems[$listItemID].FieldValues[$fieldName] -ne $null) {
                        switch ($listItems[$listItemID].FieldValues[$fieldName].ToString()) {
                            "Microsoft.SharePoint.Client.FieldLookupValue" { 
                                $listXML += "`t`t`t<Type>Microsoft.SharePoint.Client.FieldLookupValue</Type>`r`n";
                                $value = "<LookupId>" + $listItems[$listItemID].FieldValues[$fieldName].LookupId + "</LookupId>"
                                $value += "<LookupValue>" + $listItems[$listItemID].FieldValues[$fieldName].LookupValue + "</LookupValue>"
                                break;
                            }
                            "Microsoft.SharePoint.Client.FieldUserValue" { 
                                $listXML += "`t`t`t<Type>Microsoft.SharePoint.Client.FieldUserValue</Type>`r`n";
                                $value = "<LookupId>" + $listItems[$listItemID].FieldValues[$fieldName].LookupId + "</LookupId>"
                                $value += "<LookupValue>" + $listItems[$listItemID].FieldValues[$fieldName].LookupValue + "</LookupValue>"
                                break;
                            }
                            "Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue" {
                                $listXML += "`t`t`t<Type>Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue</Type>`r`n";
                                $value = "<Label>" + $listItems[$listItemID].FieldValues[$fieldName].Label + "</Label>"
                                $value += "<TermGuid>" + $listItems[$listItemID].FieldValues[$fieldName].TermGuid + "</TermGuid>"
                                $value += "<WssId>" + $listItems[$listItemID].FieldValues[$fieldName].WssId + "</WssId>"

                                break;
                            }
                            "Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValueCollection" {
                                $value = "";
                                $listXML += "`t`t`t<Type>" + $listItems[$listItemID].FieldValues[$fieldName].ToString() + "</Type>`r`n";

                                foreach ($item in $listItems[$listItemID].FieldValues[$fieldName]) {
                                    $value += "`r`n`t`t`t`t<Value>"
                                    $value += "`r`n`t`t`t`t<Label>" + $item.Label + "</Label>"
                                    $value += "`r`n`t`t`t`t<TermGuid>" + $item.TermGuid + "</TermGuid>"
                                    $value += "`r`n`t`t`t`t<WssId>" + $item.WssId + "</WssId>"
                                    $value += "`r`n`t`t`t`t</Value>"                                    
                                }
                                $value += "`r`n"                                    
                                break;
                            }
                            "Microsoft.SharePoint.Client.FieldUrlValue" {
                                $listXML += "`t`t`t<Type>Microsoft.SharePoint.Client.FieldUrlValue</Type>`r`n";
                                $value = "<Url>" + $listItems[$listItemID].FieldValues[$fieldName].Url + "</Url>"
                                $value += "<Description>" + $listItems[$listItemID].FieldValues[$fieldName].Description + "</Description>"
                                break;
                            }
                            "Microsoft.SharePoint.Client.FieldLookupValue[]" {
                                $value = "";
                                $listXML += "`t`t`t<Type>" + $listItems[$listItemID].FieldValues[$fieldName].ToString() + "</Type>`r`n";

                                foreach ($item in $listItems[$listItemID].FieldValues[$fieldName]) {
                                    $value += "`r`n`t`t`t`t<Value>"
                                    $value += "`r`n`t`t`t`t<LookupId>" + $item.LookupId + "</LookupId>"
                                    $value += "`r`n`t`t`t`t<LookupValue>" + $item.LookupValue + "</LookupValue>"
                                    $value += "`r`n`t`t`t`t</Value>"                                    
                                }
                                $value += "`r`n"                                    
                                break;
                            }
                            Default {
                                $value = $listItems[$listItemID].FieldValues[$fieldName];
                            }
                        }
                    } else {
                        $value = "";
                    }

                    $listXML += "`t`t`t<Value>" + $value + "</Value>`r`n";
                $listXML += "`t`t</Field>`r`n";            
            }
        $listXML += "`t</ListItem>`r`n";
    }
    $listXML += "</ListItems>`r`n";
    ShowMessage "." [LogLevels]::Information 
    return $listXML
}

function global:SP-GetLists($command) {
    $returnValue = @{}
    ShowMessage "Henter lister" [LogLevels]::Information $true

	$web = $global:clientContext.Web;
	$lists = $web.Lists;
    $global:clientContext.Load($web)
	$global:clientContext.Load($lists);
	$global:clientContext.executeQuery();
    foreach ($list in $lists) {
        ShowMessage "." [LogLevels]::Information $true
        $returnValue.Add($list.Title, $list)
    }
    ShowMessage "." [LogLevels]::Information

    return $returnValue 
}

function global:SP-GetListItems($list) {
    $returnValue = @{}
    ShowMessage "Henter liste elementer" [LogLevels]::Information $true

    $query = New-Object Microsoft.SharePoint.Client.CamlQuery
    $query.ViewXml = "<View/>"

	$listItems = $list.GetItems($query);
	$global:clientContext.Load($listItems);
	$global:clientContext.executeQuery();
    foreach ($listItem in $listItems) {
        ShowMessage "." [LogLevels]::Information $true
        $returnValue.Add($listItem["ID"], $listItem)
    }
    ShowMessage "." [LogLevels]::Information

    return $returnValue 
}
