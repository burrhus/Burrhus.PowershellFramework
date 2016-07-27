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

$global:LookupLists = @{}

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

function global:SP-ImportListContents($command) {
    $lists = SP-GetLists $command

	if ($lists[$command.ListName].ItemCount -gt 0) {
		ShowMessage ($command.ListName + " indeholder allerede " + $listItems.Count  + " elementer!");
		return;
	}

	$listItems = SP-GetListItems $lists[$command.ListName];
	$importFile = [System.IO.Path]::Combine($global:setupXML.Setup.ImportPath,  "Lists", $command.ListName, "Export.xml")

	# Sort the list by ID 
	$importXML = LoadSetupConfig $importFile
	[System.Xml.XmlNode]$orig = $importXML.ListItems
	$orig.ListItem | sort ID -Descending |
  	foreach { [void]$importXML.ListItems.PrependChild($_) }

	foreach ($lookupList in $command.LookupLists.LookupList) {
		if (-not $global:LookupLists.ContainsKey($lookupList)) {
			$global:LookupLists.Add($lookupList, (SP-GetListItems $lists[$lookupList]))
		}	  
	}
	$counter = 0
	foreach ($listItem in $importXML.ListItems.ListItem) {
		ShowMessage "." [LogLevels]::Information $true

		$addElement = $true

		if ($command.AvoidDuplicate -ne $null) {
			foreach ($field in $listItem.Field) {
				if ($field.Name -eq $command.AvoidDuplicate.FieldName) {
					$fieldValue = $field.Value; 
				}
			}
			$items = SP-GetListItemByField $lists[$command.ListName] $command.AvoidDuplicate.FieldName $fieldValue $command.AvoidDuplicate.FieldType
			if ($items.Count -gt 0) {
				$addElement = $false
			}	
		}

		if ($addElement) {
			SP-ImportListItem $command $lists[$command.ListName] $listItem $lists
			$counter++;

			if (($counter %= 10) -eq 0) {
				ShowMessage $counter.ToString() [LogLevels]::Information $true				
			}
		}
	}  
	ShowMessage "." [LogLevels]::Information 
	
}

function global:SP-CreateLists ($command) {
	$web = $global:clientContext.Web;

	foreach ($webDef in $installXML.Install.Webs.Web) 
	{
		$url = Get-EnvironmentVar $webDef.Url $global:setupXML
		$web = $global:clientContext.Site.OpenWeb($url)

		foreach ($list in $webDef.Lists.AddLists.List)
		{
            write-host $list.Name
			SP-CreateList $web $list $webDef
		}
	}
}

function global:SP-EditLists ($command)
{
	if ($update -ne $null)
	{
		$root = $update
	}
	else
	{
		$root = $installXML.Install;
	}

	foreach ($listDef in $root.Lists.AddLists.List)
	{
		$web = $global:clientContext.Web;
		$list = $web.Lists.GetByTitle($listDef.Name)
		$global:clientContext.Load($list)

		SP-EditList $list $listDef
	}

	foreach ($webDef in $root.Webs.Web) 
	{
		foreach ($listDef in $webDef.Lists.AddLists.List)
		{
			$url = Get-EnvironmentVar $webDef.Url $global:setupXML
			$web = $global:clientContext.Site.OpenWeb($url)
			$list = $web.Lists.GetByTitle($listDef.Name)
			$global:clientContext.Load($list)
			SP-EditList  $list $listDef 
		}
	}

}

function global:SP-EditListViews ($command)
{
	if ($update -ne $null)
	{
		$root = $update
	}
	else
	{
		$root = $installXML.Install;
	}

	foreach ($listDef in $root.Lists.AddLists.List)
	{

		$web = $global:clientContext.Web;
		$list = $web.Lists.GetByTitle($listDef.Name)
		$global:clientContext.Load($list)

		foreach ($viewDef in $listDef.Views.View)
		{
			SP-EditListView $list $listDef $viewDef 
		}
	}
	foreach ($webDef in $root.Webs.Web) #.Lists.AddLists.List)
	{
		foreach ($listDef in $webDef.Lists.AddLists.List)
		{
			$url = Get-EnvironmentVar $webDef.Url $global:setupXML
			$web = $global:clientContext.Site.OpenWeb($url)
			$list = $web.Lists.GetByTitle($listDef.Name)
			$global:clientContext.Load($list)

			foreach ($viewDef in $listDef.Views.View)
			{
				SP-EditListView $list $listDef $viewDef
			}
		}
	}
}

<#
Hjælpefunktioner
#>
function global:SP-ImportListItem($command, $list, $listItemContent, $lists) {
	$itemCreateInfo = new-object Microsoft.SharePoint.Client.ListItemCreationInformation;
	$listItem = $list.AddItem($itemCreateInfo);
	if ($command.Fields -eq $null) {
		write-host "ALLE felter er ikke implementeret endnu"	
	} else {
		foreach ($field in $command.Fields.Field) {
			$value = "";
			$inputField = $null;
			foreach ($fieldContent in $listItemContent.Field) {
				if ($fieldContent.Name -eq $field.Name -or $fieldContent.Name -eq $field.DisplayName) {
					$value = $fieldContent.Value
					$inputField = $fieldContent
				}
			}
			#write-host $field.Name ": " $value
			switch ($inputField.Type) {
				"Microsoft.SharePoint.Client.FieldUrlValue" {
					$value  = New-Object Microsoft.SharePoint.Client.FieldUrlValue
					$value.Url = $inputField.Value.Url;
					$value.Description = $inputField.Description;
					$listItem[$field.Name] = [Microsoft.SharePoint.Client.FieldUrlValue]$value;
					break;			
				}
				"Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValueCollection" {
					$taxGroupName = Get-EnvironmentVar $field.TaxGroup $global:setupXML

					$value = "";
					foreach ($inputFieldValue in $inputField.Value.Value) {
						$valueXML = (SP-GetTaxonomy $taxGroupName $field.TaxTermSet $inputFieldValue.Label)
						if ($valueXML.Id -ne $null) {
							$value += "-1;#{0}|{1};#" -f $inputFieldValue.Label, $valueXML.Id;						
						}
					}

					# write-host $value
					$value= $value.Substring(0,$value.Length-2)
					# $termValues = new-object Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValueCollection($global:clientContext, $value, $temaField);
					# $taxfield2 = [Microsoft.SharePoint.Client.ClientContext].GetMethod("CastTo").MakeGenericMethod([Microsoft.SharePoint.Client.Taxonomy.TaxonomyField]).Invoke($global:clientContext, $temaField)
					# $taxfield2.SetFieldValueByValueCollection($fileListItem,$termValues);

					$listItem[$field.Name] = $value;
					
					break;
				}
				"Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue" {
					ShowMessage "Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue er ikke implementeret" [LogLevels]::Warning
				}
				"Microsoft.SharePoint.Client.FieldLookupValue[]" {
					# Write-Host $field.LookupList;

					[Microsoft.SharePoint.Client.FieldLookupValue[]]$newvalue = $();

					foreach ($inputFieldValue in $inputField.Value.Value) {
						if ($inputFieldValue -ne $null) {
							foreach ($itemKey in $global:LookupLists[$field.LookupList].KEYS) {
								$item = $global:LookupLists[$field.LookupList][$itemKey];
								if ($item -ne $null) {
									if ($item.FieldValues["Title"] -eq $inputFieldValue.LookupValue) {
										$addValue = New-Object Microsoft.SharePoint.Client.FieldLookupValue
										$addValue.LookupId = $item.FieldValues["ID"];
										#$addValue.LookupValue = $item.FieldValues["Title"];
										$newvalue += $addValue;
									}
									
								}
							}

						}
					}
					$listItem[$field.Name] = [Microsoft.SharePoint.Client.FieldLookupValue[]]$newvalue;
					break;
				}
				"Microsoft.SharePoint.Client.FieldLookupValue" {
					# Write-Host $field.LookupList;

					[Microsoft.SharePoint.Client.FieldLookupValue[]]$newvalue = $();

					$inputFieldValue = $inputField.Value
						if ($inputFieldValue -ne $null) {
							foreach ($itemKey in $global:LookupLists[$field.LookupList].KEYS) {
								$item = $global:LookupLists[$field.LookupList][$itemKey];
								if ($item -ne $null) {
									if ($item.FieldValues["Title"] -eq $inputFieldValue.LookupValue) {
										$addValue = New-Object Microsoft.SharePoint.Client.FieldLookupValue
										$addValue.LookupId = $item.FieldValues["ID"];
										#$addValue.LookupValue = $item.FieldValues["Title"];
										$newvalue += $addValue;
									}
									
								}
							}

						
					}
					$listItem[$field.Name] = [Microsoft.SharePoint.Client.FieldLookupValue[]]$newvalue;
					break;
				}
				Default {
					if ($value -ne "") {
						$listItem[$field.Name] = $value;						
					}
				}
			}

		}
	}
	$listItem.update()
	try {
		$global:clientContext.ExecuteQuery();		
	}
	catch [System.Exception] {
		ErrorHandling
		ShowMessage $listItem["Title"] [LogLevels]::Warning
	}
}

function global:SP-ImportListDocument() {
	$importPath = $setupXML.Setup.ImportPath+$del.Path
	# setup some convenience variables to keep each line shorter
	#$path = [System.IO.Path]::Combine($Env:TEMP,"Temp.txt")
	$mode = [System.IO.FileMode]::Open
	$access = [System.IO.FileAccess]::Read
	$sharing = [IO.FileShare]::Read

}
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

function global:SP-GetListItemByField($list, $fieldName, $fieldValue, $fieldType) {
    $returnValue = @{}
	if ($fieldType -eq $null) {
		$fieldType="Text"
	}

    $query = New-Object Microsoft.SharePoint.Client.CamlQuery
    $query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='" + $fieldName + "'/><Value Type='" + $fieldType + "'>" + $fieldValue + "</Value></Eq></Where></Query></View>"

	$listItems = $list.GetItems($query);
	$global:clientContext.Load($listItems);
	$global:clientContext.executeQuery();
    foreach ($listItem in $listItems) {
        $returnValue.Add($listItem["ID"], $listItem)
    }

    return $returnValue 
}


function global:SP-CreateList($spWeb, $listDef)
{
	Debug "SP-CreateList"
	Debug $setupXML.Setup.SiteType

	$spListCollection  = $spWeb.Lists   

	Try
	{
		$spLibrary  = $spListCollection.GetByTitle($listDef.Name)
		$global:clientContext.Load($spLibrary);
		$global:clientContext.Load($spListCollection);
		$global:clientContext.Load($spWeb);
		$global:clientContext.executeQuery();
		$test = $listDef.Name;
		Write-Host  "Library " $listDef.Name " already exists in the site" -f green 
	}
	Catch
	{
		$ErrorMessage = $_.Exception.Message
		$FailedItem = $_.Exception.ItemName
		Debug $ErrorMessage ;
		Write-Host  -f  yellow "Creating  Library - "$listDef.Name

		$listTemplates = $spWeb.ListTemplates;
		$global:clientContext.Load($listTemplates);
		$global:clientContext.executeQuery();
		$templateTypeID = 100;
		foreach($lt in $listTemplates)
		{
			if ($lt.Name -eq $listDef.Template -or $lt.InternalName -eq $listDef.TemplateInternalName)
			{
				Debug $lt.Name  " " $lt.InternalName " "  $lt.ListTemplateTypeKind;
				$templateTypeID = $lt.ListTemplateTypeKind;
			}
		}

		$creationInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation;
		$creationInfo.Title = $listDef.Name;
		$creationInfo.Description = $listDef.Description;
		$creationInfo.TemplateType = $templateTypeID;
		$spList = $spWeb.Lists.Add($creationInfo);
		$global:clientContext.executeQuery();

		Write-Host  -f Green "...Success!"
	}
}

function global:SP-EditList($list, $listDef)
{
	$info = "Redigerer " + $listDef.Name
	ShowMessage $info [LogLevels]::Information

	if ($update -ne $null)
	{
		$root = $update
	}
	else
	{
		$root = $global:installXML.Install;
	}
	SP-EditListConfig $list $listDef

	foreach ($editList in $root.Lists.AddLists.List)
	{
		if ($editList.Name -eq $list.Title)
		{
			foreach ($field in $editList.Fields.Field) 
			{
				$spField = $list.Fields.GetByInternalNameOrTitle($field.Name)
				$global:clientContext.Load($spField)
				try {
					$global:clientContext.ExecuteQuery()
					SPEditField  $spField  $field $list.ParentWeb
				} catch {
					SP-AddField $list $field $list.ParentWeb  $setupXML $installXML
				}
			}
			foreach ($field in $editList.ReferenceFields.Field) 
			{
				$spField = $list.Fields.GetByInternalNameOrTitle($field.Name)
				$global:clientContext.Load($spField)
				try {
					$global:clientContext.ExecuteQuery()
					SPEditFieldReference  $spField  $field $list.ParentWeb
				} catch {
					SP-AddFieldReference $list $field $list.ParentWeb  $setupXML $installXML
				}
			}
		} else {  }
	}

}

function global:SP-EditListConfig($list, $listDef)
{
	$info = "Konfigurerer " + $listDef.Name
	ShowMessage $info [LogLevels]::Debug
	$updated = $false;
	if ($listDef.ContentTypesEnabled -ne $null) {
		if ($listDef.ContentTypesEnabled -eq "TRUE") {
			$list.ContentTypesEnabled = $true;
		} else {
			$list.ContentTypesEnabled = $false;
		}
		$updated = $true;
		$list.Update();
		$global:clientContext.ExecuteQuery();
	}

	$web = $global:clientContext.Web;
	$cts = $list.ContentTypes;
	$global:clientContext.Load($cts);
	$global:clientContext.ExecuteQuery();
	ForEach($ctDef in $listDef.ContentTypes.ContentType) {
		$found = $false;

		$contentTypes = $list.ContentTypes; 
		$clientContext.Load($contentTypes)
		$clientContext.ExecuteQuery()
		$existingContentType = Get-ContentType $contentTypes $ctDef

		if ($existingContentType -eq $null) {
			#write-host "Skal adde"
			if ($ctDef.Remove  -eq $null) {
				$contentTypes = $list.ParentWeb.AvailableContentTypes; 
				$clientContext.Load($contentTypes)
				$clientContext.ExecuteQuery()
				$existingContentType = Get-ContentType $contentTypes $ctDef
				SP-AddContentType $list $existingContentType $web
			}

			$updated = $true;
		}
		else
		{
			if ($ctDef.Remove -ne $null) {

				SP-RemoveContentType $list $existingContentType $web
			}

		}
	}

	if ($updated) {
		$list.Update();
		$global:clientContext.ExecuteQuery();
	}
}

function global:SP-EditListView($list, $listDef, $viewDef)
{
	if ($update -ne $null)
	{
		$root = $update
	}
	else
	{
		$root = $global:installXML.Install;
	}

	if ($viewDef.Name -eq "")
	{
		ShowMessage "Defaultview" [LogLevels]::Information
		$views = $list.Views
		$global:clientContext.Load($views)
		$global:clientContext.ExecuteQuery()
		foreach ($view in $views) 
		{
			if ($view.DefaultView)
			{
				$viewFields = $view.ViewFields;
				$global:clientContext.Load($viewFields)
				$global:clientContext.ExecuteQuery()
				break;
			}
		}
	}
	else
	{
		ShowMessage $viewDef.Name [LogLevels]::Information

		$view = $list.Views.GetByTitle($viewDef.Name)

		$viewFields = $view.ViewFields;
		$global:clientContext.Load($viewFields)
		$global:clientContext.Load($list)
		$global:clientContext.Load($view)
		try
		{
			$global:clientContext.ExecuteQuery()
		}
		catch
		{
			Write-Host $viewDef.Name "findes IKKE"
			$creationInfo = New-Object Microsoft.SharePoint.Client.ViewCreationInformation;
			$creationInfo.Title = $viewDef.Name;
			if ($viewDef.RowLimit -ne $null)
			{
				$creationInfo.RowLimit = $viewDef.RowLimit;
			}
			if ($viewDef.ViewFields -ne $null)
			{
				$creationInfo.ViewFields =  $viewDef.ViewFields.Split(",");
			}
			if ($viewDef.Query -ne $null)
			{
				$creationInfo.Query =  $viewDef.Query."#cdata-section";
			}

			$newView  = $list.Views.Add($creationInfo)
			$global:clientContext.Load($newView)
			$global:clientContext.ExecuteQuery()
			return;
		}
	}

	if ($viewDef.ViewFields -ne $null)
	{
		foreach ($field in $viewDef.ViewFields.Split(","))
		{
			$exists = CheckField $field $viewFields
			if (-not $exists) 
			{
				$view.ViewFields.Add($field);
			}
		}
	}
	if ($viewDef.RowLimit -ne $null)
	{
		$view.RowLimit = $viewDef.RowLimit;
	}
	if ($viewDef.Query -ne $null)
	{
		$view.ViewQuery =  $viewDef.Query."#cdata-section";
	}
	$view.Update();
	$list.Update();
	$global:clientContext.ExecuteQuery()

	return
}
