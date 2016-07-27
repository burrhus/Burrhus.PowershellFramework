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

function global:SP-CreateFields ($command)
{
	ShowMessage "Opretter felter" [LogLevels]::Flow
	$web = $global:clientContext.Web;
	$global:clientContext.Load($web)
	$global:clientContext.executeQuery();

	foreach ($field in $global:installXML.Install.Fields.AddFields.Field)
	{
		SP-AddField $web $field $web
	}

	foreach ($webDef in $global:installXML.Install.Webs.Web) 
	{
		$url = Get-EnvironmentVar $webDef.Url $global:setupXML
		$web = $global:clientContext.Site.OpenWeb($url)

		foreach ($field in $webDef.Fields.AddFields.Field)
		{
			SP-AddField $web $field $web
		}
	}
}

function global:SP-CreateAdditionalFields ($command)
{
	$web = $global:clientContext.Web;
	$global:clientContext.Load($web)
	$global:clientContext.executeQuery();

	foreach ($field in $global:installXML.Install.Fields.AddAdditionalFields.Field)
	{
		SP-AddField $web $field $web 
	}

	foreach ($webDef in $global:installXML.Install.Webs.Web) 
	{
		$url = Get-EnvironmentVar $webDef.Url $setupXML
		$web = $global:clientContext.Site.OpenWeb($url)

		foreach ($field in $webDef.Fields.AddAdditionalFields.Field)
		{
			SP-AddField $web $field $web
		}
	}
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

function global:SP-AddField($receiverObject, $fieldDef, $web) {
	ShowMessage ("Tilføjer " + $fieldDef.Name) [LogLevels]::Debug

	$spField = $receiverObject.Fields.GetByInternalNameOrTitle($fieldDef.Name);
	$global:clientContext.Load($spField)
	Try
	{
		$global:clientContext.executeQuery();
		 
		ShowMessage ($fieldDef.Name + "exists in the site - tjekker opsætning!") [LogLevels]::Information
		# SP-CheckFieldSetup $spField $fieldDef $web $setupXML $installXML
	}
	Catch
	{
		
		ShowMessage ($fieldDef.Name + " does not exist in the site") [LogLevels]::Warning
		
		ShowMessage ($fieldDef.SchemaXML."#cdata-section") [LogLevels]::Debug

		if ($fieldDef.LookupList -ne $null) {
			
			ShowMessage ("Finder lookuplist " + $fieldDef.LookupList) [LogLevels]::Debug

			$l = $web.Lists.GetByTitle($fieldDef.LookupList)
			$listID = $l.ID
			$global:clientContext.Load($l);
			$global:clientContext.ExecuteQuery()
		}

		ShowMessage ("Creating "+$field.Name) [LogLevels]::Warning
		
		$fieldXML = $fieldDef.SchemaXML; 
		if ($fieldXML."#cdata-section" -eq $null) {
			$fieldXML = $fieldXML -replace "&gt;", ">"
			$fieldXML = $fieldXML -replace "&lt;", "<"
		} else {
			$fieldXML = $fieldXML."#cdata-section"
		}
		$fieldXML = $fieldXML -replace ";ListID;", $l.ID
		$group = Get-EnvironmentVar "[Group]" $setupXML
		$fieldXML = $fieldXML -replace ";Group;", $group
		$fieldXML = $fieldXML -replace "¤Name¤", $fieldDef.Name
		ShowMessage $fieldXML [LogLevels]::Debug
		
		$spField = $receiverObject.Fields.AddFieldAsXml($fieldXML, $false, 1);
		$clientContext.Load($spField);
		$clientContext.ExecuteQuery()
		
		if ($fieldDef.FieldType -eq "TaxonomyFieldType") {
			SP-AddFieldTaxonomy $receiverObject $fieldDef $web $spField
		}
		if ($fieldDef.FieldType -eq "TaxonomyFieldTypeMulti") {
			SP-AddFieldTaxonomy $receiverObject $fieldDef $web $spField
		}
	}

	$updated = $false;
	if ($fieldDef.DisplayName -ne $null)
	{
		ShowMessage "Sætter displayname" [LogLevels]::Debug
		$spField.Title = $fieldDef.DisplayName
		$updated = $true;
	}
	if ($fieldDef.Description -ne $null)
	{
		ShowMessage "Sætter Beskrivelse" [LogLevels]::Debug
		$spField.Description = $fieldDef.Description
		$updated = $true;
	}

	if ($fieldDef.LookupList -ne $null) {

		$info = "Finder lookuplist " + $fieldDef.LookupList;
		ShowMessage $info [LogLevels]::Debug

		$l = $web.Lists.GetByTitle($fieldDef.LookupList)
		$listID = $l.ID
		$global:clientContext.Load($l);
		$global:clientContext.ExecuteQuery()

		$schemaXML = ([xml]($spfield.SchemaXml));
		$schemaXML.Field.List = $listID.ToString() 
		$spfield.SchemaXml = $schemaXML.OuterXml
		$updated = $true;
	}

	if ($fieldDef.JSLink -ne $null)
	{
		ShowMessage "Sætter JSLink" [LogLevels]::Debug
		write-host $fieldDef.JSLink
		$spField.JSLink = $fieldDef.JSLink
		$updated = $true;
	}
	$spField.Update()
	$web.Update()
	$clientContext.ExecuteQuery();
}

function global:SP-CheckField ($owner, $fieldDef)
{
	
	$field = $owner.Fields.GetByInternalNameOrTitle($fieldDef.Name);
	$global:clientContext.Load($field)
	Try
	{
		$global:clientContext.executeQuery();
		write-host -f green $fieldDef.Name "exists in " $owner.Name $owner.Title
		#SP-CheckListSetup $list $listDef
	}
	Catch
	{
		write-host -f red  $fieldDef.Name " does not exist in " $owner.Name $owner.Title
		#write-host -f red  $fieldDef.SchemaXML."#cdata-section"
	}
}

function global:SP-AddFieldTaxonomy($receiverObject, $fieldDef, $web, $taxField) {
	$session = 	[Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($global:clientContext)
	$session.UpdateCache();
	$global:clientContext.Load($session)
	try
	{
		$global:clientContext.ExecuteQuery()
	}
	catch
	{
		Write-host "Error while loading the Taxonomy Session " $_.Exception.Message -ForegroundColor Red
		exit 1
	}
  
	if($session.TermStores.Count -eq 0){
		write-host "The Taxonomy Service is offline or missing" -ForegroundColor Red
		exit 1
	}
  
	$termStores = $session.TermStores
	$global:clientContext.Load($termStores)
  
	try
	{
		$global:clientContext.ExecuteQuery()
		$termStore = $termStores[0]
		$global:clientContext.Load($termStore)
		$global:clientContext.ExecuteQuery()
		Write-Host "Connected to TermStore: $($termStore.Name) ID: $($termStore.Id)"
	}
	catch
	{
		Write-host "Error details while getting term store ID" $_.Exception.Message -ForegroundColor Red
		exit 1
	}
 
	 $groupName = Get-EnvironmentVar $fieldDef.TermGroupName $global:setupXML
	$termGroup = $termStore.Groups.GetByName($groupName);
	$global:clientContext.Load($termGroup)

	try
	{
		$global:clientContext.ExecuteQuery()
		Write-Host "Connected to TermGroup: $($termGroup.Name) ID: $($termGroup.Id)"
	}
	catch
	{
		Write-host "Error details while getting term group ID" $_.Exception.Message -ForegroundColor Red
		exit 1
	}
	$termSet = $termGroup.TermSets.GetByName($fieldDef.TermSetName);
	$global:clientContext.Load($termSet)

	try
	{
		$global:clientContext.ExecuteQuery()
		Write-Host "Connected to TermSet: $($termSet.Name)"
	}
	catch
	{
		Write-host "Error details while getting term group ID" $_.Exception.Message -ForegroundColor Red
		exit 1
	}

	$taxfield2 = [Microsoft.SharePoint.Client.ClientContext].GetMethod("CastTo").MakeGenericMethod([Microsoft.SharePoint.Client.Taxonomy.TaxonomyField]).Invoke($global:clientContext, $taxField)
    $taxfield2.SspId = $termStore.Id;
    $taxfield2.TermSetId = $termSet.Id;
	$taxField2.Open = $true;
	$taxField2.CreateValuesInEditForm = $true;
    $taxfield2.Update();
    $global:clientContext.ExecuteQuery();
}

function global:SP-AddFieldReference($receiverObject, $fieldDef, $web) {
	$field = $receiverObject.Fields.GetByInternalNameOrTitle($fieldDef.Name);
	$global:clientContext.Load($field)
	Try
	{
		$global:clientContext.executeQuery();
		write-host -f green $fieldDef.Name "exists in " $receiverObject.Title $receiverObject.Name
	}
	Catch
	{
		write-host -f red  $fieldDef.Name " does not exist in" $receiverObject.Title $receiverObject.Name
		$field = $web.AvailableFields.GetByInternalNameOrTitle($fieldDef.Name);
		$global:clientContext.Load($field)
		$global:clientContext.executeQuery();
		$creationInfo = New-Object Microsoft.SharePoint.Client.FieldLinkCreationInformation;
		$creationInfo.Field = $field;

		$receiverObject.FieldLinks.Add($creationInfo);
		$receiverObject.Update($true);
		$clientContext.ExecuteQuery();
	}
}

function global:CheckField($fieldName, $viewFields)
{
	foreach ($viewField in $viewFields)
	{
		if ($viewField -eq $fieldName)
		{
			return $true
		}
	}
	return $false;	
}