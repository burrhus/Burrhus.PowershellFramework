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

$global:taxonomyTreeXML = "";
$global:taxonomyTree = @{}
ShowMessage "Loading SharePoint Online Taxonomy Version 3.0.0" [LogLevels]::Information

<#
Menufunktioner
#>
function global:SP-TestTaxonomies($command) {
    SP-GetTaxonomies 
    $xml = New-Object -TypeName System.Xml.XmlDocument
    $xml.LoadXML($global:taxonomyTreeXML);

    write-host $xml.Taxonomies.TaxonomyGroup.count
}

function global:SP-ExportTaxonomies($command) {
    SP-GetTaxonomies 
	$taxonomyFile = [System.IO.Path]::Combine($global:logPath, "taxonomy.xml")
	$taxonomyFile = (Input "Sti til taxonomifil" $taxonomyFile)

	$global:taxonomyTreeXML | out-file $taxonomyFile 
}

function global:SP-ImportTaxonomies($command) {
	$taxXMLFile = [System.IO.Path]::Combine($global:setupXML.Setup.ImportPath, "taxonomy.xml");
	$taxXML = LoadSetupConfig $taxXMLFile

	if ($taxXML.Taxonomies.TaxonomyGroup.Name.ToString().StartsWith("[")) {
		$taxGroupName = Get-EnvironmentVar $taxXML.Taxonomies.TaxonomyGroup.Name $global:setupXML
	} else {
		$taxGroupName =$taxXML.Taxonomies.TaxonomyGroup.Name
	}

	SP-CreateTaxonomyGroups $command $taxXML
	Write-Host $taxGroupName
	Read-Host $taxXMLFile
}

<#
Hjælpefunktioner
#>

function global:SP-GetTaxonomies($taxonomyGroupDef, $termSetDef, $returnValue)
{
	$termStore = SP-GetTermStoreInfo $null 
	SP-GetTaxonomyTree $global:taxonomyTree $termStore
}

function global:SP-GetTermStoreInfo($command)
{
	$spTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($global:clientContext)
	$spTaxSession.UpdateCache();
	$global:clientContext.Load($spTaxSession)
  
	try
	{
		$global:clientContext.ExecuteQuery()
	}
	catch
	{
		ShowDebug ("Error while loading the Taxonomy Session " + $_.Exception.Message) [LogLevels]::Error
		exit 1
	}
  
	if($spTaxSession.TermStores.Count -eq 0){
		ShowDebug "The Taxonomy Service is offline or missing" [LogLevels]::Error
		exit 1
	}
  
	$termStores = $spTaxSession.TermStores
	$global:clientContext.Load($termStores)
  
	try
	{
		$global:clientContext.ExecuteQuery()
		$termStore = $termStores[0]
		$global:clientContext.Load($termStore)
		$global:clientContext.ExecuteQuery()
		ShowMessage "Connected to TermStore: $($termStore.Name) ID: $($termStore.Id)" [LogLevels]::Information
	}
	catch
	{
		ShowMessage ( "Error details while getting term store ID" + $_.Exception.Message) [LogLevels]::Error
		exit 1
	}
 
	return $termStore	
}

function global:SP-GetTaxonomyTree($taxonomyTree, $termStore)
{
	ShowMessage "Get TaxonomyTree" [LogLevels]::Flow
	$groups = @{}
	$global:taxonomyTreeXML += "<Taxonomies>"

	$termGroups = $termStore.Groups;
	$global:clientContext.Load($termGroups);
	$global:clientContext.ExecuteQuery();
	foreach ($group in $termGroups)
	{
		if (-not $group.IsSystemGroup)
		{
			ShowMessage ("GroupName: " + $group.Name) [LogLevels]::Information
			$groupList = @{};
			$global:taxonomyTreeXML += "<TaxonomyGroup>"
			$global:taxonomyTreeXML += "<Name>" + $group.Name + "</Name>";
			$global:taxonomyTreeXML += "<Description>" + $group.Description + "</Description>"
			$global:taxonomyTreeXML += "<Id>" + $group.Id + "</Id>"
			$groups.Add("NAME"+$group.Id, $group.Name);
			$global:taxonomyTreeXML += "<TermSets>"
			SP-GetTermSets $group $groupList
			$global:taxonomyTreeXML += "</TermSets>"
			$groups.Add($group.Id, $groupList)
			$global:taxonomyTreeXML += "</TaxonomyGroup>"
		}			
	}

	$taxonomyTree.Add("Root", $groups);

	$global:taxonomyTreeXML += "</Taxonomies>"
}

function global:SP-GetTermSets($taxGroup, $returnValue)
{	
	$termSets = $taxGroup.TermSets;
	$global:clientContext.Load($termSets);
	$global:clientContext.ExecuteQuery();

	$returnValue.Add($taxGroup.Name, $taxGroup.Id)

	foreach ($termSet in $termSets)
	{
		$global:taxonomyTreeXML += "<TermSet>"
		$global:taxonomyTreeXML += "<Name>" + $termSet.Name + "</Name>"
		$global:taxonomyTreeXML += "<Description>" + $termSet.Description + "</Description>"
		$global:taxonomyTreeXML += "<Id>" + $termSet.Id + "</Id>"
		$global:taxonomyTreeXML += "<IsAvailableForTagging>true</IsAvailableForTagging>"
        $global:taxonomyTreeXML += "<IsOpenForTermCreation>true</IsOpenForTermCreation>"

		$termsetList = @{};
		SP-GetTerms $termSet $termsetList
		$global:taxonomyTreeXML += "</TermSet>"
		$returnValue.Add($termSet.Id, $termsetList)
		ShowMessage ("TermsetName: " + $group.Name) [LogLevels]::Debug
	}
	#$returnValue
}

function global:SP-GetTerms($termSet, $returnValue)
{
	$terms = $termSet.Terms;
	$global:clientContext.Load($terms);
	$global:clientContext.ExecuteQuery();

	$returnValue.Add($termSet.Name, $termSet.Id)
	#$returnValue;
	$global:taxonomyTreeXML += "<Terms>"
	foreach ($term in $terms)
	{
		$termsetList = @{};
		$global:taxonomyTreeXML += "<TermSet>"
		$global:taxonomyTreeXML += "<Name>" + $term.Name + "</Name>"
		$global:taxonomyTreeXML += "<Id>" + $term.Id + "</Id>"
		$global:taxonomyTreeXML += "<IsAvailableForTagging>true</IsAvailableForTagging>"
		SP-GetTerms $term $termsetList
		$global:taxonomyTreeXML += "</TermSet>"
		$returnValue.Add($term.Id, $termsetList)
		ShowMessage ("TermName: " + $term.Name) [LogLevels]::Debug
		ShowMessage (".") [LogLevels]::Information $true
	}
	$global:taxonomyTreeXML += "</Terms>"
	#$returnValue;
	
}

function global:SP-CreateTaxonomyGroups($command, $taxXML) {
	$termStore = SP-GetTermStoreInfo $command
	ShowMessage $taxXML.Taxonomies.TaxonomyGroup.Name [LogLevels]::Flow
    foreach($taxonomyGroupDef in $taxXML.Taxonomies.TaxonomyGroup)
    {
		$groupName = Get-EnvironmentVar $taxonomyGroupDef.Name $global:setupXML
		ShowMessage ($groupName + " --- " + $taxonomyGroupDef.Name) [LogLevels]::Flow
		$group = SP-CreateTaxGroup $taxonomyGroupDef $termStore
		foreach ($termSetDef in $taxonomyGroupDef.TermSets.TermSet)
		{
			$termSet = Create-TermSet $termSetDef $termStore $group 
			$term = Create-Terms $termSetDef $termStore $termSet
		}
	 }
}

function global:SP-CreateTaxGroup($taxonomyGroupDef, $termStore)
{
	$groupName = Get-EnvironmentVar $taxonomyGroupDef.Name $global:setupXML
	$description = $taxonomyGroupDef.Description;

	ShowMessage ("Processing Group: " + $groupName + "...") [LogLevels]::Information $true
  
	$groups = $termStore.Groups;
	$global:clientContext.Load($groups);
	$global:clientContext.ExecuteQuery();

	foreach($group in $groups)
	{
		if ($group.Name -eq $groupName) 
		{
			return $group;
		}
	}

	$groupGuid = [System.Guid]::NewGuid()
	ShowMessage ($groupGuid) [LogLevels]::Debug

	$group = $termStore.CreateGroup($groupName, $groupGuid);
    $global:clientContext.Load($group);
    try
    {
        $global:clientContext.ExecuteQuery();
	# 	$group;
	# 	read-host "TEST"
	 	return $group;
        ShowMessage ( "Inserted" ) [LogLevels]::Information
    }
    catch
    {
        ShowMessage ( "Error creating new Group " + $name + " " + $_.Exception.Message ) [LogLevels]::Error
        exit 1
    }
}

function global:Create-TermSet($termSetDef, $termStore, $group) 
{
    $errorOccurred = $false
	$name = $termSetDef.Name;
    $description = $termSetDef.Description;
    $customSortOrder = $termSetDef.CustomSortOrder;
    ShowMessage ("Processing TermSet " + $name + "... " ) [LogLevels]::Information $true
    $termSets = $group.TermSets; #GetTermSet($id);
    $global:clientContext.Load($termSets);
	$global:clientContext.ExecuteQuery();
	foreach ($termSet in $termSets)
	{
		ShowMessage ($termSet.Name) [LogLevels]::Information $true
		if ($termSet.Name -eq $name)
		{
			return $termSet;
		}
	}


	$id = [System.Guid]::NewGuid()
    $termSet = $group.CreateTermSet($name, $id, $termStore.DefaultLanguage);
    $termSet.Description = $description;
 
    if($customSortOrder -ne $null)
    {
        $termSet.CustomSortOrder = $customSortOrder
    }
 
    $termSet.IsAvailableForTagging = [bool]::Parse($termSetDef.IsAvailableForTagging);
    $termSet.IsOpenForTermCreation = [bool]::Parse($termSetDef.IsOpenForTermCreation);
 
    if($termSetDef.CustomProperties -ne $null)
    {
        foreach($custProp in $termSetDef.CustomProperties.CustomProperty)
        {
            $termSet.SetCustomProperty($custProp.Key, $custProp.Value)
        }
    }
  
    try
    {
        $global:clientContext.ExecuteQuery();
    }
    catch
    {
        ShowMessage ("Error occured while create Term Set" + $name + $_.Exception.Message) [LogLevels]::Error
        $errorOccurred = $true
    }
	ShowMessage ("") [LogLevels]::Information
	return $termSet;
}

function global:Create-Terms($termDef, $termStore, $termSet) 
{
	foreach ($term in $termDef.Terms.TermSet)
	{
		$subTerms = $termSet.Terms
		$global:clientContext.Load($subTerms);
		$global:clientContext.ExecuteQuery();
		$subTerm = Create-Term $term $termStore $termSet $subTerms
	}
}

function global:Create-Term($termDef, $termStore, $termSet, $subTerms) 
{
	foreach ($subTerm in $subTerms)
	{
 		if ($subTerm.Name -eq $termDef.Name)
		{
			Create-Terms $termDef $termStore $subTerm
			return $subTerm
		}			
	}

	$id = [System.Guid]::NewGuid()
	$subTerm = $termSet.CreateTerm($termDef.Name, $termStore.DefaultLanguage, $id);

	$customSortOrder = $termDef.CustomSortOrder;
    $description = $termDef.Description;
    $subTerm.SetDescription($description, $termStore.DefaultLanguage);
    $subTerm.IsAvailableForTagging = [bool]::Parse($termDef.IsAvailableForTagging);
 
 
    if($customSortOrder -ne $null)
    {
        $subTerm.CustomSortOrder = $customSortOrder
    }
 
    if($termDef.CustomProperties -ne $null)
    {
        foreach($custProp in $termDef.CustomProperties.CustomProperty)
        {
            $subTerm.SetCustomProperty($custProp.Key, $custProp.Value)
        }
    }
 
    if($termDef.LocalCustomProperties -ne $null)
    {
        foreach($localCustProp in $termDef.LocalCustomProperties.LocalCustomProperty)
        {
            $subTerm.SetLocalCustomProperty($localCustProp.Key, $localCustProp.Value)
        }
    }
 
    try
    {
        $global:clientContext.Load($subTerm);
        $global:clientContext.ExecuteQuery();
        ShowMessage ("Created " + $termDef.Name) [LogLevels]::Information
    }
    catch
    {
        ShowMessage ( "Error occured while create Term" + $name + $_.Exception.Message) [LogLevels]::Error
        $errorOccurred = $true
    }

	Create-Terms $termDef $termStore $subTerm $setupXML $installXML $outFile	
	return $subTerm;
}
