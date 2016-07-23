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

<#
Menufunktioner
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
