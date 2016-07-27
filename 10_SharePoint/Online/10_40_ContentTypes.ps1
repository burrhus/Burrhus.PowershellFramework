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

function global:SP-CreateContentTypes($command) {
	$web = $global:clientContext.Web;
	$global:clientContext.Load($web)
	$global:clientContext.executeQuery();

	foreach ($contentType in $installXML.Install.ContentTypes.AddContentTypes.ContentType)
	{
		SP-CreateContentType $web $contentType $web
	}

	foreach ($webDef in $installXML.Install.Webs.Web) 
	{
		$url = Get-EnvironmentVar $webDef.Url $global:setupXML
		$web = $global:clientContext.Site.OpenWeb($url)

		$global:clientContext.Load($web)
		$global:clientContext.executeQuery();

		foreach ($contentType in $webDef.ContentTypes.AddContentTypes.ContentType)
		{
			SP-CreateContentType $web $contentType $web 
		}
	}
}

<#
Hjælpefunktioner
#>

function global:Get-ContentType ($level, $ctDef)
{
	Debug $level.Count;
	foreach($contentType in $level)
	{
		if ($ctDef.ContentTypeID -ne $null) {
		} else {
			if ($contentType.Name -eq $ctDef.Name)
			{
				Debug "------------- " $contentType.Name
				return $contentType;
			}
		}
	}
	return $null;
}

function global:SP-CreateContentType($receiverObject, $contentTypeDef, $web) {
	$availableContentTypes = $receiverObject.AvailableContentTypes;
	$global:clientContext.Load($availableContentTypes)
	$global:clientContext.executeQuery();
	$contentType = Get-ContentType $availableContentTypes $contentTypeDef
	if ($contentType -eq $null)
	{
		$contentTypes = $receiverObject.ContentTypes;
		$global:clientContext.Load($contentTypes)
		$global:clientContext.executeQuery();

		$creationInfo = New-Object Microsoft.SharePoint.Client.ContentTypeCreationInformation;

		if ($contentTypeDef.ParentContentTypeID -ne $null){
			Debug "nedarver fra "$contentTypeDef.ParentContentTypeID
			$parentContentType = $availableContentTypes.GetById($contentTypeDef.ParentContentTypeID);
			$global:clientContext.Load($parentContentType)

			$global:clientContext.executeQuery();

			$creationInfo.ParentContentType  = $parentContentType;
		} else {
			foreach ($availableContentType in $availableContentTypes) {
				if ($availableContentType.Name -eq $contentTypeDef.ParentContentType) {
					$creationInfo.ParentContentType  = $availableContentType;
				}
			}
		}
		$creationInfo.Name = $contentTypeDef.Name
		$group = Get-EnvironmentVar $contentTypeDef.Group $global:setupXML
		$creationInfo.Group = $group
		write-host $contentTypeDef.Name
		Write-Host $group

		$ctReturn = $receiverObject.ContentTypes.Add($creationInfo);
		$global:clientContext.Load($ctReturn)

		$global:clientContext.executeQuery();
		$web.Update()
		foreach ($field in $contentTypeDef.ReferenceFields.Field)
		{
			SP-AddFieldReference $ctReturn $field $web 
		}
		foreach ($field in $contentTypeDef.AddFields.Field)
		{
			SP-CheckField $ctReturn $field 
        }
	}
	else
	{
		write-host -f yellow $contentTypeDef.Name "is available in the site - and will not be created"

		foreach ($field in $contentTypeDef.ReferenceFields.Field)
		{
			SP-AddFieldReference $contentType $field $web  $setupXML $installXML
		}
		foreach ($field in $contentTypeDef.AddFields.Field)
		{
			SP-CheckField $contentType $field $setupXML $installXML
		}
	}
}

function global:SP-RemoveContentType($receiverObject, $contentTypeDef, $web) {
	Write-Host "Fjerner "$contentTypeDef.Name
	$contentTypes = $receiverObject.ContentTypes;
	$global:clientContext.Load($contentTypes)
	$global:clientContext.executeQuery();
	$contentType = Get-ContentType $contentTypes $contentTypeDef
	if ($contentType -ne $null)
	{
		$contentType.DeleteObject();
		$global:clientContext.executeQuery();
	}
	else
	{
	}
}

function global:SP-AddContentType($receiver, $contentType, $web)
{
	$ctReturn = $receiver.ContentTypes.AddExistingContentType($contentType)
	$clientContext.Load($ctReturn)
	try
	{
		$clientContext.ExecuteQuery()
	}
	catch 
	{

	}
}