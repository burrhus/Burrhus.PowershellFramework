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
ShowMessage "Loading Menu Version 3.0.0" [LogLevels]::Information

function ShowMenu
{
    if ($DirectRun -ne "") {
        RunMenuV3 $DirectRun
    }
    if ($Interactive -eq $false) {
        return;
    }

    read-host "Initialiseret"

   	$Menu = ""
	while ($menu -ne "X")
	{
  		cls

        ShowMessage ("Url: " + $global:setupXML.Setup.SiteUrl) [LogLevels]::Flow $true
		ShowMessage ("Environment: " + $Environment) [LogLevels]::Flow $true
		ShowMessage ("Solution: " + $Solution) [LogLevels]::Flow $true
		ShowMessage ("SiteType: " + $global:setupXML.Setup.SiteType) [LogLevels]::Flow $false   
		Write-host "" 
		Write-host "Menu: " 
		Write-host "" 
        $menu = "X"

        if ($global:installXML.Install.Version -eq $null) {
            MenuV2
        } else {
            MenuV3
        }

        Write-host "X. Afslut" 
		$Menu = Input "Valg" "X"
		
        if ($global:installXML.Install.Version -eq $null) {
            RunMenuV2 $Menu
        } else {
            RunMenuV3 $Menu
        }
    }
}

function MenuV3 {
    foreach ($menu in $global:installXML.Install.MenuV3.Menu) {
        write-host $menu.ShortCut"." $menu.Name
    }
}

function RunMenuV3($selectedMenu) {    
    $blnFound = $false
    foreach ($menu in $global:installXML.Install.MenuV3.Menu) {
        if ($selectedMenu -eq $menu.ShortCut) {
            $blnFound = $true
            foreach ($command in $menu.Commands.Command) {
                RunCommand $command
            }
        }
    }
	if ($blnFound -eq $false) {
		write-host "Kan ikke finde $selectedMenu"
	}
    Read-Host "."
}

function MenuV2
{
   		if($Global:installXML.Install.Menu.FullCheck -ne $null){
			Write-host "A. SharePoint Full Check" 
		}
		if($Global:installXML.Install.Menu.Report -ne $null){
			Write-host "B. Report" 
		}
		if ($Addon -ne "") {
			Write-host "C. Custom addon";
		}
		if($Global:installXML.Install.Menu.Backup -ne $null){
			Write-host "D. Backup" 
		}
		if($Global:installXML.Install.Menu.Restore -ne $null){
			Write-host "E. Restore" 
		}
		if($Global:installXML.Install.Menu.FillLists -ne $null){
			Write-host "F. Fill lists" 
		}
		if($Global:installXML.Install.Menu.EditLists -ne $null){
			Write-host "G. Edit listitems" 
		}
		if($Global:installXML.Install.Menu.PreCheck -ne $null){
			Write-host "H. Check installation" 
		}
		if($Global:installXML.Install.Menu.Configure -ne $null){
			Write-host "I. Configure installation" 
		}
		if($Global:installXML.Install.Menu.Import -ne $null){
			Write-host "J. Import data"
		}
		if($Global:installXML.Install.Menu.UploadFile -ne $null){
			Write-host "P. Upload File"
		}
		if($Global:installXML.Install.Menu.SpecificUpdate -ne $null){
			Write-host "S. Specific update" 
		}
		
		if($Global:installXML.Install.Menu.UnConfigure -ne $null){
			Write-host "U. UnConfigure installation" 
		}
		
		#Write-host "T. Test" 
		if($Global:setupXML.Setup.SiteType -ne "SPOnline"){
			Write-host "R. Restart services" 
			Write-host "R1. Stop services" 
			Write-host "R2. Start services" 
		}
		if($Global:installXML.Install.Menu.UnInstall -ne $null){
			Write-host "Y. UnInstall" 
		}
		if($Global:installXML.Install.Menu.Install -ne $null){
			Write-host "Z. Install/Upgrade" 
		}

		if($Global:installXML.Install.Menu.ExportTaxonomies -ne $null){
			Write-host "EXT. Export taxonomier" 
		}
		if($Global:installXML.Install.Menu.ImportTaxonomies -ne $null){
			Write-host "IMT. Import taxonomier" 
		}
		if($Global:installXML.Install.Menu.ExportSite -ne $null){
			Write-host "EXS. Export Site" 
		}
		Write-host "" 

}

function RunMenuV2($selectedMenu) {
	switch ($selectedMenu)
	{
		"A" 
		{
			SharePointFullCheck $global:setupXML  $global:installXML;
		}
		"B"
		{
			SP-CreateReport $global:setupXML $global:installXML
		}
		"C"
		{
			RunAddon $global:setupXML $global:installXML
		}
		"D"
		{
			SP-RunBackup $null $global:setupXML $global:installXML
		}
		"E"
		{
			SP-RunRestore $null $global:setupXML $global:installXML
		}
		"F"
		{
			SP-FillLists $null $global:setupXML $global:installXML
		}
		"G"
		{
			SP-EditListItems $null $global:setupXML $global:installXML
		}
		"H" 
		{
			SP-PreCheckInstallation $global:setupXML $global:installXML
		}
		"I"
		{
			ShowMessage "Configuring" [LogLevels]::Flow
			SP-ConfigureInstallation $global:setupXML $global:installXML
		}
		"J"
		{
			ShowMessage "Importing" [LogLevels]::Flow
			SP-ImportData $global:setupXML $global:installXML
		}
		"P" 
		{
			SP-UploadFiles  $global:setupXML $global:installXML;
		}
		"R" 
		{
			SP-Restart;
		}
		"R1" 
		{
			SP-Stop;
		}
		"R2" 
		{
			SP-Start;
		}
		"S" 
		{
			if ($Global:UpdateID -eq "") {
				$Global:UpdateID = Input "Update ID" $Global:UpdateID
			}

			SPRunUpdate $Global:UpdateID $global:setupXML $global:installXML;

		}
		"T" 
		{
			SP-UploadFiles2 $global:setupXML $global:installXML;
		}
		"U"
		{		
			ShowMessage "UnConfiguring will remove all data made with the solution" [LogLevels]::Flow
			ContinueYesNo
			SP-UnConfigureInstallation $global:setupXML $global:installXML
		}
		"Y" 
		{
			ShowMessage "UnInstall will remove all data made with the solution" [LogLevels]::Flow
			ContinueYesNo
			SP-UnInstallSolution $global:setupXML $global:installXML;
			#SP-UnConfigureInstallation $global:setupXML $global:installXML
			#SP-UnInstallPackages $command $global:setupXML $global:installXML;
			#read-host "Tryk enter for at fortsætte!"
		}
		"Z" 
		{
			SP-InstallSolution $global:setupXML $global:installXML;
		}
		"IMT" 
		{
			SP-ImportTaxonomy $global:setupXML $global:installXML;
		}
		"EXT" 
		{
			SP-ExportTaxonomy $global:setupXML $global:installXML;
		}
		"EXS" 
		{
			SP-ExportSite $global:setupXML $global:installXML;
		}
		"Default" {}
				
	}

	if ($Interactive) 
	{
		read-host "Tryk enter for at fortsætte"
	}
}

function RunCommand ($command) {
    ShowMessage $command.Log [LogLevels]::Flow

    switch ($command.Function) {
        "SP-TestTaxonomies" {  
            SP-TestTaxonomies $command
        }
		"SP-ExportTaxonomies"{
			SP-ExportTaxonomies $command
		}
        Default {
			read-host $command.Function
		}
    }
}