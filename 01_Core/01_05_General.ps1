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

ShowMessage "Loading General functions Version 3.0.0" [LogLevels]::Information

function ContinueYesNo(){
    <#
    <Summary>
        <Description>Generel fortsæt funktion</Description>
    </Summary>    
    #>
    if ((Input "Continue!" "y") -ne "y") {
        exit;
    }
}

function Debug([string] $prompt){
    <#
    <Summary>
        <Description>Viser en debugbesked, hvis Debug er slået til</Description>
        <Parameters>
            <Parameter Name="$prompt" Description="Den tekst der vises" />
        </Parameters>
    </Summary>    
    #>
	if($Global:ShowDebug) {
		if ($Global:PauseOnDebug) {
			Read-Host -Prompt $prompt;
		} else {
			ShowMessage $prompt [LogLevels]::Debug
		}
	}
}

function Input([string] $prompt, [string] $default) {
    <#
    <Summary>
        <Description>Generel indtastnings funktion</Description>
        <Parameters>
            <Parameter Name="$prompt" Description="Den tekst der vises" />
            <Parameter Name="$default" Description="Default værdi der vælges hvis der ikke skrives andet" />
        </Parameters>
        <Returns>string</Returns>
    </Summary>    
    #>
    if ($default -ne "") {
        $prompt = $prompt + " (default: " + $default + ")";
    }


    $returnValue = Read-Host -Prompt $prompt;
    if ($returnValue -eq "") {
        $returnValue = $default;
    }
    return $returnValue;
}

function Get-EnvironmentVar($key, $setupXML)
{
    <#
    <Summary>
        <Description>Henter variable fra setupfilen</Description>
        <Parameters>
            <Parameter Name="$key" Description="Nøglenavn til variablen" />
            <Parameter Name="$setupXML" Description="XMLObject der indeholder current setup" />
        </Parameters>
        <Returns>string</Returns>
    </Summary>    
    #>
	foreach ($vp in $setupXML.Setup.EnvironmentConfiguration.KeyValuePair)
	{
		if ($vp.Key -Eq $key)
		{
			ShowMessage ("Get-EnvironmentVar: " + $key + "=" + $vp.Value) [LogLevels]::Debug
			return $vp.Value;
		}
	}
	return $key
}
