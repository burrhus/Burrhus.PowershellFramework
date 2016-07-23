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
ShowMessage "Loading errorhandling Version 3.0.0" [LogLevels]::Information

$Global:AskOnError=$true;
$Global:StopOnError=$false;

function ErrorHandling
{
    <#
    <Summary>
        <Description>Håndterer fejl </Description>
    </Summary>
    #>
    if ($Error.Count -gt 0){ 
		ShowMessage $Error[0] [LogLevels]::Error
        
        if ($Global:StopOnError) {
            exit;
        }

        if ($Global:AskOnError) {
		    $t = Input "Error - Continue?" "Y";
		    if ($t -ne "Y")
		    {
			    exit;
		    }
        }
        $Error.Clear();
	}
}