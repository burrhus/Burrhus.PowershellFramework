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
ShowMessage "Loading SharePoint Online Files 3.0.0" [LogLevels]::Information

<#
Menufunktioner
#>


<#
Hjælpefunktioner
#>

function global:SP-DownloadAttachments($listItem, $path) {
    if ($listItem.FieldValues.ContainsKey("Attachments")) {
        if ($listItem.FieldValues["Attachments"]) {
            $attachments  = $listItem.AttachmentFiles
            $global:clientContext.Load($attachments);
            $global:clientContext.ExecuteQuery();

            if ($attachments.count -gt 0) {
                $listFile = [System.IO.Path]::Combine($path, $listItem.FieldValues["ID"])
                if (-not (Test-Path $listFile)) {
                    md $listFile                 
                }                    
            }
            foreach ($attachment in $attachments) {
                $fileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($global:clientContext, $attachment.ServerRelativeUrl);
                $fileName = [System.IO.Path]::Combine($listFile, $attachment.FileName)
                $fileStream = [System.IO.File]::Create($fileName)
                $fileInfo.Stream.CopyTo($fileStream);

                $fileStream.Close()            
            }                    
        }        
    } else {
        $listFile = [System.IO.Path]::Combine($path, $listItem.FieldValues["ID"])
        if (-not (Test-Path $listFile)) {
            md $listFile                 
        }                    
        $file = $listItem.File
        $Global:clientContext.Load($file)
        $Global:clientContext.ExecuteQuery()
        write-host $file.Name
        $fileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($global:clientContext, $file.ServerRelativeUrl);
        $fileName = [System.IO.Path]::Combine($listFile, $file.Name)
        Write-Host $fileName
        $fileStream = [System.IO.File]::Create($fileName)
        $fileInfo.Stream.CopyTo($fileStream);
        $fileStream.Close()            
    }

}