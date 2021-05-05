#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.SharePoint.Client\v4.0_16.0.0.0__71e9bce111e9429c\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.SharePoint.Client.Runtime\v4.0_16.0.0.0__71e9bce111e9429c\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.SharePoint.Client.Publishing\v4.0_16.0.0.0__71e9bce111e9429c\Microsoft.SharePoint.Client.Publishing.dll"
Add-Type -Path "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.SharePoint.Client.Search\v4.0_16.0.0.0__71e9bce111e9429c\Microsoft.SharePoint.Client.Search.dll"
Add-Type -Path "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.SharePoint.Client.DocumentManagement\v4.0_16.0.0.0__71e9bce111e9429c\Microsoft.SharePoint.Client.DocumentManagement.dll"

#Set parameter values
$SiteURL="https://clarios.sharepoint.com/sites/Partners/"
$SourceFilePath="D:\Files\upload\filesname1.xlsx"
$TargetFolderRelativeURL ="/sites/Partners/Shared Documents/subfolder"
$User = "sender@domain.com"
$Password = "Password"  | ConvertTo-SecureString -AsPlainText -Force

#Bind to site collection
$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User,$Password)
$Ctx.Credentials = $Credentials
  
Try {
    #Setup the context
    $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $Ctx.Credentials = $Credentials
       
    #Get the Target Folder to upload
    $Web = $Ctx.Web
    $Ctx.Load($Web)
    $TargetFolder = $Web.GetFolderByServerRelativeUrl($TargetFolderRelativeURL)
    $Ctx.Load($TargetFolder)
    $Ctx.ExecuteQuery() 
 
    #Get the source file from disk
    $FileStream = ([System.IO.FileInfo] (Get-Item $SourceFilePath)).OpenRead()
    #Get File Name from source file path
    $SourceFileName = Split-path $SourceFilePath -leaf  
    $TargetFileURL = $TargetFolderRelativeURL+"/"+$SourceFileName
 
    #Upload the File to SharePoint Library Folder
    $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
    $FileCreationInfo.Overwrite = $true
    $FileCreationInfo.ContentStream = $FileStream
    $FileCreationInfo.URL = $TargetFileURL
    $FileUploaded = $TargetFolder.Files.Add($FileCreationInfo)  
    $Ctx.ExecuteQuery()  
 
    #Close file stream
    $FileStream.Close()
    Write-host "File '$TargetFileURL' Uploaded Successfully!" -ForegroundColor Green
}
catch {
    write-host "Error Uploading File to Folder: $($_.Exception.Message)" -foregroundcolor Red
}
