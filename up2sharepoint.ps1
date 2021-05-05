#Specify tenant admin and site URL
$User = "sender@domain.net"
$Password = "Password"  | ConvertTo-SecureString -AsPlainText -Force
$SiteURL = "https://abc.sharepoint.com/sites/subfolder/"
$Folder = "D:\files\sbfolder\"
$DocLibName = "Documents"
$SubFolderName = "subfolder"

#Add references to SharePoint client assemblies
Add-Type -Path "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.SharePoint.Client\v4.0_16.0.0.0__71e9bce111e9429c\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.SharePoint.Client.Runtime\v4.0_16.0.0.0__71e9bce111e9429c\Microsoft.SharePoint.Client.Runtime.dll"

#Bind to site collection
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User,$Password)
$Context.Credentials = $Creds

    #Get the Target Folder to upload
    $Web = $Context.Web
    $Context.Load($Web)
    $TargetFolder = $Web.GetFolderByServerRelativeUrl($TargetFolderRelativeURL)
    $Context.Load($TargetFolder)
    $Context.ExecuteQuery() 

#Retrieve list
$List = $Context.Web.Lists.GetByTitle($DocLibName)
$FolderToBindTo = $List.RootFolder.Folders
$Context.Load($FolderToBindTo)
$Context.ExecuteQuery()
$FolderToUpload = $FolderToBindTo | Where {$_.Name -eq $SubFolderName}

$files = ([System.IO.DirectoryInfo] (Get-Item $Folder)).GetFiles()

#Upload file
ForEach($file in $files)
{
$FileStream = New-Object IO.FileStream($File.FullName,[System.IO.FileMode]::Open)
$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
$FileCreationInfo.Overwrite = $true
$FileCreationInfo.ContentStream = $FileStream
$FileCreationInfo.URL =$File
$Upload = $FolderToUpload.Files.Add($FileCreationInfo)

$Context.ExecuteQuery()
}
