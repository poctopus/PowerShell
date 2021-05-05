$subject ="报表-$((Get-Date).ToString("yyyyMMdd"))"
$fileName =  "D:\FILES\FILENAME-$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
$fileName1 = "D:\FILES\FILENAME1-$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
$fileName2 = "D:\FILES\FILENAME2-$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
$fileName3 = "D:\FILES\FILENAME3-$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
$fileName4 = "D:\FILES\FILENAME4-$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
$Database = 'ORDER'
$Server = '127.0.01'
$UserName = 'sq'
$Password = 'Query2020'
#数据库查询脚本
$SqlQuery = "
SQLSCRIPT-1;
SQLSCRIPT-2;
SQLSCRIPT-3;
SQLSCRIPT-4;
"
function Export_Excel {
# Accessing Data Base
$SqlConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Data Source=$Server;Initial Catalog=$Database;user id=$UserName;pwd=$Password"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$set = New-Object data.dataset
# Filling Dataset
$SqlAdapter.Fill($set)
# Consuming Data
$Table = $Set.Tables[0] 
$Table | Export-CSV -encoding utf8 -NoTypeInformation $fileName
$Table = $Set.Tables[1] 
$Table | Export-CSV -encoding utf8 -NoTypeInformation $fileName1
$Table = $Set.Tables[2] 
$Table | Export-CSV -encoding utf8 -NoTypeInformation $fileName2
$Table = $Set.Tables[3] 
$Table | Export-CSV -encoding utf8 -NoTypeInformation $fileName3
$Table = $Set.Tables[4] 
$Table | Export-CSV -encoding utf8 -NoTypeInformation $fileName4
#$Table
}
function SendMail {
#发邮件
$smtpServer = "smtpdm.aliyun.com"
$smtpUser = "sender@domain.com"
$smtpPassword = "Password"
#$smtp.Send($mail)
$ss=ConvertTo-SecureString -String "$smtpPassword" -AsPlainText -force
$ss|Write-Host
$cre= New-Object System.Management.Automation.PSCredential("$smtpUser",$ss)
Send-MailMessage -to test@domain.com,test1@domain.com -Cc test3@domain.com,test4@domain.com  -from 自动邮件<sender@domain.com> -Subject $subject -SmtpServer "$smtpServer" -Port 25 -Encoding UTF8 -Attachments $fileName,$fileName1,$fileName2,$fileName3,$fileName4 -Credential $cre
}
Export_Excel
Start-Sleep -Seconds 10
SendMail
echo "remove-item -Force $fileName"
#remove-item -Force $fileName #删除7天前的文件
