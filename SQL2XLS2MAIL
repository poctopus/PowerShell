$Subject ="报表-$((Get-Date).ToString("yyyyMMdd_HHmmss"))"
$Directory="D:\FILES\"
$csvfilename="CSVFILENAME-_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
$xlsxfilename="XLSFILENAME-_$((Get-Date).ToString("yyyyMMdd_HHmmss")).xlsx"
$Sendfile="$Directory$xlsxfilename"
$Database = 'ORDER'
$Server = '127.0.01'
$UserName = 'sq'
$Password = 'Query2020'

#数据库查询脚本
$SqlQuery = "SQLSCRPT-1"
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
$Table | Export-CSV -encoding utf8 -NoTypeInformation $Directory$csvfileName
#$Table
}

function To_Excel {
### Create a new Excel Workbook with one empty sheet
$excel = New-Object -ComObject excel.application 
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)

### Build the QueryTables.Add command
### QueryTables does the same as when clicking "Data ? From Text" in Excel
$TxtConnector = ("TEXT;" + "$Directory$csvfileName")
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)
#$delimiter = "\r\n" #Specify the delimiter used in the file
### Set the delimiter (, or ;) according to your regional settings
$query.TextFileOtherDelimiter = $Excel.Application.International(5)

### Set the format to delimited and text for every column
### A trick to create an array of 2s is used with the preceding comma
$query.TextFilePlatform = 65001
#$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType  = 1
$query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1

### Execute & delete the import query
$query.Refresh()
$query.Delete()

### Save & close the Workbook as XLSX. Change the output extension for Excel 2003
$Workbook.SaveAs("$Directory$xlsxfilename",51)
$excel.Quit()
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
Send-MailMessage -to test@domain.com,test1@domain.com -Cc test3@domain.com,test4@domain.com  -from 自动邮件<sender@domain.com> -Subject $subject -SmtpServer "$smtpServer" -Port 25 -Encoding UTF8 -Attachments $Sendfile  -Credential $cre
}
Export_Excel
To_Excel
Start-Sleep -Seconds 35
SendMail
echo "remove-item -Force $csvfilename"
#remove-item -Force "D:\ZHANGYU\SW\TM14\*.*"
#remove-item -Force "$Directory$xlsxfilename" #删除两个临时文件
