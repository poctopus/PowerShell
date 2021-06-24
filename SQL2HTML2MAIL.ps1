# Database info
$Database = 'test'
$Server = '127.0.0.1'
$UserName = 'sa'
$Password = 'sa'

# Email params
$EmailParams = @{
    To         = 'aladown@gmail.com'
    From       = '自动邮件<abc@abc.com>'
    Smtpserver = 'smtpdm.aliyun.com'
    Subject    = "检查结果-$((Get-Date).ToString("yyyyMMdd_HHmmss"))"
}
$smtpServer = "smtpdm.aliyun.com"
$smtpUser = "abc@abc.com"
$smtpPassword = "Password"
$ss=ConvertTo-SecureString -String "$smtpPassword" -AsPlainText -force
$ss|Write-Host
$cre= New-Object System.Management.Automation.PSCredential("$smtpUser",$ss)

# Function to get data from SQL server
function Get-SQLData {
    param($Query)
    $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = "Data Source=$Server;Initial Catalog=$Database;user id=$UserName;pwd=$Password"
    $connection.Open()
    
    $command = $connection.CreateCommand()
    $command.CommandText = $Query
    $reader = $command.ExecuteReader()
    $table = New-Object -TypeName 'System.Data.DataTable'
    $table.Load($reader)
    
    $connection.Close()
    
    return $Table
}

# Define the SQL Query
$Query = "--SQL脚本
select * from order"

# Html CSS style
$Style = @"
<style>
table { 
    border-collapse: collapse;
}
td, th { 
    border: 1px solid #ddd;
    padding: 8px;
}
th {
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: center;
    background-color: #4286f4;
    color: white;
}
</style>
"@

# Run the SQL Query
$Results = Get-SQLData $Query

# If results are returned
If ($Results.count -ne 0) {

    # Convert results into html format
    $Html = $Results |
        ConvertTo-Html -Property '单号','平台单号','订单时间','付款时间','状态' -Head $style -Body "<h2>内部检查结果</h2>" -CssUri "http://www.w3schools.com/lib/w3.css"  | 
        Out-String

    # Send the email
    Send-MailMessage @EmailParams -Body $Html -BodyAsHtml -Credential $cre -Encoding UTF8
}
