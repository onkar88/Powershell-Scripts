<# Admin credentials are already saved in encrypted form in a text file. 
Get the contents of encryption key and encrypted password. 
This admin account is local admin on all windows servers#>

$AESKeyFilePath = "C:\AESKey.txt"
$credentialFilePath = "C:\Credential.txt"

$username = "<admin@domain.com>"
$AESKey = Get-Content $AESKeyFilePath
$pwdTxt = Get-Content $credentialFilePath

<#Decrypt password using encryption key#>
$securePwd = $pwdTxt | ConvertTo-SecureString -Key $AESKey


<#Create credential object to store admin username and password#>
$cred = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $securePwd


<#Define required variables#>
$freespace = @()
$size=@()
$objects1 = @()
$data = @()
$ErrorActionPreference = 'silentlycontinue'



<#Read server IP and get drive capacity alongwith drive letter. Convert capacity into GB's .#>

ForEach($s in Get-Content '<Path to text file for Server IP List>')
{

$drive=gwmi Win32_Volume -ComputerName $s -Filter "drivetype=3"|?{$_.driveletter -ne $null}

foreach($item in $drive)
{
$sv=$s
$cp1=$item|Select-Object -ExpandProperty capacity
$dr=$item|Select-Object -ExpandProperty driveletter
$fp1=$item|Select-Object -ExpandProperty freespace

$cp="{0:N1}" -f($cp1/1gb)
$fp="{0:N1}" -f($fp1/1gb)
$ft="{0:P0}" -f($fp1/$cp1)


<#Create object to hold drive details for each server #>

$objects1 += New-Object -Type PSObject -Prop @{'Server'=$sv;'Drive'=$dr;'Size(GB)'=$cp;'Freespace(GB)'=$fp;'FreeSpacePerCent'=$ft}


}

}


<#Find free space percentage#>
$objects =  $objects1 |sort -Property {[INT]($_.freespacepercent -replace '%')}


<#Export details to excel file. Excel file name will show date & time#>
$objects |export-excel "C:\diskspacereport_$((Get-date).Tostring("dd_MM_yyyy HH_mm")).xlsx"


<#Create object to hold html table parameters#>
$convertParams = @{ 

 
 head = @"
 <Title>Disk Space Report</Title>
<style>
body { background-color:#E5E4E2;
       font-family:Monospace;
       font-size:10pt; }
TD{border: 1px solid black; padding: 5px; }
th { color:white;
     background-color:black; }
table, tr, td, th { padding: 2px; margin: 0px ;white-space:pre; }
tr:nth-child(odd) {}
table { border: 1px solid black; border-collapse: collapse;}
h2 {
 font-family:Tahoma;
 color:#6D7B8D;
}
.alert {
 
 background-color: red;
 }
.footer 
{ color:green; 
  margin-left:10px; 
  font-family:Tahoma;
  font-size:8pt;
  font-style:italic;
}
</style>
"@
}


<#Convert object to xml#>
[xml]$html = $objects | convertto-html -Fragment


<#Add red color to the row if free space is less than 10%. Red color is indicated by "Alert" type of value#>
for ($i=1;$i -le $html.table.tr.count-1;$i++) {
  if ([int]($html.table.tr[$i].td[0]).trim('%') -lt 10) {
    $class = $html.CreateAttribute("class")
    $class.value = 'alert'
    $html.table.tr[$i].attributes.append($class) | out-null
  }
}


<#Add to html table parameter object#>
$convertParams.add("body",$html.InnerXml)


<#Create CSV file for the report & save it in a variable#>
convertto-html @convertParams |Out-file "<CSV file path>"
$emailbody1 = Get-Content "<CSV file path>"|Out-String

<#Create email message body#>
$EmailBody = 'Dear Team, <br /> <br />Please create space on highlighted servers <br /> <br />' + $emailbody1


<#Add TO & CC addresses#>
[string[]]$recipients = "address1","address2"
[string[]]$ccrecipients= "address3","address4"

<#Send new email message#>
Send-MailMessage -From '<sender address>' -To $recipients -Cc $ccrecipients -SmtpServer '<smtp endpoint which receives emails>' -Credential $cred  -Body $EmailBody  -Subject "Disk space report" -BodyAsHtml


<#Remove CSV file which was temporary for creating report#>
remove-item '<CSV file path>'




