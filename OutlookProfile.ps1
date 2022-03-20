
<#New Profile name is "NewOutlookProfile"#>


<#Function to check if registry key exists or not
---------------------------------------------------------------
---------------------------------------------------------------
#>
Function Test-Value{

Param (
[parameter(Mandatory=$true)]$path,
[parameter(Mandatory=$true)]$value
)

try {

Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null


return $true
}
catch{
return $false
}
}


<#Find SID value for logged-in user account
--------------------------------------------------------------
--------------------------------------------------------------
#>
$info = tasklist /v /FI "IMAGENAME eq explorer.exe" /FO list | find "User Name:"
$usr=$info.Split('\')[1]
$obj=New-Object System.Security.Principal.NTAccount($usr)
$abc=$Obj.Translate([System.Security.Principal.SecurityIdentifier])
$sid=$abc.value




<#Remove existing registry key for New Profile (NewOutlookProfile)
-------------------------------------------------------------------------
-------------------------------------------------------------------------#>
$pt = "Registry::HKEY_USERS\$sid\software\microsoft\office\16.0\outlook\profiles\NewOutlookProfile"

if(test-path $pt)
{
Remove-Item "$pt" -Force -Recurse
}



<#Remove existing registry key for NewOutlookProfile from "Deleted Profiles"
----------------------------------------------------------------------------
---------------------------------------------------------------------------
#>
$pf = "Registry::HKEY_USERS\$sid\Software\Microsoft\Office\16.0\Outlook\Deleted Profiles\NewOutlookProfile"
if(test-path $pf)
{
Remove-Item "$pf" -Force -Recurse
}


<#Remove existing OST files for New Profile Name (NewOutlookProfile)
--------------------------------------------------------------------------------
--------------------------------------------------------------------------------
#>


Get-ChildItem -Path C:\Users\$usr\AppData\Local\Microsoft\Outlook -Filter *NewOutlookProfile*|Remove-Item 



<#Create registry key for new profile. Registry key name is "NewOutlookProfile"
-----------------------------------------------------------------------
-----------------------------------------------------------------------
#>
New-Item -Path Registry::HKEY_USERS\$sid\software\microsoft\office\16.0\outlook\profiles\NewOutlookProfile




<#Create "DefaultProfile" key if it does not exists. 
-------------------------------------------------------------------------
--------------------------------------------------------------------------
#>

$name = "DefaultProfile"
$vl = "NewOutlookProfile"
$registrypath = "Registry::HKEY_USERS\$sid\software\microsoft\office\16.0\outlook"
if((Test-Value -path $registrypath -value $name) -eq $false)
{

New-Item -Path $registrypath -Value $name
}


<#Set "DefaultProfile" key value as "NewOutlookProfile"
--------------------------------------------------------------------------
--------------------------------------------------------------------------
#>
set-ItemProperty -Path $registrypath -Name $name -Value $vl


