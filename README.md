# Powershell-Scripts

•	**DiskSpaceReport.ps1**  –  Creates html report for drive capacity on windows servers and sends this report through email to IT admins. Additionally it saves report in a separate excel file. It highlights volumes which have free space lower than 10 %. <br />Prerequisites:
<br />3 text files are required - 1. Server IP List file   &nbsp;&nbsp;2. File that stores encryption key    &nbsp;&nbsp;3. File that stores encrypted admin password.
<br /> Email will be sent as shown below - 
<br />
![DiskSpaceReport](Images/DiskSpaceReport.png)

•	**OutlookProfile.ps1**  - Creates new outlook profile. New profile name is "NewOutlookProfile". It first deletes existing registry keys & ost file for newoutlookprofile. Then it registers new profile name in windows registry & sets it as default profile. After running this script, outlook will connect with new local outlook profile. <br /> 
