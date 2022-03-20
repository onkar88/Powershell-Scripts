# Powershell-Scripts
Scripts that automate different IT tasks to save time/money. Most of the scripts need a domain admin account to run. 

•	**DiskSpaceReport.ps1**  –  Creates html report for drive capacity on windows servers. Then sends this report through email to IT admins. Additonally it saves report in a separate excel file. It highlights drives which have free space lower than 10 %. This script can be scheduled to run once a day.  
3 text files are required - 1. Server IP List file   &nbsp;&nbsp;2. File that stores encryption key    &nbsp;&nbsp;3. File that stores encrypted admin password. <br /> Report will be created as shown below - 
<br />
![DiskSpaceReport](images/DiskSpaceReport.png)
