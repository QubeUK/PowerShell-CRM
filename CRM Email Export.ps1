# Powershell with SQL Query to retrieve all unique Email addresses from CRM database.
# Needs to be run as Domain Administrator for permission to SQL database.
# Created Date:  20/11/2015	                   Created by:  Lee Williams
# Modified Date: 12/01/2016	                   Modified by: Lee Williams
$Cuser = [Environment]::UserName
$dbname = "CRM"
$dbsrv = "MBSSQL\MMS2011"
$TmpFile = "C:\CRM Email Temp.csv"
$Final = "C:\CRM Email.csv"
$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop               # Needed to generate Popup
$Query= @" 
 SELECT DISTINCT Emai_EmailAddress 
 FROM crm.dbo.Email 
 WHERE Emai_EmailAddress LIKE '%@%' 
 AND Emai_EmailAddress NOT LIKE '%*%' 
 ORDER BY Emai_EmailAddress;
"@                                                       
Clear-Host
If ($Cuser -eq "Administrator")                                               # Checks if user is Administrator
{Write-Host "Running SQL Query..."                                            # Runs the SQL Query using Invoke-SQLcmd
 Invoke-Sqlcmd -ServerInstance $dbsrv -Database $dbname -Query $Query | Export-CSV $TmpFile
 Write-Host "Generating CSV File..."
 $File = Get-Content $TmpFile                                                 # Loads content of File into TmpFile
 $File = $File[2..$($File.Count - 1)]                                         # Drop the first two lines from temp CSV
 $File > $Final                                                               # Save it to a new file
 Remove-Item -path $TmpFile                                                   # Removes Temp CSV File
 $Rows = @(Get-Content $Final).Length                                         # Counts row in CSV File
 $wshell.Popup("$Rows Rows have been saved to $Final",0,"Task Complete",64+0) # Generates Popup
 Clear-Host}
Else {
Clear-Host
$wshell.Popup("You are trying to excute the script as $Cuser. `n
You need to be logged in as Domain Administrator for the script to run.`n
Please try again.",0,"Error",16+0)                                            # Generates Popup
Clear-Host
}                                                                       