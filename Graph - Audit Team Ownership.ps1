Set-ExecutionPolicy RemoteSigned -Scope CurrentUser

#Add the CPSTeamsManagement module
Remove-Module GraphTeamsManagement
Import-Module -Name ".\GraphTeamsManagement" -Verbose

#Add the Microsoft Graph module
Install-Module -Name Microsoft.Graph

# Switch to beta profile
Select-MgProfile -Name beta
$tenantId_E5 = "3aad59e2-dbed-4fc0-9780-91156f9e0bc4"

# Sign-in with the permissions needed for grop management
Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All" -TenantId $tenantId_E5

#Set variables
$CSVPath = "D:\TeamsManagement\"

#Step 1 - Check
$CSVFilePath_TeamsCompliance = Check-GraphTeamsCompliance `
    -CSVPath $CSVPath `
    -ShowDebug $false

#Step 2 - Identify remediation actions
#  Update the CSV file  found at $CSVFilePath_TeamsCompliance 

#Step 3 - Remediate
Update-TeamsCompliance `
    -CSVFilePath $CSVFilePath_TeamsCompliance `
    -ShowDebug $true