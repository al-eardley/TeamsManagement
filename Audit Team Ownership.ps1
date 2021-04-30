Set-ExecutionPolicy RemoteSigned -Scope CurrentUser

#Add the CPSTeamsManagement module
Remove-Module TeamsManagement
Import-Module -Name ".\TeamsManagement" -Verbose

#Add the MicrosoftTeams module
Install-Module -Name MicrosoftTeams -Force
Import-Module MicrosoftTeams

#Connect to Microsoft Teams
Connect-MicrosoftTeams

#Set variables
$CSVPath = "D:\TeamsManagement\"

#Step 1 - Check
$CSVFilePath_TeamsCompliance = Check-TeamsCompliance `
    -CSVPath $CSVPath `
    -ShowDebug $false

#Step 2 - Identify remediation actions
#  Update the CSV file  found at $CSVFilePath_TeamsCompliance 

#Step 3 - Remediate
Update-TeamsCompliance `
    -CSVFilePath $CSVFilePath_TeamsCompliance `
    -ShowDebug $true