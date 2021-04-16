Set-ExecutionPolicy RemoteSigned

#Add the CPSTeamsManagement module
Remove-Module CPSTeamsManagement
Import-Module -Name ".\CPSTeamsManagement" -Verbose

#Add the MicrosoftTeams module
Install-Module -Name MicrosoftTeams
Import-Module MicrosoftTeams

#Connect to Microsoft Teams
Connect-MicrosoftTeams

#Set variables
$CSVPath = "D:\CPSTeamsManagement\"

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