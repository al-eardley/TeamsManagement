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
$userUPN = "alan.eardley@company.com"

#Step 1 - Check
$CSVFilePath_TeamMembership = Get-TeamMembership `
    -UserUPN $userUPN `
    -CSVPath $CSVPath `
    -ShowDebug $false

#Step 2 - Identify remediation actions
#  Update the CSV file  found at $CSVFilePath_TeamMembership 

#Step 3 - Remediate
Update-TeamsCompliance `
    -CSVFilePath $CSVFilePath_TeamMembership `
    -ShowDebug $true

#Step 4 - Remove rows where the user should remain a member of the Team

#Remove user from Teams remaining in CSV file
Remove-TeamMember `
    -UserUPN $userUPN `
    -CSVFilePath $CSVFilePath_TeamMembership `
    -ShowDebug $false