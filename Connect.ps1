Set-ExecutionPolicy RemoteSigned

#Add the CPSTeamsManagement module
Remove-Module TeamsManagement
Import-Module -Name ".\TeamsManagement" -Verbose

#Add the MicrosoftTeams module
Install-Module -Name MicrosoftTeams
Import-Module MicrosoftTeams

#Connect to Microsoft Teams
Connect-MicrosoftTeams