Set-ExecutionPolicy RemoteSigned

#Add the CPSTeamsManagement module
Remove-Module GraphTeamsManagement
Import-Module -Name ".\GraphTeamsManagement" -Verbose

#Add the Microsoft Graph module
Install-Module -Name Microsoft.Graph
Install-Module -Name Microsoft.Graph.Authentication -RequiredVersion 1.3.1.0

$tenantId = "<tenantId>"

# Sign-in with the permissions needed for grop management
Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All" -TenantId $tenantId

# Switch to beta profile
Select-MgProfile -Name beta