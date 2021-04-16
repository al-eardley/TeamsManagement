$userUPN = "user1@company.com"
$newOwnerUPN = "user2@company.com"
$CSVPath = "D:\CPSTeamsManagement\"

<#
$CSVFilePath_TeamMembership = Get-TeamMembership `
    -UserUPN $userUPN `
    -CSVPath $CSVPath `
    -ShowDebug $false

Remove-TeamMember `
    -UserUPN $userUPN `
    -CSVFilePath $CSVFilePath_TeamMembership `
    -ShowDebug $false

Replace-TeamOwner `
    -OldOwnerUPN $userUPN  `
    -NewOwnerUPN $newOwnerUPN  `
    -CSVFilePath $CSVFilePath `
    -ShowDebug $true
#>

$CSVFilePath_TeamsCompliance = Check-TeamsCompliance `
    -CSVPath $CSVPath `
    -ShowDebug $false

Update-TeamsCompliance `
    -CSVFilePath $CSVFilePath_TeamsCompliance `
    -ShowDebug $true

Get-Team

$privateChannels = Get-TeamChannel -GroupId 8d85e949-f756-428b-ae1b-08b446b07875 -MembershipType Private

$privateChannels

Get-TeamChannel -