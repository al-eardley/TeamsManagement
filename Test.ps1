$userUPN = "user1@company.com"
$newOwnerUPN = "user2@company.com"
$CSVPath = "C:\Repo\CPSTeamsManagement\"

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

$CSVFilePath_TeamsCompliance = Check-TeamsCompliance `
    -CSVPath $CSVPath `
    -ShowDebug $false

Add-TeamOwner `
    -CSVFilePath "$CSVPath\Teams - AddTeamOwner.csv" `
    -ShowDebug $true