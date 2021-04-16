#region Write-Status
<#
    .Synopsis
    Writes a status message

    .Description
    Assess the type of message, the $ShowDebug and Level and writes an appropriate message

    .Parameter Message
    The message to be written

    .Parameter Level
    The level of indentation to be used

    .Parameter Type
    The type of message: Start, Progress, Debug, Warning or Success

    .Parameter ShowDebug
    Indicates whether or not to show debug messages

    .Example
    $Message = Write-Status `
        -Message "This is the message"  `
        -Level 1 `
        -Type "Progress" `
        -ShowDebug $false
#>

function Write-Status {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$Message,

        [Parameter(Mandatory=$true)]
        [int]$Level,

        [Parameter(Mandatory=$true)]
        [ValidateSet("Start", "Debug", "Progress", "Warning", "Success", "Error")]
        [String]$Type,

        [Parameter(Mandatory=$false)]
        [Bool]$ShowDebug = $false
    )

    $displayMessage = $true

    switch ($Type) {
        "Start" {
            $ForegroundColours = @{
                ForegroundColor = "Blue"
            }
            break
        }
        "Debug" {
            $Level = ($Level+2)
            $displayMessage = $ShowDebug
            $ForegroundColours = @{
                ForegroundColor = "Cyan"
            }
            break
        }
        "Progress" {
            $ForegroundColours = @{
                ForegroundColor = "DarkMagenta"
            }
            break
        }
        "Warning" {
            $ForegroundColours = @{
                ForegroundColor = "DarkOrange"
            }
            break
        }
        "Success" {
            $ForegroundColours = @{
                ForegroundColor = "Green"
            }
            break
        }
        "Error" {
            $Level = 0
            $ForegroundColours = @{
                ForegroundColor = "Red"
            }
            break
        }
    }

    $Padding = 4*$Level + $Message.Length
    
    $MessageOutput = $Message.PadLeft($Padding," ")

    
    if ($displayMessage) {
        Write-Host $MessageOutput @ForegroundColours
    }
}
Export-ModuleMember -Function "Write-Status"
#endregion Write-Status

#region Check-GraphTeamMembership
<#
    .Synopsis
    Creates a CSV file containing a list of all of the Teams that a user is a member or owner of

    .Description
    Iterates through all Teams that a user is a member of
        Adds a row to the output collection if the user is a member or an owner.
    Outputs the collection to a CSV file named TeamMembership_alan_eardley_company_com.csv
    Returns the the path and name of the CSV file that is created
    
    .Parameter UserUPN
    The UPN of the user to check for - this is case sensitive

    .Parameter CSVPath
    The path and name of the CSV file to create

    .Parameter ShowDebug
    If true, will display debug output

    .Example
    $CSVFilePath = Check-GraphTeamMembership `
        -UserUPN "alan.eardley@company.com"  `
        -CSVPath "C:\Scratch\" `
        -ShowDebug $false
#>

function Check-GraphTeamMembership {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$UserUPN,

        [Parameter(Mandatory=$true)]
        [String]$CSVPath,

        [Parameter(Mandatory=$false)]
        [bool]$ShowDebug = $false
    )

    Write-Status -Message "Setting local variable" -Level $StatusLevel -Type Progress -ShowDebug $ShowDebug
    $outputCollection = @()
    $processItem = 0
    $statusLevel = 1
    $CSVFileName = "TeamMembership_$($UserUPN.Replace(".","_").Replace("@","_at_")).csv"
    $outputCSVFile = "$CSVPath$CSVFileName"

    Write-Status -Message "Get Names" -Level $statusLevel -Type Progress -ShowDebug $ShowDebug
    $teams = Get-Team 
    $processItemCount = $teams.count

    Write-Status -Message "Processing $processItemCount Teams" -Level $statusLevel -Type Progress -ShowDebug $ShowDebug

    $teams | ForEach-Object {
        $processItem++

        Write-Status -Message "Processing Team $processItem of $processItemCount - $_.DisplayName" -Level ($statusLevel + 1) -Type Progress -ShowDebug $ShowDebug

        Write-Status -Message "Get Owners, members and guests" -Level ($statusLevel + 1) -Type Debug -ShowDebug $ShowDebug
        $TeamUsers = Get-TeamUser -GroupId $_.GroupID
        $TeamOwnerCount = ($TeamUsers | Where-Object {$_.Role -like "owner"}).count
        $TeamMemberCount = ($TeamUsers | Where-Object {$_.Role -like "member"}).count
        $TeamGuestCount = ($TeamUsers | Where-Object {$_.Role -like "guest"}).count
    
        $TeamOwnerNames = ""
        if ($TeamOwnerCount -gt 0) {
            Write-Status -Message "Get list of Owners" -Level ($statusLevel + 1) -Type Debug -ShowDebug $ShowDebug
            $TeamOwners = (Get-TeamUser -GroupId $_.GroupID | Where-Object {$_.Role -like "owner"}) | Select-Object -Property User
            $TeamOwners | ForEach-Object {
                $TeamOwnerNames += ";$($_.User)" 
            }
        }

        Write-Status -Message "Check against $userUPN" -Level ($statusLevel + 1) -Type Debug -ShowDebug $ShowDebug
        $IsOwner = ($TeamUsers | Where-Object {($_.Role -like "owner") -and ($_.User -eq $UserUPN)}).count
        $IsMember = ($TeamUsers | Where-Object {($_.Role -like "member") -and ($_.User -eq $UserUPN)}).count
    
        if (($IsOwner -ne 0) -or ($IsMember -ne 0)) {
            Write-Status -Message "$userUPN is owner or member" -Level ($statusLevel + 1) -Type Debug -ShowDebug $ShowDebug

            $output = New-Object -TypeName PSobject 
            $output | add-member NoteProperty "GroupId" -value $_.GroupId
            $output | add-member NoteProperty "DisplayName" -value $_.DisplayName
            $output | add-member NoteProperty "Description" -value $_.Description
            $output | add-member NoteProperty "Visibility" -value $_.Visibility
            $output | add-member NoteProperty "Archived" -value $_.Archived
            $output | Add-Member NoteProperty "OwnerCount" -Value $TeamOwnerCount
            $output | Add-Member NoteProperty "MemberCount" -Value $TeamMemberCount
            $output | Add-Member NoteProperty "GuestCount" -Value $TeamGuestCount
            $output | add-member NoteProperty "OwnerUPNs" -value $TeamOwnerNames.TrimStart(";")
            $output | add-member NoteProperty "IsOwner" -value $IsOwner
            $output | add-member NoteProperty "IsMember" -value $IsMember
            $output | Add-Member NoteProperty "Action" -Value "No action"
            $output | Add-Member NoteProperty "NewOwners" -Value ""
            $output | Add-Member NoteProperty "OldOwners" -Value ""
        
            $outputCollection += $output
        }
    }

    Write-Status -Message "Writing output to $outputCSVFile" -Level $statusLevel -Type Progress -ShowDebug $ShowDebug
    $outputCollection | Export-Csv -Path $outputCSVFile -Force

    return $outputCSVFile
}
Export-ModuleMember -Function "Check-GraphTeamMembership"
#endregion Check-GraphTeamMembership

#region Remove-GraphTeamMember
<#
    .Synopsis
    Reads a CSV file with a list of Teams and removes a user from the membership of each Team 

    .Description
    Provide the UPN of a user and a CSV with a column called GroupID
    For each row
        Check that the user is a Member of the Team
        Remove the user from the Team

    .Parameter UserUPN
    The UPN of the user to check for - this is case sensitive

    .Parameter CSVFilePath
    The path and name of the CSV file to create

    .Parameter ShowDebug
    If true, will display debug output

    .Example
    Remove-GraphTeamMember `
        -UserUPN "alan.eardley@company.com"  `
        -CSVFilePath "C:\Scratch\TeamMembership_alan_eardley_at_company_com.csv" `
        -ShowDebug $false
#>

function Remove-GraphTeamMember {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$UserUPN,

        [Parameter(Mandatory=$true)]
        [String]$CSVFilePath,

        [Parameter(Mandatory=$false)]
        [bool]$ShowDebug = $false
    )

    Write-Status -Message "Setting local variable" -Level $StatusLevel -Type Progress -ShowDebug $ShowDebug
    $processItem = 0
    $statusLevel = 1

    Write-Status -Message "Import CSV $CSVPath" -Level $StatusLevel -Type Progress -ShowDebug $ShowDebug
    $csvRows = Import-Csv -Path $CSVFilePath

    Write-Status -Message "Processing items" -Level $StatusLevel -Type Progress -ShowDebug $ShowDebug
    foreach ($row in $csvRows) {
        $processItem++
        $isMember = $false

        Write-Status -Message "Processing row $processItem" -Level ($StatusLevel+1) -Type Progress -ShowDebug $ShowDebug
        $groupId = $row.GroupId

        Write-Status -Message "Get Team" -Level ($StatusLevel+1) -Type Debug -ShowDebug $ShowDebug
        $team = Get-Team -GroupId $groupId

        Write-Status -Message "Check if $userUPN is a member of $($team.DisplayName)" -Level ($StatusLevel+1) -Type Debug -ShowDebug $ShowDebug
        $isMember = (Get-TeamUser -GroupId $groupId | Where-Object {($_.Role -like "member") -and ($_.User -eq $userUPN)}).count

        if ($isMember -eq 1) {
            Write-Status -Message "Removing $userUPN from $($team.DisplayName)" -Level ($StatusLevel+1) -Type Progress -ShowDebug $ShowDebug
            Remove-TeamUser -GroupId $groupId -User $userUPN -ErrorAction Stop
        } else {
            Write-Status -Message "$userUPN is not a member of $($team.DisplayName)" -Level ($StatusLevel+1) -Type Warning -ShowDebug $ShowDebug
        }
    }
}
Export-ModuleMember -Function "Remove-GraphTeamMember"
#endregion Remove-GraphTeamMember

#region Replace-GraphTeamOwner
<#
    .Synopsis
    Reads a CSV file with a list of Teams. For each Team add a new owner and remove the user from the owner role 

    .Description
    Provide the UPN of a user to remove and a user to add and a CSV with a column called GroupID
    For each row
        Check that the user is an Owner of the Team
        Add the new user as an Owner
        Remove the old user from the Owner role

    .Parameter OldOwnerUPN
    The UPN of the user to check for - this is case sensitive

    .Parameter NewOwnerUPN
    The UPN of the new user to be set as owner - this is case sensitive

    .Parameter CSVFilePath
    The path and name of the CSV file to read

    .Parameter ShowDebug
    If true, will display debug output

    .Example
    Replace-GraphTeamOwner `
        -OldOwnerUPN "user1@company.com"  `
        -NewOwnerUPN "user2@company.com"  `
        -CSVFilePath "C:\Scratch\TeamMembership_alan_eardley_at_company_com.csv" `
        -ShowDebug $false
#>

function Replace-GraphTeamOwner {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$OldOwnerUPN,

        [Parameter(Mandatory=$true)]
        [String]$NewOwnerUPN,

        [Parameter(Mandatory=$true)]
        [String]$CSVFilePath,

        [Parameter(Mandatory=$false)]
        [bool]$ShowDebug = $false
    )

    Write-Status -Message "Setting local variable" -Level $StatusLevel -Type Progress -ShowDebug $ShowDebug
    $processItem = 0
    $statusLevel = 1

    Write-Status -Message "Import CSV $CSVPath" -Level $StatusLevel -Type Progress -ShowDebug $ShowDebug
    $csvRows = Import-Csv -Path $CSVFilePath

    Write-Status -Message "Processing items" -Level $StatusLevel -Type Progress -ShowDebug $ShowDebug
    foreach ($row in $csvRows) {
        $processItem++
        $isOwner = $false

        Write-Status -Message "Processing row $processItem" -Level ($StatusLevel+1) -Type Progress -ShowDebug $ShowDebug
        $groupId = $row.GroupId

        Write-Status -Message "Get Team" -Level ($StatusLevel+1) -Type Debug -ShowDebug $ShowDebug
        $team = Get-Team -GroupId $groupId

        Write-Status -Message "Check if $userUPN is a member of $($team.DisplayName)" -Level ($StatusLevel+1) -Type Debug -ShowDebug $ShowDebug
        $isOwner = (Get-TeamUser -GroupId $groupId | Where-Object {($_.Role -like "owner") -and ($_.User -eq $OldOwnerUPN)}).count

        if ($isOwner -eq 1) {
            Write-Status -Message "Removing $userUPN from $($team.DisplayName)" -Level ($StatusLevel+1) -Type Progress -ShowDebug $ShowDebug
            Add-TeamUser -GroupId $groupId -User $NewOwnerUPN -Role "owner" -ErrorAction Stop
            Remove-TeamUser -GroupId $groupId -User $userUPN -Role "owner" -ErrorAction Stop
        } else {
            Write-Status -Message "$OldOwnerUPN is not a member of $($team.DisplayName)" -Level ($StatusLevel+1) -Type Warning -ShowDebug $ShowDebug
        }
    }
}
Export-ModuleMember -Function "Replace-GraphTeamOwner"
#endregion Replace-GraphTeamOwner

#region Check-GraphTeamsCompliance
<#
    .Synopsis
    Creates a CSV containing a list of all of the Teams in a tenant with a count of owners, members and guests

    .Description
    Iterates through all Teams in a tenante.
        Adds a row to the output collection.
    Outputs the collection to a CSV file named TeamCompliance.csv
    Returns the path and name of the CSV file

    .Parameter CSVPath
    The path and name of the CSV file to create

    .Parameter ShowDebug
    If true, will display debug output

    .Example
    $CSVFilePath = Check-GraphTeamsCompliance `
        -CSVPath "C:\Scratch\" `
        -ShowDebug $false
#>

function Check-GraphTeamsCompliance {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$CSVPath,

        [Parameter(Mandatory=$false)]
        [bool]$ShowDebug = $false
    )

    Write-Status -Message "Setting local variable" -Level $StatusLevel -Type Progress -ShowDebug $ShowDebug
    $outputCollection = @()
    $processItem = 0
    $statusLevel = 1
    $CSVFileName = "TeamsCompliance.csv"
    $outputCSVFile = "$CSVPath$CSVFileName"

    Write-Status -Message "Get Teams" -Level $statusLevel -Type Progress -ShowDebug $ShowDebug
    $groups = Get-MgGroup -All
    $teams = $groups | Where-Object {$_.ResourceProvisioningOptions -Contains "Team"}
    $processItemCount = $teams.count

    Write-Status -Message "Processing $processItemCount Teams" -Level $statusLevel -Type Progress -ShowDebug $ShowDebug

    $teams | ForEach-Object {
        $processItem++

        Write-Status -Message "Processing Team $processItem of $processItemCount - $($_.DisplayName)" -Level ($statusLevel + 1) -Type Progress -ShowDebug $ShowDebug

        Write-Status -Message "Get Owners, members and guests" -Level ($statusLevel + 1) -Type Debug -ShowDebug $ShowDebug
        $TeamUsers = Get-MgTeamMember -TeamId $_.Id
        $TeamOwnerCount = ($TeamUsers | Where-Object {$_.Roles -contains "owner"}).count
        $TeamMemberCount = ($TeamUsers | Where-Object {$_.Roles -contains "member"}).count
        $TeamGuestCount = ($TeamUsers | Where-Object {$_.Roles -contains "guest"}).count
    
        $TeamOwnerNames = ""
        if ($TeamOwnerCount -gt 0) {
            Write-Status -Message "Get list of Owners" -Level ($statusLevel + 1) -Type Debug -ShowDebug $ShowDebug
        #    $TeamOwners = (Get-TeamUser -GroupId $_.GroupID | Where-Object {$_.Role -like "owner"}) | Select-Object -Property User
            $TeamOwners | ForEach-Object {
                $TeamOwnerNames += ";$($_.User)" 
            }
        }
    
        Write-Status -Message "$userUPN is owner or member" -Level ($statusLevel + 1) -Type Debug -ShowDebug $ShowDebug

        $output = New-Object -TypeName PSobject 
        $output | add-member NoteProperty "GroupId" -value $_.GroupId
        $output | add-member NoteProperty "DisplayName" -value $_.DisplayName
        $output | add-member NoteProperty "Description" -value $_.Description
        $output | add-member NoteProperty "Visibility" -value $_.Visibility
        $output | add-member NoteProperty "Archived" -value $_.Archived
        $output | Add-Member NoteProperty "OwnerCount" -Value $TeamOwnerCount
        $output | Add-Member NoteProperty "MemberCount" -Value $TeamMemberCount
        $output | Add-Member NoteProperty "GuestCount" -Value $TeamGuestCount
        $output | add-member NoteProperty "OwnerUPNs" -value $TeamOwnerNames.TrimStart(";")
        $output | Add-Member NoteProperty "Action" -Value "No action"
        $output | Add-Member NoteProperty "NewOwners" -Value ""
        $output | Add-Member NoteProperty "OldOwners" -Value ""
        
        $outputCollection += $output
    }

    Write-Status -Message "Writing output to $outputCSVFile" -Level $statusLevel -Type Progress -ShowDebug $ShowDebug
    $outputCollection | Export-Csv -Path $outputCSVFile -Force

    return $outputCSVFile
}
Export-ModuleMember -Function "Check-GraphTeamsCompliance"
#endregion Check-GraphTeamsCompliance

#region Add-GraphTeamOwner
<#
    .Synopsis
    Reads a CSV file with a list of Teams and adds a user as the owner 

    .Description
    Provide a CSV with a column called GroupID to identify Teams and a column called NewOwnerUPN containing the UPN of a user to be added as Owner

    .Parameter NewOwnerUPN
    The UPN of the new user to be set as owner - this is case sensitive

    .Parameter CSVFilePath
    The path and name of the CSV file to read

    .Parameter ShowDebug
    If true, will display debug output

    .Example
    Add-GraphTeamOwner `
        -CSVFilePath "C:\Scratch\" `
        -ShowDebug $false
#>

function Add-GraphTeamOwner {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$CSVFilePath,

        [Parameter(Mandatory=$false)]
        [bool]$ShowDebug = $false
    )

    Write-Status -Message "Setting local variable" -Level $StatusLevel -Type Progress -ShowDebug $ShowDebug
    $processItem = 0
    $statusLevel = 1

    Write-Status -Message "Import CSV $CSVPath" -Level $StatusLevel -Type Progress -ShowDebug $ShowDebug
    $csvRows = Import-Csv -Path $CSVFilePath

    Write-Status -Message "Processing items" -Level $StatusLevel -Type Progress -ShowDebug $ShowDebug
    foreach ($row in $csvRows) {
        $processItem++

        Write-Status -Message "Processing row $processItem" -Level ($StatusLevel+1) -Type Progress -ShowDebug $ShowDebug
        $groupId = $row.GroupId
        $newOwnerUPN = $row.NewOwnerUPN

        Write-Status -Message "Get Team" -Level ($StatusLevel+1) -Type Debug -ShowDebug $ShowDebug
        $team = Get-Team -GroupId $groupId

        Write-Status -Message "Adding $newOwnerUPN as Owner of $($team.DisplayName)" -Level ($StatusLevel+1) -Type Progress -ShowDebug $ShowDebug
        Add-TeamUser -GroupId $groupId -User $NewOwnerUPN -Role "owner" -ErrorAction Stop
    }
}
Export-ModuleMember -Function "Add-GraphTeamOwner"
#endregion Add-GraphTeamOwner

#region Update-GraphTeamsCompliance
<#
    .Synopsis
    Reads a CSV file from Check-GraphTeamsCompliance and checks the Action column to determine whether to add new owners, archive or remove a Team 

    .Description
    Provide a CSV with a column called GroupID to identify Teams and a column called Action to identify the action to take:
        UpdateOwners - Add the semi-colon separated list of UPNs in the NewOwners column and remove those in the OldOwners column
        Archive - Set the Team to archived
        Delete - Delete the Team

    .Parameter CSVFilePath
    The path and name of the CSV file to read

    .Parameter ShowDebug
    If true, will display debug output

    .Example
    Update-GraphTeamsCompliance `
        -CSVFilePath "C:\Scratch\" `
        -ShowDebug $false
#>

function Update-GraphTeamsCompliance {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$CSVFilePath,

        [Parameter(Mandatory=$false)]
        [bool]$ShowDebug = $false
    )

    Write-Status -Message "Setting local variable" -Level $StatusLevel -Type Progress -ShowDebug $ShowDebug
    $processItem = 0
    $statusLevel = 1

    Write-Status -Message "Import CSV $CSVPath" -Level $StatusLevel -Type Progress -ShowDebug $ShowDebug
    $csvRows = Import-Csv -Path $CSVFilePath

    Write-Status -Message "Processing items" -Level $StatusLevel -Type Progress -ShowDebug $ShowDebug
    foreach ($row in $csvRows) {
        $processItem++
        $continue = $true

        Write-Status -Message "Processing row $processItem" -Level ($StatusLevel+1) -Type Progress -ShowDebug $ShowDebug
        $action = $row.Action
        $groupId = $row.GroupId
        $newOwnersString = $row.NewOwners
        $oldOwnersString = $row.OldOwners

        try {
            $team = Get-Team -GroupId $groupId
        } catch {
            Write-Status -Message "Team $($team.DisplayName) does not exist" -Level ($StatusLevel+2) -Type Progress -ShowDebug $ShowDebug
            $continue = $false
        }

        if ($continue) {
            switch ($action) {
                "UpdateOwners" {
                    Write-Status -Message "Changing Owners of Team $($team.DisplayName)" -Level ($StatusLevel+2) -Type Progress -ShowDebug $ShowDebug

                    if ($newOwnersString -ne "") {
                        $newOwners = $newOwnersString.split(";")
                        
                        foreach ($newOwner in $newOwners) {
                            $newOwnerExists = (Get-TeamUser -GroupId $groupId | Where-Object {($_.Role -like "owner") -and ($_.User -eq $newOwner)}).count
                            if ($newOwnerExists -eq 0) {
                                Write-Status -Message "Adding user $newOwner" -Level ($StatusLevel+2) -Type Progress -ShowDebug $ShowDebug
                                Add-TeamUser -GroupId $groupId -User $newOwner -Role "owner" -ErrorAction Stop
                            } else {
                                Write-Status -Message "User $newOwner is already an Owner" -Level ($StatusLevel+2) -Type Success -ShowDebug $ShowDebug
                            }
                        }
                    }

                    if ($oldOwnersString -ne "") {
                        $oldOwners = $oldOwnersString.split(";")
                        
                        foreach ($oldOwner in $oldOwners) {
                            $oldOwnerExists = (Get-TeamUser -GroupId $groupId | Where-Object {($_.Role -like "owner") -and ($_.User -eq $oldOwner)}).count
                            if ($oldOwnerExists -eq 1) {
                                Write-Status -Message "Removing owner $oldOwner" -Level ($StatusLevel+2) -Type Progress -ShowDebug $ShowDebug
                                Remove-TeamUser -GroupId $groupId -User $oldOwner -Role "owner" -ErrorAction Stop
                            } else {
                                Write-Status -Message "User $oldOwner is not an Owner" -Level ($StatusLevel+2) -Type Success -ShowDebug $ShowDebug
                            }
                        }
                    }
                    break
                }
                "Archive" {
                    if ($team.Archived) {
                        Write-Status -Message "Team $($team.DisplayName) is already archived" -Level ($StatusLevel+2) -Type Success -ShowDebug $ShowDebug
                    } else {
                        Write-Status -Message "Archiving Team $($team.DisplayName)" -Level ($StatusLevel+2) -Type Progress -ShowDebug $ShowDebug
                        Set-TeamArchivedState -GroupId $groupId -Archived $true
                    }
                    break
                }
                "Delete" {
                    Write-Status -Message "Deleting Team $($team.DisplayName)" -Level ($StatusLevel+2) -Type Progress -ShowDebug $ShowDebug
                    Remove-Team -GroupId $groupId
                    
                    break
                }
                Default {
                    Write-Status -Message "No action required" -Level ($StatusLevel+2) -Type Progress -ShowDebug $ShowDebug
                }
            }
        }
    }
}
Export-ModuleMember -Function "Update-GraphTeamsCompliance"
#endregion Update-GraphTeamsCompliance