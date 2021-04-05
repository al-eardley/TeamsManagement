# Introduction 
Whenever someone leaves an organisation or changes roles, the permissions of their user account should be checked to ensure that
1. No Teams are left without an owner (every Team should really have two owners)
2. Their access to Teams they no longer need access to is revoked.

There are many ways to do this but most organisations will consider this when the need arises. This project contains a PowerShell module called TeamsManagement that provides commandlets to carry out some essential actions to support the management of Teams permissions. 

# Getting Started
1.	Run the Connect.ps1 script to load the required modules
2.	Take a copy of the Test.ps1 script and adapt it to your specific needs

# The Commandlets

## Check-TeamMembership
Creates a CSV file containing a list of all of the Teams that a user is a member or owner of.

`$CSVFilePath = Check-TeamMembership -UserUPN "alan.eardley@company.com" -CSVPath "C:\Scratch\" -ShowDebug $false `

The path and name of the output file is returned from the commandlet and take the form of TeamMembership_alan_eardley_at_comapny_com.csv

## Replace-TeamOwner
Reads a CSV file with a list of Teams. For each Team add a new owner and remove the user from the owner role

The CSV requires the following column
- GroupId - contains the GroupID of the Teams where the owner will be replaced

`Replace-TeamOwner -OldOwnerUPN "user1@company.com" -NewOwnerUPN "user2@company.com" -CSVFilePath "C:\Scratch\TeamMembership_alan_eardley_at_company_com.csv" -ShowDebug $false`

## Remove-TeamMember
Reads a CSV file with a list of Teams and removes a user from the membership of each Team 

The CSV requires the following column
- GroupId - contains the GroupID of the Teams from with the user will be removed

`Remove-TeamMember -UserUPN "alan.eardley@company.com" -CSVFilePath "C:\Scratch\TeamMembership_alan_eardley_at_comapny_com.csv" -ShowDebug $false`

## Check-TeamsCompliance
Creates a CSV containing a list of all of the Teams in a tenant with a count of owners, members and guests

`$CSVFilePath = Check-TeamsCompliance -CSVPath "C:\Scratch\" -ShowDebug $false`

## Add-TeamOwner
Reads a CSV file with a list of Teams and adds a user as the owner

The CSV requires the following column
- GroupId - contains the GroupID of the Teams where the owners will be added
- NewOwnerUPN - The UPN of the user to add as an owner

`Add-TeamOwner -CSVFilePath "C:\Scratch\AddOwner.csv" -ShowDebug $false`

# How to use the Commandlets
## A user is changing role
Aim: Find all the Teams that a user is an Owner of, review them and replace them with someone else

Steps:
`$CSVFilePath = Check-TeamMembership -UserUPN "alan.eardley@company.com" -CSVPath "C:\Scratch\" -ShowDebug $false `

Review the CSV file that is created and save a copy that contains only the Teams where the replacement is needed remain

Review the CSV file that is created and save a copy that contains only the Teams where the user should no longer have access

`Replace-TeamOwner -OldOwnerUPN "alan.eardley@company.com" -NewOwnerUPN "new.owner@company.com" -CSVFilePath "C:\Scratch\ReplaceOwner_alan_eardley_at_company_com.csv" -ShowDebug $false`

`Remove-TeamMember -UserUPN "alan.eardley@company.com" -CSVFilePath "C:\Scratch\RemoveMember_alan_eardley_at_comapny_com.csv" -ShowDebug $false`

## Add owners to Teams with no owners
Aim: find all Teams with no owners and add an owner

Steps:
`$CSVFilePath = Check-TeamsCompliance -CSVPath "C:\Scratch\" -ShowDebug $false`

Review the CSV file and add a column called NewOwnerUPN containing the UPN of the user to be added to the Team as an owner.

`Add-TeamOwner -CSVFilePath "C:\Scratch\AddOwner.csv" -ShowDebug $false`
