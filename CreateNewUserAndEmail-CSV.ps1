<#
.DESCRIPTION
    This script is meant to be used to create new users from a specially formatted CSV file.

.USAGE

  PREPARE CSV FILE WITH USERS
    Prepare a CSV like the table below. Required fields are indicated above the table column. Column headers must match exactly. There is a template included with the scripts.

      Required  Required   Required   Required                                                                                                         Required                               
    +----------+----------+----------+--------------+--------+---------------------+-----------------+-----------+--------------------+--------------+----------+---------------------------+
    | fName    | lName    | alias    | tempPassword | office | officePhone         | description     | title     | department         | managerAlias | extAttr3 | userToCloneSecurityGroups |
    +----------+----------+----------+--------------+--------+---------------------+-----------------------------+--------------------+--------------+----------+---------------------------+
    | Taylor   | Swift    | tswift   | Letit$n0w    | Summit | (909) 866-5766 x140 | Country Singer? | testerrrr | Testing Department | sstewart     | F1;      | jjonas                    |
    +----------+----------+----------+--------------+--------+---------------------+-----------------+-----------+--------------------+--------------+----------+---------------------------+

  SET PARAMETERS
    + CSV File Path: Either use the template file included with the scripts or update the filepath parameter after this comment block.
    + Additional resort related parameters are in "HelperFunctions.ps1" file. The helper functions file should be in the same directory (folder) as this script.

  RUN SCRIPT
    To run the script, open Exchange Online Powershell Module. Change directories to where this script is saved, and execute the script be typing the script name in the prompt (use Tab complete if you like).
    The script will prompt you first for your Exchange Online admin credentials. This is used to check the tenant for whether the desired alias is available or not.
    Next, the script will prompt you for your Exchange On-Prem admin credentials. Use your full login (eg jdoe@idirectory.itw). This is used to create the remote mailbox and set attributes in Active Directory.

    If the script completed with no errors, the new users are placed in the OU designated in the resort config settings. Don't forget to move these new users to their appropriate OU!


.NOTES
    
.TODO

.AUTHOR
    Spencer Stewart, Big Bear Mountain Resort

.DATE
    Created: 2019-10-04
    Last modified: 2019-10-04
#>



# BEGIN: Parameters. Please specify the absolute path to the formatted CSV with new user data

$newUsersFile = ".\NewUsers.csv"

# END: Parameters.


# IMPORT HELPER FUNCTIONS
. .\HelperFunctions.ps1


# BEGIN: executing the script


# Get users from csv file
Write-Host "### Getting users from $newUsersFile ###" -ForegroundColor DarkGray
$newUsers = Import-Csv $newUsersFile

# Check alias availability from the tenant
Write-Host "`n### Connecting to Exchange Online to verify alias availability ###" -ForegroundColor DarkGray
Connect-EXOPSSession -WarningAction SilentlyContinue
foreach ($user in $newUsers)
{
    $user.alias = CheckAliasAvailability -alias $user.alias
}
Get-PSSession | Remove-PSSession # Disconnect from Exchange Online PowerShell


# Create remote user, assign attributes, clone security groups
Write-Host "`n### CONNECTING TO ON PREM EXCHANGE ###" -ForegroundColor DarkGray
$UserCredential = $host.ui.PromptForCredential("Exchange On-Prem Credentials", "Please enter your Exchange On-Prem Creds.", "", "NetBiosUserName")
ConnectToOnPremExchange $ExchangeServerName

foreach ($user in $newUsers)
{
    # Validate input
    $user.managerAlias = Get-RealADUser -userType "Manager" -isFromCSV $true -samAccountName $user.managerAlias
    $user.userToCloneSecurityGroups = Get-RealADUser -userType "Source user to copy security groups" -isFromCSV $true -samAccountName $user.userToCloneSecurityGroups
    $user.extAttr3 = GetExtAttr3 -extAttr3 $user.extAttr3
    
    # Create User
    CreateRemoteMailboxUser $user.fname $user.lname $user.name $user.alias $user.tempPassword $newUserOU -isFromCSV $true
    AddAdditionalAttributesToUser $user.alias $user.description $user.office $user.officePhone $user.title $user.department $user.managerAlias $user.extAttr3
    if (![string]::IsNullOrWhiteSpace($user.userToCloneSecurityGroups)) {
        CloneSecurityGroupMembership -sourceUser $user.userToCloneSecurityGroups -targetUser $user.alias
    }

    Write-Host ""
}

# Disconnect remote PowerShell Session
Get-PSSession | Remove-PSSession