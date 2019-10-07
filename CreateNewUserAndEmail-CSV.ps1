<#
.DESCRIPTION
    This script is meant to be used to create new users from a specially formatted CSV file.

.USAGE

  PREPARE CSV FILE WITH USERS
    Prepare a CSV like the table below. Required fields are indicated above the table column. Column headers must match exactly. There will be a template available on M365 Resort Admin Team's PowerShell Channel.

      Required    Required    Required        Required                                                                                                                                 Required   Required                               Required    Required  
    +-----------+-----------+---------------+--------------+--------+---------------------+-----------------------------+-----------+--------------------+------------+--------------+----------+----------+---------------------------+-----------+--------------+
    | firstName | lastName  | alias         | tempPassword | office | officePhone         | description                 | title     | department         | company    | managerAlias | extAttr2 | extAttr3 | userToCloneSecurityGroups | upnSuffix | resortSuffix |
    +-----------+-----------+---------------+--------------+--------+---------------------+-----------------------------+-----------+--------------------+------------+--------------+----------+----------+---------------------------+-----------+--------------+
    | Test      | Spencer 8 | test_spencer8 | Summit18     | Summit | (909) 866-5766 x140 | 487 - Testing Dept - Tester | testerrrr | Testing Department | BBMR Rocks | sstewart     | BBMR     | F1;      | test_spencer3             | bbmr.com  | (SS)         |
    +-----------+-----------+---------------+--------------+--------+---------------------+-----------------------------+-----------+--------------------+------------+--------------+----------+----------+---------------------------+-----------+--------------+

  SET PARAMETERS
    + CSV File Path: Save the CSV file in an easy to find place. In the parameters section, set $newUsersFile equal to the absolute file path to the CSV file.
    + New User OU: Set this to equal the Name, Canonical Name, Distinguished Name (DN), or GUID of the OU you want these new users to show up in
    + Local Exchange Server: Set $ExchangeServerName equal to the FQDN of the on-prem Exchange server that your users live in

  RUN SCRIPT
    To run the script, open Exchange Online Powershell Module. Change directories to where this script is saved, and execute the script be typing the script name in the prompt (use Tab complete if you like).
    The script will prompt you first for your Exchange Online admin credentials. This is used to check the tenant for any existence of aliases that match the users you're trying to make
    Next, the script will prompt you for your on-prem Exchange admin credentials. This is used to create the remote mailbox and set attributes in Active Directory.

    If the script completed with no errors, the new users are placed in the OU you designated in the parameters. Don't forget to move these new users to their appropriate OU!


.NOTES
    

.AUTHOR
    Spencer Stewart, Big Bear Mountain Resort

.DATE
    Created: 2019-10-04
    Last modified: 2019-10-04
#>


# BEGIN: Parameters. Please update for your resort

$ExchangeServerName = "bbmr-exch2013.bbmr.local"
$newUserOU = "OU=New Users,OU=Standard User,OU=BBMR Users,OU=Snow_Summit_LLC,DC=BBMR,DC=local"
$newUsersFile = "C:\temp\CreateNewUsers.csv"

# END: Parameters. Please update for your resort





# BEGIN: Helper functions used by the script


# Checks availability in the tenant for the alias. Exits if alias already exists and prints the existing user's displayName.
function CheckAliasAvailability
{
    param($alias)
    $DisplayName =  ( Get-Recipient -RecipientTypeDetails GroupMailbox, SharedMailbox, UserMailbox, MailContact, MailUniversalDistributionGroup, DynamicDistributionGroup, MailUser, RoomMailbox -filter "alias -eq '$alias'" ).Name

    if ($DisplayName -ne $null)
    {
        Write-Error "[$alias] Alias is taken by $DisplayName"
        exit
    } else {
        Write-Output "[$alias] Alias is available!"
    }
}


# Connects to on-prem Exchange Server
function ConnectToOnPremExchange
{
    param($ExchangeServerName)

    # Assemble URI
    $ExchangeServerUri = "http://" + $ExchangeServerName + "/PowerShell/"

    # Sets-up a new session
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeServerUri -Authentication Kerberos -Credential $UserCredential

    # Imports new session
    Write-Host "### CONNECTING TO ON PREM EXCHANGE ###"
    Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null
}


# Creates the remote mailbox user
function CreateRemoteMailboxUser
{
    param($firstName, $lastName, $alias, $tempPassword, $upnSuffix, $resortSuffix, $newUserOU)

    # Convert password to secure string
    $password = ConvertTo-SecureString $tempPassword -AsPlainText -Force
    $name = "$firstName $lastName $resortSuffix"
    $upn = $alias + "@" + $upnSuffix

    # Create Remote Mailbox & User
    Write-Host "[$alias] Creating User"
    New-RemoteMailbox -Name $name `
         -FirstName $firstName `
         -LastName $lastName `
         -UserPrincipalName $upn `
         -Password $password `
         -OnPremisesOrganizationalUnit $newUserOU `
         -ResetPasswordOnNextLogon:$true `
         | Out-Null
}

# Adds some optional and required additional attributes to the user
function AddAdditionalAttributesToUser
{
    param($alias, $description, $office, $officePhone, $title, $department, $company, $managerAlias, $extAttr2, $extAttr3)


    # Add additional attributes in AD
    Write-Host "[$alias] Setting additional attributes"

    $command = "Set-ADUser -Credential `$UserCredential -Identity $alias -Replace @{
        extensionAttribute2=`"$extAttr2`"
        extensionAttribute3=`"$extAttr3`"
    }"

    if (![string]::IsNullOrWhiteSpace($description)) { $command += " -description `"$description`"" }
    if (![string]::IsNullOrWhiteSpace($office)) { $command += " -office `"$office`"" }
    if (![string]::IsNullOrWhiteSpace($officePhone)) { $command += " -officePhone `"$officePhone`"" }
    if (![string]::IsNullOrWhiteSpace($title)) { $command += " -title `"$title`"" }
    if (![string]::IsNullOrWhiteSpace($department)) { $command += " -department `"$department`"" }
    if (![string]::IsNullOrWhiteSpace($company)) { $command += " -company `"$company`"" }
    if (![string]::IsNullOrWhiteSpace($manager)) { $command += " -manager `"$managerAlias`"" }

    Invoke-Expression $command

    # Try to set manager, else write a message
    try {
        Set-ADUser -Credential $UserCredential -Identity $alias -Manager $managerAlias
    } catch {
        Write-Host "[$alias] [error] Could not find manager based on alias provided!"
    }

}



# Clone group membership if sourceUser provided
function CloneSecurityGroupMembership
{
    param($sourceUser, $targetUser)
    
    # Copy Security Groups
    if (![string]::IsNullOrWhiteSpace($sourceUser))
    {
        try {
            Get-ADUser -Identity "$sourceUser" -Properties memberof |
                Select-Object -ExpandProperty memberof |
                    Add-ADGroupMember -Credential $UserCredential -Members "$targetUser"

            Write-Host "[$targetUser] Security group membership cloned from $sourceUser"
        } catch {
            Write-Host "[$targetUser] [error] Could not find source user: $sourceUser"            
        }
    } else {
        Write-Host "[$targetUser] No source user provided. Group membership was NOT cloned."
    }

    

    Write-Host
}

# END: helper functions




# BEGIN: executing the script


# Get users from csv file
$newUsers = Import-Csv $newUsersFile

# Check alias availability from the tenant
Write-Host "### CONNECTING TO EXCHANGE ONLINE ###"
#Connect-EXOPSSession -WarningAction SilentlyContinue
foreach ($user in $newUsers)
{
    #CheckAliasAvailability $user.alias
}
# Disconnect Remote PowerShell session to Exchange Online
Get-PSSession | Remove-PSSession


# Create remote user, assign attributes, clone security groups
$UserCredential = Get-Credential -Credential $null
ConnectToOnPremExchange $ExchangeServerName

foreach ($user in $newUsers)
{
    CreateRemoteMailboxUser $user.firstName $user.lastName $user.alias $user.tempPassword $user.upnSuffix $user.resortSuffix $newUserOU
    AddAdditionalAttributesToUser $user.alias $user.description $user.office $user.officePhone $user.title $user.department $user.company $user.managerAlias $user.extAttr2 $user.extAttr3
    CloneSecurityGroupMembership -sourceUser $user.userToCloneSecurityGroups -targetUser $user.alias
}

# Disconnect remote PowerShell Session
Get-PSSession | Remove-PSSession