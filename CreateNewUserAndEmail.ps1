<#
.DESCRIPTION
    This script is meant to create an individual new user. There is a separate script to import a CSV file of users.

.USAGE

  SET PARAMETERS
    Update the Resort Config settings just after this section.

  RUN SCRIPT
    To run the script, open Exchange Online Powershell Module. Change directories to where this script is saved, and execute the script be typing the script name in the prompt (use Tab complete if you like).
    The script will prompt you first for your Exchange Online admin credentials. This is used to check the tenant for any existence of aliases that match the users you're trying to make
    Next, the script will prompt you for your on-prem Exchange admin credentials. This is used to create the remote mailbox and set attributes in Active Directory.

    If the script completed with no errors, the new users are placed in the OU you designated in the parameters. Don't forget to move these new users to their appropriate OU!


.NOTES
    
.TODO
    Validate extensionAttribute3 value


.AUTHOR
    Spencer Stewart, Big Bear Mountain Resort

.DATE
    Created: 2019-10-07
    Last modified: 2019-10-07
#>

#####################################################
# BEGIN: RESORT CONFIG. Please update for your resort
#####################################################

# newUserOU: Can be Name, Canonical Name, Distinguished Name, or GUID for where new users will be created
$newUserOU = "OU=New Users,OU=Standard User,OU=BBMR Users,OU=Snow_Summit_LLC,DC=BBMR,DC=local"

# Local Exchange Server FQDN
$ExchangeServerName = "bbmr-exch2013.bbmr.local"

# Domain Controller. Use FQDN
$DomainController = "bbmrdc1-2012.bbmr.local"
#$DomainController = "vm-den-dc01.iDirectory.itw"

# UPN Suffix
$upnSuffix = "bbmr.com"

# Display Name Suffix. Include the parentheses - eg "(DEN)" for the user "John Doe (DEN)"
$resortSuffix = "(SS)"

# Resort code for Extension Attribute 2
$resortCode = "BBMR"

# Company for Active Directory
$company = "BBMR"

#####################################################
# END: RESORT CONFIG.
#####################################################




# BEGIN: Helper functions used by the script


# Checks availability in the tenant for the alias. Exits if alias already exists and prints the existing user's displayName.
function GetUnusedAlias
{

    $alias = Read-Host "Alias [required]"

    Write-Host "### Connecting to Exchange Online to verify alias availability ###" -ForegroundColor DarkGray

    Connect-EXOPSSession -WarningAction SilentlyContinue


    do
    {
        $DisplayName =  ( Get-Recipient -RecipientTypeDetails GroupMailbox, SharedMailbox, UserMailbox, MailContact, MailUniversalDistributionGroup, DynamicDistributionGroup, MailUser, RoomMailbox -filter "alias -eq '$alias'" ).Name

        if ($DisplayName -ne $null)
        {
            $aliasIsAvailable = $false
            Write-Host "[$alias] Alias is taken by $DisplayName" -ForegroundColor Red
            $alias = Read-Host "Enter a different alias"
        } else {
            Write-Host "[$alias] Alias is available!" -ForegroundColor Green
            $aliasIsAvailable = $true
        }
    } until ($aliasIsAvailable)

    # Disconnect Remote PowerShell session to Exchange Online
    Get-PSSession | Remove-PSSession

    $alias
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
    Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null
}


# Creates the remote mailbox user
function CreateRemoteMailboxUser
{
    param($firstName, $lastName, $name, $alias, $tempPassword, $upnSuffix, $resortSuffix, $newUserOU)

    # Convert password to secure string
    # $password = ConvertTo-SecureString $tempPassword -AsPlainText -Force # Used for CSV powershell

    $upn = $alias + "@" + $upnSuffix

    # Create Remote Mailbox & User
    Write-Host "[$alias] Creating User" -ForegroundColor Green
    New-RemoteMailbox -Name $name `
         -FirstName $firstName `
         -LastName $lastName `
         -UserPrincipalName $upn `
         -Password $tempPassword `
         -OnPremisesOrganizationalUnit $newUserOU `
         -ResetPasswordOnNextLogon:$true `
         -DomainController $DomainController `
         | Out-Null
}

# Adds some optional and required additional attributes to the user
function AddAdditionalAttributesToUser
{
    param($alias, $description, $office, $officePhone, $title, $department, $company, $managerAlias, $extAttr2, $extAttr3)


    # Add additional attributes in AD
    Write-Host "[$alias] Setting additional attributes" -ForegroundColor Green

    $command = "Set-ADUser -Credential `$UserCredential -Identity $alias -Server $DomainController `
        -Replace @{
        extensionAttribute2=`"$extAttr2`"
        extensionAttribute3=`"$extAttr3`"
    }"

    if (![string]::IsNullOrWhiteSpace($description)) { $command += " -description `"$description`"" }
    if (![string]::IsNullOrWhiteSpace($office)) { $command += " -office `"$office`"" }
    if (![string]::IsNullOrWhiteSpace($officePhone)) { $command += " -officePhone `"$officePhone`"" }
    if (![string]::IsNullOrWhiteSpace($title)) { $command += " -title `"$title`"" }
    if (![string]::IsNullOrWhiteSpace($department)) { $command += " -department `"$department`"" }
    if (![string]::IsNullOrWhiteSpace($company)) { $command += " -company `"$company`"" }
    if (![string]::IsNullOrWhiteSpace($managerAlias)) { $command += " -manager `"$managerAlias`"" }

    Invoke-Expression $command

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
                    Add-ADGroupMember -Credential $UserCredential -Server $DomainController -Members "$targetUser"

            Write-Host "[$targetUser] Security group membership cloned from $sourceUser" -ForegroundColor Green
        } catch {
            Write-Host "[$targetUser] [error] Could not find source user: $sourceUser"            
        }
    } else {
        Write-Host "[$targetUser] No source user provided. Group membership was NOT cloned."
    }

    

    Write-Host
}

function GetExtAttr3
{
    
    do
    {
        $extAttr3 = Read-Host "ExtensionAttribute3 (licensing) [required]"
        
        if ($extAttr3 -eq "F1;" -or $extAttr3 -eq "E1;" -or $extAttr3 -eq "E3;") {
            $validExtAttr3 = $true;
        } else {
            Write-Host "ExtensionAttribute 3 is invalid. Please use 'F1;', 'E1;', or 'E3;'." -ForegroundColor Red
            $validExtAttr3 = $false;
        }

    } until ($validExtAttr3)

    $extAttr3
}

function Get-RealADUser
{
    param($userType)

    $samAccountName = Read-Host "$($userType)'s Alias [optional]"

    do
    {
        try {
            if ([string]::IsNullOrWhiteSpace($samAccountName)) {
                $validUser = $true # Break out of loop if no entry specified
            } else {
                $ADUser = Get-ADUser -Identity $samAccountName
                Write-Host "Using $($ADUser.Name) ($samAccountName) as $userType" -ForegroundColor DarkGray
                $validUser = $true
            }
        }
        catch {
            Write-Host "Unable to find $userType with alias '$samAccountName'." -ForegroundColor Red
            $samAccountName = Read-Host "Enter a different alias or type 'S' to skip"
            if ($samAccountName.ToLower() -eq "s") {
                $samAccountName = ""
                $validUser= $true
            } else {
                $validUser = $false
            }
        }
    } until ($validUser)

    $samAccountName
}


# END: helper functions




# BEGIN: executing the script


# Write configuration information
Write-Host "
################################################################# `
    This script will attempt to create a user with the following config `
    - Exchange Server: $ExchangeServerName `
    - ExtensionAttribute2: $resortCode `
    - Resort UPN Suffix: $upnSuffix `
    - Company in AD: $company `
    `
    If these values should be changed, please edit this script in the 'Resort Config' section.
    You can cancel the script by entering Ctrl + C.
    You can accept the default values by simply pressing Enter.
#################################################################
    " -ForegroundColor DarkGray



# Collect new user information
$user = New-Object -TypeName psobject
$user | Add-Member -NotePropertyName fname -NotePropertyValue ""
$user | Add-Member -NotePropertyName lname -NotePropertyValue ""
$user | Add-Member -NotePropertyName name -NotePropertyValue ""
$user | Add-Member -NotePropertyName alias -NotePropertyValue ""
$user | Add-Member -NotePropertyName tempPassword -NotePropertyValue ""
$user | Add-Member -NotePropertyName extAttr2 -NotePropertyValue $resortCode
$user | Add-Member -NotePropertyName extAttr3 -NotePropertyValue ""
$user | Add-Member -NotePropertyName title -NotePropertyValue ""
$user | Add-Member -NotePropertyName department -NotePropertyValue ""
$user | Add-Member -NotePropertyName managerAlias -NotePropertyValue ""
$user | Add-Member -NotePropertyName description -NotePropertyValue ""
$user | Add-Member -NotePropertyName office -NotePropertyValue ""
$user | Add-Member -NotePropertyName officePhone -NotePropertyValue ""
$user | Add-Member -NotePropertyName userToCloneSecurityGroups -NotePropertyValue ""
$user | Add-Member -NotePropertyName company -NotePropertyValue $company
$user | Add-Member -NotePropertyName upnSuffix -NotePropertyValue $upnSuffix
$user | Add-Member -NotePropertyName resortSuffix -NotePropertyValue $resortSuffix

# Required details
Write-Host "### REQUIRED DETAILS ###" -ForegroundColor DarkGray

$user.fname = Read-Host "First Name [required]"
$user.lname = Read-Host "Last Name [required]"
$user.name = "$($user.fname) $($user.lname) $resortSuffix"
$prompt = Read-Host "Display Name [default: $($user.name)]"
$user.name = ($user.name,$prompt)[[bool]$prompt]
$user.alias = GetUnusedAlias
$user.extAttr3 = GetExtAttr3
$user.tempPassword = Read-Host "Temporary Password [required]" -AsSecureString

# Strongly suggested details
Write-Host "### OPTIONAL BUT SUGGESTED DETAILS ###" -ForegroundColor DarkGray

$user.managerAlias = Get-RealADUser "Manager"
$user.userToCloneSecurityGroups = Get-RealADUser "Source user to copy security groups"
$user.title = Read-Host "Title [optional]"
$user.department = Read-Host "Department [optional]"
$user.description = Read-Host "Description [optional]"
$user.office = Read-Host "Office [optional]"
$user.officePhone = Read-Host "Office Phone [optional]"

# Default details
Write-Host "### DEFAULT DETAILS ###" -ForegroundColor DarkGray
$prompt = Read-Host "extensionAttribute2 [default: $($user.extAttr2)]"
$user.extAttr2 = ($user.extAttr2,$prompt)[[bool]$prompt]
$prompt = Read-Host "Company [default: $($user.company)]"
$user.company = ($user.company,$prompt)[[bool]$prompt]
$prompt = Read-Host "UPN Suffix [default: $($user.upnSuffix)]"
$user.upnSuffix = ($user.upnSuffix,$prompt)[[bool]$prompt]




# Create remote user, assign attributes, clone security groups
Write-Host "### CONNECTING TO ON PREM EXCHANGE ###" -ForegroundColor DarkGray
# $UserCredential = Get-Credential -Credential $null -Message "Enter your Exchange On-Prem credentials"
$UserCredential = $host.ui.PromptForCredential("Exchange On-Prem Credentials", "Please enter your Exchange On-Prem Creds.", "", "NetBiosUserName")
ConnectToOnPremExchange $ExchangeServerName


CreateRemoteMailboxUser $user.fname $user.lname $user.name $user.alias $user.tempPassword $user.upnSuffix $user.resortSuffix $newUserOU


#Add additional attributes
AddAdditionalAttributesToUser $user.alias $user.description $user.office $user.officePhone $user.title $user.department $user.company $user.managerAlias $user.extAttr2 $user.extAttr3


# Clone security group membership if sourceUser provided
if (![string]::IsNullOrWhiteSpace($user.userToCloneSecurityGroups)) {
    CloneSecurityGroupMembership -sourceUser $user.userToCloneSecurityGroups -targetUser $user.alias
}


# Disconnect remote PowerShell Session
Get-PSSession | Remove-PSSession