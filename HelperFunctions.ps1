﻿<#
.DESCRIPTION
    These are helper functions used by the create new user scripts.

.USAGE

  SET PARAMETERS
    Update the Resort Config settings just after this section.

  RUN SCRIPT
    Run the other scripts and ensure they are in the same directory (folder) as this script.
    The other scripts reference functions from here.


.NOTES
    
.TODO
    Cleaner fail if Exchange On-Prem username is incorrect

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
# vm-den-exch01.iDirectory.itw | msa-exch2013.mmsa.local | bbmr-exch2013.bbmr.local

# Domain Controller FQDN
$DomainController = "bbmrdc1-2012.bbmr.local"
# vm-den-dc01.iDirectory.itw | msa-ad1.mmsa.local | bbmrdc1-2012.bbmr.local

# UPN Suffix (without the @ sign)
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



# Write configuration information to user
Write-Host "
################################################################# `
    This script will attempt to create user objects with the following `
    - Exchange Server: $ExchangeServerName `
    - Domain Controller: $DomainController `
    - ExtensionAttribute2: $resortCode `
    - Resort Suffix: $resortSuffix `
    - Resort UPN Suffix: $upnSuffix `
    - Company in AD: $company `
    `
    If these values should be changed, please edit this script in the 'Resort Config' section.
    You can cancel the script by entering Ctrl + C.
    You can accept the default values by simply pressing Enter.
#################################################################
    " -ForegroundColor DarkGray


# BEGIN: Helper functions used by the script


# Checks availability in the tenant for the alias. Returns an available alias.
function CheckAliasAvailability
{
    param($alias)


    # Do...Until loop to test multiple aliases if the first one is already used
    do
    {
        $DisplayName =  ( Get-Recipient -RecipientTypeDetails GroupMailbox, SharedMailbox, UserMailbox, MailContact, MailUniversalDistributionGroup, DynamicDistributionGroup, MailUser, RoomMailbox -filter "alias -eq '$alias'" ).Name

        if ($DisplayName -ne $null)
        {
            $aliasIsAvailable = $false
            Write-Host "[$($user.alias)] Alias is taken by $DisplayName" -ForegroundColor Red
            $alias = Read-Host "[$($user.alias)] Enter a different alias"
        } else {
            Write-Host "[$alias] Alias is available!" -ForegroundColor Green
            $aliasIsAvailable = $true
        }
    } until ($aliasIsAvailable)


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
    param($firstName, $lastName, $name, $alias, $tempPassword, $newUserOU, $isFromCSV)

    # If from CSV,convert password to secure string
    if ($isFromCSV) {
        $tempPassword = ConvertTo-SecureString $tempPassword -AsPlainText -Force
    }

    if ([string]::IsNullOrWhiteSpace($name)) {
        $name = "$firstName $lastName $resortSuffix"
    }

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
    param($alias, $description, $office, $officePhone, $title, $department, $managerAlias, $extAttr3)


    # Add additional attributes in AD
    Write-Host "[$alias] Setting additional attributes" -ForegroundColor Green

    $command = "Set-ADUser -Credential `$UserCredential -Identity $alias -Server $DomainController -Replace @{
        extensionAttribute2=`"$resortCode`"
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
    
    Get-ADUser -Identity $sourceUser -Properties memberof |
        Select-Object -ExpandProperty memberof |
            Add-ADGroupMember -Credential $UserCredential -Server $DomainController -Members "$targetUser"

    Write-Host "[$targetUser] Security group membership cloned from $sourceUser" -ForegroundColor Green
}

# Validate Extension Attribute 3
function GetExtAttr3
{
    param($extAttr3 = "")
    
    do
    {
        # Get input for extAttr3 if not from CSV
        if ([string]::IsNullOrWhiteSpace($extAttr3)) {
            $extAttr3 = Read-Host "ExtensionAttribute3 for licensing"
        }

        if ($extAttr3 -eq "F1;" -or $extAttr3 -eq "E1;" -or $extAttr3 -eq "E3;") {
            $extAttr3IsValid = $true;
        } else {
            Write-Host "[$($user.alias)] ExtensionAttribute 3 is invalid. You provided $extAttr3. Please use 'F1;', 'E1;', or 'E3;'." -ForegroundColor Red
            $extAttr3 = Read-Host "[$($user.alias)] ExtensionAttribute3"
            $extAttr3IsValid = $false;
        }

    } until ($extAttr3IsValid)

    $extAttr3
}

function Get-RealADUser
{
    param($userType, $isFromCSV = $false, $samAccountName = "")

    # Get samAccountName if not provided from CSV
    if (!$isFromCSV) {
        $samAccountName = Read-Host "$($userType)'s alias"
    }

    do
    {
        try {
            if ([string]::IsNullOrWhiteSpace($samAccountName)) {
                $isValidUser = $true # Break out of loop if no entry specified
            } else {
                $ADUser = Get-ADUser -Identity $samAccountName
                Write-Host "[$($user.alias)] Using $($ADUser.Name) ($samAccountName) as $userType" -ForegroundColor DarkGray
                $isValidUser = $true
            }
        }
        catch {
            Write-Host "[$($user.alias)] Unable to find $userType with alias '$samAccountName'." -ForegroundColor Red
            $samAccountName = Read-Host "[$($user.alias)] Enter a different alias for $userType or type 'S' to skip"
            if ($samAccountName.ToLower() -eq "s") {
                $samAccountName = ""
                $isValidUser= $true
            } else {
                $isValidUser = $false
            }
        }
    } until ($isValidUser)

    $samAccountName
}


# END: helper functions