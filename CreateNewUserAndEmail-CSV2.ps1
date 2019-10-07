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


# IMPORT HELPER FUNCTIONS
. .\CreateNewUserHelperFunctions.ps1


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
$user.alias = CheckAlias -signIn $true
#$user.alias = Read-Host "Alias [required]" # For testing only
$user.extAttr3 = GetExtAttr3
$user.tempPassword = Read-Host "Temporary Password [required]" -AsSecureString

# Strongly suggested details
Write-Host "### OPTIONAL BUT SUGGESTED DETAILS ###" -ForegroundColor DarkGray

$user.managerAlias = Get-RealADUser -userType "Manager"
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


CreateRemoteMailboxUser $user.fname $user.lname $user.name $user.alias $user.tempPassword $user.upnSuffix $user.resortSuffix $newUserOU -isFromCSV $false


#Add additional attributes
AddAdditionalAttributesToUser $user.alias $user.description $user.office $user.officePhone $user.title $user.department $user.company $user.managerAlias $user.extAttr2 $user.extAttr3


# Clone security group membership if sourceUser provided
if (![string]::IsNullOrWhiteSpace($user.userToCloneSecurityGroups)) {
    CloneSecurityGroupMembership -sourceUser $user.userToCloneSecurityGroups -targetUser $user.alias
}


# Disconnect remote PowerShell Session
Get-PSSession | Remove-PSSession