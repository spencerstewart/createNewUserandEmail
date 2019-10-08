<#
.DESCRIPTION
    This script is meant to create an individual new user. There is a separate script to import a CSV file of users.

.USAGE

  SET PARAMETERS
    Update the Resort Config settings in the helper functions file: "HelperFunctions.ps1".

  RUN SCRIPT
    To run the script, open Exchange Online Powershell Module. Change directories to where this script is saved, and execute the script by typing the script name in the prompt (use Tab complete if you like).
    The script will prompt you first for your Exchange Online admin credentials. This is used to check the tenant for whether the desired alias is available or not.
    Next, the script will prompt you for your Exchange On-Prem admin credentials. Use your full login (eg jdoe@idirectory.itw). This is used to create the remote mailbox and set attributes in Active Directory.

    If the script completed with no errors, the new users are placed in the OU designated in the resort config settings. Don't forget to move these new users to their appropriate OU!


.NOTES
    This script relies on a separate file with helper functions, called "HelperFunctions.ps1". It should be in the same directory as this script.

.TODO

.AUTHOR
    Spencer Stewart, Big Bear Mountain Resort

.DATE
    Created: 2019-10-07
    Last modified: 2019-10-07
#>


# IMPORT HELPER FUNCTIONS
. .\HelperFunctions.ps1


# BEGIN: executing the script


# Assemble new user object and set defaults
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

### GET REQUIRED DETAILS ###
Write-Host "### REQUIRED DETAILS ###" -ForegroundColor DarkGray

$user.fname = Read-Host "(REQUIRED) First Name"
$user.lname = Read-Host "(REQUIRED) Last Name"

# Display Name with suggested default
$user.name = "$($user.fname) $($user.lname) $resortSuffix"
$prompt = Read-Host "(REQUIRED) Display Name [$($user.name)]"
$user.name = ($user.name,$prompt)[[bool]$prompt]

# Alias. Connects to Exchange Online to verify availability.
$user.alias = Read-Host "(REQUIRED) Alias"
Write-Host "### Connecting to Exchange Online to verify alias availability ###" -ForegroundColor DarkGray
Connect-EXOPSSession -WarningAction SilentlyContinue
$user.alias = CheckAlias -alias $user.alias
Get-PSSession | Remove-PSSession

$user.extAttr3 = GetExtAttr3
$user.tempPassword = Read-Host "(REQUIRED) Temporary Password" -AsSecureString


### GET SUGGESTED DETAILS ###
Write-Host "### OPTIONAL BUT SUGGESTED DETAILS ###" -ForegroundColor DarkGray

$user.managerAlias = Get-RealADUser -userType "Manager"
$user.userToCloneSecurityGroups = Get-RealADUser "Source user to copy security groups"
$user.title = Read-Host "Title [optional]"
$user.department = Read-Host "Department [optional]"
$user.description = Read-Host "Description [optional]"
$user.office = Read-Host "Office [optional]"
$user.officePhone = Read-Host "Office Phone [optional]"


### DO THE WORK ###
# Create remote mailbox user
Write-Host "### CONNECTING TO ON PREM EXCHANGE ###" -ForegroundColor DarkGray
$UserCredential = $host.ui.PromptForCredential("Exchange On-Prem Credentials", "Please enter your Exchange On-Prem Creds.", "", "NetBiosUserName")
ConnectToOnPremExchange $ExchangeServerName
CreateRemoteMailboxUser $user.fname $user.lname $user.name $user.alias $user.tempPassword $user.upnSuffix $user.resortSuffix $newUserOU -isFromCSV $false


# Add additional attributes
AddAdditionalAttributesToUser $user.alias $user.description $user.office $user.officePhone $user.title $user.department $user.company $user.managerAlias $user.extAttr2 $user.extAttr3


# Clone security group membership if sourceUser provided
if (![string]::IsNullOrWhiteSpace($user.userToCloneSecurityGroups)) {
    CloneSecurityGroupMembership -sourceUser $user.userToCloneSecurityGroups -targetUser $user.alias
}


# Disconnect remote PowerShell Session
Get-PSSession | Remove-PSSession