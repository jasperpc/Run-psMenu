<#
.NAME
 Set-ADUserPwd.ps1
.SYNOPSIS
 This will prompt you for a username for whom you wish to change the password options
.DESCRIPTION
 2019-06-19
 Written by Jason Crockett - Laclede Electric Cooperative
 This will prompt you for a username for whom you wish to change the password options
.PARAMETERS
 N/A
.EXAMPLE
 .\Set-ADUserPwd.ps1 []
.SYNTAX
 .\Set-ADUserPwd.ps1 []
.REMARKS
 To see the examples, type: help Set-ADUserPwd.ps1 -examples
 To see more information, type: help Set-ADUserPwd.ps1 -detailed
.TODO
#>
# clear-host
$global:AnotherPW = "No"
Write-Output = "This will prompt you for a username for whom you wish to change the password options"
$PromptADUser =Read-Host "Enter samaccountname here"
$ADUFilter = "$PromptADUser"
$OK2SetPW=$null
$OK2ChgPW=$null
$OK2ChgPW = Read-Host "`nType yes if you wish to force this user to change the password : $ADUFilter "
$SearchBase = Read-Host "Type the Search Base in the following format: CN=OUisOptional,DC=subdomain,DC=domain"
$PWRecommendedAge = Read-Host "What is the preferred Password age? default = 999"

function PromptChgPwd ()
{
    # The following line gets the information for the user
    Get-ADUser -Filter 'Samaccountname -like $ADUFilter'-SearchBase $SearchBase -Properties *| select SamAccountName,mail,CannotChangePassword,passwordlastset,PasswordExpired,LockedOut,Enabled
    
    # The following line gets the information and also prompts you to expire the password and prompt for a change of password at next logon
    # ** Important ** leave the -confirm option on incase the filter includes more that just the username(s) you wish to modify
    Get-ADuser -filter 'SamAccountname -like $ADUFilter'-SearchBase $SearchBase -properties * |set-aduser -PasswordNeverExpires $false -ChangePasswordAtLogon $true -Confirm
    
    # Allow time for processing of the previous command
    Read-Host "Wait a few seconds and then press Enter to continue."
    
    # Verify that the changes were made
    Get-ADUser -Filter 'Samaccountname -like "$ADUFilter"'-SearchBase $SearchBase -Properties *| select SamAccountName,mail,CannotChangePassword,passwordlastset,PasswordExpired,LockedOut
}

function Set-Pwd()
{
    $newPassword = Read-Host -Prompt "`nEnter new Password" -AsSecureString
    Set-ADAccountPassword -Identity $ADUFilter -NewPassword $newPassword -reset -confirm
}

if ($OK2ChgPW -eq "yes")
    {
        PromptChgPwd
    }
Else
    {
        $OK2SetPW = Read-Host "`nType yes if YOU wish to create a new password NOW for $ADUFilter?"    
        if ($OK2SetPW -eq "yes")
        {
            Set-Pwd
        }
    
    }

$list = Read-Host "`nIf you want to see the list of changed and unchanged users / passwords, type yes"
if ($list -eq "yes")
{
    $Completed = get-aduser -filter '(mail -notlike "healthmailbox*")' -SearchBase $SearchBase  -Properties * |where {$_.enabled -eq $true} |select Samaccountname,passwordlastset,cannotchangepassword,passwordexpired,lockedout,mail|sort passwordlastset |ft -AutoSize
    $completed
    $CCount = $Completed.count
    "Completed: $CCount"
    $PWRecommendedAge = Read-Host "What is the preferred Password age? Default: -999"
    $GD1 = (get-date -date ($Global:CDateTime.adddays($PWRecommendedAge)) -uformat "%m/%d/%y")
    $Waiting = get-aduser -filter '(PasswordLastSet -lt $GD1) -AND (mail -notlike "healthmailbox*")' -SearchBase $SearchBase  -Properties * |where {$_.enabled -eq $true} |select Samaccountname,passwordlastset,cannotchangepassword,passwordexpired,lockedout,mail|sort passwordlastset,SamAccountName |ft -AutoSize
    $Waiting
    
    (get-date -date ($Global:CDateTime.adddays(-365)) -uformat "%m/%d/%y")
    
    $WCount = $Waiting.count
    "Waiting: $WCount"
}

$global:AnotherPW = Read-host "`nIf you wish to change another password now type Yes "
Clear-Host
$global:AnotherPW
