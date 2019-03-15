# clear-host
$global:AnotherPW = "No"
Write-Output = "This will prompt you for a username for whom you wish to change the password options"
$PromptADUser =Read-Host "Enter samaccountname here"
$ADUFilter = "$PromptADUser"
$OK2SetPW=$null
$OK2ChgPW=$null
$OK2ChgPW = Read-Host "`nType yes if you wish to force user:$ADUFilter to create a new password now."
$SearchBase = "CN=Users,DC=lec,DC=lan"

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
        $OK2SetPW = Read-Host "`nType yes if you wish to create a new password now for $ADUFilter?"    
        if ($OK2SetPW -eq "yes")
        {
            Set-Pwd
        }
    
    }

$list = Read-Host "`nIf you want to see the list of changed and unchanged users / passwords, type yes"
if ($list -eq "yes")
{
    # $Completed = get-aduser -filter <#(PasswordLastSet -gt "5/16/2016") -AND #> '(mail -notlike "healthmailbox*")' -SearchBase $SearchBase  -Properties * |where {$_.enabled -eq $true} |select Samaccountname,passwordlastset,cannotchangepassword,passwordexpired,lockedout,mail|sort passwordlastset |ft -AutoSize
    $Completed = get-aduser -filter <#(PasswordLastSet -gt "5/16/2016") -AND #> '(mail -notlike "healthmailbox*")' -SearchBase $SearchBase  -Properties * |where {$_.enabled -eq $true} |select Samaccountname,passwordlastset,cannotchangepassword,passwordexpired,lockedout,mail|sort passwordlastset |ft -AutoSize
    $completed
    "Completed: $Completed.count"
    $Waiting =  get-aduser -filter '(PasswordLastSet -lt "5/15/2016") -AND (mail -notlike "healthmailbox*")' -SearchBase $SearchBase  -Properties * |where {$_.enabled -eq $true} |select Samaccountname,passwordlastset,cannotchangepassword,passwordexpired,lockedout,mail|sort passwordlastset |ft -AutoSize
    $Waiting
    "Waiting: $Waiting.count"
}

$global:AnotherPW = Read-host "`nIf you wish to change another password now type Yes "
Clear-Host
$global:AnotherPW
