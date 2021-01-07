<#
.NAME
    Manage-ADUsers.ps1
.SYNOPSIS
 Automate AD User creation
.DESCRIPTION
 (c) 2019
 Written by Jason Crockett - Laclede Electric Cooperative
.PARAMETERS
.EXAMPLE
    .\Manage-ADUsers.ps1 []
.SYNTAX
    .\Manage-ADUsers.ps1 []
.REMARKS
    To see the examples, type: help Manage-ADUsers.ps1 -examples
    To see more information, type: type: help Manage-ADUsers.ps1 -detailed
#>
# Enable AD and PS Update modules
Import-Module ActiveDirectory
# set script variables
$Script:PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
Write-Warning $Script:PSScriptRoot
$Global:CDate = Get-Date -UFormat "%Y%m%d"
$Global:CDateTime = [datetime]::ParseExact($Global:CDate,'yyyyMMdd',$null)
$LT_PWLastSetDate = $Null
$GT_PWLastSetDate = $Null
$Global:TomorrowDateTime = ($Global:CDateTime).AddDays(1)
#if ($Global:UserList -eq $null){
Clear-Host
. .\Get-UsersList.ps1 # load functions in .ps1
#}
Clear-Host

# Prepare for menu
[BOOLEAN]$Global:ExitSession=$false
$Script:t4 = "`t`t`t`t"
$Script:d4 = "-------------------------------------------------------------------------------"
$Script:NC = $null
$Script:p1 = $null
$Script:p2 = $null
$Script:p3 = $null
function chcolor($p1,$p2,$p3,$NC){
            
            Write-host $t4 -NoNewline
            Write-host "$p1 " -NoNewline
            Write-host $p2 -ForegroundColor $NC -NoNewline
            Write-host "$p3" -ForegroundColor White
            $script:NC = $null
            $script:p1 = $null
            $script:p2 = $null
            $script:p3 = $null
}
$MADUmenuArray = @($Exit," $MMADUNum. ","Exit to Main Menu"," ","Red"), #@(SwitchVar,p1,p2,p3,nc)
                @($GCred," $MMADUNum.  Enter ","Credentials ","for use in script.","Cyan"),
                @($FWStatus," $MMADUNum.  Add/Create ","New User ","wizard","yellow"),
                @($ListChangePW," $MMADUNum.  List PW info or force a User to ","change password ","on his/her AD account","yellow"),
                @($adUserPW," $MMADUNum. `tUnlock, Reset, or Change","the password ","on a user's PC","yellow"),
                @($Disable_ADUser," $MMADUNum. `tMove user(s) to the disabled OU and ","Disable User(s) "," ","yellow")
    
function MADUmenu()
{
while ($MADUmenuselect -lt 1 -or $MADUmenuselect -gt 7)
{
    Trap {"Error: $_"; Break;}        
        $MMADUNum = 0;Clear-host |out-null
        Start-Sleep -Milliseconds 100
        Write-host $t4 $d4 -ForegroundColor Green

        # Exit back to previous menu
        $MMADUNum ++;$Exit=$MMADUNum;$p1 = " $MMADUNum. ";$p2 = "Exit to Previous Menu";
            $p3 = " ";$NC = "Red";chcolor $p1 $p2 $p3 $NC
        # Enter Credentials
        $MMADUNum ++;$GCred=$MMADUNum;$p1 =" $MMADUNum.  Enter ";
            $p2 = "Credentials ";$p3 = "for use in script."; $NC = "Cyan";chcolor $p1 $p2 $p3 $NC
        # AD Users List
        $MMADUNum ++;$Get_ADUserList = $MMADUNum;$p1 =" $MMADUNum.  AD";$p2 = "User List ";
		    $p3 = " (SamAccountName,Name,CanonicalName,LastLogonDate,Created,mail,address)"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Either list User/Password information or force a user to change the account's password
        $MMADUNum ++;$ListChangePW=$MMADUNum;$p1 =" $MMADUNum.  List PW info or force a User to ";
            $p2 = "change password ";$p3 = "on his/her AD account"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Wizard to Add/Create a new user
        $MMADUNum ++;$FWStatus=$MMADUNum;$p1 =" $MMADUNum.  Add/Create ";
            $p2 = "New User ";$p3 = "wizard"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Unlock, Reset, or Change a user's password
        $MMADUNum ++;$adUserPW = $MMADUNum;$p1 =" $MMADUNum. `tUnlock, Reset, or Change";
		    $p2 = "the password ";$p3 = "on a user's PC"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
# Disable selected ADUser(s) and move objects to a new AD container (default is the Disabled container)
                $MMADUNum ++;$Disable_ADUser = $UnGMNum;$p1 =" $MMADUNum.`tMove and ";
		    $p2 = "Disable ADUser ";$p3 = "(with -Confirm)"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        Write-Host "$t4$d4"  -ForegroundColor Green
        [int]$MADUmenuselect = Read-Host "`n`tPlease select the menu item you would like to do now"
        # Write-Output `n "You picked item $MADUmenuselect"
}
switch($MADUmenuselect)
    {
        $Exit{}
        $GCred{Clear-Host;Get-Cred;$MADUmenuselect = $null;Reload-PromptMADUmenu}
        $Get_ADUserList{$MADUmenuselect=$null;get-aduserlist;Reload-PromptMADUmenu}
        $ListChangePW{Force-NewADPassword;Reload-PromptMADUmenu}
        $FWStatus{Create-NewADUser;Reload-PromptMADUmenu}
        $adUserPW{Reset-ADUserPW;Reload-PromptMADUmenu}
        $Disable_ADUser{Disable-ADUser;Reload-PromptMADUmenu}
        default
        {
        cls
        $Global:ExitSession=$true;break;clear-host
        }
    }    
}
function Reload-PromptMADUmenu()
{
        Read-Host "Hit Enter to continue"
        $MADUmenuselect = $null
        MADUmenu
}

function Identify-UserName
{
    # This if statement could potentially be moved to inside the Get-UsersList function, but it might need to be modified to do so
    if ($UserLine.name -eq $null)
    {
        $Global:User = $UserLine
    }
    else
    {
        $Global:User = $UserLine.SamAccountName
    }
}
function Create-NewADUser()
{
# Collect the following user information to create a new ADUser
    $FullName = $null
    $FullName = Read-Host "Enter the Full Name"
    $Names = ($FullName -split " ") # |%{$_[0] }) -join ""
    $NamesCount = ($Names.Count)-1
    $Given = $Names[0]
    If ($NamesCount -gt 0)
    {
        $Surname = $Names[$NamesCount]
    }
    else 
    {
        $Surname = $null
        Write-warning "Surname is null!!"
        $NoSNPrompt = Read-Host "Type Y to continue without a surname"
        If ($NoSNPrompt -ne "Y"){break}
    }

    $Given
    $Surname

# The first name must start with a letter
    If (($Given -match '^([a-zA-Z])') -eq $true)
    {
        $GI = $Given[0]
        $SAM = $GI+$Surname
        $UseName = $Given+" "+$Surname
    }
# Create the User Principal Name (UPN)
    $EmailDomain = Read-Host "Please enter the email domain name (something like gmail.com)"
    $Email = $null
    $Email = $SAM+"@$EmailDomain"

# Use the following user as a template
    $TemplateSAM = $null
    $TemplateSAM = Read-Host "Enter the SAMAccountName of the account to use as a template"
    $InstanceUser = $null
    $InstanceUser = Get-ADUser -Identity $TemplateSAM |select SamAccountName
    $InstanceUser = $InstanceUser.SamAccountName
    $ASS = "Account Password"
    Write-Output "Creating a new user $SAM ($Given $Surname - $Email) based on user $InstanceUser"
    <#
    New-ADUser -Instance $InstanceUser -Name "$Given $Surname" -GivenName $Given -Surname $Surname `
    -SamAccountName $SAM -AccountPassword (Read-Host -AsSecureString "Account Password") `
    -UserPrincipalName $Email `
    -PassThru -WhatIf | Enable-ADAccount -WhatIf
    #>

    $Script:NA1 = $null;$Script:NA1 = "-Instance $InstanceUser -Name $UseName -GivenName $Given -Surname $Surname "
    $Script:NA2 = $null;$Script:NA2 = "-SamAccountName $SAM -AccountPassword (Read-Host -AsSecureString $ASS) "
    $Script:NA3 = $null;$Script:NA3 = "-PassThru -WhatIf "
    $Script:NA4 = $null;$Script:NA4 = "| Enable-ADAccount -WhatIf"
    $NAUCMD = $Script:NA1+$Script:NA2+$Script:NA3
    $NAEnable = $Script:NA4
    $NAUCMD
    New-ADUser (($NAUCMD).ToString()) -Confirm
}
# Force a user to change his/her AD Password
function Force-NewADPassword()
{
    $ObjectRslt=$null
    $GAD=$null
    Get-UserList
    Write-Warning "`nIf changes are enabled, this function will allow password to expire and force a password change at logon.`n"
    $Action = Read-Host "Type L to List, P to Preview changes, or S to set changes" # "List or Set"
    # Determine Less than portion of the between dates query
    $TDT = ($Global:TomorrowDateTime).ToString("yyyy/MM/dd")
    $TDTLessAge = $Global:TomorrowDateTime.AddDays(-999).ToString("yyyy/MM/dd")
    if (($LT_PWLastSetDate = Read-Host "Enter date for less than PW last set date in the format yyyy/MM/dd [Enter] to use current default: $TDT`n`t - OR - 999 days ago as: $TDTLessAge") -eq "") 
        {$LT_PWLastSetDateTime = $Global:TomorrowDateTime}
        Else
        {$LT_PWLastSetDateTime = [datetime]::ParseExact($LT_PWLastSetDate,'yyyy/MM/dd',$null)}
    # Determine Greater than portion of the between dates query
    if (($GT_PWLastSetDate = Read-Host "Enter date for greater than PW last set date in the format yyyy/MM/dd [Enter] to use current default: 2001/07/29") -eq "") {$GT_PWLastSetDate = "2001/07/29"}
    $GT_PWLastSetDateTime = [datetime]::ParseExact($GT_PWLastSetDate,'yyyy/MM/dd',$null)
    Write-output "Searching for PasswordLastSet greater than $GT_PWLastSetDate but less than $LT_PWLastSetDate"
    $GLULC = $Global:UserList.Count
    "Searching for $GLULC users"
    $ObjectRslt = foreach ($UserLine in $Global:UserList)
    {
        Identify-UserName
        $U = $Global:User
        $GAD = get-aduser -filter "samaccountname -like `"$U`"" -Properties *
        # Report on all users regardless of current password expiration settings
        if ($GAD)
        {
            # Only make changes under the following conditions
            if ($Action -eq "S" -or $Action -eq "P")
            {
                $GADs = $GAD |where {($_.PasswordLastSet -lt $LT_PWLastSetDateTime -and $_.PasswordLastSet -gt $GT_PWLastSetDateTime )}
                if ($GADs)
                {
                    if ($Action -eq "S")
                    {
                        #$GADs|set-aduser -PasswordNeverExpires:$false #-Confirm
                        $GADs|set-aduser -ChangePasswordAtLogon:$true -PasswordNeverExpires:$false -Confirm
                    }
                    $ActStat = "Marking for Change"
                }
                else
                {
                    $ActStat = ""
                }
                $GADs =$null
            }
        $OBProperties = @{'SamAccountName'=$GAD.SamAccountName;
                    'Enabled'=$GAD.Enabled;
                    'CannotChangePassword'=$GAD.CannotChangePassword;
                    'LockedOut'=$GAD.LockedOut;
                    'PasswordExpired'=$GAD.PasswordExpired;
                    'passwordlastset'=$GAD.passwordlastset;
                    'Name'=$GAD.Name;
                    'DistinguishedName'=$GAD.DistinguishedName;
                    'PasswordNeverExpires'=$GAD.PasswordNeverExpires;
                    'ActStat'=$ActStat;
                    'LastLogonDate'=$GAD.LastLogonDate}
            New-Object -TypeName PSObject -Prop $OBProperties
        }
        $GAD=$null
    }
    # Detail Password expiry information about all requested users found 
    $ObjectRslt|sort PasswordNeverExpires,passwordlastset,SamAccountName,LastLogonDate|ft SamAccountName,CannotChangePassword,PasswordExpired,PasswordNeverExpires,passwordlastset,LastLogonDate,ActStat,LockedOut,Enabled
    $ORC = $ObjectRslt.Count
    "Found $ORC users"
}
function Reset-ADUserPW()
{
    Clear-Host
    Write-host "Note that the Change option will provide a list of users and dates their passwords were last set (with or without options)"
    $Choice = Read-Host "Enter List, Unlock, Reset, Change`nReset and Change accomplish the same thing slightly differently "
    
    if ($Choice -eq "Unlock")
    {
        $ADUser = Read-Host "Enter part of the user's name "
        Get-ADUser -Filter ('name -like "*' + $ADUser + '*"') |Unlock-ADAccount -Confirm
    }
    Elseif ($Choice -eq "Reset")
    {
        Invoke-Expression ".\Set-ADUserPwd.ps1"
        while($global:AnotherPW -eq "Yes")
            {
                Invoke-Expression ".\Set-ADUserPwd.ps1"
                Clear-host
            }
    }    
# Need to verify this option
# This is like Set-Pwd function in Set-ADUserPwd.ps1 which can be iterated to in the change choice, but this is much quicker
    Elseif ($Choice -eq "Change")
    {
        $ADUser = Read-Host "Enter part of the user's name "
        # $ADUserResult = Get-ADUser -Filter ('name -like "*' + $ADUser + '*"')
        $ADUserResult = Get-ADUser -filter 'samaccountname -eq $ADUSer'
        $ADUserSAM = $ADUserResult.SamAccountName
        $NewCred = Read-Host "Enter a new password?" -AsSecureString
        $ConfirmNewCred = Read-Host "Re-enter the new password." -AsSecureString
        if ($NewCred -eq $ConfirmNewCred)
        {
            $ADUserSAM |Set-ADAccountPassword -NewPassword $newCred -confirm -Reset
            <#
              You must set the OldPassword and the NewPassword parameters to set the password unless you specify the Reset parameter. 
              When you specify the Reset parameter, the password is set to the NewPassword value that you provide and the OldPassword 
              parameter is not required.
            #>
        }
        Else
        {
            Write-Warning "The two passwords do not match"
        }
    }
    Elseif ($Choice -eq "List")
    {
        $SearchBase = "DC=lec,DC=lan" # Previously included "CN=Users," in the SearchBase
        $Completed = get-aduser -filter <#(PasswordLastSet -gt "5/16/2016") -AND #> '(mail -notlike "healthmailbox*")' -SearchBase $SearchBase  -Properties * |where {$_.enabled -eq $true} |select Samaccountname,passwordlastset,cannotchangepassword,passwordexpired,lockedout,mail|sort passwordlastset |ft -AutoSize #|Out-GridView
        $completed
        $cc = $Completed.count
        Write-Warning "Completed: $cc"
        Start-Sleep -Seconds 2
        $Waiting =  get-aduser -filter '(PasswordLastSet -lt "5/15/2016") -AND (mail -notlike "healthmailbox*")' -SearchBase $SearchBase  -Properties * |where {$_.enabled -eq $true} |select Samaccountname,passwordlastset,cannotchangepassword,passwordexpired,lockedout,mail|sort passwordlastset |ft -AutoSize # |Out-GridView
        $Waiting 
        $wc = $Waiting.count
        Write-Warning "Waiting: $wc"
    }
}
function List-ADUserSID
{

gwmi Win32_UserAccount |ft
}
function Disable-ADUser
{
# Disable user and move to the Disabled container
$Move2TargetPath = "OU=Disabled,DC=lec,DC=lan" # Create this as a default for a parameter variable
$SearchText = Read-Host "Enter search criteria for all or part of the samaccountname"
$ADUsersList = (get-aduser -filter 'samaccountname -like $SearchText' -Properties *|select SamAccountName,Enabled,Name,DistinguishedName,emailaddress,Description)
foreach ($ADUserTarget in $ADUsersList) 
    {
        $SAN = $ADUserTarget.SamAccountName
        $ADUserTarget
        $ADUserDescription = $ADUserTarget.Description
        # Get-adobject -Identity $ADUserTarget.distinguishedname
        Write-Warning "Moving $SAN to $Move2TargetPath using Move-ADObject"
        Move-ADObject -Identity $ADUserTarget.distinguishedname -TargetPath $Move2TargetPath -Confirm
        Write-Warning "Disabling $SAN"
        Set-ADUser -Identity $SAN -Enabled $false -Description "D_$ADUserDescription" -Confirm
    }
}
Function get-aduserlist()
{
    Clear-Host
    $SearchText = $null
    # $ADUserList = get-aduser -Filter * |Select-Object SamAccountName, Name # |Export-Csv 20151015ADUsers.csv
    if(($SearchText = Read-Host "Enter search criteria for all or part of the samaccountname using * as wildcard [ENTER] for all") -eq ""){$SearchText = "*"}
    $Fltr = "samaccountname -like `"$SearchText`""
    $ADUL = get-aduser -Filter $Fltr -Properties *|where {$_.enabled -eq "true"} |Select-Object SamAccountName, Name,CanonicalName,LastLogonDate,Created,mail,address,UserPrincipalName
    Start-Sleep 1
    $ADUL|Select SamAccountName,Name,CanonicalName,LastLogonDate,Created,mail,address,UserPrincipalName |Out-GridView 
    # $ADUL|ft SamAccountName,Name,CanonicalName,LastLogonDate,Created,mail,address -AutoSize -Wrap
}
MADUmenu