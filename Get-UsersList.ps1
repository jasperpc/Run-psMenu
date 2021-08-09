##  ***** This script is not ready yet!!! It is being adopted from the Get-SystemsList.ps1 script to address users not computers
<#
.NAME
    Get-UsersList.ps1
.SYNOPSIS
 Return a list of systems to the script that calls it
.DESCRIPTION
 2019-06-14
 Written by Jason Crockett - Laclede Electric Cooperative
 Function Get-UserList returns a list of systems to the script that calls it
 Function Identify-UserName takes the list of pcnames and makes sure it can be 
 handled correctly
.PARAMETERS
 N/A
.EXAMPLE
 .\Get-UsersList.ps1 []
.SYNTAX
 .\Get-UsersList.ps1 []
.REMARKS
 To see the examples, type: help Get-UsersList.ps1 -examples
 To see more information, type: help Get-UsersList.ps1 -detailed
.TODO
#>
function chcolor($p1,$p2,$p3,$NC){
            Write-host $t4 -NoNewline
            Write-host "$p1 " -NoNewline
            Write-host "$p2 " -ForegroundColor $NC -NoNewline
            Write-host "$p3" -ForegroundColor White
            $script:NC = $null
            $script:p1 = $null
            $script:p2 = $null
            $script:p3 = $null
}
function Get-UserList()
{$ADorCSL = $null
    $p1 = "Type AD to use enabled Users listed in Active Directory or";
    $p2 = " C ";
    $p3 = "to type your own comma separated list";
    $NC = "Yellow";
    chcolor $p1 $p2 $p3 $NC;$ADorCSL = Read-Host "`t`t"
<# 
    # Option to add a userlist from a file as a source
    $userlist use a file as an optional data source
    $userlist = Get-Content .\pl.txt # Get CSV list from this file listed in the current directory
    $UserList 
    $UserList .GetType()
#>
<#
            # Run in a "foreach ($UserLine in $UserList ){}" statement
            # This if statement could potentially be moved to inside the Get-UserList function, but it might need to be modified to do so
            if ($UserLine.name -eq $null)
            {
                $User = $UserLine
            }
            else
            {
                $User = $UserLine.name
            }
#>
if ($ADorCSL -eq "AD")
    {
        # Get the list from users listed in Active Directory
        # $Global:userlist = Get-ADUser -Filter 'Enabled -eq "True"'
        # Filter out AD for Enabled systems in the Domain users and Domain Controllers OUs
        $Global:userlist = get-aduser -Filter '(Enabled -eq "True")' -Properties *|where {$_.mail -notlike "healthmailbox*"} # |select SamAccountName,Enabled,Name,DistinguishedName,UserPrincipalName
    }
elseif ($ADorCSL -eq "C")
    {
        # Get the list from users listed in comma-separated list typed by the user
        $Global:UserListPrompt = Read-Host "Type in a comma separated list of users (using SamAccountName)"
        [array]$Global:Userlist = $UserListPrompt.split(",") |%{$_.trim()}
        # need to standardize on how this list is presented throughout the script then remove this line: Use just $userlist throughout the script.
        $users = $Global:UsersList
    }
Else
    {
    $opt1 = $null
        $p1 = "That doesn't match the options type ";
        $p2 = "yes ";
        $p3 = "to run against the current user only";
        $NC = "Yellow";chcolor $p1 $p2 $p3 $NC;$opt1 = Read-Host "`t`t"
    If ($opt1 -eq "yes")
        {
        # Use the local username as the list of users 
        $pscn = $env:COMPUTERNAME #$proc.pscomputername
        # Get the username for the currently logged-in user
            # $proc = gwmi win32_process -Filter "Name = 'explorer.exe'"
            # $proc.GetOwner().user
        $pgo = (gwmi win32_process -Filter "Name = 'explorer.exe'" -ComputerName $pscn).GetOwner().user
        $Global:userlist = $pgo
        # need to standardize on how this list is presented throughout the script then remove this line: Use just $userlist throughout the script.
        $users = $Global:userlist
        }
    Else
        {
                $ErrorActionPreference = "continue"
        }
    }
}

function Identify-UserName
{
    # This if statement could potentially be moved to inside the Get-UserList function, but it might need to be modified to do so
    if ($Global:UserLine.name -eq $null)
    {
        $Global:User = $Global:UserLine
    }
    else
    {
        $Global:User = $Global:UserLine.name
    }
}