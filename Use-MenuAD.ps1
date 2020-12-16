<#
.NAME
    Run-ADMenuFunctions.ps1
.SYNOPSIS
    Run Active Directory / Network Administrative tasks
.DESCRIPTION
    Called from Run-psMenu to load functions for Active Directory / Network Administrative tasks
 2018-06-04

 Written by Jason Crockett - Laclede Electric Cooperative

 .PARAMETERS
    This parameter doesn't exist (Comps = list of computernames)
.EXAMPLE
    "a command with a parameter would look like: (.\Get-DiskInfo.ps1 -comps dc)"
    .\Use-ADMenuFunctions.ps1 []
.SYNTAX
    .\Use-ADMenuFunctions.ps1 []
.REMARKS
    
.TODo
    
#>

<#
[CmdletBinding()] #(allows the use of parameter attributes)
param(
        #[Parameter(Mandatory=$True)]
            #[string[]]$comps = "$env:COMPUTERNAME",
)
#>

# Enable AD modules
Import-Module ActiveDirectory

# Enable Exchange cmdlets (On an Exchange Server)
# add-pssnapin *exchange* -erroraction SilentlyContinue 

Clear-Host
# set script variables
# . .\Get-SystemsList.ps1
[BOOLEAN]$Global:ExitSession=$false
# Get Credentials for later wmi requests
# $cred = Get-Credential "$env:USERDOMAIN\Administrator"

$Script:t4 = "`t`t`t`t"
$Script:ADd4 = "------------------------------Use-ADMenuFunctions-----------------------------------"
$Script:NC = $null
$Script:p1 = $null
$Script:p2 = $null
$Script:p3 = $null
$CDate = Get-Date -UFormat "%Y%m%d"
[int]$admenuselect = 0

Function chcolor($p1,$p2,$p3,$NC){
            
            Write-host $t4 -NoNewline
            Write-host "$p1 " -NoNewline
            Write-host $p2 -ForegroundColor $NC -NoNewline
            Write-host "$p3" -ForegroundColor White
            $script:NC = $null
            $script:p1 = $null
            $script:p2 = $null
            $script:p3 = $null
}
# The following is an example of combining -Property and -ExpandProperty
# (Get-ADUser -filter 'SamAccountName -like $ADUser1' -Properties * |Select-Object -property @{n="ADUserInfo";e={$_.SamAccountName,$_.memberof}}|select -ExpandProperty ADUserInfo)
Function ADmenu()
{
    while ($admenuselect -lt 1 -or $admenuselect -gt 12)
    {
        Trap {"Error: $_"; Break;}        
        $MADNum = 0;Clear-host |out-null
        Start-Sleep -m 50
        Write-host $t4 $Script:ADd4 -ForegroundColor Green
# Exit to Main Menu
        $MADNum ++;$adExit = $MADNum; 
		    $p1 =" $MADNum. `tSelect to ";
		    $p2 = "Exit ";
		    $p3 = "to Main Menu"; $NC = "Red";chcolor $p1 $p2 $p3 $NC
# Prompt for credential with permissions across the network
        $MADNum ++;$GCred=$MADNum;$p1 =" $MADNum.  `t Enter";
            $p2 = "Credentials ";$p3 = "for use in script."; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
# New AD Objects List
        $MADNum ++;$Get_NewADObjects = $MADNum;
		    $p1 =" $MADNum. `tFilter ";
		    $p2 = "New AD Objects ";
		    $p3 = "by date created."; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
# Test LDAP connections to the AD DCs
        $MADNum ++;$Run_Test_LDAP = $MADNum;
		    $p1 =" $MADNum. `tTest";
		    $p2 = "LDAP Connections ";
		    $p3 = "to the Active Directory Domain Controllers "; $NC = "green";chcolor $p1 $p2 $p3 $NC
# Check for Primary subnet of AD Site
        $MADNum ++;$Get_ADSitSubnet=$MADNum;
            $p1 = " $MADNum. View ";
            $p2 = "AD Site and subnet";
            $p3 = " for the current Active Directory context"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# Group Policy Results for Domain PC/User
        $MADNum ++;$Get_GpReport = $MADNum;
            $p1 =" $MADNum. `tShow ";
		    $p2 = "Group Policy Results ";
		    $p3 = "for Domain PC/User"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
# List the LAPS Passwords for all domain member PCs that have them
        $MADNum ++;$List_DomPCsLAPSPw = $MADNum;
		    $p1 =" $MADNum. `tList the";
		    $p2 = "LAPS Password for ";
		    $p3 = "member PCs in this domain (Requires Credentials)"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
# Unlock, Reset, or Change a user's password
        $MADNum ++;$Reset_ADUserPW = $MADNum;
            $p1 =" $MADNum. `tUnlock, Reset, or Change";
		    $p2 = "the password ";
		    $p3 = "on a user's PC"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
# Delete Old AD Computer objects from AD
        $MADNum ++;$Run_ADCleanup = $MADNum;
            $p1 =" $MADNum. `tBulk Remove by age";
		    $p2 = "Old AD Computer Objects ";
		    $p3 = "from Active Directory (with confirm/suspend)"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
# Decommission Old AD Computer object(s) from AD
        $MADNum ++;$Decommission_ADPC = $MADNum;
            $p1 =" $MADNum. `tBulk ";
		    $p2 = "Decommission,Disable AD Computer - Requires RunAs Admin ";
		    $p3 = "Objects in AD and prepend description with _D (with confirm/suspend)"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
# TestFix-PCSecureChannel (trust relationship) 
        $MADNum ++;$TestFix_PCSecureChannel = $MADNum;
            $p1 =" $MADNum. `tTest or ";
		    $p2 = "Fix the secure channel (trust relationship) ";
		    $p3 = "between computer and domain (with confirm/suspend)"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
# Search an Exchange mailbox `n$t4`t (this will only work on Exchange and is included for reference only here.)
        $MADNum ++;$Search_ExchMbox = $MADNum;
		    $p1 =" $MADNum. `tSearch";
		    $p2 = "Exchange Mailbox ";
		    $p3 = "(reference only)"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        Write-Host "$t4$Script:ADd4"  -ForegroundColor Green
        # Setting up the menu in this way allows you to move menu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $MADNum ++;$p1 =" $MADNum.  FIRST";$p2 = "HIGHLIGHTED ";$p3 = "END"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        [int]$admenuselect = Read-Host "`n`tPlease select the menu item you would like to do now"
    }
    switch($admenuselect)
    {
        $adExit{$admenuselect=$null;reloadADmenu}
        $GCred{Get-Cred;reloadadmenu}
        $Get_NewADObjects{$admenuselect=$null;Get-NewADObjects;reloadadmenu}
        $Run_Test_LDAP{$admenuselect=$null;Invoke-Expression ".\Run-Test-LDAP.ps1";reloadadmenu}
        $Get_ADSitSubnet{Get-ADSiteSubnet;reloadnetmenu}
        $Get_GpReport{$admenuselect=$null;Get-GPReport;reloadadmenu}
        $List_DomPCsLAPSPw{$admenuselect=$null;List-DomPCsLAPSPw;reloadadmenu}
        $Reset_ADUserPW{$admenuselect=$null;Reset-ADUserPW;reloadadmenu}
        $Run_ADCleanup{$admenuselect=$null;Run-ADCleanup;reloadadmenu}
        $Decommission_ADPC{$admenuselect=$null;Decommission-ADPC;reloadadmenu}
        $TestFix_PCSecureChannel{$admenuselect=$null;TestFix-PCSecureChannel;reloadadmenu}
        $Search_ExchMbox{$admenuselect=$null;Search-ExchMbox;reloadadmenu}
        
        default
        {
            $admenuselect = $null
            reloadADmenu
        }
    }
}


Function reloadadmenu()
{
        Read-Host "Hit Enter to continue"
        # Start-Sleep -m 250
        $admenuselect = $null
        admenu
}

# list AD objects where "WhenCreated" date is greater than ...
Function Get-NewADObjects
{
$WCDate = Read-Host "Enter a start date using the format mm/dd/yyyy to list AD Objects created since this date."

(get-adobject -filter * -properties *|where {$_.WhenCreated -gt $WCDate}) |select Name,canonicalname,whencreated,modified |sort WhenCreated,modified
}

Function Get-ADSiteSubnet()
{
    get-adreplicationsite -Filter * -Properties * |select Name,siteObjectBL,Subnets,Group |group siteObjectBL |select @{Name='Subnet';Expression={$_.Name.split(',')[0].Trim('{CN=')}},@{Name='Name';Expression={$_.Group.Name}} |ft
}




Function EmptyGroups ()
{
    Write-Verbose "Getting AD Groups that have no members"
    Get-ADGroup -Filter * | where { -Not ($_ | Get-ADGroupMember)} |select name |FT -Autosize
}
Function OldDCPCs() # View AD Computer last logon times
{
# adopted from module by Tilo 2013-08-27
# https://gallery.technet.microsoft.com/scriptcenter/Get-Inactive-Computer-in-54feafde
    Import-Module activedirectory
    $domain = $env:USERDOMAIN
    $DaysInactive = Read-host "Inactive for how many days?"
    $time = (Get-Date).AddDays(-($DaysInactive))

# Search for all AD computers with lastLogonTimestamp less than our time
    Get-ADComputer -filter {lastlogontimestamp -lt $time} -Properties lastlogontimestamp |

# output hostname and lastlogonTimestamp into a CSV
    Select-Object Name,@{Name="Time Stamp"; Expression={[DateTime]::FromFileTime($_.lastlogonTimeStamp)}}, Enabled |
    Sort-object Enabled, "Time Stamp" |
    <# Where-object Name -eq "pcname" | ft #> 
    Where-object Enabled -eq $True | ft
    # Add the following line if you want the result saved into a .CSV file
    #| Export-Csv OLD_Computer.csv -NoTypeInformation
}

Function GroupsOfMembers()
{
 [array]$ADUsers=(Read-Host "Enter a sam account name or * to see what groups the user(s) belong to").split(",") | %{$_.trim()}
    foreach ($ADUser in $ADUsers)
    {
        Write-Host $ADUser -ForegroundColor Yellow
        $GADUList1 = (Get-ADUser -filter 'SamAccountName -like $ADUser'|select -expandproperty SamAccountName)
        foreach ($ADUser1 in $GADUList1)
        {
            $GADUList = (Get-ADUser -filter 'SamAccountName -like $ADUser1' -Properties * |Select-Object select -ExpandProperty memberof)
            Write-Warning "$ADUser1 is a member of the following group(s)"
            foreach ($GADU in $GADUList)
            {
                ($GADU.split(",")[0].substring(3))
            }
        }
        
    }
   # Invoke-Expression ./Review-ADGroupMembership.ps1
}

Function List-DomPCsLAPSPw()
{
    Clear-Host
    Get-PCList
    foreach ($Global:PCLine in $Global:PCList)
    {
        Identify-PCName
        $PC = $Global:PC
        #$LDLAPsFltr = 'Name -like $PC' #+ '"*$PC*"'
        Get-ADComputer -filter {Name -like $PC} -Credential $Global:cred -Properties dnshostname,ms-Mcs-AdmPwd,ms-Mcs-AdmPwdExpirationTime | where {$_.'ms-Mcs-AdmPwd' -ne $null}|
            select Name, @{n="Admin LAPS PW (ms-Mcs-AdmPwd)";e={($_ | select -ExpandProperty ms-Mcs-AdmPwd)-join ','}},@{n="ms-Mcs-AdmPwdExpirationTime";e={w32tm /ntte ($_ | select -ExpandProperty ms-Mcs-AdmPwdExpirationTime)}} |ft # <#@{n="FQDN";e={($_ | select -ExpandProperty dnshostname)-join ','}},#> 
    }
}
 Function SysEvents()
{ 
    Get-PCList
    $WinLogName= Read-host "Enter Application, Security, Setup, or System"
    $EntryLevel_Type= Read-host " Enter Critical, Error, Warning, Information, FailureAudit, SuccessAudit"
    $MaxEvents = Read-host "How many log records do you want to search through?"
    $cnt = Read-host "How many log results do you want to see of this type?" 
    $id = Read-host "If applicable, enter a specific Event ID to search for"
    $Source = Read-host "Enter source application if desired (ex: wininit for chkdsk)`n`t(Only works with Get-Eventlog portion of script)"
        if ($pclist -eq "") {$pclist = [environment]::machinename}
        if ($WinLogName -eq "") {$WinLogName = "System"} 
        if ($EntryLevel_Type -eq "") {$EntryLevel_Type = "Error"}
        if ($MaxEvents -eq "") {$MaxEvents = 50}
        if ($cnt -eq "") {$cnt = 10}
        if ($id -eq "") {$id = "*"}
        if ($Source -eq "") {$Source = "*"}
    Invoke-Expression ".\Get-SysEvents.ps1 -PCList $Global:pclist -WinLogName $WinLogName -EntryLevel_Type $EntryLevel_Type -MaxEvents $MaxEvents -cnt $cnt -id $id -Source $Source"
    $Global:pclist=$null
}
Function UsrOnPC ()
{
# Report the Username logged into any active computer listed in Active Directory
    Clear-Host
    Get-PCList
    foreach ($Global:PCLine in $Global:PCList)
    {
        Identify-PCName
		#Get explorer.exe processes
		$proc = gwmi win32_process -computer $Global:PC -Filter "Name = 'explorer.exe'" -Credential $cred -ErrorAction SilentlyContinue
        #Search collection of processes for username
            if ($proc -ne $null)
            {
		        ForEach ($p in $proc) 
                {
	    	        $Script:PCUser = ($p.GetOwner()).User
                    Write-Host "$Global:PC :`t$Script:PCUser is logged in" -ForegroundColor Green
                }
            }
            else
            {
                Write-host "$Global:PC :`tUnable to get process" -ForegroundColor Red # $script:computer
            } 
	}
    if ($CalledFromFuction -eq $True)
        {
            return                
        }
 }
Function UsrsOnPCs()
{
# Report the IP Address, ComputerName and Username logged into any active computer listed in Active Directory
Clear-Host
Get-PCList
Write-Host "IPAddress `tMachine Name `tUser Logged on"
foreach ($Global:PCLine in $Global:pcList)
	{
	Identify-PCName
    $CompIP = $null
    $CompName = $null
    $ErrorActionPreference = "silentlycontinue"
    # Get IP from HostName if possible
    $CompIP = ([system.net.dns]::GetHostAddresses($Global:pc)).IPAddressToString
    If ($CompIP)
        {
            $ping = test-connection $Global:pc -Count 1 -Quiet -ErrorAction SilentlyContinue
            # Verify Computername, hostname, with IP
            $CompName = [System.Net.Dns]::GetHostByAddress($CompIP).HostName
            If ($CompName -notlike "*$Global:pc*")
            # If ($Global:pc -notlike "*$CompName*")
                {
                    Write-host "Cannot use $Global:pc as a valid computer name " -ForegroundColor Red
                }
        }
    Else{
            $ping = $false
            $CompIP = "   .   .   .   "
        }
    if($ping -eq $true)
        {
            # $pcinfo = @()
		    #Get explorer.exe process
		    $proc = gwmi win32_process -computer $Global:pc -Filter "Name = 'explorer.exe'" -Credential $cred -ErrorAction SilentlyContinue
		    #Search collection of processes for process owner username
            if ($proc -ne $null)
            {
		        ForEach ($p in $proc) {
                    $pscn = $p.pscomputername
                    $pgo = $p.GetOwner().user
                    Write-Host "$pscn :`t$CompIP`t$pgo`tis logged in." -ForegroundColor Green
           
		        }
            }
            else
            {
                Write-host "$Global:PC :`tUnable to get process" -ForegroundColor Red # $script:computer
            }
        }
        else
        {
            Write-Host "$Global:pc :`t$CompIP`tUNK`tNONE" -ForegroundColor Red
        }
    }
    $ErrorActionPreference = "continue"
    Write-Output "`n"
    if ($CalledFromFuction -eq $True)
        {
            return                
        }
}

Function FilterADUsers()
{
    $Fltr = Read-Host "Type All or part of a users name to filter the list. `nUse * for all users and as a wild-card at the beginning or end of your search criteria."
    $ADUserList = get-aduser -filter 'Name -like $Fltr' -Properties Samaccountname, ObjectGUID, passwordlastset, Name, EmailAddress,DisplayName,LastLogonDate,Created,Mail,Enabled |where {$_.enabled -eq $true} |select Samaccountname, passwordlastset, Name, EmailAddress,DisplayName,LastLogonDate, ObjectGUID,Created,Mail,Enabled |sort SamAccountName, ObjectGUID # |Out-GridView
    $ADUserList.count
    $ADUserList |FT -AutoSize -Wrap
    $adgm = $ADUserList|select Name,DisplayName
    $adgm |FT -AutoSize 
}
Function ADGrps-n-Usrs()
{
$groupsusers=get-adgroup -filter * | 
    ForEach-Object{
                $settings=@{Group=$_.DistinguishedName;Member=$null}
        $_ | get-adgroupmember |
		ForEach-Object
                {
                    $settings.Member=$_.DistinguishedName
                    New-Object PsObject -Property $settings
                }
    }
    
$ScriptDIR = Split-Path $Script:MyInvocation.MyCommand.Path
$GrpUsrcsv = "\GroupsUsers.csv"
$groupsusers | Export-Csv $scriptdir$GrpUsrcsv -NoTypeInformation
Write-host "The information has been saved to: " $ScriptDIR$GrpUsrcsv 
}

Function ADNonExpiring()
{
    # adapted from Netwrix
    Search-ADAccount -PasswordNeverExpires |Select-Object ObjectClass, Name, SAMAccountName, PasswordNeverExpires |ft # | Export-Csv c:\temp\users_password_expiration_false.csv
}
Function Search-ExchMbox()
{
    # https://technet.microsoft.com/en-us/library/dd298173(v=exchg.160).aspx
    Clear-Host
    # https://technet.microsoft.com/en-us/library/dd335083(v=exchg.160).aspx
    # How to create and import session to Exchange to use exchange commands from local computer
    # Import-PSSession $pssExch
    Write-Host "`nRemote into Exchange13 and start a Windows PowerShell or Exchange Managment Shell`n" -ForegroundColor Yellow
    Write-Host "Run the Recover-Messages.ps1 script from the default profile folder`n" -ForegroundColor Yellow
    <#
        # This code is not quire ready for prime-time yet. It needs work invoking the Recover-Messages.ps1 on the remote system or locally (exchange addin needed)
        $pssExch = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange13/PowerShell -Authentication Kerberos -Credential $cred
        Import-PSSession $pssExch
        $recMessages = ".\Recover-Messages.ps1"
        invoke-command -FilePath "$recMessages"
        Read-Host "Touch a Enter to continue"
        Remove-PSSession $pssExch
    #>

}

Function ADUser-Logons()
{
    $nowlessnum =((get-date).AddDays(-30))
    $usecnt = Get-ADUser -Filter {lastlogondate -lt $nowlessnum} |ft
    $usecnt

    
}

Function Reset-ADUserPW()
{
    Clear-Host
    Write-host "Note that the Change option will provide a list of users and dates their passwords were last set (with or without options)"
    $Choice = Read-Host "Enter List, Unlock, Reset, or Change "
    
    if ($Choice -eq "Unlock")
    {
        $ADUser = Read-Host "Enter part of the user's name "
        Get-ADUser -Filter ('name -like "*' + $ADUser + '*"') |Unlock-ADAccount -Confirm

     
    }    
# Need to verify this option
    if ($Choice -eq "Reset")
    {
        $ADUser = Read-Host "Enter part of the user's name "
        # $ADUserResult = Get-ADUser -Filter ('name -like "*' + $ADUser + '*"')
        $ADUserResult = Get-ADUser -filter 'samaccountname -eq $ADUSer'
        $ADUserSAM = $ADUserResult.SamAccountName
        $newCred = Get-Credential "$ADUserSAM"
         #|Set-ADAccountPassword -NewPassword
     
    }
    if ($Choice -eq "Change")
    {
        Invoke-Expression ".\Set-ADUserPwd.ps1"
        while($global:AnotherPW -eq "Yes")
            {
                Invoke-Expression ".\Set-ADUserPwd.ps1"
                Clear-host
            }

     
    }
    if ($Choice -eq "List")
    {
        $SearchBase = "CN=Users,DC=lec,DC=lan"
        # $Completed = get-aduser -filter <#(PasswordLastSet -gt "5/16/2016") -AND #> '(mail -notlike "healthmailbox*")' -SearchBase $SearchBase  -Properties * |where {$_.enabled -eq $true} |select Samaccountname,passwordlastset,cannotchangepassword,passwordexpired,lockedout,mail|sort passwordlastset |ft -AutoSize
        $Completed = get-aduser -filter <#(PasswordLastSet -gt "5/16/2016") -AND #> '(mail -notlike "healthmailbox*")' -SearchBase $SearchBase  -Properties * |where {$_.enabled -eq $true} |select Samaccountname,passwordlastset,cannotchangepassword,passwordexpired,lockedout,mail|sort passwordlastset |ft -AutoSize #|Out-GridView
        $completed
        $cc = $Completed.count
        Write-host "Completed: $cc" -ForegroundColor Yellow
        $Waiting =  get-aduser -filter '(PasswordLastSet -lt "5/15/2016") -AND (mail -notlike "healthmailbox*")' -SearchBase $SearchBase  -Properties * |where {$_.enabled -eq $true} |select Samaccountname,passwordlastset,cannotchangepassword,passwordexpired,lockedout,mail|sort passwordlastset |ft -AutoSize # |Out-GridView
        $Waiting 
        $wc = $Waiting.count
        Write-host "Waiting: $wc" -ForegroundColor Yellow
    }
        
}

Function Get-GPReport()
{
    Clear-Host
    $PSScriptRoot
        $CalledFromFuction = $True
    $rptpath = "$PSScriptRoot\GPReport.htm"
    UsrsOnPCs
    # $Global:PC = Read-Host "Type the name of a PC"
    $User = Read-Host "Enter the username of the policy holder"
    Get-GPResultantSetOfPolicy -Computer $Global:PC -User $User -reporttype html -path $rptpath
    start chrome.exe $rptpath
    $CalledFromFuction = $False
}

Function Run-ADCleanup ()
{
<#
 Function depends on an import of the ActiveDirectory module
 Use this function to delete computers from Active Directory
 Designed to remove inactive ADComputer accounts from the Active Directory
#>
Clear-Host
$ADDays = Read-Host "Go back how many days? [Default -1095 (3 years)]"
# modify $ADSBase and/or add a Read-Host before the first quote to prompt for a different search base
$ADSBase = "OU=Disabled,DC=LEC,DC=LAN" 
$DaysAgo = (Get-Date).AddDays($ADDays)
# Apply filter
$ADPCList = get-adcomputer -Filter {(SamAccountName -like "*") -and ((PwdLastSet -lt $DaysAgo) -and (LastLogonTimeStamp -lt $DaysAgo)) -AND (CN -ne "EDGETRANSPORT")}`
    -SearchBase $ADSBase -Properties * 
# List filtered results
$ADPCList|Sort LastLogonTimeStamp,PwdLastSet,CN |`
    select Enabled,CN,SamAccountName,DistinguishedName,`
    @{Name="PwdLastSet";Expression={[datetime]::FromFileTime($_.PwdLastSet)}},`
    @{Name="LastLogonTimeStamp";Expression={[datetime]::FromFileTime($_.LastLogonTimeStamp)}}|FT #Out-GridView
Write-Warning "Ready to delete the above systems from $ADSBase`nIt appears they haven't logged into AD or changed their password since $DaysAgo"
# Process ADObjects recursively and prompt to Delete
Foreach ($PCline in $ADPCList)
    {
        $pc = $pcline.name
        $AD1 = get-adobject -filter * -SearchBase "CN=$pc,OU=Disabled,DC=LEC,DC=LAN"
        $AD1.Count
        if ($AD1.Count -gt 0)
            {
                Write-Warning "Using -Recursive because $pc has leaf objects"
                $AD1
                $AD1|sort -Descending DistinguishedName|Remove-ADObject -Recursive -Confirm
            } 
        Else
            {
                $AD1 |Remove-ADObject -confirm
            }
    }
# List all deleted objects
$GADOPrompt = Read-Host "Type Y to list all deleted ADObjects"
If ($GADOPrompt -eq "Y")
    {
     
        Get-ADObject -Filter 'isDeleted -eq $true -and name -ne "Deleted Objects"' -IncludeDeletedObjects -Properties *|sort ModifyTimeStamp |select isRecycled,SamAccountName,ModifyTimeStamp,ObjectGUID,DistinguishedName |ft -Wrap -AutoSize
    }
}

Function Decommission-ADPC
{
    Clear-Host
    Get-PCList
    $Global:pclist|Get-ADComputer -Properties * |where {$_.enabled -eq $true} |select Name,Enabled,DistinguishedName,description |Out-GridView
    # change *Disable* in the filter if your naming convention is different or create an OU named Disabled, etc...
    $GADOUDisable = (Get-ADOrganizationalUnit -Filter * |where {$_.Name -like "*Disable*"}).DistinguishedName
    if (($DestTargetOU = Read-Host "Enter the destination OU. Or [Enter] to use default: $GADOUDisable") -eq "") {$DestTargetOU = $GADOUDisable}
    foreach ($Global:PCline in $Global:PCList)
    {
        Identify-PCName
        $ADPCObj = get-adcomputer -Identity $Global:PC -Properties *
        $ADPCObj|select Name,Enabled,DistinguishedName |FT
        # Pre-Pend D_ to computer object description
        $ADPCO_D = $ADPCObj.Description
        $ADPCO_D = "D_"+$ADPCO_D
        # Disable the computer object and change its description (perform before moving the object)
        Set-ADComputer $ADPCObj -Description $ADPCO_D -Enabled $false -Confirm
        # Move the system to the destination target OU
        $ADPCObj|Move-ADObject -TargetPath $DestTargetOU -Confirm
        Get-ADComputer -Identity $Global:PC -Properties * |select Name,Enabled,DistinguishedName |FT -HideTableHeaders
    }
}
Function TestFix-PCSecureChannel
{
     # See 2015/01/18 post by Ahmed Raza for other options (https://superuser.com/questions/555297/re-joining-a-computer-to-domain)
     $CredUser = $Cred.UserName
     if ((Read-Host "Current credentials are for: `"$CredUser`" , type C to change or [Enter] to continue")-eq "C") {Get-Cred}
     if ((Test-ComputerSecureChannel -Credential $Cred -verbose) -eq $false)
     {
        $FixSCPrompt = Read-Host "Type yes to fix the secure channel (trust relationship) between the local computer and the domain $env:USERDNSDOMAIN "
        if ($FixSCPrompt -eq "yes")
        {
            # Account credentials must be authorized to join computers to the domain
            Test-ComputerSecureChannel -Credential $Cred -Repair -Confirm -Verbose
        }
     }
}
admenu