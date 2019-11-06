<#
.NAME
    Run-psMenu.ps1
.SYNOPSIS
    Useful Active Directory related PS tasks
    Some script functions will only work domain-wide when the script is run with administrative privileges
.DESCRIPTION
    Menu design patterned after: https://quickclix.wordpress.com/2012/08/14/making-powershell-menus/
    2016-04-20 -> 2019-07-23
    Written by Jason Crockett - Laclede Electric Cooperative

 use Ctrl+J to bring up "snippets" that can be inserted into code
 Tips from MVA
    - Scripting ToolMaking
    - www.powershell.org
    - #PowerShell on Twitter & @JSnover
.PARAMETERS
    None
.EXAMPLE
    .\Run-psMenu.ps1 []
.SYNTAX
    .\Run-psMenu.ps1 []
.REMARKS
    To see the examples, type: help Run-psMenu.ps1 -examples
    To see more information, type: help Run-psMenu.ps1 -detailed
.TODO
    Enable Script block logging
#>
<#
[CmdletBinding()] #(allows the use of parameter attributes)
param(
    [Parameter(Mandatory=$True)]
    [string[]]$comps = "$env:COMPUTERNAME", #was 'localhost' This can change it's just here to hold the place and illustrate how/where to insert parameters & for CmdletBinding to not fail
    [string]$Non-Array = "Something"
)
#>

# Enable Exchange cmdlets (On an Exchange Server)
# add-pssnapin *exchange* -erroraction SilentlyContinue 

# Utilize Script Analyzer on scripts
# Script Analyzer example:
# 

# Enable AD and PS Update modules
# See Getting Started with PowerShell 3.0 Jump Start in MS Virtual Academy for loading modules from other systems
# $ADPS=new-pssession -computername $ADSRV
# AD module is available on a non-server OS after installing the Remote Server Administration Tool (RSAT) WindowsTH-RSAT_WS2016-x64.msu
Import-Module ActiveDirectory
Import-Module PSWindowsUpdate
# Import-Module AzureAutomationAuthoringToolkit
Clear-Host
# Verify Powershell version:
$PSVersionTable.PSVersion
. .\Get-SystemsList.ps1 # load functions in .ps1
# Identify-PCName was in Get-SystemsList.ps1 but doesn't seem to be getting pulled in to be used so I'm loading it here - perhaps remove this after troubleshooting
function Identify-PCName
{
    # This if statement could potentially be moved to inside the Get-PCList function, but it might need to be modified to do so
    if ($Global:pcline.name -eq $null)
    {
        $Global:pc = $Global:pcline
    }

    else
    {
        $Global:pc = $Global:pcline.name
    }
}
# List all the available colors
# Thanks: Gary Siepser @ https://blogs.technet.microsoft.com/gary/2013/11/20/sample-all-powershell-console-colors/
[enum]::GetValues([System.ConsoleColor]) |ForEach-Object {Write-Host $_ -ForegroundColor $_ -nonewline}
<#
# List all available forground and background color combinations
# Thanks: nattydread2009 @ https://blogs.technet.microsoft.com/gary/2013/11/20/sample-all-powershell-console-colors/
$colors = [enum]::GetValues([System.ConsoleColor])
foreach( $fcolor in $colors )
{
    foreach( $bcolor in $colors )
    {
        Write-Host "ForegroundColor is : $fcolor BackgroundColor is : $bcolor "-ForegroundColor $fcolor -BackgroundColor $bcolor
    }
}
#>
<#
Find all of the association classes (See: https://devblogs.microsoft.com/scripting/use-powershell-cim-cmdlets-to-discover-wmi-associations/)
get-cimclass -QualifierName association |? cimclassname -Match '^win32'|sort CimClassName
-OR-
get-cimclass -QualifierName association |%{get-cimclass $_.CimClassName -QualifierName dynamic} |sort CimClassName
Find out which classes make up the association
get-cimclass -QualifierName association |%{get-cimclass $_.CimClassName -qualifiername dynamic}|%{$_ |select -ExpandProperty cimclassproperties}|group name |sort count -descending
#>
# set script variables
[BOOLEAN]$Global:ExitSession=$false
# Get Credentials for later wmi requests
$cred = Get-Credential "$env:USERDOMAIN\Administrator"
$Script:t4 = "`t`t`t`t"
$Script:d4 = "-------------------------------------------------------------------------------"
$script:NC = $null
$script:p1 = $null
$script:p2 = $null
$script:p3 = $null
$Script:PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
Write-Warning $Script:PSScriptRoot
$Global:CDate = Get-Date -UFormat "%Y%m%d"
$Global:CDateTime = [datetime]::ParseExact($Global:CDate,'yyyymmdd',$null)
[int]$menuselect = 0

function chcolor($p1,$p2,$p3,$NC){
            Write-host $t4 -NoNewline
            Write-host "$p1 " -NoNewline
            Write-host "$p2" -ForegroundColor $NC -NoNewline
            Write-host "$p3" -ForegroundColor White
            $script:NC = $null
            $script:p1 = $null
            $script:p2 = $null
            $script:p3 = $null
}
# The following is an example of combining -Property and -ExpandProperty
# (Get-ADUser -filter 'SamAccountName -like $ADUser1' -Properties * |Select-Object -property @{n="ADUserInfo";e={$_.SamAccountName,$_.memberof}}|select -ExpandProperty ADUserInfo)
function ADmenu()
{
    while ($admenuselect -lt 1 -or $admenuselect -gt 27)
    {
        Trap {"Error: $_"; Break;}        
        $MNum = 0;Clear-host |out-null
        Start-Sleep -m 250
        Write-host $t4 $d4 -ForegroundColor Green
        # Exit to Main Menu
        $MNum ++;$adExit = $MNum; 
		    $p1 =" $MNum. `tSelect to ";
		    $p2 = "Exit ";
		    $p3 = "to Main Menu"; $NC = "Red";chcolor $p1 $p2 $p3 $NC
        # User Members in a group
        $MNum ++;$adGroup_ID = $MNum;
		    $p1 =" $MNum. `tView ";
		    $p2 = "Members ";
		    $p3 = "in a group."; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Group Memberships assigned to a user
        $MNum ++;$adUserGroups = $MNum;
		    $p1 =" $MNum. `tView ";
		    $p2 = "Groups ";
		    $p3 = "assigned to users."; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Group Memberships by individual"
        $MNum ++;$adGrpMbrXInd = $MNum;
            $p1 =" $MNum. `tView ";
		    $p2 = "Group Memberships ";
		    $p3 = "by individual "; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # PC Information and Service Tag
        $MNum ++;$adPCInfoSTag = $MNum;
		    $p1 =" $MNum. `tPC";
		    $p2 = "Information ";
		    $p3 = "and Service Tag"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Domain PCs that haven't logged in for 90 days
        $MNum ++;$adPCLog90 = $MNum;
		    $p1 =" $MNum. `tDomain PCs:";
		    $p2 = "90 days ";
		    $p3 = "since previous logon"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # List all domain member PCs
        $MNum ++;$adListAllPCs = $MNum;
		    $p1 =" $MNum. `tList all";
		    $p2 = "member PCs ";
		    $p3 = "in this domain"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # List all domain member PCs
        $MNum ++;$adUserAssignedPC = $MNum;
		    $p1 =" $MNum. `tList primary user(s)";
		    $p2 = "PC assignments ";
		    $p3 = "in this domain"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # List which user(s) is/are logged onto selected computer(s) listed in the domain
        $MNum ++;$adUserLoggedOn = $MNum;
		    $p1 =" $MNum. `tList";
		    $p2 = "user Logged in ";
		    $p3 = "to a computer in this domain"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # List which user(s) are logged onto selected computer(s) listed in the domain
        $MNum ++;$adAllUserLoggedOn = $MNum;
		    $p1 =" $MNum. `tList ";
		    $p2 = "users logged into, `n$t4 ";
		    $p3 = "selected computers in this domain"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Perform one of the following: Copy a file, fix links, view mapped drives to/on selected computers in this domain
        $MNum ++;$adManageADUFiles = $MNum;
		    $p1 =" $MNum. `tPerform one of the following: ";
		    $p2 = "Copy a file, fix links, view mapped drives, `n$t4`tor read data from a file ";
		    $p3 = "to/on selected computers in this domain"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # List Empty Active Directory Groups
        $MNum ++;$adEmptyGroups = $MNum;
		    $p1 =" $MNum. `tList all";
		    $p2 = "Empty Groups ";
		    $p3 = "in the active directory"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Event logs
        $MNum ++;$adEventLogs = $MNum;
		    $p1 =" $MNum. `tView";
		    $p2 = "Event Logs ";
		    $p3 = ""; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Active Directory Groups and Members -> .CSV
        $MNum ++;$adGrpMembers = $MNum;
		    $p1 =" $MNum. `tAD";
		    $p2 = "Groups and Members ";
		    $p3 = "-> .CSV"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # AD Users List
        $MNum ++;$ADMenuUserList = $MNum;
		    $p1 =" $MNum. `tAD";
		    $p2 = "User List ";
		    $p3 = ""; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # AD Users List
        $MNum ++;$ADMenuFltrUserList = $MNum;
		    $p1 =" $MNum. `tFilter AD";
		    $p2 = "User List ";
		    $p3 = "and list workstation"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # AD accounts that will never expire
        $MNum ++;$adNoExpiration = $MNum;
		    $p1 =" $MNum. `tAD accounts";
		    $p2 = "never to expire ";
		    $p3 = ""; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # AD Privileged Group Membership and password age
        $MNum ++;$adPrivGrpMemAge = $MNum;
		    $p1 =" $MNum. `tAD Privileged";
		    $p2 = "Group Membership and Password ";
		    $p3 = "age"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Search an Exchange mailbox `n$t4`t (this will only work on Exchange and is included for reference only here.)
        $MNum ++;$adSearchExch = $MNum;
		    $p1 =" $MNum. `tSearch";
		    $p2 = "Exchange Mailbox ";
		    $p3 = "(reference only)"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Architecture list of all Domain PCs
        $MNum ++;$adPCArchList = $MNum;
		    $p1 =" $MNum. `tList";
		    $p2 = "Architecture ";
		    $p3 = "of select Domain PCs"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Unlock, Reset, or Change a user's password
        $MNum ++;$adUserPW = $MNum;
            $p1 =" $MNum. `tUnlock, Reset, or Change";
		    $p2 = "the password ";
		    $p3 = "on a user's PC"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Disable selected ADUser(s) and move objects to a new AD container (default is the Disabled container)
        $MNum ++;$adDisableADUser = $MNum;
            $p1 =" $MNum. `t";
		    $p2 = "Disable ADUser ";
		    $p3 = "and move to a new AD container (with -Confirm)"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Group Policy Results for Domain PC/User
        $MNum ++;$adGPReport = $MNum;
            $p1 =" $MNum. `tShow ";
		    $p2 = "Group Policy Results ";
		    $p3 = "for Domain PC/User"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # View PC Startups
        $MNum ++;$adPCStartups = $MNum;
            $p1 =" $MNum. `tView ";
		    $p2 = "Startups ";
		    $p3 = "on PCs "; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # View WinSAT / Windows Experience Index...
        $MNum ++;$adWinSat = $MNum;
            $p1 =" $MNum. `tView ";
		    $p2 = "WinSAT ";
		    $p3 = "/ Windows Experience Index... "; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # List OpenFiles on Servers
        $MNum ++;$adOpFiles = $MNum;
            $p1 =" $MNum. `tList ";
		    $p2 = "open files ";
		    $p3 = "on AD file servers... "; $NC = "yellow";chcolor $p1 $p2 $p3 $NC   
        # Delete Old AD Computer objects from AD
        $MNum ++;$ADM_ADCleanup = $MNum;
            $p1 =" $MNum. `tRemove ";
		    $p2 = "Old AD Computer Objects ";
		    $p3 = "from Active Directory "; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        Write-Host "$t4$d4"  -ForegroundColor Green
        # Setting up the menu in this way allows you to move menu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $MNum ++;$p1 =" $MNum.  FIRST";$p2 = "HIGHLIGHTED ";$p3 = "END"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        [int]$admenuselect = Read-Host "`n`tPlease select the menu item you would like to do now"
    }
    switch($admenuselect)
    {
        $adExit{$admenuselect=$null;reloadmenu}
        $adGroup_ID{$admenuselect=$null;GroupMembership;reloadadmenu}
        $adUserGroups{$admenuselect=$null;GroupsOfMembers;reloadadmenu}
        $adGrpMbrXInd{Write-Warning "get-GrpMbrXInd has not been imported from adutils";reloadadmenu}
        $adPCInfoSTag{$admenuselect=$null;pcinfo;reloadadmenu}
        $adPCLog90{$admenuselect=$null;OldDCPCs;reloadadmenu}
        $adListAllPCs{$admenuselect=$null;ListDomPCs;reloadadmenu}
        $adUserLoggedOn{$admenuselect=$null;UsrOnPC;reloadadmenu}
        $adUserAssignedPC{$admenuselect=$null;Find-UserPCs;reloadadmenu}
        $adAllUserLoggedOn{$admenuselect=$null;UsrsOnPCs;reloadadmenu} # Good one to multi-thread
        $adManageADUFiles{$admenuselect=$null;Manage-ADUserFiles;reloadadmenu}
        $adEmptyGroups{$admenuselect=$null;EmptyGroups;reloadadmenu}
        $adEventLogs{$admenuselect=$null;SysEvents;reloadadmenu}
        $adGrpMembers{$admenuselect=$null;ADGrps-n-Usrs;reloadadmenu}
        $ADMenuUserList{$admenuselect=$null;alladusers;reloadadmenu}
        $ADMenuFltrUserList{$admenuselect=$null;Filteradusers;reloadadmenu}
        $adNoExpiration{$admenuselect=$null;ADNonExpiring;reloadadmenu}
        $adPrivGrpMemAge{$admenuselect=$null;Get-ADPrivGrpMbrshp;reloadadmenu}
        $adSearchExch{$admenuselect=$null;SearchExchMbox;reloadadmenu}
        $adPCArchList{$admenuselect=$null;ADPcArch;reloadadmenu}
        $adUserPW{$admenuselect=$null;Reset-ADUserPW;reloadadmenu}
        $adDisableADUser{$admenuselect=$null;Disable-ADUser;reloadadmenu}
        $adGpReport{$admenuselect=$null;GPReport;reloadadmenu}
        $adPCStartups{$admenuselect=$null;Get-ADPcStarts;reloadadmenu}
        $adWinSat{$admenuselect=$null;get-winSAT;reloadadmenu}
        $adOpFiles{$admenuselect=$null;get-OpenFiles;reloadadmenu}
        $ADM_ADCleanup{$admenuselect=$null;Run-ADCleanup;reloadadmenu}
        
        default
        {
            $admenuselect = $null
            reloadmenu
        }
    }
}

function NetMenu()
{
while ($NetMenuSelect -lt 1 -or $NetMenuSelect -gt 17)
    {
        Trap {"Error: $_"; Break;}        
        $MNum = 0;Clear-host |out-null
        Start-Sleep -m 250
        Write-host $t4 $d4 -ForegroundColor Green
        # Exit
        $MNum ++;$Exit=$MNum;
            $p1 = " $MNum. ";
            $p2 = "Exit to Main Menu";
            $p3 = " ";
            $NC = "Red";chcolor $p1 $p2 $p3 $NC
        # Test PC(s) for network connectivity
        $MNum ++;$testConnections=$MNum;
            $p1 =" $MNum.  Run";
            $p2 = "Test-Connection ";
            $p3 = "on select PCs or all PCs in Domain";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Firewall Status
        $MNum ++;$FWStatus=$MNum;
            $p1 =" $MNum.  View";
            $p2 = "Firewall Status ";
            $p3 = "on select PCs or all PCs in Domain";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Test Ports
        $MNum ++;$testports=$MNum;
            $p1 =" $MNum.  Test";
            $p2 = "ports ";
            $p3 = "on a system or list of systems";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Ping
        $MNum ++;$PingTest=$MNum;
            $p1 =" $MNum.  Test";
            $p2 = "(Ping) Network ";
            $p3 = "Connection Persistence";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # IP from MAC
        $MNum ++;$GetIP=$MNum;
            $p1 =" $MNum. ";
            $p2 = "MAC 2 IP ";
            $p3 = "- Use a known MAC address to find an IP address";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Status of NIC on PC
        $MNum ++;$CallGetNICStatus=$MNum;
            $p1 =" $MNum. Get ";
            $p2 = "NIC ";
            $p3 = "Status and information";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # View IPConfig-like information on select systems
        $MNum ++;$GetIPConfig=$MNum;
            $p1 = " $MNum. View ";
            $p2 = "IPConfig";
            $p3 = " information for select systems"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC   
        # HostName from IP
        $MNum ++;$Gethostname=$MNum;
            $p1 =" $MNum.  Find";
            $p2 = "hostname from IP ";
            $p3 = "address";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Validate IP
        $MNum ++;$TestValidIPPrompt=$MNum;
            $p1 =" $MNum.  Identify";
            $p2 = "valid IP ";
            $p3 = "address from input";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # NetStat
        $MNum ++;$ViewNetStat=$MNum;
            $p1 =" $MNum.  View";
            $p2 = "NetStat results ";
            $p3 = "for a system on the network";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # nlasvc restart
        $MNum ++;$RestartNLASvc=$MNum;
            $p1 = " $MNum.  Restart";
            $p2 = "nlasvc";
            $p3 = ", the local Network Location Awareness Service";
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
        # Set NIC location
        $MNum ++;$GetSetncp=$MNum;
            $p1 = " $MNum. Get or change";
            $p2 = "Location";
            $p3 = " for a local NIC";
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
        # Wait for live port
        $MNum ++;$WaitLive=$MNum;
            $p1 = " $MNum. Wait for ";
            $p2 = "System port number";
            $p3 = " to come alive";
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
                # Wait for live port
        $MNum ++;$TestSrvConnects=$MNum;
            $p1 = " $MNum. Check ";
            $p2 = "Active Status";
            $p3 = " for critical systems";
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
        # Check for Active RDP sessions on LEC servers
        $MNum ++;$GetRDPStats=$MNum;
            $p1 = " $MNum. Check ";
            $p2 = "for Active RDP";
            $p3 = " sessions on LEC servers";
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
        # Check for Primary subnet of AD Site
        $MNum ++;$GetADSitSubnet=$MNum;
            $p1 = " $MNum. View ";
            $p2 = "AD Site and subnet";
            $p3 = " for the current Active Directory context"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC     
            
        # Setting up the menu in this way allows you to move menu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $MNum ++;$p1 =" $MNum.  FIRST";$p2 = "HIGHLIGHTED ";$p3 = "END"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        Write-Host "$t4$d4"  -ForegroundColor Green
        $Script:MNum = 0;[int]$NetMenuSelect = Read-Host "`n`tPlease select the menu item you would like to do now"
    }


    switch($NetMenuSelect)
        {
            # Variables at the beginning of these lines in the switch are numbers assigned by their position in the above
            # configured menu layout. This allows them to be moved around and let the numbers automatically adjust
            $Exit{$NetMenuSelect=$null;reloadmenu}
            $testConnections{test-connections;reloadnetmenu}
            $testports{test-ports}
            $RestartNLASvc{Restart-NLASvc}
            $GetSetncp{GetSet-ncp}
            $PingTest{Test-PCUp} #PingTest;reloadnetmenu}
            $GetIP{Get-IP}
            $GetIPConfig {Get-IPConfig;reloadnetmenu}
            $CallGetNICStatus{CallGet-NICStatus;reloadnetmenu}
            $Gethostname{Get-hostname}
            $TestValidIPPrompt{clear-host|Test-ValidIPPrompt}
            $ViewNetStat{View-NetStat}
            $FWStatus{FWStatus}
            $WaitLive{Wait-Live}
            $TestSrvConnects{test-srvconnects;reloadnetmenu}
            $GetRDPStats{Get-RDPStats}
            $GetADSitSubnet{Get-ADSiteSubnet;reloadnetmenu}


            default
            {
                $NetMenuSelect = $null
                $global:NC = " "
                $global:p1 = " "
                $global:p2 = " "
                $global:p3 = " "
                $Script:MNum = 0
                ReloadNetMenu
            }
        }
}

function menu()
{
while ($menuselect -lt 1 -or $menuselect -gt 28)
{
    Clear-Host
        Trap {"Error: $_"; Break;}        
        $MNum = 0;Clear-host |out-null
        Start-Sleep -m 250
        Write-host $t4 $d4 -ForegroundColor Green
#    Write-Host "$t4  1. Exit Now" -ForegroundColor Red
        # Exit
        $MNum ++;$Exit=$MNum;
            $p1 = " $MNum. `t";
            $p2 = "Exit to Main Menu";
            $p3 = " ";
            $NC = "Red";chcolor $p1 $p2 $p3 $NC
#    Write-Host "$t4  2. Active Directory Menu" -ForegroundColor Yellow
            # Go to the Active Directory Menu
        $MNum ++;$ADmenu=$MNum;
            $p1 = " $MNum. `t Go to ";
            $p2 = "Active Directory";
            $p3 = " menu"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC

#    Write-Host "$t4  3. Net Utilities Menu" -ForegroundColor Yellow
            # Go to the Net Utilities Menu
        $MNum ++;$NetMenu=$MNum;
            $p1 = " $MNum. `t Go to ";
            $p2 = "Net Utilities";
            $p3 = " menu"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
#    Write-Output "$t4  4. RDP - Enable on a workstation"
    # Enable RDP on a system
        $MNum ++;$EnableRDP=$MNum;
            $p1 = " $MNum. `t";
            $p2 = "Enable RDP";
            $p3 = " on a Windows system"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
#    Write-Output "$t4  5. View Non-Domain PC Information and Service Tag"
    # View Non-Domain PC Information and Service Tag
        $MNum ++;$NonDomPCinfo=$MNum;
            $p1 = " $MNum. `t View ";
            $p2 = "non-domain PC Service Tag ";
            $p3 = "and generic pc information"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
#    Write-Output "$t4  6. Event log view"
    # View Event log information
        $MNum ++;$SysEventlogs=$MNum;
            $p1 = " $MNum. `t View specific";
            $p2 = "Event log ";
            $p3 = "information."
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
#    Write-Output "$t4  7. Create a new Remote PowerShell connection"
        $MNum ++;$RmtPwr=$MNum;
            $p1 = " $MNum. `t Create a new ";
            $p2 = "Remote PowerShell connection ";
            $p3 = "(disconnect from this menu)"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
#    "Discover boot times (from select systems)"
        $MNum ++;$BootTimes=$MNum;
            $p1 = " $MNum. `t Discover ";
            $p2 = "boot times ";
            $p3 = "(from select systems)"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
    Write-Output "$t4  9. Java versions installed on PC"
    Write-Output "$t4 10. Process status on a PC"
    Write-Output "$t4 11. Update PowerShell Helps"
    Write-Output "$t4 12. PC Info and Service Tag"
    Write-Output "$t4 13. List programs installed on a pc"
    Write-Output "$t4 14. Uninstall (a) program(s) on a pc"
    Write-Output "$t4 15. Service Status Check"
    Write-Output "$t4 16. MS KB update installation status"
    # Check for Most Recent Windows Update on LEC servers
    Write-Output "$t4 17. List most recent MS update for the LEC servers"
    Write-Output "$t4 18. View associated logon names used to start a service"
    Write-Output "$t4 19. convert-Decimal to Binary to Hexadecimal"
    Write-Output "$t4 20. Get text URL Web Page Info"
    Write-Output "$t4 21. Get a list of Files Modified in the last xx days"
    Write-Output "$t4 22. NTFS Folder - Change Ownership and Add FullControl permissions then prompt to delete."
    Write-Output "$t4 23. Browse-Share"
    Write-Output "$t4 24. Find Duplicate Files"
    Write-Output "$t4 25. View a file hash"
    Write-Output "$t4 26. View-Certificates"
    Write-Output "$t4 27. Screen saver status on pc to see if a user is AFK"            
        # Setting up the menu in this way allows you to move menu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $MNum ++;$p1 =" $MNum.  FIRST";$p2 = "HIGHLIGHTED ";$p3 = "END"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        Write-Host "$t4$d4"  -ForegroundColor Green
        $Script:MNum = 0;[int]$MenuSelect = Read-Host "`n`tPlease select the menu item you would like to do now"


    <#
    Write-Host "$t4$d4"  -ForegroundColor Green
    [int]$menuselect = Read-Host "`n`tPlease select the menu item you would like to do now"
    #>
}

switch($menuselect)
    {
        # 1{}
        $Exit{$MenuSelect=$null}
        #2{ADmenu}
        $ADmenu{ADmenu}
        $NetMenu{Clear-Host;NetMenu}
        $EnableRDP{Enable-RDP}
        $NonDomPCinfo{NoDomPCinfo;reloadmenu}
        $SysEventlogs{SysEvents;reloadmenu}
        $RmtPwr{New-PSRemote;reloadmenu}
        $BootTimes{. .\Get-BootTime.ps1;Discover-BootTime;reloadmenu}
        9{ID-Java}
        10{Get-ProcessStatus}
        11{updthlp}
        12{pcinfo;reloadmenu}                  # Is this the same as 5 - NoDomPCInfo?
        13{get-apps}
        14{Uninstall-apps}
        15{Get-Svc;reloadmenu}
        16{get-KB;reloadmenu}
        17{List-RecentUpdate}
        18{Get-ServiceLogons}
        19{convert-Numbers}
        20{wGet-URL}
        21{Get-ModifiedFileList}
        22{Clear-Host;Manage-NTFS;reloadmenu}
        23{Browse-Share}
        24{Compare-Files}
        25{run-get-filehash}
        26{View-Certificates;reloadmenu}
        27{Active-ScrnSav;reloadmenu}
        default
        {
            cls
            $MenuSelect = $null
            $global:NC = " "
            $global:p1 = " "
            $global:p2 = " "
            $global:p3 = " "
            $Script:MNum = 0
            $Global:ExitSession=$true;clear-host;break
        }

    }    
}
function reloadmenu()
{
        Read-Host "Hit Enter to continue"
        $menuselect = $null
        $Global:pclist = $null
        menu
}
function reloadadmenu()
{
        Read-Host "Hit Enter to continue"
        $admenuselect = $null
        $Global:pclist = $null
        admenu
}
function ReloadNetMenu()
{
        Read-Host "Hit Enter to continue"
        $NetMenuSelect = $null
        $Global:pclist = $null
        NetMenu
}
function EmptyGroups ()
{
    Write-Verbose "Getting AD Groups that have no members"
    Get-ADGroup -Filter * | where { -Not ($_ | Get-ADGroupMember)} |select name
    
        # reloadadmenu
}

function OldDCPCs()
{
# View AD Computer last logon times
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
    <# Where-object Name -eq LEC12 | ft #> 
    Where-object Enabled -eq $True | ft
    # Add the following line if you want the result saved into a .CSV file
    #| Export-Csv OLD_Computer.csv -NoTypeInformation

       # reloadadmenu

}
function GroupsOfMembers()
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

function Associate-UserAndPC
{
    Write-Output "The following are the workstations found for each user"
    $AUnPResult = $AUnP|Where {$_.Name}|sort Name -Unique
    $OBClass1 = $AUnP.OBClass
    # "OBClass is $OBClass"
    Start-sleep 1
    $UsrPCAssn = foreach ($line in $AUnPResult)
    {
        $LName = ($line.Name)
        If ($ObClass1 -eq "User")
        {
            $A1 = (get-adcomputer -Properties * -filter 'description -like $LName' -ErrorAction SilentlyContinue |Where {$_.enabled -eq $true}|select Name,Description -ErrorAction SilentlyContinue)
            $A1SAN = $A1.SamAccountName
            $A1Name = $A1.Name
            Write-Host "$A1Name`t" -NoNewline
            Write-Host "$Lname`t" -NoNewline
            Write-Host "$A1SAN`t"
        }
        Elseif ($ObClass1 -eq "Computer")
        {
            $A1 = (Get-aduser -Filter *  -ErrorAction SilentlyContinue |where {$_.Name -like $line.Description}|select SamAccountName,Name -ErrorAction SilentlyContinue)
            $A1SAN = $A1.SamAccountName
            $A1Name = $A1.Name
            Write-Host "$Lname`t" -NoNewline
            Write-Host "$A1SAN`t`t" -NoNewline
            Write-Host "$A1Name`t"

            # "$LName `t $A1"
        }
    }
    $UsrPCAssn |ft -autosize
    $AUnP=$null
}

function GroupMembership()
{
param(
[Parameter(Mandatory=$False)]
$ADG = ((Read-Host "Enter name of group of which you wish to see members. (ex. Domain Admins)").split(",") | %{$_.trim()})
)
    $adgm = $null
    Write-Warning "`nRunning:`tGet-ADGroupMember -Identity $ADG"
    $adgm = Get-ADGroupMember -Identity $ADG -Recursive |sort ObjectClass,Name,Description # |seleqct * # Name,SamAccountname,distinguishedName,ObjectClass
    $AUnP = foreach ($adgmRec in $adgm)
    {
        <#  # - Just use the -Recursive on get-adGroupMember and this if is not necessary
        if ($adgmRec.ObjectClass -eq "group")
        {
            foreach ($Groupline in $adgmRec) #Continue to drill down into groups as long as one exists
                { # This needs to be expanded / configured to passresults of group membership back to the 
                  # $AUnP variable that eventually passes to the Associate-UserAndPC function 
                  # (Groups don't work with that function, just users)
                            $SAN = ($Groupline |select SamAccountName).SamAccountName
                            GroupMembership -ADG $SAN
                }
        }
        #>
        If ($adgmRec.ObjectClass -eq "user" -or $adgmRec.ObjectClass -eq "computer")
        {
            $adgmpcname = $adgmRec.Name
            $adgmRecObClass = $adgmRec.ObjectClass
            # $adgmpcname
            $ADPCDescription = Get-ADcomputer -Filter 'Name -eq $adgmpcname' -Properties * -ErrorAction SilentlyContinue |select Name,Description
            $SAN = ($adgmRec |select SamAccountName).SamAccountName
            $OBName =($adgmRec |select Name).Name
            $OBDNAme = ($adgmRec |select DistinguishedName).DistinguishedName
            $OBDescription =($ADPCDescription |select Description).Description
            $ObClass = $adgmRecObClass #).Class
            If ($adgmRec.ObjectClass -eq "user")
            {
                Write-Host "($SAN)`t$OBName" -ForegroundColor Green
            }
            If ($adgmRec.ObjectClass -eq "computer")
            {
                Write-Host "`t$OBName`t$OBDescription" -ForegroundColor Green
            }
            $OBProperties = @{'SamAccountName'=$SAN;
                        'Name'=$OBName;
                        'DistinguishedName'=$OBDNAme;
                        'Description' = $OBDescription;
                        'OBClass' = $ObClass}
        }
        New-Object -TypeName PSObject -Prop $OBProperties
        $ADPCDescription = $null
    }
        Associate-UserAndPC
}
function Test-PCUp {
# make $PC a parameter that gets passed to this function
$PC= Read-host "`nEnter a system name to test the connection"
$a = 0
$ProgressPreference
Write-Host "Pinging $PC " -NoNewline
    :DoLoop Do
    {
        Write-Host ". " -NoNewline
        # Set $PP to Powershell's default $ProgressPreference
        $PP = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        $TNC = Test-NetConnection $PC -InformationLevel Quiet |select PingSucceeded
        Start-Sleep -Seconds 3
        # $TNC
        # $TNC.PingSucceeded
        If ($TNC.PingSucceeded -eq $false){$a++}
    }
    until ($a -gt 2)
    Write-Warning "Unable to ping lec151"
    # Set Powershell's default $ProgressPreference back to what it started out as using $PP
    $ProgressPreference = $PP
}
function ADPcArch()
{
    Clear-Host
    Get-PCList
    $CLineNum = $MyInvocation.ScriptLineNumber
    Write-Warning "Current line in Script is: $CLineNum and _Script:PCList is $Global:pclist"
    # Display the command that will be run for each pc listed in Active directory Get OS information from pc list
    Write-Host '$sOS = Get-WmiObject -class Win32_OperatingSystem -Impersonation Impersonate -Credential $cred -computername $pc'
    foreach($Global:pcline in $Global:pclist)
    {
            Identify-PCName
            <#
            # This if statement could potentially be moved to inside the Get-PCList function, but it might need to be modified to do so
            if ($pcline.name -eq $null)
            {
                $pc = $pcline
            }
            else
            {
                $pc = $pcline.name
            }
            #>
        If (Test-Connection $pc -Count 1 -Quiet)
        {
        # Get OS information from pc list
            $sOS = Get-WmiObject -class Win32_OperatingSystem -Impersonation Impersonate -Credential $cred -computername $pc
            foreach($sProperty in $sOS)
            {
                $sProperty |select @{Name="System Name"; e=$sProperty.PSComputerName},PSComputerName,@{Name="OS"; Expression={__Class}},__Class, Caption, Description, @{Name="Architecture"; Expression={OSArchitecture}},OSArchitecture
                Write-Output "`nPC: $sProperty.CSNAme" #PSComputerName"
                write-host "`nOS: "$sProperty.Caption
                write-host "Architecture: "$sProperty.OSArchitecture
                write-host "Registered to: "$sProperty.Description
                write-host "Version #: " $sProperty.ServicePackMajorVersion
                Write-Output "`n$t4$d4$d4"
            }
        }
    }
}

# Get and display a variety of printer information for listed system
function Get-PrinterInfo()
{
param(
[Parameter(Mandatory=$False)]
[string[]] $pclist = ((Read-Host "Enter a comma separated list of computernames").split(",") | %{$_.trim()})
)
    foreach($pc in $pclist)
    {
    # Get and display printer information for listed system
        Write-Host 'Running: `nGet-WmiObject -Class win32_printer -Credential $cred -ComputerName $pc | select name, drivername, portname, shared, sharename |ft -AutoSize'
        # $prin = Get-WmiObject -Class win32_printer -Credential $cred -ComputerName $pc | select name, drivername, portname, shared, sharename |ft -AutoSize
        $Prin = Get-WmiObject -Class Win32_printer -Credential $cred -ComputerName $pc | select Name, drivername, portname, shared, Location, PrinterState, PrinterStatus, ShareName, SystemName, DetectedErrorState, ExtendedDetectedErrorState
        foreach
        ($PrinItem in $Prin) 
        {
        switch($PrinItem.DetectedErrorState)
            { 
                0{$Status="0/On line"}
                1{$Status="1/Paused"}
                2 {$Status="2/Pending Deletion"}
                3{$Status="3/Error"}
                4{$Status="4/Paper Jam"}
                5{$Status="5/Paper Out"}
                6{$Status="6/Manual Feed"}
                7{$Status="7/Paper Problem"}
                8{$Status="8/Offline"}
                9{$Status="9/Service Requested"}
                256{$Status="256/IO Active"}
                512{$Status="512/Busy"}
                1024{$Status="1024/Printing"}
                2048{$Status="2048/Output Bin Full"}
                4096{$Status="4096/Not Available"}
                8192{$Status="8192/Waiting"}
                6384{$Status="6384/Processing"}
                32768{$Status="32768/Initializing"}
                65536{$Status="65536/Warming Up"}
                131072{$Status="131072/Toner Low"}
                262144{$Status="262144/No Toner"}
                524288{$Status="524288/Page Punt"}
                1048576{$Status="1048576/User Intervention"}
                2097152{$Status="2097152/Out of Memory"}
                4194304{$Status="4194304/Door Open"}
                8388608{$Status="8388608/Server Unknown"}
                16777216{$Status="16777216/Power Save"}
                default{$Status="UNKNOWN"}
            }
        ($PrinItem.DetectedErrorState) = $Status
        }
        foreach ($PrinItem in $Prin) 
        {
            switch($PrinItem.PrinterState)
            { 
                0{$Status="0/On line"}
            }
        ($PrinItem.PrinterState) = $Status
        }
        $prin |fl #t -AutoSize -Wrap
    # Testing another Printer Info command
        Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Network=True" | select Sharename,name | export-csv -path "$PSScriptRoot\Win32_Printer.csv" -NoTypeInformation
        Write-warning "File saved to: $PSScriptRoot\Win32_Printer.csv"
    # Get and display Printer Port information for listed system
        Write-Host 'get-printerport -computer $pc |select ComputerName,Name,Description,PrinterHostAddress |ft -A'
        $prin2 = get-printerport -computer $pc |select ComputerName,Name,Description,PrinterHostAddress |ft -A
        # get-printerport -computer $pc |select Caption, CommunicationStatus,ComputerName,Description, DetailedStatus, ElementName, HealthState, InstallDate, InstanceID, Name, OperatingStatus, PortMonitor, PrimaryStatus, PSComputerName, Status, StatusDescriptions |fl
        $prin2
    # Get and display Printer Driver information for listed system
        Write-Host 'get-printerDriver -computer $pc |select * |ft -A # ComputerName,Name,Description,PrinterHostAddress,OEMUrl,Monitor,Path,provider |ft -A'
        $prin3 = get-printerDriver -computer $pc |select ComputerName,Name|Sort Name|ft -A # Manufacturer
        $prin3
    }
        if(!$NoMenu)
        {
            reloadmenu
        }
}

function Get-Uptime
 {
 param(
        [string[]]$pc = "$env:COMPUTERNAME"
       )
    # Get the date and time of the system's last boot
        $WH = "gwmi win32_operatingsystem -ComputerName $pc |select csname,@{Label='LastBootUpTime';Expression={$_.ConverttoDateTime($_.lastbootuptime)}}"
        Write-Host 'Running: `n$WH'
        gwmi win32_operatingsystem -ComputerName $pc |select csname,@{Label='LastBootUpTime';Expression={$_.ConverttoDateTime($_.lastbootuptime)}}
	# Uptime in days, hours, and minutes
        $os = Get-wmiobject win32_operatingsystem -ComputerName $pc
	    $uptime = (Get-date) - ($os.ConvertToDateTime($os.lastbootuptime))
	    $Display = "$pc Uptime: " + $Uptime.Days + " days, " + $Uptime.Hours + " hours, " + $Uptime.Minutes + " minutes"
	    Write-Output $Display
}

function pcinfo()
{
    Clear-Host
    # if (!($Global:pclist))
    #{
        Get-PCList
    #}
    
    foreach($Global:pcline in $Global:pclist)
    {
            Identify-PCName
    # Get the Dell Service Tag # from pc list
        get-wmiobject -Credential $cred -computername $Global:pc -class win32_bios |select "$Global:pc", SerialNumber |fl
        Write-Host 'Running: `nget-wmiobject -Credential $cred -computername $Global:pc -class win32_bios |select "$Global:pc", SerialNumber |fl'
    # Get the date and time of the system's last boot
        Get-Uptime $Global:pc
    # Get OS information from pc list
        Write-Host 'Running: `nGet-WmiObject -class Win32_OperatingSystem -Impersonation Impersonate -Credential $cred -computername $Global:pc'
        $sOS = Get-WmiObject -class Win32_OperatingSystem -Impersonation Impersonate -Credential $cred -computername $Global:pc
        foreach($sProperty in $sOS){
            write-host "`nOS: "$sProperty.Caption
            write-host "Architecture: "$sProperty.OSArchitecture
            write-host "Registered to: "$sProperty.Description
            write-host "Version #: " $sProperty.ServicePackMajorVersion
        }
        
    # Get CPU Information for listed system
        $cpuinfo = get-wmiobject win32_processor <# -Impersonation Impersonate #> -Credential $cred -ComputerName $Global:pc |select-object role, name, numberofcores,numberoflogicalprocessors,
        @{Label = "CPU Status" ;Expression={
                    switch ($cpuinfo.cpustatus)
                    {
                        0 {"The status of the CPU is unknown"}
                        1 {"The CPU is Enabled"}
                        2 {"The CPU is Disabled by User via BIOS Setup"}
                        3 {"The CPU is Disabled by BIOS (POST Error)"}
                        4 {"The CPU is Idle"}
                        5 {"The status of the CPU is Reserved"}
                        6 {"The status of the CPU is Reserved"}
                        7 {"The status of the CPU is Other (unk)"}
                    }
         }}
    # @{Label = "CPU Status" ;Expression={$cpuinfo}}
        Write-Host "`nThe following is the information on the CPU`n"
        $cpuinfo


$W32FanItems = get-wmiobject -class "Win32_Fan" -namespace "root\CIMV2" -computername $Global:pc 
Write-Host "System Fan Readings" -ForegroundColor Green
foreach ($W32FanItem in $W32FanItems) { 
    write-host "`tSystem Name: "`t`t`t`t $W32FanItem.SystemName 
    write-host "`tActive Cooling: "`t`t`t $W32FanItem.ActiveCooling 
        Switch ($W32FanItem.Availability)
        {
        1{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Other)"}
        2{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Unknown)"}
        3{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Running/Full Power)"}
        4{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Warning)"}
        5{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (In Test)"}
        6{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Not Applicable)"}
        7{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Power Off)"}
        8{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Off Line)"}
        9{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Off Duty)"}
        10{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Degraded)"}
        11{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Not Installed)"}
        12{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Install Error)"}
        13{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Power Save - Unknown)"}
        14{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Power Save - Low Power Mode)"}
        15{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Power Save - Standby)"}
        16{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Power Cycle)"}
        17{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Power Save - Warning)"}
        18{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Paused)"}
        19{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Not Ready)"}
        20{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Not Configured)"}
        21{write-host "`tAvailability: "`t`t`t`t $W32FanItem.Availability " (Quiesced)"}
        }
    write-host "`tCaption: "`t`t`t`t`t $W32FanItem.Caption 
    write-host "`tConfig Mgr Error Code: "`t $W32FanItem.ConfigManagerErrorCode 
    write-host "`tConfig Mgr User Config: "`t $W32FanItem.ConfigManagerUserConfig 
    write-host "`tCreation Class Name: "`t`t $W32FanItem.CreationClassName 
    write-host "`tDescription: "`t`t`t`t $W32FanItem.Description 
    write-host "`tDesired Speed: "`t`t`t`t $W32FanItem.DesiredSpeed 
    write-host "`tDevice ID: "`t`t`t`t $W32FanItem.DeviceID 
    write-host "`tError Cleared: "`t $W32FanItem.ErrorCleared 
    write-host "`tError Description: "`t $W32FanItem.ErrorDescription 
    write-host "`tInstallation Date: "`t $W32FanItem.InstallDate 
    write-host "`tLast Error Code: "`t $W32FanItem.LastErrorCode 
    write-host "`tName: "`t`t`t`t`t`t $W32FanItem.Name 
    write-host "`tPNP Device ID: "`t`t`t $W32FanItem.PNPDeviceID 
    write-host "`tPower Management Capabilities: "`t $W32FanItem.PowerManagementCapabilities 
    write-host "`tPower Management Supported: "`t $W32FanItem.PowerManagementSupported 
    write-host "`tStatus: "`t`t`t`t`t $W32FanItem.Status 
Switch ($W32FanItem.StatusInfo)
        {
        1{write-host "`tStatus Information: "`t`t $W32FanItem.StatusInfo " (Other)"}
        2{write-host "`tStatus Information: "`t`t $W32FanItem.StatusInfo " (Unknown)"}
        3{write-host "`tStatus Information: "`t`t $W32FanItem.StatusInfo " (Enabled)"}
        4{write-host "`tStatus Information: "`t`t $W32FanItem.StatusInfo " (Disabled)"}
        5{write-host "`tStatus Information: "`t`t $W32FanItem.StatusInfo " (Not Applicable)"}
        }
    write-host "`tSystem Creation Class Name: " $W32FanItem.SystemCreationClassName 
    write-host "`tVariable Speed: "`t`t`t $W32FanItem.VariableSpeed 
    Write-Host "`n"
} 

    # Get and display memory information for listed system
        Write-Host 'Running: `nGet-WmiObject cim_physicalmemory -Impersonation 3 -Credential $cred -ComputerName $Global:pc | select banklabel,capacity,caption,devicelocator,partnumber,speed |ft -AutoSize'
        $mem = Get-WmiObject cim_physicalmemory -Impersonation 3 -Credential $cred -ComputerName $Global:pc | select banklabel,capacity,caption,devicelocator,partnumber,speed |ft -AutoSize
        $mem
    $NoMenu = $true
    Get-PrinterInfo $Global:pc
    $NoMenu = $false
    # Get Disk information for system
        Write-Host '`nRunning: get-wmiobject win32_logicaldisk -Credential $cred -Computername $Global:pc|
        fl  systemname, Name, deviceid, description, Drivetype, mediatype, filesystem, @{n="MB freespace";e={[string]::Format("{0:0,000} MB",$_.FreeSpace/1MB)}}, @{n="MB Size";e={[string]::Format("{0:0,000} MB",$_.Size/1MB)}}, volumedirty, volumename, volumeserialnumber'
        get-wmiobject win32_logicaldisk -Credential $cred -Computername $Global:pc |
            fl  systemname, Name, deviceid, description, Drivetype, mediatype, filesystem, @{n='MB freespace';e={[string]::Format("{0:0,000} MB",$_.FreeSpace/1MB)}}, @{n='MB Size';e={[string]::Format("{0:0,000} MB",$_.Size/1MB)}}, volumedirty, volumename, volumeserialnumber
        Write-Host '`nGet-WmiObject win32_diskdrive -Credential $cred -Computername $Global:pc |where {$_.model -match "SSD"}'
        Get-WmiObject win32_diskdrive -Credential $cred -Computername $Global:pc |where {$_.model -match 'SSD'}
    #Get attached monitor information
        Write-Output "Attached Monitor Information:"
        gwmi WmiMonitorID -Namespace root\wmi -ComputerName $Global:pc | ForEach-Object {($_.UserFriendlyName -notmatch 0 | foreach {[char]$_}) -join ""; ($_.SerialNumberID -notmatch 0 | foreach {[char]$_}) -join ""}
        <# 
        Get-CimInstance -Namespace root\wmi -ClassName wmimonitorid -ComputerName $Global:pc | foreach {
	        New-Object -TypeName psobject -Property @{
                Manufacturer = ($_.ManufacturerName -notmatch '^0$' | foreach {[char]$_}) -join ""
                Name = ($_.UserFriendlyName -notmatch '^0$' | foreach {[char]$_}) -join ""
                Serial = ($_.SerialNumberID -notmatch '^0$' | foreach {[char]$_}) -join ""
            }
        }
        #>
    #Get process from pc
        $logonInfo = get-wmiobject win32_process -computername $Global:pc -Credential $cred| select Name,sessionId,Creationdate  |Where-Object {($_.Name -like "Logon*")} 
        Write-Host "`nRunning: wmi search to see if the logon processes is running...`n"
        $logonInfo
    #Get Network information from pc
        Write-host "Checking for NIC status"
        # Get-NICStatus
        $PCNICStatus = Get-NICStatus
        $PCNICStatus |ft
        # Write-host "Checking for NIC status"    
        # Get-WmiObject -Class win32_networkadapterconfiguration -ComputerName $Global:pc -Credential $cred|where {$_.IPAddress -ne $null}|select PSComputerName,MACAddress,IPSubnet,DNSServerSearchOrder, @{n="IP Address";e={($_ | select -ExpandProperty IPAddress)-join ','}},DHCPEnabled|ft
    }
        $Global:pclist=$null    
}
#Get Network information from pc
function Get-NICStatus()
{
    Get-WmiObject -Class win32_networkadapterconfiguration -ComputerName $Global:pc -Credential $cred|where {$_.IPAddress -ne $null}|select PSComputerName,MACAddress,IPSubnet,DNSServerSearchOrder, @{n="IP Address";e={($_ | select -ExpandProperty IPAddress)-join ','}},DHCPEnabled # |ft #out-grid-view
}
function NoDomPCinfo()
{
    Clear-Host
    Get-PCList
    #$pclist = Read-Host "Enter a computername to see details"
    foreach($Global:pcline in $Global:pclist)
    {
            Identify-PCName
        $DomName = Read-Host "Enter the domain name"
        $sOS = Get-WmiObject -class Win32_OperatingSystem -Impersonation 3 -Credential "$DomName\administrator" -computername $pc
        # Get the Dell Service Tag # from pc list
            get-wmiobject  -Impersonation 3 -Credential "$DomName\administrator" -computername $pc -class win32_bios <#-erroraction SilentlyContinue#> |select "$pc", SerialNumber |fl
        foreach($sProperty in $sOS){
        # Get OS information from pc list
            write-host "`nOS: "$sProperty.Caption
            write-host "Architecture: "$sProperty.OSArchitecture
            write-host "Registered to: "$sProperty.Description
            write-host "Version #: " $sProperty.ServicePackMajorVersion
        }
        $mem = Get-WmiObject cim_physicalmemory -Impersonation 3 -Credential "$DomName\administrator" -ComputerName $pc | select banklabel,capacity,caption,devicelocator,partnumber,speed |ft -AutoSize    
        $mem
# This is similar to the following from a DOS cli - wmic:root\cli>/node:"10.127.10.79" /user:workgroup\administrator memorychip get /all /format:list
    }
}


function ListDomPCs()
{
    cls
    Write-warning 'Running: Get-ADComputer -filter ''objectclass -eq "Computer"'' | Where {$_.enabled -eq $true}| select enabled,name,@{n="dnshostname";e={($_ |select -ExpandProperty dnshostname)}},distinguishedname | ft -autosize'
    Get-ADComputer -filter 'objectclass -eq "Computer"' | Where {$_.enabled -eq $true}| select enabled,name,@{n="dnshostname";e={($_ |select -ExpandProperty dnshostname)}},distinguishedname | ft -autosize
}
# Where Users's first and last names are in the description of a Computer object
# this will list the User and the associated computer
function Find-UserPCs
{
    [string[]] $Names = ((Read-Host "Enter a comma separated list of all or part of People's first or last names").split(",") | %{$_.trim()})
    foreach ($LName in $Names) 
    {
        get-adcomputer -Filter * -Properties *|where {$_.enabled -eq $true -and $_.Description -like "*$LName*"}|ft Description,Name
    }
}

function UsrOnPC()
{
    Clear-Host
    Get-PCList
        $PCUserList= foreach ($Global:pcline in $Global:pclist)
        {
            Identify-PCName
		    #Get explorer.exe processes
            $PCDesc = get-adcomputer $Global:PC -Properties *|select Description
		    $proc = gwmi win32_process -computer $Global:PC -Filter "Name = 'explorer.exe'" -Credential $cred -ErrorAction SilentlyContinue
            #Search collection of processes for username
                if ($proc -ne $null)
                {
		           ForEach ($p in $proc) 
                    {
	    	            $Script:PCUser = ($p.GetOwner()).User
                        $Script:PCUserEmail = (Get-aduser -Identity $Script:PCUser -Properties * |select emailaddress)
                        # PowerShell equivalent of "wmic useraccount where name=$Script:PCUser get sid"
                        $PCUserSID = get-wmiobject -class Win32_useraccount |where {$_.name -like "$Script:PCUser"}|select SID
                        $PCUserInfo=[pscustomobject]@{
                            Hostname = $Global:PC #.GetValue(0)
                            SamAccountName = $Script:PCUser
                            emailaddress = $Script:PCUserEmail.emailaddress
                            UserSID = $PCUserSID.SID
                            PCDescription = $PCDesc.Description
                        }
                        $pcUserInfo |select HostName,SamAccountName,EmailAddress,PCDescription,UserSID
                    }
                }
                else
                {
                    Write-host "Unable to get process for $Global:PC"
                }
		}
        $PCUserList |ft
        if ($CalledFromFuction -eq $True)
            {
                return                
            }
}

function Manage-ADUserFiles()
{
    $Choose2CopyFile = $null
        $p1 = "Type the word - ";
        $p2 = "copy ";
        $p3 = " - if you wish to copy a file to each user's desktop";
        $NC = "Yellow";chcolor $p1 $p2 $p3 $NC;$Choose2CopyFile = Read-Host "`t`t"
    $ReplaceLink = $null
        $p1 = "Type ";
        $p2 = "dlink or plink ";
        $p3 = "to fix desktop links - OR - Links in a path";
        $NC = "Yellow";chcolor $p1 $p2 $p3 $NC;$ReplaceLink = Read-Host "`t`t"
    $Opt2 = $null
        $p1 = "Type the word - ";
        $p2 = "map ";
        $p3 = "- if you wish to see mapped drives on each users's system";
        $NC = "Yellow";chcolor $p1 $p2 $p3 $NC;$Opt2 = Read-Host "`t`t"
    $CollectFile = $null
        $p1 = "Type the word - ";
        $p2 = "collect ";
        $p3 = " - if you wish to collect the contents of a pre-determined file from each system's public download folder";
        $NC = "Yellow";chcolor $p1 $p2 $p3 $NC;$CollectFile = Read-Host "`t`t"
    # Name of old file to replace
    # FNExample1: Shared Files (LEC).lnk
    # FNExample2:  WorkStudio INFOCENTER.url
    # Set the value of $NPconfirm to "N" as default for "Shall I continue to fix all links at $Computer Y/N?"
    $global:NPconfirm = "N"
    if ($Choose2CopyFile -eq "Copy")
        {
        $NamedFiles = Read-Host "Enter a file name to remove. Ex: SharedFiles (S).lnk or *"
        # Location of new file to copy
        # Example: \\lec.lan\dfs\Shared Files\NetAdmin\New PC Setup\New Profile
        # Folder where icons are pinned to the TaskBar
        #  ls "C:\Users\jasoncrockett\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
        $NAmedFNLocation = Read-Host "`nEnter the path to the new file to copy (exclude trailing \);
            `nExample: \\lec.lan\dfs\sharedfiles\NetAdmin\New PC Setup\New Profile"
        # Name of new file
        # I moved this line into the $Choose2CopyFile if statement. If it is needed by something else too, it 
        # will need to be moved back out or copied to another if statement
        $ReplaceWithFN = Read-Host "Enter the new file name to copy `nExample: SharedFiles (L).lnk" # "SharedFiles (L).lnk"
        }
    if ($CollectFile)
        {
            $NamedFiles = Read-Host "Enter a file name to read. Ex: printers-"
            # need to de-personalize this -OR- get the .\ or script folder to work
            $OutFile = "$PSScriptRoot\$NamedFiles"+"Results.csv"
            $AppendOverwrite = Read-Host "Type O to overwrite $OutFile`nAll other options will Append"
            Write-Warning "Collect Function Results will be saved in: $OutFile"
        }


    # Concatonate new filename and location
    $copyFile = $NAmedFNLocation + "\" + $ReplaceWithFN

    $ADorCSL = $null
    $p1 = "Type AD to use active computers listed in Active Directory or ";
    $p2 = "custom ";
    $p3 = "to type your own comma separated list";
    $NC = "Yellow";chcolor $p1 $p2 $p3 $NC;$ADorCSL = Read-Host "`t`t"

        
        # This section commented out in-lieu-of using the Get-PCList function
        if ($ADorCSL -eq "AD")
            {
                # Report the Username logged into any active computer listed in Active Directory
                Get-PCList
                # $Global:pclist = Get-ADComputer -Filter 'Enabled -eq "True"'
            }
        elseif ($ADorCSL -eq "custom")
            {
                # $pclist = Read-Host "Type in a comma separated list of computers"
                Get-PCList
                # [array]$computers = $Global:pclist.split(",") |%{$_.trim()}
            }
        Else
            {
                $opt1 = $null
                $p1 = "That doesn't match the options type ";
                $p2 = "yes ";
                $p3 = "to run against this system only";
                $NC = "Yellow";chcolor $p1 $p2 $p3 $NC;$opt1 = Read-Host "`t`t"
        
            If ($opt1 -eq "yes")
                {
                $proc = gwmi win32_process -Filter "Name = 'explorer.exe'"
                $pscn = $proc.pscomputername
                $pgo = $proc.GetOwner().user
                "pscn is $pscn"
                iterate-Profiles
                }
            Else
                {
                        $ErrorActionPreference = "continue"
                }
             }

    Write-Host "Logged In Computer`tUser" -ForegroundColor Yellow
    Write-Host "IPAddress `tMachine Name `tUser Logged on"
    #foreach ($comp in $computers)
    foreach ($PCline in $PCList)
        {
            # This if statement could potentially be moved to inside the Get-PCList function, but it might need to be modified to do so
            if ($pcline.name -eq $null)
            {
                $pc = $pcline
            }
            else
            {
                $pc = $pcline.name
            }
<#
        if ($ADorCSL -eq "AD")
            {
                $Computer = $comp.Name
            }
	    else
            {
                $Computer = $comp
            }
#>
    $CompIP = $null
    $CompName = $null
    $PingOutCSV = "C:\Users\JasonCrockett\Documents\WindowsPowerShell\"+$Global:CDate+"PingOut.csv"
    $ErrorActionPreference = "silentlycontinue"
    $CompIP = ([system.net.dns]::GetHostAddresses($pc)).IPAddressToString |where  {$_ -like "*.*.*.*"}
    # Was: $CompIP = ([system.net.dns]::GetHostAddresses($Computer)).IPAddressToString # but it was pulling IPv6 & IPv4
    # Write-host "CompIP $CompIP pre-if" -ForegroundColor Red
    If ($CompIP)
        {
            #1. Get ping results into object
            $ping = test-connection $pc -Count 1 -Quiet -ErrorAction SilentlyContinue
            $PingRslt = New-Object -type PSObject -Property ([Ordered]@{
                Hostname = $pc
                Online = [bool]($ping)
                PingDate = $Global:CDate
                })
            $PingRslt|select * | Export-Csv $PingOutCSV -NoTypeInformation -Append
            # Verify Computername, hostname, with IP
            # "$Computer IP is $CompIP"
            $CompName = [System.Net.Dns]::GetHostByAddress($CompIP).HostName
            $CN_GHBA = if($pc){$pc}Else{"nothing"}
            # Write-host "Computer is $Computer , CompName is $CompName" -BackgroundColor DarkYellow
            If ($pc -notlike "*$pc*")
                {
                    Write-host "Testing $pc, but reverse lookup returns: $CN_GHBA" -ForegroundColor Red
                }
        }
        Else
        {
            $ping = test-connection $pc -Count 1 -Quiet -ErrorAction SilentlyContinue
            $PingRslt = New-Object -type PSObject -Property ([Ordered]@{
                Hostname = $pc
                Online = [bool]($ping)
                PingDate = $Global:CDate
                })
            $PingRslt|select * | Export-Csv $PingOutCSV -NoTypeInformation -Append
            $ping = $false
            $CompIP = "___.___.___.___"
            # Write-host "No IP $CompIP found for: for $Computer" -foreground Red
        }
    # $ping = test-connection $computer -Count 1 -Quiet -ErrorAction SilentlyContinue

  	if($ping -eq $true)
        {
        $pcinfo = @()
		#Get explorer.exe processes
		$proc = gwmi win32_process -computer $pc -Filter "Name = 'explorer.exe'"

		#Search collection of processes for username
		ForEach ($p in $proc) 
            {
	    	    # $usr = ($p.GetOwner()).User
                $pscn = $p.pscomputername
                $pgo = $p.GetOwner().user
                Write-Output "$CompIP`t$pscn`t$pgo`t" |ft
            
                if ($Choose2CopyFile -eq "copy" -or $ReplaceLink -eq "dlink" -or $ReplaceLink -eq "plink" -or $CollectFile -eq "collect") # If the user previously confirmed he wants to copy by typing the word "copy" then run the following function
                    {
                        iterate-Profiles
                    }
                if ($Opt2 -eq "map") # If the user previously confirmed he wants to see mapped drives by typing the word "map" then run the following function
                    {
                        List-MappedDrives
                    }            
            # ***********Return************
           
		    }
        }
    else
        {
            Write-Host "$CompIP`t$pc`tUNK/NONE" -ForegroundColor Red
        }
    }
        $ErrorActionPreference = "continue"
    if ($CollectFile){Write-Warning "Collect Function Results has been saved in: $OutFile"}
}

function alladusers()
{
    $ADUserList = get-aduser -filter * -Properties Samaccountname, ObjectGUID, passwordlastset, Name, EmailAddress,DisplayName,LastLogonDate,Created,Mail,Enabled |where {$_.enabled -eq $true}<#{$_.ObjectGUID -like "*"}#> |select Samaccountname, ObjectGUID, passwordlastset, Name, EmailAddress,DisplayName,LastLogonDate,Created,Mail,Enabled |sort SamAccountName |Out-GridView
    $ADUserList
}
function FilterADUsers()
{
    $Fltr = Read-Host "Type All or part of a users name to filter the list. `nUse * for all users and as a wild-card at the beginning or end of your search criteria."
    $ADUserList = get-aduser -filter 'Name -like $Fltr' -Properties Samaccountname, ObjectGUID, passwordlastset, Name, EmailAddress,DisplayName,LastLogonDate,Created,Mail,Enabled |where {$_.enabled -eq $true} |select Samaccountname, ObjectGUID, passwordlastset, Name, EmailAddress,DisplayName,LastLogonDate,Created,Mail,Enabled |sort SamAccountName, ObjectGUID # |Out-GridView
    $ADUserList
    # $AUnP = $ADUserList|select Name #,DisplayName
    $adgm = $ADUserList|select Name,DisplayName
    <#
    $adgm |gm
    $adgm.Name
    $adgm.DisplayName
    # $AUnP = "$adgm.DisplayName"
    $AUnP.Name = $ADUserList.DisplayName
    Associate-UserAndPC
    #>
}

function SysEvents()
{ 
    Get-PCList
    #$pclist = Read-host "Enter Computer list"
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
function New-PSRemote()
{
    Clear-Host
    # Connect to Server Remotely
    # Still needs some work for correct authorization / force credentials again to connect to remote machines
        $userCredential = Get-Credential
        $PC = Read-Host "Enter the computername to which you want to open a PowerShell Remote Session"

    # Connect to Exchange
    # Enter-PSSession -ComputerName exchange13 -Authentication Kerberos -Credential $userCredential
    # Add-PSSnapin *exchange*

        Enter-PSSession -ComputerName $PC -Authentication Kerberos -Credential $userCredential
        Clear-host
break
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

function Active-ScrnSav()
{
    Clear-Host
    Get-PCList
    # [array]$pclist=(Read-host "Enter a comma separated list of computernames or IP addresses to see details").split(",") | %{$_.trim()}
    foreach ($Global:pcline in $Global:pclist)
    {
        Identify-PCName
        # Search for the screen saver process on a pc as a way of seeing if a user is active
         $scrpresent = Get-WmiObject Win32_process -ComputerName $Global:pc -credential $cred |Where-Object {($_.processname -like "scrn*" )} |Select-Object processname,@{n='CreatedTime';e={$_.ConvertToDateTime($_.CreationDate)}}
          
         if ($scrpresent.processname -like "scrn*") 
         {
            Write-Host "`nThe screen saver has been active on $Global:pc since "$scrpresent.createdtime "`n"
         }
         else
         {
            Write-Host "The screen saver is not active on $Global:pc"
         }

         $getwmiobject = Get-WmiObject -class Win32_computersystem -computername $Global:pc -Credential $Cred
         $Username = $Getwmiobject.username

         if($UserName -eq $NULL) {
         Write-host "`nThere is no Current user logged onto $Global:pc`n"
         }
            else {write-host "$Username is logged in`n"
         }
    }
$Global:pclist=$null
}
function Enable-RDP()
{
#   $Cred = Get-Credential
    Clear-Host
    $pc = Read-Host "Enter PC to enable RDP on"
    (Get-WmiObject Win32_TerminalServiceSetting -ComputerName $pc -Credential $Cred -Namespace root\cimv2\TerminalServices).SetAllowTsConnections(1,1)|Out-Null
    # (get-wmiobject -Class Win32_terminalServiceSetting -Namespace root\CIMV2\TerminalServices -ComputerName $pc -Authentication 6).SetAllowTSConnections(1,1)
    (Get-WmiObject -class "Win32_TSGeneralSetting" -ComputerName $pc -Credential $Cred -Namespace root\cimv2\TerminalServices -Filter "TerminalName='RDP-tcp'").SetUserAuthenticationRequired(0)|Out-Null
    Write-Output "You may still need to modify the `"Remote Desktop Users`" group members if you will not be logging in as Administrator"
    # DOS cli:> net localgroup "Remote Desktop Users" /add "domain\Username"

    Start-Sleep -Seconds 5
reloadmenu
}
function ADNonExpiring()
{
    # adapted from Netwrix
    Search-ADAccount -PasswordNeverExpires |Select-Object SAMAccountName, ObjectClass, PasswordNeverExpires, Name |sort SAMAccountName |ft # | Export-Csv c:\temp\users_password_expiration_false.csv
    <#
        # This results in the inverse of the above and prompts for a change from PasswordNeverExpires $False to $True
        $ADUsers = get-aduser -Filter * -Properties * |where {$_.enabled -eq $true -and $_.passwordNeverExpires -eq $false}
        foreach ($ADUser in $ADUsers)
        {
            Write-host $ADUser
            # Set-ADUser -identity $ADUser.SamAccountName -PasswordNeverExpires $true -Confirm
        }
    #>

}

function Get-ADPrivGrpMbrshp()
{
        Invoke-Expression ".\Get-ADPrivGrpMbrshp.ps1"
        $InActiveUsersFile = ".\" + $Global:CDate +"InactiveUsers.html"
        Write-Warning "`nSending a list users with old passwords to $InActiveUsersFile`n"
        get-aduser -Filter * -prop PasswordLastSet | Where {$_.PasswordLastSet -eq $null -or $_.PasswordLastSet -lt (Get-Date).AddDays(-365)}|Where {$_.Enabled -eq $true}|Sort PasswordLastSet |ConvertTo-Html -PreContent 'Possibly-inactive users' -Prop SamAccountName,Name,PasswordLastSet,Enabled|Out-File $InActiveUsersFile # Change Get-ADUsers to Get-ADComputers to look for computers not users
}

function ID-Java()
{
    Clear-Host
    Get-PCList
    # [array]$pclist=(Read-host "Enter a comma separated list of computernames or IP addresses to see details").split(",") | %{$_.trim()}
    #$pclist = Read-Host "Enter a computername to see details"
    #$pclist = get-content pclist.csv
    foreach($Global:pcline in $Global:pclist)
    {
        Identify-PCName
        $tcpc1 = test-connection -computername $Global:pc -Count 2 -Quiet #-ErrorAction SilentlyContinue
        if ($tcpc1 -eq $true)
        {
#        if ($tcpc1.PrimaryAddressResolutionStatus -eq 0) {
            #use WMI to identify Java versions on pc
            Write-Host "`nGet-WmiObject -Class win32_product -Filter "name like 'Java%%'" -ComputerName $Global:pc -Credential $cred |select -expand version |ft"
            Write-Host "`nDetermining (slowly) what versions of Java are installed on "$Global:pc" :`n"
            # Get-WmiObject -Class win32_product -Filter "name like 'Java%%'" -ComputerName $pc -Credential $cred |select PSComputerName, Name, version, __Path |ft
            Get-WmiObject -Class win32_installedwin32program -Filter "name like 'Java%%'" -ComputerName $Global:pc -Credential $cred |select PSComputerName, Name, version, __Path |ft
            <# Get Java Uninstall String
            $javinstprop = get-itemproperty "HKLM:\SOFTWARE\microsoft\windows\currentversion\Installer\UserData\S-1-5-18\Products\4EA42A62D9304AC4784BF2381208370F\InstallProperties"
            $javinstprop |Select-Object uninstallstring
            $javinstprop = ""
            #>
        }
        else 
        {
        write-host "`n"$Global:pc " is not available right now.`n"
        }
    }
    
        reloadmenu    
}

Function Get-ProcessStatus()   # not ready yet!!!
    {
        $pc = Read-host "`nWhat is the computername or IP address? "
        $ReadProc = Read-Host "`nWhat is the ID Name or Number of the process you wish to get information about? "
    if ($readproc -match "^[0-9]+$")
        {
            $gwmio32p = Get-WmiObject win32_process -ComputerName $pc |Where-Object {$_.ProcessID -eq $ReadProc} |Select-Object Path
            # param($id=$pid)
            $gwmio32p
            Get-WmiObject -Class win32_process -Filter "ParentprocessId=$id"

        }
    else
        {
            $gwmio32p = Get-WmiObject win32_process -ComputerName $pc -Credential $cred| Where-Object {$_.ProcessName -like "*$ReadProc*"}
            $gwmio32p  | Select-Object PSComputerName, ProcessID, ProcessName, Path |ft
            # param($id=$pid)
            $gwmio32p
            Get-WmiObject -Class win32_process -Filter "ParentprocessId=$id"
            # $gwmichild = Get-WmiObject win32_process -ComputerName $pc -Credential $cred| Where-Object {$_.ProcessName -eq $ReadProc} -filter "ParentProcessID=$ID"

        }
 
        reloadmenu    
}



function PSWindowsUpdate()
{
    #https://gallery.technet.microsoft.com/scriptcenter/2d191bcd-3308-4edd-9de2-88dff796b0bc#content
    #available functions are listed at the above website.
}

function updthlp()
{
    clear
    write-host "Running `"update-help -Force`" to update PowerShell Help"
    update-help -Force
    reloadmenu
}

function SearchExchMbox()
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
function test-connections()
{
    Clear-Host
    $PP = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    # Get-PCList
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    $PCCnt = ($Global:PCList).Count
    Write-Warning "Running a PING test on $PCCnt Systems"
    $ObjectOut =""
    $ObjectOut = foreach ($Global:pcline in $Global:pclist)
        
        {
            $TCAdd=""
            $TCv4Add=""
            # This if statement could potentially be moved to inside the Get-PCList function, but it might need to be modified to do so
            Identify-PCName
            # Similar code using test-netconnection except in a script a pop-up UI stayed open making it useless
            # {$TC = Test-netconnection $Global:PC -InformationLevel Quiet;$TC |Select ComputerName,RemoteAddress,PingSucceeded}
            $TC = Test-connection $Global:PC -Count 1 -ErrorAction SilentlyContinue -ErrorVariable TCError 
            if ($TC.StatusCode -eq 0)
            {
                Write-host "." -NoNewline -ForegroundColor Green
                $OBProperties = @{'pscomputername'=$Global:PC;
                    # 'Address'=$TC.Address;
                    'IPV4Address'=$TC.IPV4Address;
                    'StatusCode'=$TC.StatusCode;
                    'ResponseTime'=$TC.ResponseTime;
                    'Result'="$Global:PC : Success"}
                New-Object -TypeName PSObject -Prop $OBProperties
            }
            Else
            {
                Write-host "." -NoNewline -ForegroundColor Red
                $TCAdd = $TC.Address
                $TCv4Add = $TC.IPV4Address
                $TCSC = $TC.StatusCode
                foreach ($TCErrorLine in $TCError)
                {
                 $OBProperties = @{'pscomputername'=$Global:PC;
                    # 'Address'=$TCAdd;
                    'IPV4Address'=$TC.IPV4Address;
                    'StatusCode'=$TC.StatusCode;
                    'ResponseTime'=$TC.ResponseTime;
                    'Result'=$TCErrorLine}
                    New-Object -TypeName PSObject -Prop $OBProperties
                }
            }
        }
    $ObjectOut|Sort -Descending ResponseTime,Result |FT PSComputerName,IPV4Address,ResponseTime,StatusCode,Result
    $Global:PCList = $null
    # Set Powershell's default $ProgressPreference back to what it started out as using $PP
    $ProgressPreference = $PP
}
function test-ports()
{
    Clear-Host
    $pc = Read-Host "Enter a pc name or list of pcs "
    $port = Read-Host "`nEnter a port, CSV port list, or list..range "
    Invoke-Expression ".\Test-Ports.ps1 -computer $pc -port @($port) | ft"
    
    ReloadNetMenu
}

function get-apps()
{
    Clear-Host
    $ObjectOut = $null
    Get-PCList
    $prglist = Read-Host "Enter a search term or nothing to find a program or all programs "
    $ObjectOut = foreach ($Global:pcline in $Global:pclist) # $pclist)
    {
            # This if statement could potentially be moved to inside the Get-PCList function, but it might need to be modified to do so
            Identify-PCName
        # $WMIOApps = Get-WmiObject -class Win32_installedwin32program -ComputerName $Global:PC -Credential $cred -ErrorAction SilentlyContinue -ErrorVariable E1 |where {$_.name -like "*$prglist*"} |select pscomputername,Name,Version # |Sort-Object name |ft #|ConvertTo-Csv -NoTypeInformation |select -Skip 1 |Out-File -Append ".\ProgsInPC.csv"
        $WMIOApps = Get-WmiObject -class Win32_product -ComputerName $Global:PC -Credential $cred -ErrorAction SilentlyContinue -ErrorVariable E1 |where {$_.name -like "*$prglist*"} |select PSComputerName,Name,Version
        $WAC = $WMIOApps.count
        # Check for No records returned from the Get-WmiObject command
        If ($WAC -eq 0)
        {
            # check for an error result from the Get-WmiObject command
            If ($E1)
            {
                $OBPCName = $Global:PC
                $OBName = "Error searching for $prglist"
                $OBVersion = "Error"
                Write-Host "." -ForegroundColor Red -nonewline
                Write-Host "$OBPCName,$OBName,$OBVersion" -ForegroundColor Red

                $OBProperties = @{'SamAccountName'=$OBPCName;
                            'Name'=$OBName;
                            'Version'=$OBVersion}
                New-Object -TypeName PSObject -Prop $OBProperties
            }
            # No error from the Get-WmiObject command, but no result either
            Else
            {
                $OBPCName = $Global:PC
                $OBName = "$prglist not found"
                $OBVersion = "UNK"
                Write-Host "$OBPCName,$OBName,$OBVersion" -ForegroundColor DarkYellow # -nonewline
                $OBProperties = @{'SamAccountName'=$OBPCName;
                'Name'=$OBName;
                'Version'=$OBVersion}
                New-Object -TypeName PSObject -Prop $OBProperties
            }
        }
        foreach ($WMIOApp in $WMIOApps)
        {
            $OBPCName = ($WMIOApp |select pscomputername).PSComputerName
            $OBName = ($WMIOApp |select Name).Name
            $OBVersion = ($WMIOApp |select Version).Version
            # Write-Warning "$OBPCName"
            Write-Host "$OBPCName,$OBName,$OBVersion" -ForegroundColor Green #.$OBPCName"
            $OBProperties = @{'SamAccountName'=$OBPCName;
                'Name'=$OBName;
                'Version'=$OBVersion}
            New-Object -TypeName PSObject -Prop $OBProperties
        }
        #**HERE**
    }

    $ObjectOut|select SamAccountName,Name,Version|Where {$_.Name -ne $null} | Sort SamAccountName,Name |Out-GridView #|FT
        reloadmenu 
}

function Uninstall-apps()
{
    Invoke-Expression ".\Uninstall-Application.ps1"
    reloadmenu
}

function ADUser-Logons()
{
    $nowlessnum =((get-date).AddDays(-30))
    $usecnt = Get-ADUser -Filter {lastlogondate -lt $nowlessnum} |ft
    $usecnt
}
function Get-Svc()
{
    Clear-Host
    $Global:pclist = $null
    Get-PCList
    $Global:pclist
    $SvcName = Read-Host "Enter part of the service name (Ex: CRCPluginServer)"
foreach($Global:pcline in $Global:pclist)
    {
    Identify-PCName
    $WMI_Service = Get-WmiObject win32_service -ComputerName $Global:PC -Credential $cred -ErrorAction SilentlyContinue -ErrorVariable WMI_SErr |select pscomputername,processid,name,state,StartMode,pathname|Where-Object {$_.Name -like "*$SvcName*"} #|ft
    # Need to add a prompt for service type "Automatic", Disabled,System, etc
    # Use the following line or similar to set a service type to Automatic
    # Set-Service -StartupType Automatic -ComputerName "LEC138" -Name RemoteRegistry
    if ($WMI_SErr)
        {
            Write-Warning "$Global:PC $WMI_SErr"
            $WMI_SErr = $null
        }
    if ($WMI_Service)
        {
        $WMI_Service |select PSComputerName,ProcessID,Name,State,StartMode,PathName
        $gsout = Get-Service -name $svcName -ComputerName $Global:PC # |Select-Object DisplayName,MachineName,ServiceName,Status
        if ($gsout)
            {
                $StatusChoice = Read-Host "Type yes to change the status of this service or any other key to continue."
            If ($StatusChoice -eq "yes")
                {
                Write-Host -nonewline "`nEnter the new status for the"$gsout.ServiceName"service`n" -ForegroundColor Green
                Write-Warning "Service may not start unless it is first enabled"
                Write-Host -nonewline " currently"$gsout.Status -ForegroundColor Red 
                Write-Host " on"$gsout.MachineName"." -ForegroundColor Green
                $NewSvcStat = Read-Host "Your options are Stopped, Paused, or Running"
                If ($NewSvcStat -eq "Stopped")
                    {
                        $gsout.stop() # |Set-Service -Status Stopped -ComputerName $Global:PC
                    }
                ElseIf ($NewSvcStat -eq "Running")
                    {
                        if ($gsout.StartType -eq "Disabled")
                            {
                                $SrvStartType = Read-Host "Type the new StartType for ($gsout.name). (Options are Automatic,Boot,Disabled,Manual,System)"
                                $gsout|Set-Service -StartupType $SrvStartType -ErrorAction Inquire
                                Get-WmiObject win32_service -ComputerName $Global:PC -Credential $cred |select pscomputername,processid,name,state,pathname|Where-Object {$_.Name -like "*$SvcName*"} #|ft
                            }
                    
                        $gsout.start() # |Set-Service -Status Running -ComputerName $Global:PC
                    }
                ElseIf ($NewSvcStat -eq "Paused")
                    {
                        $gsout.Pause() # |Set-Service -Status Paused -ComputerName $Global:PC
                    }
                Else
                    {
                        Write-Host "That is not a valid option"
                    }
                }

            }
        }
    }
    $Global:pclist=$null
}
# In Windows 10, the nlasvc service must sometimes be restarted to force the network location to identify itself correctly
# When it does not know where it is, Outlook cannot connect to the exchange server even if the pc has an IP address and can get to the internet

function Restart-NLASvc()
{
    clear-host
    Write-Host "`nRestarting the Network Location Awareness service (nlasvc)" -ForegroundColor Green
    Restart-Service nlasvc -Force
    $getnlas = Get-Service nlasvc,netprofm
    $getnlas
    ReloadNetMenu
}
function Reset-ADUserPW()
{
    Clear-Host
    Write-host "Note that the Change option will provide a list of users and dates their passwords were last set (with or without options)"
    $Choice = Read-Host "Enter List, Unlock, Reset, Change "
    
    if ($Choice -eq "Unlock")
    {
        $ADUser = Read-Host "Enter part of the user's name "
        Get-ADUser -Filter ('name -like "*' + $ADUser + '*"') |Unlock-ADAccount -Confirm
    }    
# Need to verify this option
    Elseif ($Choice -eq "Reset")
    {
        $ADUser = Read-Host "Enter part of the user's name "
        # $ADUserResult = Get-ADUser -Filter ('name -like "*' + $ADUser + '*"')
        $ADUserResult = Get-ADUser -filter 'samaccountname -eq $ADUSer'
        $ADUserSAM = $ADUserResult.SamAccountName
        $newCred = Get-Credential "$ADUserSAM" |Set-ADAccountPassword -NewPassword -Confirm
    }
    Elseif ($Choice -eq "Change")
    {
        Invoke-Expression ".\Set-ADUserPwd.ps1"
        while($global:AnotherPW -eq "Yes")
            {
                Invoke-Expression ".\Set-ADUserPwd.ps1"
                Clear-host
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

# This function basically pings a system to see if it is alive. It logs the data and time and whether it is accessible or not
Function PingTest()
{
    clear-host
    $pc= Read-host "`nEnter a system name to test the connection"
        # [bool]$hideKeysStrokes
        [bool]$hideKeyStrokes = $false #$true
    try
    {
    Write-host "`nPress x to quit" -foregroundcolor yellow
        while(($true) -or ($ky -ne "X"))
        {
            
            "$($ky = $key.key;Write-Output "$ky key As of ")$(Get-Date)$(Write-Output ", $pc ")$(if($(Test-connection -ComputerName $pc -Quiet)){Write-output "is accessible"}else{Write-Output "is inaccessible"})" |ft
            # Test-PCUp
            
            # while(($true) -or ($key.Key -ne "x")){Test-NetConnection lec137 |select ComputerName, PingSucceeded,RemoteAddress}
            $key = [System.Console]::Readkey($true)
            "ky is $ky"
        }
    }
    finally
    {
            "Done:finally:ReloadNetMenu"
    }
    Write-Host "W-H:ReloadNetMenu"
}
Function get-KB()
{
    clear-host
    Get-PCList
    $ghf = $null
    $HotFixErr=$null
    # [array]$ghf_total = $null
    $kb = Read-host "`nEnter a KB number or a partial number with wildcards not including the KB"
    $ObjectOut = foreach($Global:pcline in $Global:pclist)
    {
        Identify-PCName
        get-hotfix -computername $Global:pc |where {$_.HotFixID -like "KB$kb"}|sort InstalledOn |ft
        $ghf = get-hotfix -ComputerName $Global:pc -ErrorAction SilentlyContinue -ErrorVariable HotFixErr |where {$_.Description -like "*Security*"} |sort InstalledOn |select PSComputerName, Description, HotFixID, InstalledBy, InstalledOn -Last 1
        if ($ghf)
        {
            $HFDescription = $ghf.Description
            $HFHotFixID = $ghf.HotFixID
            $HFInstalledBy = $ghf.InstalledBy
            $HFInstalledOn = $ghf.InstalledOn
            $OBProperties = @{'pscomputername'=$Global:PC;
                'Description'=$HFDescription;
                'HotFixID'=$HFHotFixID;
                'InstalledBy'=$HFInstalledBy;
                'InstalledOn'=$HFInstalledOn;
                'HotFixErr'=$HotFixErr}
                New-Object -TypeName PSObject -Prop $OBProperties
        }
    }
    Write-Output "The following are the most recent security updates applied to these systems."
    $ObjectOut|sort InstalledOn,PSComputerName -Descending|select pscomputername,HotFixID,Description,InstalledOn,installedBy,HotFixErr|ft -AutoSize -Wrap
    $HotFixErr=$null
    $Global:pclist = $null
}
Function Get-IP()
{
    Clear-Host
    Write-Host "`nYou may need to ping or communicate with a host prior to running this.`nTry pinging the broadcast address first."
    $macaddr = Read-Host "`nEnter some or all of a MAC address in the format AC-86-...."
    arp -a | select-string $macaddr #|% {$_.ToString().Trim().Split(" ")[0]}
    ReloadNetMenu
}
function GPReport()
{
    Clear-Host
    $PSScriptRoot
        $CalledFromFuction = $True
    $rptpath = "$PSScriptRoot\GPReport.htm"
    UsrOnPC
    $Computer = Read-Host "Type the name of a Server or Workstation to run GPResult against"
    # $Script:PCUser  # User does not update for local user?
    $User = Read-Host "Enter the username of the policy holder"
    Get-GPResultantSetOfPolicy -Computer $Computer -User $User -reporttype html -path $rptpath
    start chrome.exe $rptpath
    $CalledFromFuction = $False
}
function FWStatus()
{
    # !!!TO DO !!! If the service is disabled it must also be enabled before it can be started
    # !!!TO DO !!! Add an option to query whether you want all domain pcs or just a specific one or few
    # If we cannot ping PC could be off or Firewall could be on
    # If we can ping but cannot get profile status RemoteRegistry Service could be off

    Clear-Host
    $RegErr = $false
    $FWprofile = Read-Host "Type firewall profile name to query - default is DomainProfile"
    $FWprofile = "DomainProfile"
    # $PCChoice = Read-Host "Type L to enter your own comma-separated list of PCs to query."
    # $pclist = Read-Host "Enter System name"
    Get-PCList
    <#
    $pclist=Get-ADComputer -Filter * -Properties enabled,name,lastlogondate | Where-object {$_.Enabled -ne $false} |
        Sort-Object lastlogondate -Descending |select-object -Property @{label = "name";expression = {$_.name}},enabled,lastlogondate
    #>
    foreach($pcline in $Global:pclist)
    {
            # This if statement could potentially be moved to inside the Get-PCList function, but it might need to be modified to do so
            if ($pcline.name -eq $null)
            {
                $pc = $pcline
            }
            else
            {
                $pc = $pcline.name
            }
        Write-Warning "`npc is $pc`n"
        $GSRemReg = Get-Service Remoteregistry -ComputerName $pc -ErrorAction SilentlyContinue
        # Saving the status of the RemoteRegistry Service to the $RegStat variable
        $RegStat = $GSRemReg.Status
        if ($RegStat -eq $null)
            {
                Write-host "Unable to get RemoteRegistry Status for $pc" -ForegroundColor Red
            }
        else
            {
                if ($RegStat -eq "stopped")
                    {
                        Write-host "$pc RemoteRegistry Status is $RegStat " -foregroundcolor green
                        # Attempt to start the registry service prior to continuing on (could add a verification of success here)
                        write-host "Trying to start Remote Registry Service on $pc " -ForegroundColor Yellow
                        $GSRemReg |Start-Service |Out-Null
                    }
                 
        # Write-Host "RegStat" -ForegroundColor Green
        try
            {
                $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$pc)
            }
        catch
            {
                $RegErr = $true
            }
        finally
            {
                # Write-Host "RegErr is $RegErr" -ForegroundColor Red
                if ($RegErr -eq $true)
                    {
                        Write-Host "Error connecting to registry on $pc" -ForegroundColor Red
                        #Reset RegErr to false for next catch
                        $RegErr = $false
                    }
                ElseIf($RegErr -eq $null)
                    {
                        Write-Host "Possibly not a Windows (R) system"
                    }
                Else
                    {
                        $firewallEnabled = $reg.OpenSubKey("System\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\$FWprofile").GetValue("EnableFirewall")
                        # $firewallEnabled
                        $FWstatus = [bool]$firewallEnabled
                        # $FWstatus
                    }
            }
            If ($FWstatus -eq $false)
                {
                    
                    Write-Output "$pc $FWprofile firewall is off "
                }
            Elseif($FWstatus -eq $true)
                {
                    Write-host "$pc $FWprofile firewall is on " -ForegroundColor DarkYellow
                }
            Else
                {
                    "System is off or Firewall is blocking communication"
                }

            }
        #Reset FWStatus to true for next query
        $FWstatus = $null # $true
    }
    ReloadNetMenu
}

# Use Reset-IE to reset IE to its defaults and clear browsing history. Requires some user interaction
# Not on menu yet
Function Reset-IE (){
    RunDll32.exe InetCpl.cpl,ResetIEtoDefaults
    RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351
}
# Test input for valid IP Address
# Manually this can be called from Test-ValidIPPrompt
# This function illustrates parameter validation and using RegEx to validate IPv4 IP addresses including broadcast address
function Test-ValidIP()
    {
        param ([ValidatePattern('^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$')]
        [string]$gip
        )
        Write-Output "$gip is a valid IP address. One moment please."
    }
# Prompt for IP Address from User. Combine calls Test-ValidIP to test input for valid IP Address
function Test-ValidIPPrompt
    {
    try {
            $gip = Read-host "`nType in a valid IP address to find information about"
            Test-ValidIP $gip -erroraction "stop"
            ReloadNetMenu
        }
    catch
        {
            $Another = Read-Host "`nInvalid IP Address. Type Y to enter another."
            while($Another -eq "Y")
                {
                    Test-ValidIPPrompt
                }
            ReloadNetMenu
        }
    }
# View NetStat results for a system including the process name and active ports
function View-NetStat
    {
        Invoke-Expression .\Get-NetStat.ps1
        ReloadNetMenu
    }
function Get-ServiceLogons
    {
        $uname = Read-Host "Enter all or part of a valid possible Username that a service may use to start"
        [array]$pclist=(Read-Host "Enter a comma separated list of computernames").split(",") | %{$_.trim()}
        Get-WmiObject -Class Win32_Service -ComputerName $pclist |Sort-Object StartName,PSComputerName |Where-Object {$_.startName -like "*$uname*"}|ft StartName,pscomputername,Name,State,Displayname -AutoSize
        reloadmenu
    }

function Get-RDPStats()
    {
    # Compare with code at: https://superuser.com/questions/587737/powershell-get-a-value-from-query-user-and-not-the-headers-etc#702328
    # Possibly add something similar to take each column to a variable of an object and clean up the output
    Clear-Host
    Write-host "Searching for active connections to LEC servers"
    $InvkCmd=$null
    $Active=$false
    $SN=$null
    $servers = Get-ADComputer -Filter {OperatingSystem -like "*server*" -and Enabled -eq $true}|select Name,SamAccountName,Enabled -ErrorAction SilentlyContinue
    foreach ($server in $servers)
        {
            $SN = $Server.Name        
            if (Test-Connection $sn -Count 1 -ErrorAction SilentlyContinue )
                {
                $InvkCmd = invoke-command {query user /server:$sn 2>$1} -ErrorAction SilentlyContinue
                if ($InvkCmd -like "*Active*")
                    {
                        Write-host "$SN : Active" -NoNewline -ForegroundColor Green
                        ($InvkCmd) -replace '\s{2,}', ','|convertfrom-csv|ft
                        ""
                        $Active=$true
                        start-sleep -Milliseconds 500
                    }
                elseif ($InvkCmd -like "*Disc*")
                    {
                        Write-host "$SN : Disconnected" -NoNewline -ForegroundColor Yellow
                        ($InvkCmd) -replace '\s{2,}', ','|convertfrom-csv |ft
                        start-sleep -Milliseconds 500
                    }
                else
                    {
                        Write-warning "$SN : Error"
                        start-sleep -Milliseconds 500
                    }
                }

            else 
                {
                    Write-Warning "$SN is a server but cannot be reached "
                    start-sleep -Milliseconds 500
                }
        }
    If ($Active -eq $false)
        {
            Write-Host "No active RDP sessions found" -ForegroundColor Red
            start-sleep -Milliseconds 500
        }
    reloadnetmenu
    }

function convert-Numbers()
    {
        clear-host
        [int64]$Num = Read-host "Enter a base 10 number to convert)"
        $BinNum = [convert]::ToString($Num,2)
        "Binary: $BinNum"
        $DecNum = [convert]::ToInt64($BinNum,2)
        "Decimal: $DecNum"
        $hexNum = $Num |Format-Hex
        $hexNum = $hexNum -replace "^0{8}\s+","0x"
        "Hexadecimal: $hexNum"
        # Number with 2 places after the decimal
        $Numd2 = "{0:N2}" -f $Num
        $Numd2
        # Formatted with Currency
        $Cur = "{0:C2}" -f $Num
        $Cur
        # Formatted in Binary (8 places)
        $BinNum = "{0:D8}" -f $BinNum
        $BinNum
        # Formatted as a percent
        $Percent = "{0:P0}" -f ($DecNum/100)
        $Percent
        # Formatted in Traditional 0x Hex
        $Hex = "{0:X0}" -f $Num # $hexNum
        "Hex: $Hex"
        $hexNum
        reloadmenu
    }


function wGet-URL() # Initiate a web-request for a URL and return it in text form
    {
        Clear-Host
        $URL = Read-Host "Please type a URL to investigate"
        $R = Invoke-WebRequest $URL -UseBasicParsing # (wget)
        # Include Images
        $I = $R.Images |ft src # |select -ExpandProperty images |ft
        # Include links
        $L = ($R.Links).href |ft href
        # Show Raw html content
        # $RC = $R.RawContent
        # Show html headers
        # $RH = $R.Headers
        $I+$L |ft
        reloadmenu
    }

function Get-ADSiteSubnet()
{
    get-adreplicationsite -Filter * -Properties * |select Name,siteObjectBL,Subnets,Group |group siteObjectBL |select @{Name='Subnet';Expression={$_.Name.split(',')[0].Trim('{CN=')}},@{Name='Name';Expression={$_.Group.Name}} |ft
}


function Get-hostname()
    {
        Clear-Host
        [array]$IPlc=(Read-Host "Enter a comma separated list of IP Addresses").split(",") | %{$_.trim()}
        Write-Warning "[System.Net.DNS]::GetHostByAddress(IPVariableHere)"
        foreach($IP in $IPlc)
        {
            $GHAddress = [System.Net.DNS]::GetHostByAddress($IP)
            $GHAddress | select * |ft
        }
        ReloadNetMenu
    }

function Get-ADPcStarts()
{
    Clear-Host
    $pclist=Get-ADComputer -Filter 'enabled -eq "true"' -Properties Enabled,Name,IPv4Address |Where {$_.ipv4address -ne $null} |select Name,IPv4Address
    $CSVFileNam = $Global:CDate+"startups.csv"
    [array]$CSVOutObj = "System_Name,Relative_Path,Local_Path,Command,Description,Location,Name,User,UserSID"
    Write-host "File will be exported to $CSVFileNam" -ForegroundColor Cyan
    # Display the command that will be run for each pc listed in Active directory to get StartUp information from PCs in list
    Write-Host '$sStart = gwmi -Class win32_startupcommand -Impersonation Impersonate -Credential $cred -ComputerName $pc |Select-Object PSComputerName,Command,Description,Location, Name,User' # ,UserSID
    foreach($pc in $pclist)
    {
        $pcn = $pc.name
        $pcip = $pc.ipv4address
        # Only attempt to get information from computers that are reachable with a ping
        If (Test-Connection $pcn -Count 1 -Quiet -ErrorAction SilentlyContinue)
        {
            # Get Startup information from pc list
            $sStart = gwmi -Class win32_startupcommand -Impersonation Impersonate -Credential $cred -ComputerName $pcn |Select-Object PSComputerName,Command,Description,Location, Name,User # ,UserSID
            $CSVOutObj += $sStart
            # Use the following line to view last boot time of System
            $OS = Get-CimInstance -ClassName win32_operatingsystem -ComputerName $pcn -ErrorAction SilentlyContinue | select csname, lastbootuptime
            $lbut = $os.LastBootUpTime
            Write-Output "$pcn with IP $pcip was booted at $lbut"

        }
    }
    $CSVOutObj |Select-Object @{n="System_Name";e="PSComputerName"},@{n="Relative_Path";e="__RelPath"},@{n="Local_Path";e="__Path"},Command,Description,Location,Name,User,UserSID| Export-Csv $CSVFileNam -NoClobber -NoTypeInformation

    Start-Sleep 1
}
function GetSet-ncp
{
    $II = ""
    $NetCategory = ""
    Get-NetConnectionProfile
    $II = Read-Host "Enter the number from the above network Interface to change that adapter"
    $NetCategory = Read-Host "Enter Public, Private, (or!!UNTESTED DomainAuthenticated!! - just restart the nlasvc service...) for the connection profile"
    write-host $II
    Start-Sleep 10
    If ($II -ne "")
    {
        write-host "II is $II"
        Set-NetConnectionProfile -InterfaceIndex $II -NetworkCategory $NetCategory -ErrorAction SilentlyContinue
    }
    reloadnetmenu
}

function Wait-Live()
{
    Clear-Host
    $pc = Read-host "Enter the name of a pc to wait for"
    $port = Read-host "Enter the port number (like 3389)"
    "`n"
    (Invoke-Expression ".\Test-Ports.ps1 -computer $pc -port @($port)")
    while ((Invoke-Expression ".\Test-Ports.ps1 -computer $pc -port @($port)").Open -eq "False")
    {
        Write-host "." -NoNewline
        start-sleep 5
    }
    "`n"
    (Invoke-Expression ".\Test-Ports.ps1 -computer $pc -port @($port)")

    reloadnetmenu
}

function Get-ModifiedFileList()
{
    Clear-Host
    # Set up Variables
        $Dir2Scan = "" # Dir2Scan is the folder that will be the root folder for the search$DaySpan = $null # DaySpan is the number of days since a file was modified
        $Ext2Scan = $null # Ext2Scan is a specific file extension to filter by
        $FFilter = $null # FFilter uses a typical filter syntax to find a file
        $tdir = "$env:TMP"+"\" # This is the user's temporary folder - change it to ".\" or PWD for current folder
    # Prompt for optional search criteria
        $DaySpan = Read-Host "Enter the number of days of modified files to list"
        # Set number of days to 100 years if nothing was entered
        If($DaySpan -eq ""){$DaySpan = "36500"}
        $Dir2Scan = Read-Host "Hit enter to scan Shared Files on DCLEC or enter an alternate path"
    # Set a default Dir2Scan value if nothing was entered in the prompt
        If($Dir2Scan -eq ""){$Dir2Scan = "\\Dclec\Shared Files"}
        $Ext2Scan = Read-Host "`nType an extension to restrict search"
        $FFilter = Read-Host "`nEnter a file search filter if necessary"
    # Inform user of selections made
        Write-Warning "`nSearching $Dir2Scan for files up to $DaySpan day(s) old."
    # $ExCSVFileName = ((Get-date -Format "yyyy_M_d_hms") + "_Files.csv") #Use this option if you want the file saved in the current script folder
        $ExCSVFileName = "$tdir" + ((Get-date -Format "yyyy_M_d_hms") + "_Files.csv")
        "`n"
    # Run different searches depending on whether a file age in days was declared
    Get-ChildItem -Filter $FFilter -Recurse $Dir2Scan -ErrorAction SilentlyContinue |Where-Object {$_.Extension -like "*$Ext2Scan"}|
        Where-Object {$_.LastWriteTime -gt (get-date).AddDays(-$DaySpan)}|
        ForEach-Object {$_ |add-member -name "Owner" -MemberType NoteProperty -Value (get-acl $_.FullName).Owner -PassThru}|
        Sort-Object fullname |select Fullname,Name,psdrive,Directory,Extension,CreationTime,LastWriteTime,Length,Owner|
        Export-Csv $ExCSVFileName -Force -NoTypeInformation 
    $YNPrompt = Read-Host "Output at $ExCSVFileName is temporary. `nBe certain to save it elsewhere if you want a permanent copy.`nDo you wish to open it now in Excel? - Y or N"
    if($YNPrompt -eq "Y"){Invoke-Item $ExCSVFileName}else{Import-Csv $ExCSVFileName|Out-GridView -Title $ExCSVFileName}
    reloadmenu
    # Delete the CSV temp file when user exits out of menu
    # File will not be deleted if it is still open at this time
    Remove-Item $ExCSVFileName -Force
}
function run-get-filehash()
{
    # View the following for ideas
    # https://docs.microsoft.com/en-us/powershell/module/Microsoft.PowerShell.Utility/Get-FileHash?view=powershell-5.1
    Clear-Host
    $FullFileName = Read-Host "`nType the path and filename for the file for which you wish to view the hash`n Example: C:\Users\Public\Downloads\Software\Tools\Java\jdk-10.0.1_windows-x64_bin.exe`n"
    $HashAlgorithm = Read-Host "`nEnter the hashing algorithm you wish to use. `n`tChoose from SHA1 | SHA256 | SHA384 | SHA512 | MACTripleDES | MD5 | RIPEMD160"
    # Use of Get-FileHash: Get-FileHash [-Path] <String[]> [-Algorithm {SHA1 | SHA256 | SHA384 | SHA512 | MACTripleDES | MD5 | RIPEMD160}]
    Write-Warning "Running the following command:`n`nget-filehash -Path $FullFileName -Algorithm $HashAlgorithm -Verbose`n"
    get-filehash -Path $FullFileName -Algorithm $HashAlgorithm -Verbose

    reloadmenu
}
function get-WinSAT()
{
    $wsatResults = @()
    $pcselect = Read-Host "Enter domain pc name or * for all"
    $pclist = Get-ADComputer -Filter 'name -like $pcselect' |select name
    foreach ($pc in $pclist)
    {
        $pc.name
        $wsat = get-wmiobject win32_winsat -ComputerName $pc.name -ErrorAction SilentlyContinue |select pscomputername,cpuscore,d3dscore,diskscore,graphicsscore,memoryscore,winsatassessmentstate,winsprlevel
        $wsatResults += $wsat
    }
    $wsatResults |sort pscomputername,winsprlevel,cpuscore,diskscore,memoryscore,graphicsscore |ft
}

function test-srvconnects()
{
    Clear-Host
    $err=@()
    Get-PCList
    Write-Host "Testing.." -nonewline
    $ObjectOut = foreach ($Global:PCLine in ($Global:PCList))
        {
        Identify-PCName
        $tc =  test-connection $Global:PC -Count 1 -ErrorVariable err -ErrorAction SilentlyContinue
        Write-Host "." -nonewline
        $EE = $err.Exception
            $OBProperties = @{'pscomputername'=$Global:PC;
                'Source'=$tc.PSComputerName;
                'IPV4Address'=$tc.IPV4Address;
                'Bytes'=$tc.ReplySize;
                'ResponseTime'=$tc.ResponseTime;
                'DateChanged'=$Global:CDate;
                'Error'=$EE}
            New-Object -TypeName PSObject -Prop $OBProperties
        }
    $ObjectOut |select pscomputername, IPV4Address, Bytes, ResponseTime, Source, DateChanged, Error  |ft # |export-csv -Path $ObCSVFile -Append -NoTypeInformation
    Write-output "Check port 3389 for 10.127.10.85 rather than viewing the PING test`n"
}

function get-OpenFiles()
{
    Clear-Host
    [array]$srvs = @("DCLEC,LECvFiles").split(",") |%{$_.trim()} # DCBackup
    foreach ($pc in $srvs)
    {
        $srvobj = invoke-command {c:\Windows\System32\openfiles.exe /query /s $pc /fo csv /V} |convertfrom-csv |
        select Hostname,ID,Type,@{Name="UserName"; Expression="Accessed By"},@{Name="Path_FileName"; Expression="Open File (Path\executable)"} |
        Where {$_.Path_FileName -ne "D:\\"}|
        sort Path_FileName |ft
        $srvobj
    }
}

function iterate-Profiles()
{
    # from calling function: $proc = gwmi win32_process -computer $pscn -Filter "Name = 'explorer.exe'"
    if ($Choose2CopyFile -eq "copy" -or $ReplaceLink -eq "dlink")
    {
        # make $userdir variable equal to current user SamAccountName
        $userdir = $pgo
        # Set $ProfileList to contain all contents of a systems Users folder (Usually just a list of Profile folder names)
        $ProfileList = gci "\\$pscn\C$\Users\" -Directory -Force
    }
    if ($CollectFile -eq "collect")
    {
        # make $userdir variable equal to current user SamAccountName
        $userdir = $pgo
        # Set $ProfileList to contain all contents of a systems Users folder (Usually just a list of Profile folder names)
        $ProfileList = gci "\\$pscn\C$\Users\" -Directory -Force
        if ($AppendOverwrite -eq "O")
            {
                $sw = New-Object -TypeName System.IO.StreamWriter -ArgumentList $OutFile
                # Change value of $AppendOverwrite it so that it would overwrite initially and then append from then on.
                $AppendOverwrite = "NoMoreO"
            }
        else
            {
                $sw = New-Object -TypeName System.IO.StreamWriter -ArgumentList $OutFile ,'Append' #$CollectFile Out
            }
        $sw.WriteLine("Workstation,Name,Location,Printer,PrinterState,PrinterStatus,ShareName,SystemName")#$CollectFile Out header
    }
    if ($ReplaceLink -eq "dlink" -or $ReplaceLink -eq "plink")
    {
        $value0 = (Read-Host "Enter new target Drive Letter in the following format L:\") # "L:\" # Change this into a = Read-Host "Enter new target Drive Letter in the following format L:\"
        $DLValue = (Read-Host "Enter old target Drive Letter in the following format O:\") # "S:\" # Change this into a = Read-Host "Enter old target Drive Letter in the following format O:\"
    }
    # comment out this foreach and the next $userdir variable declaration to only apply the following action to the currently loged on user
    if ($ReplaceLink -eq "dlink")
    {
        # Should create a do-dlink function with the rest of this if statement and loop back to it on the NPconfirm -ne "A"
        if ($Global:NPconfirm -eq "Y" -or $Global:NPconfirm -eq "N")
        {
            $Global:NPconfirm = Read-Host "Type A to fix all, Y to fix next, N to NOT fix next, or None to NOT fix any."
        }
        Elseif ($Global:NPconfirm -eq "None")
        {
            # No Need to do anything here
            # Is there another place that we need to test for None rather than just N?
        }
        Elseif ($Global:NPconfirm -ne "A")
        {
            Write-Host "Your options are A,Y, or N. Please try again"
        }
    }
    if ($ReplaceLink -eq "plink")
    {
        $PathToLink = "\\$pscn\C$\Users" # $PathToLink = Read-Host "Type the root of the path to search. Do not end in \.`nUse a local drive letter or UNC path."
        Replace-Links $PathToLink $value0 $DLValue
    }
    foreach ($PL in $ProfileList)
    {
    $userdir = $PL.Name
    $PathToLink = "\\$pscn\C$\Users\$userdir\Desktop"
        if ($ReplaceLink -eq "dlink")
            {
                 Replace-Links $PathToLink $value0 $DLValue
            }
        if ($Choose2CopyFile -eq "copy") # If the user previously confirmed he wants to copy by typing the word "copy" then run the following function
            {
                # Don't copy to the following profiles
                if ($userdir -ne "Public" -and $userdir -ne "All Users" -and $userdir -ne "Default User" -and $userdir -ne "Default" -and $userdir -ne "Default.Migrated")
                {
                    replace-DesktopFiles
                }
            }
        # Process the contents of a text file on a system. Output results to a text file.
        if ($CollectFile -eq "collect")
            {
                # Currently set to a static location. This could be automated
                $P2File = "\\$pscn\C$\Users\Public\Downloads\" + $NamedFiles + $userdir + ".txt"
                if (Test-Path $P2File)
                {
                    $P2File
                    Write-Warning "AppendOverwrite $AppendOverwrite"
                    $content = New-Object -TypeName System.IO.StreamReader -ArgumentList $P2File
                    $FL = GC $P2File -first 1
                    # Comment out the belowpattern you do not want otherwise, the second pattern will succeed
                        $Pattern1 = '^([A-Z][a-z]*\s*)([A-Z][a-z]*\s*)([A-Z][a-z]*\s*)([A-Z][a-z]*\s*)([A-Z][a-z]*\s*)([A-Z][a-z]*\s*)'
                        $Pattern1 = '^(Location*\s*)(Name*\s*)(PrinterState*\s*)(PrinterStatus*\s*)(ShareName*\s*)(SystemName*\s*)'
                    # Comment out the above pattern you do not want otherwise, the second pattern will succeed
                    $C1Width = ($ColLocationWidth = [regex]::Matches($FL, $Pattern1).Groups[1].Value).length
                    $C2Width = ($ColNameWidth = [regex]::Matches($FL, $Pattern1).Groups[2].Value).length
                    $C3Width = ($ColPrinterStateWidth = [regex]::Matches($FL, $Pattern1).Groups[3].Value).length
                    $C4Width = ($ColPrinterStatusWidth = [regex]::Matches($FL, $Pattern1).Groups[4].Value).length
                    $C5Width = ($ColShareNameWidth = [regex]::Matches($FL, $Pattern1).Groups[5].Value).length
                    $C6Width = ($ColSystemNameWidth = [regex]::Matches($FL, $Pattern1).Groups[6].Value).length
                    $0to1 = $C1Width
                    $0to2 = $0to1 + $C2Width
                    $0to3 = $0to2 + $C3Width
                    $0to4 = $0to3 + $C4Width
                    $0to5 = $0to4 + $C5Width
                    $0to6 = $0to5 + $C6Width
                    while($null -ne ($line = $content.ReadLine())) {
                        $line
                        $Location = ($line.SubString(0,$C1Width).trim())
                        $Location
                        $sw.Write($pscn);$sw.Write(",")
                        $sw.Write($userdir);$sw.Write(",")
                        if ($Location){$sw.Write($line.SubString(0,$C1Width).trim())} else {$sw.Write("Local")};$sw.Write(",")
                        $sw.Write($line.SubString($0to1,$C2Width).trim());$sw.Write(",")
                        $sw.Write($line.SubString($0to2,$C3Width).trim());$sw.Write(",")
                        $sw.Write($line.SubString($0to3,$C4Width).trim());$sw.Write(",")
                        $sw.Write($line.SubString($0to4,$C5Width).trim());$sw.Write(",")
                        $sw.Write($line.SubString($0to5,$C6Width).trim());$sw.Write("`n")
        
                    }
                    <#
                    while($null -ne ($line = $content.ReadLine())) {
                        $line
                        $Location = ($line.SubString(0,10).trim())
                        $sw.Write($pscn);$sw.Write(",")
                        $sw.Write($userdir);$sw.Write(",")
                        if ($Location){$sw.Write($line.SubString(0,10).trim())} else {$sw.Write("Local")};$sw.Write(",")
                        $sw.Write($line.SubString(11,31).trim());$sw.Write(",")
                        $sw.Write($line.SubString(43,14).trim());$sw.Write(",")
                        $sw.Write($line.SubString(57,14).trim());$sw.Write(",")
                        $sw.Write($line.SubString(71,12).trim());$sw.Write(",")
                        $sw.Write($line.SubString(83).trim());$sw.Write("`n")
                    }
                    #>
                }
            }
    }
Write-output "GC OutFile: $OutFile"
Get-content $OutFile # Collect OutFile
$sw.close() # Close Collect OutFile
}


# Function replaces / copies a file to a folder on specified systems
# This can be designed to be called from different areas.
# It is currently called from function iterate-Profiles via UsrsOnPCs to run against certain active PCs
function Replace-DesktopFiles()
{
    $ciResult = gci $PathToLink |where {$_.Name -like $NamedFiles} |select Name, FullName, LastWriteTime, @{Name="ComputerName"; Expression={$pscn}} # look for deleting shared files link "SharedFiles (S)"
    Write-warning "CiResult is: $ciResult"
    # $ciResult = gci $PathToLink |where {$_.Name -like "*share*.lnk"} |select Name,LastWriteTime, @{Name="ComputerName"; Expression={$pscn}} # look for deleting shared files link
    # $ciResult = gci $PathToLink -ErrorAction SilentlyContinue |where {$_.Name -like "*INFOCENTER*.url"} |select Name,LastWriteTime, @{Name="ComputerName"; Expression={$pscn}} # look for infocenter.url link
    if ($ciResult -ne $null)
    {
        Write-Output "$CiResult" # -WarningVariable a ;$a
        foreach($cr in $ciResult)
        {
            # Determine Path and Filename to delete
            $fn = $cr.name 
            $rmfile = $PathToLink +"\"+$fn 
            # uncomment to delete the old file found first prior to copying its replacement.
            del $rmfile # -Confirm
            # Export a list of all shortcuts marked for replacement
            $CSVName = ".\_"+$Global:CDate+"DesktopLinks.csv"
            $cr |Export-Csv $CSVName -NoTypeInformation -Append
        }
    }
    # copy $copyFile to new location on $desktop with filename: $ReplaceWithFN
    Write-Warning "Copying $copyFile to $PathToLink\$ReplaceWithFN"
    copy $copyFile "$PathToLink\$ReplaceWithFN" -ErrorAction Inquire -ErrorVariable ER # -ErrorAction SilentlyContinue 
    # $cr = $error[0].exception.message
    $CSVName = ".\_"+$Global:CDate+"DesktopLinks.csv"
    # if ($cr -ne "")
    if ($ER -ne "")
    {
        $ER|Out-File $CSVName -Append -NoClobber
        # Write-Output "Was trying to copy $copyFile to $PathToLink\$ReplaceWithFN" |Out-File $CSVName -Append -NoClobber
        $ER = ""    
    }
    Else
    {
        Write-Output "Copied $copyFile to $PathToLink\$ReplaceWithFN" |Out-File $CSVName -Append -NoClobber
    }
    

}


# This can be designed to be called from different areas.
# It is currently called from function UsrsOnPCs to run against all active PCs
function List-MappedDrives()
{
    # from calling function: $proc = gwmi win32_process -computer $pscn -Filter "Name = 'explorer.exe'"
    # make $userdir variable equal to current user SamAccountName
    $domain = $env:USERDOMAIN
    $userdir = $pgo
    $psdname = "$pscn"+"hku"
    $npsd = New-PSDrive -PSProvider Registry -Name $psdname -root HKEY_USERS
    # comment out this foreach and the next $userdir variable declaration to only apply the following action to the currently loged on user
    $ProfileList = gci "\\$pscn\C$\Users\" -Directory -Force
    # $NameList = get-aduser -Filter 'samaccountname -like "*"' |select SamAccountName
    # [array]$NameList = @("JasonCrockett,CHazlewood").split(",") |%{$_.trim()}

    foreach ($PL in $ProfileList)
    {
        $Name = $PL.Name
        # $NChar = $Name.length
        # Write-host "$Name is $NChar Long"
        if ($Name -ne "All Users" -and $Name -notlike "Default*" -and $Name -ne "Public")
        {
            $strw32acct = Get-wmiobject -Class win32_useraccount -Filter "Domain = '$domain' AND Name = '$Name'" |select Name,FullName,SID
            $strSID = $strw32acct.SID
            Write-Host $strw32acct.Name,$strSID,"`n" -NoNewline
            $path = $psdname + ":\" + $strSID + "\Network"
            if (test-path -Path $path) {Write-host "`t$Name,TRUE,$strSID,network,$pscn`t" -ForegroundColor Green} else {Write-Host "`t$Name,FALSE,$strSID,network,$pscn`t" -ForegroundColor Red}
            if (Test-Path -Path $path"\s") {Write-host "`t$Name,TRUE,$strSID,network\s,$pscn`t" -ForegroundColor Green} else {Write-host "`t$Name,FALSE,$strSID,network\s,$pscn`t" -ForegroundColor Red}
            # "Path:$path"
            # gci $path |select Name,@{n="Remote Path";e={($_.Property | select -ExpandProperty )-join ','}}
            gci $path |select @{n="DriveLetter";e={$_.pschildname}},@{n="RegistryKey";e={$_.Name}}|ft
            $strSID=$null
        }
    }
    Remove-PSDrive -Name $psdname
    $NameList = $null
    $pc = $null

}


# Get-NewADObjects
# list AD objects where "WhenCreated" date is greater than ...
function Get-NewADObjects
{
$WCDate = Read-Host "Enter a start date in format mm/dd/yyyy"

(get-adobject -filter * -properties *|where {$_.WhenCreated -gt $WCDate}) |select Name,canonicalname,whencreated,modified |sort WhenCreated,modified
}


# Get-Shortcut called from Replace-Links function
function Get-Shortcut {
  param(
    $path = $null
  )
  # https://stackoverflow.com/questions/484560/editing-shortcut-lnk-properties-with-powershell
  $obj = New-Object -ComObject WScript.Shell

  if ($path -eq $null) {
    $pathUser = [System.Environment]::GetFolderPath('StartMenu')
    $pathCommon = $obj.SpecialFolders.Item('AllUsersStartMenu')
    $path = dir $pathUser, $pathCommon -Filter *.lnk -Recurse 
  }
  if ($path -is [string]) {
    $path = dir $path -Filter *.lnk
  }
  $path | ForEach-Object { 
    if ($_ -is [string]) {
      $_ = dir $_ -Filter *.lnk
    }
    if ($_) {
      $link = $obj.CreateShortcut($_.FullName)
      $link |select *
      $info = @{}
      $info.Hotkey = $link.Hotkey
      $info.TargetPath = $link.TargetPath
      $info.LinkPath = $link.FullName
      $info.Arguments = $link.Arguments
      $info.Target = try {Split-Path $info.TargetPath -Leaf } catch { 'n/a'}
      $info.Link = try { Split-Path $info.LinkPath -Leaf } catch { 'n/a'}
      $info.WindowStyle = $link.WindowStyle
      $info.IconLocation = $link.IconLocation
      $info.WorkingDirectory = try {Split-Path $info.WorkingDirectory -Leaf } catch { 'n/a'}
      # $info.RelativePath = try {Split-Path $info.RelativePath -Leaf } catch { 'n/a'}

      New-Object PSObject -Property $info
    }
  }
}
function Set-Shortcut {
  param(
  [Parameter(ValueFromPipelineByPropertyName=$true)]
  $LinkPath,
  $Hotkey,
  $IconLocation,
  $Arguments,
  $TargetPath,
  $WorkingDirectory
  )
    # https://stackoverflow.com/questions/484560/editing-shortcut-lnk-properties-with-powershell
  begin {
    $shell = New-Object -ComObject WScript.Shell
  }

  process {
    $link = $shell.CreateShortcut($LinkPath)

    $PSCmdlet.MyInvocation.BoundParameters.GetEnumerator() |
      Where-Object { $_.key -ne 'LinkPath' } |
      ForEach-Object { $link.$($_.key) = $_.value }
    $link.Save()
  }
}

function Replace-Links()
{
<#
.EXAMPLE
    Use the following example to assign F11 hotkey to PowerShell (uncomment line, make sure you have full admin privileges: 
    Get-Shortcut | Where-Object { $_.LinkPath -like '*Windows PowerShell.lnk' } | Set-Shortcut -Hotkey F11 
#>
param(
  [Parameter(ValueFromPipelineByPropertyName=$true)]
  $PathToLink,
  $value0 = (Read-Host "Enter new target Drive Letter in the following format L:\"), # "L:\" # Change this into a = Read-Host "Enter new target Drive Letter in the following format L:\"
  $DLValue = (Read-Host "Enter old target Drive Letter in the following format O:\") # "S:\" # Change this into a = Read-Host "Enter old target Drive Letter in the following format O:\"
  )
    #TO-DO: Add the following as parameters and do the Read-Host before the profile list get's collected
        # Designate the new Drive to be used in the new TargetPath
            # $value0 = Read-Host "Enter new target Drive Letter in the following format L:\" # "L:\" # Change this into a = Read-Host "Enter new target Drive Letter in the following format L:\"
            # Change $DLValue and $Pattern1 at the same time
            # $DLValue = Read-Host "Enter old target Drive Letter in the following format O:\"# "S:\" # Change this into a = Read-Host "Enter old target Drive Letter in the following format O:\"
        # $DLValue = "\\SERVER\Share Name\"
    #EndTO-DO
        # Set the WorkingDirectory driveletter to the same driveletter as the new target drive letter
        $WDvalue0 = $value0
    $DLCSVOut = "C:\Users\%username%\Documents\WindowsPowerShell\"+$Global:CDate+"DesktopLinks.csv"
    $LinksFound = ls $PathToLink *.lnk -Recurse # \\PCName\C$\Users\ProfileName\Desktop
    # Write-warning "Writing CSV to $DLCSVOut"
    $LinksFound |select DirectoryName, Name |export-csv -Path $DLCSVOut -Append -NoTypeInformation # select Directory, DirectoryName, FullName, Name, BaseName 
    foreach ($lf in $LinksFound)
    {
        $source = $lf.fullname
        $TP_GSc_qry = $DLValue+"*"
        $GSc = Get-shortcut $source |select FullName, RelativePath,TargetPath,WorkingDirectory |where {$_.TargetPath -like "$TP_GSc_qry" -and $_.WorkingDirectory -ne "n/a"}
        $LSc = Get-shortcut $source |select FullName, RelativePath,TargetPath,WorkingDirectory |where {$_.WorkingDirectory -ne "n/a"}
        $alldsktpLnkPropertiesOut = "C:\Users\JasonCrockett\Documents\WindowsPowerShell\"+$Global:CDate+"AllDsktpLnkPropertiesOut.csv"
        $LSc |select * |export-csv -Path $alldsktpLnkPropertiesOut -Append -NoTypeInformation # -Verbose
        # $desktoplnkPropertiesOut = "C:\Users\JasonCrockett\Documents\WindowsPowerShell\"+$Global:CDate+"DesktopLinkProperties.csv"
        $desktoplnkPropertiesOut = $Script:PSScriptRoot+"\"+$Global:CDate+"DesktopLinkProperties.csv"
        $GSc |select * |export-csv -Path $desktoplnkPropertiesOut -Append -NoTypeInformation # -Verbose
        # Zero out variables
        $GSclp = $null
        $GSctp = $null
        $GScwd = $null
        $NewTargetPath = $null
        $NewWDPath = $null
        if ($GSc)
        {
            # Get LinkPath to variable $GSclp
            $GSclp = $GSc.FullName
            # Get TargetPath to variable $GSctp
            $GSctp = $GSC.TargetPath
            # Get WorkingDirectory to variable $GScwd
            $GScwd = $GSC.WorkingDirectory
            # Create a regex pattern to break the TargetPath into two groups: Drive and Path This is currently case sensitive
            <#This line workd, testing the line below with variable $DLValue#> #$Pattern1 = '^(S:\\)(.*?)$' #'^($DLValue\\)(.*?)$' (w/UNC: '^(\\\\dclec\\Shared\sFiles\\)(.*?) *$'
            $Pattern1 = "^($DLValue\)(.*?)$" # (w/UNC: '^(\\\\dclec\\Shared\sFiles\\)(.*?) *$'
            # $Pattern1 = '^(\\\\server\\share\sname\\)(.*?) *$'
            # $Pattern1 = '^(\\\\SERVER\\Share\sName\\)(.*?) *$'
            # Match the drive letter and path into two separate variables
            
            $value1 = [regex]::Matches($GSctp, $Pattern1).Groups[1].Value # This is case sensitive using -match is case insensitive
            $value2 = [regex]::Matches($GSctp, $Pattern1).Groups[2].Value # This is case sensitive using -match is case insensitive
            
            <#
            # If I can tweak this style of regex, it would be case insensitive
            $value1 = ($GSctp -match $Pattern1).Groups[1].Value
            $value2 = ($GSctp -match $Pattern1).Groups[2].Value
            #>
            # Display the values showing the results of the Regex matches
            Write-warning "Re: $source"
            Write-warning "0(NewDrive TargetPath) = $value0"
            Write-warning "1(Matches Group1) = $value1"
            Write-warning "2(Matches Group2)= $value2"
            
            if ($GScwd)
            {
                $WDvalue1 = [regex]::Matches($GScwd, $Pattern1).Groups[1].Value # This is case sensitive using -match is case insensitive
                $WDvalue2 = [regex]::Matches($GScwd, $Pattern1).Groups[2].Value # This is case sensitive using -match is case insensitive
                <#
                # Display the values showing the results of the Regex matches
                $WDvalue0
                $WDvalue1
                $WDvalue2
                #>
            }
            if ($value1 -eq $DLValue) #"S:\")
            {
                $NewTargetPath = $value0+$value2
                $NewTargetPath
                if ($GScwd)
                {
                    $NewWDPath = $WDvalue0+$WDvalue2
                    $NewWDPath
                }
                Write-host "The new path on $PSCN is:`n`t $NewTargetPath `nand the WorkingDirectory is `n`t$NewWDPath" -ForegroundColor Green
                $lnkPropertiesNewOld = $Script:PSScriptRoot+"\"+$Global:CDate+"lnkPropertiesNewOld.csv"
                $GSc |select * |export-csv -Path $lnkPropertiesNewOld -Append -NoTypeInformation # -Verbose
                If ($Global:NPconfirm -eq "Y" -or "A") # NPconfirm is not N or None
                {
                    Set-Shortcut -TargetPath $NewTargetPath -LinkPath $GSclp -WorkingDirectory $NewWDPath
                }
                else # NPconfirm is N or None so DO NOTHING
                {
                    Write-Warning "I am not replacing $GSctp with $NewTargetPath"
                }
            }
        }
    }
}

function List-RecentUpdate()
{
    Clear-Host
    $SN=$null
    $RecentUpdateList=$null
    Get-PCList
    # $servers = Get-ADComputer -Filter {OperatingSystem -like "*server*" -and Enabled -eq $true}
    # $servers = Get-ADComputer -Filter {Enabled -eq $true}
    $RecentUpdateList = foreach ($PCline in $Global:pclist)
    {
            # This if statement could potentially be moved to inside the Get-PCList function, but it might need to be modified to do so
            if ($pcline.name -eq $null)
            {
                $pc = $pcline
            }
            else
            {
                $pc = $pcline.name
            }
        Write-Host ". " -NoNewline
        # Write-Warning $pc
        $ping = test-connection $pc -Count 1 -Quiet -ErrorAction SilentlyContinue
        if ($ping -eq $true){
            $RUL = Get-WMIObject -Class "Win32_QuickFixEngineering" -ComputerName $pc -Credential $Cred |select PSComputerName,InstalledOn,HotFixID -Last 1

            $RULProperties = @{'pscomputername'=$pc;
                    'Name'=$RUL.PSComputerName;
                    'InstalledOn'=$RUL.InstalledOn;
                    'HotFixID'=$RUL.HotFixID}
            New-Object -TypeName PSObject -Prop $RULProperties
        }
    }
    $RecentUpdateList |sort InstalledOn |ft
    reloadmenu
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
function Combine-TextFiles
{
Invoke-Expression -command '.\Combine-TextFiles.ps1 -ext ".trf" -Path "C:\Users\Public\Downloads\_Temp Clear Phone\TRAFFIC"'
}
function Browse-Share
{
param(
  [Parameter(ValueFromPipelineByPropertyName=$true)]
  [string[]] $ShareName = ($ShareName=if ($SN = Read-Host -Prompt "Please enter a UNC server and sharename ('\\lec.lan\dfs\SharedFiles')") {$SN} else {"\\lec.lan\dfs\SharedFiles"})
  # $pclist = ((Read-Host "Enter a comma separated list of computernames").split(",") | %{$_.trim()})
  )
    $gPd = Get-PSDrive|where {$_.Root -like "?:\"} |select Name
    $gpdcsvList = $gpd.Name -join ", "
    $NewDL = Read-Host "Enter a drive letter to map to $ShareName (other than $gpdcsvList)"
    $NewPath = Read-Host "The path to the location you would like to list starting with \"
    New-PSDrive -Name "$NewDL" -PSProvider "FileSystem" -Root "$ShareName" -Credential $cred
    Get-PSDrive |ft -AutoSize
    $NewDLPath = "$NewDL"+":"
    $DL = Get-ChildItem $NewDLPath
    $DL = Get-ChildItem ("$NewDLPath"+"$NewPath")
    $DL |select PSDrive,CreationTime,FullName,LastAccessTime,Mode <#| Sort Mode -Descending -property @{E="FullName";Descending=$false} @{Expression="Mode";Descending=$true}|sort FullName #>|ft -AutoSize -Wrap # Root,Parent,PSChildName,
    Remove-PSDrive -Name "$NewDL"
    reloadmenu
}
Function Compare-Files
{
$BaseDirectory = Read-Host "Enter the root of the path to search for duplicate files"
$wpsdir = gci $BaseDirectory -Recurse
$ObjectOut = foreach ($wpsdir1 in $wpsdir.fullname)
{
        $GFH = Get-FileHash $wpsdir1 -Algorithm SHA1 -Verbose
        if ($GFH.Hash -ne $null)
        {
            $GFH |Select *
            $FilePath = $GFH.Path
            $FileHash = $GFH.Hash
            $FileAlgorithm = $GFH.Algorithm
            $OBProperties = @{'FilePath'=$FilePath;
                        'FileHash'=$FileHash;
                        'Algorithm'=$FileAlgorithm}
                New-Object -TypeName PSObject -Prop $OBProperties
        }
}

$ObjectOutNonDups = $ObjectOut.FileHash |select -Unique
$ObjectOut|Where {$_.FileHash -ne $null} |sort FileHash |Group-Object FileHash |? {$_.count -gt 1} | %{$_.Group} |ft

$ObCSVFile = $PSScriptRoot+"\"+$Global:CDate+"_Compare_Files.csv"
$ObjectOut|Where {$_.FileHash -ne $null} |sort FileHash |Group-Object FileHash |? {$_.count -gt 1} | %{$_.Group} |Export-Csv $ObCSVFile -NoClobber -NoTypeInformation -Append
Write-Host "All lines: " -NoNewline
    $ObjectOut.Count
Write-Host "Non-Duplicate lines: " -NoNewline
    $ObjectOutNonDups.Count

# $ObjectOut |select FileHash, Algorithm, FilePath
}
function Manage-NTFS
{
<#
.NAME
    Manage-NTFS
.SYNOPSIS
    Change NTFS permissions and perform actions to target file/folder
.DESCRIPTION
     Change NTFS permissions so that the user will have permissions to 
     delete a target file/folder. This could be modified to prompt for and perform other actions
.PARAMETERS
    $NewOwn is the name of the target owner 
    $NewPermUser is the name of the target user to receive full permissions
    $Targetfolders is the name of the Folder(s) (wild-card accepted) to which the permission change will be applied
.EXAMPLE
    Manage-NTFS -NewOwn domain\NewOwnerName -NewPermUser domain\NewPermUserName -TargetFolders C:\TEMP - Copy*
.SYNTAX
    Manage-NTFS [-NewOwn] <string[]> [-NewPermUser] <string[]> [-TargetFolders] <string[]>
.REMARKS
    NTFSSecurity module can be downloaded from the following site:
    https://gallery.technet.microsoft.com/scriptcenter/1abd77a5-9c0b-4a2b-acef-90dbb2b84e85
    Instructions for the module's use are found at:
    https://blogs.technet.microsoft.com/fieldcoding/2014/12/05/ntfssecurity-tutorial-1-getting-adding-and-removing-permissions/
    See >help *ntfs* for other cmdlets in the ntfssecurity module
#>
[CmdletBinding()] #(allows the use of parameter attributes)
    param
    (
            #[Parameter(Mandatory=$True)]
            $NewOwn =(Read-Host "Enter a valid new Owner dom\UserName"), 
            $NewPermUser = (Read-Host "Enter a valid new User for full permissions dom\UserName"),
            $TargetFolders = (Read-Host "Enter a valid Target Folder like: C:\TEMP - Copy*" )
    )
Import-Module ntfssecurity
$folders = ls $Targetfolders # $Targetfolders # List direcotry to variable If user doesn't have list permissions this ls may fail
$FStart = $folders.Count
    foreach ($folder in $folders)
    {
        Get-NTFSOwner $folder
        # Change $folder ownership
        Set-NTFSOwner -path $folder -Account $NewOwn
        # Add FullControl to $folder for the $NewPermUser user
        Get-Item $folder|Add-NTFSAccess -Account $NewPermUser -AccessRights FullControl
        # Remove the $folder directory and all contents
        rmdir $folder -Recurse -Confirm
    }
$FEnd = ls $Targetfolders 
$dif = $FEnd.Count
Write-Warning "Out of $FStart Objects, only $dif remain."
}

Function CallGet-NICStatus()
{
    Clear-Host
    # [array]$pclist=(Read-Host "Enter a comma separated list of computernames").split(",") | %{$_.trim()}
    if ($Global:pclist -eq $null)
    {
        Get-PCList
    }
    Write-host "Checking for NIC status"
    $PCNICStatus = foreach($Global:pcLine in $Global:pclist)
    {
        Identify-PCName
        Get-NICStatus
    }
    $PCNICStatus |ft
    $Global:pclist=$null
}

Function View-Certificates()
{
    Get-PCList
    foreach ($PCline in $PCList)
    {
            # This if statement could potentially be moved to inside the Get-PCList function, but it might need to be modified to do so
            if ($pcline.name -eq $null)
            {
                $pc = $pcline
            }
            else
            {
                $pc = $pcline.name
            }
    $certs = invoke-command {gci Cert:\LocalMachine\My -Recurse} -computername $pc
    # Set-Location Cert:\LocalMachine\My
    Write-Host = "Asset ID: "$pc
    $certs |select PSComputerName, Issuer, Subject, NotBefore, notafter|sort NotAfter |FT PSComputerName,Issuer,Subject,NotBefore,NotAfter
    #Get-ChildItem -Recurse cert: |select Subject, notafter
    Write-host "`n"
    }
}

function Get-IPConfig()
{
<#
.Synopsis
 Runs 'Get-WmiObject Win32_NetworkAdapterConfiguration' against a list of Windows systems with enabled NICs
.Description
 Runs 'Get-WmiObject Win32_NetworkAdapterConfiguration' against a list of Windows systems with enabled NICs
 Based on code by Techibee: https://techibee.com/powershell/powershell-get-ip-address-subnet-gateway-dns-serves-and-mac-address-details-of-remote-computer/1367
 Adapted by Jason Crockett - Laclede Electric Cooperative
 2019-05-02
.Parameter
 pclist (one or more computernames)
.Example 
.\Get-IPConfig.ps1

.ToDo
    Fix RegEx for matching IP addresses

#>
    <#
    [cmdletbinding()]
    param (
     [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [string[]]$ComputerName = $env:computername
    )
    #>
    Clear-Host 
    $ObCSVFile = $PSScriptRoot+"\"+$Global:CDate+"_IPConfig.csv"
    Get-PCList
    
# $OutputObj  = New-Object -Type PSObject
    $OutputObj = foreach($pcline in $Global:pclist) {
            # This if statement could potentially be moved to inside the Get-PCList function, but it might need to be modified to do so
            if ($pcline.name -eq $null)
            {
                $pc = $pcline
            }
            else
            {
                $pc = $pcline.name
            }
      if(Test-Connection -ComputerName $pc -Count 1 -ea 0) {
       try {
        $Networks = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $pc -EA Stop | ? {$_.IPEnabled}
       } catch {
            Write-Warning "Error occurred while querying $pc."
            Continue
       }
       foreach ($Network in $Networks) {
        $IPAddress  = $Network.IpAddress[0]
        $SubnetMask  = $Network.IPSubnet[0]
        $DefaultGateway = $Network.DefaultIPGateway
        $DNSServers  = $Network.DNSServerSearchOrder
        $WINS1 = $Network.WINSPrimaryServer
        $WINS2 = $Network.WINSSecondaryServer   
        $WINS = @($WINS1,$WINS2)         
        $IsDHCPEnabled = $false
        $NError = $Stop
        If($network.DHCPEnabled) {
         $IsDHCPEnabled = $true
        }
        $MACAddress  = $Network.MACAddress
        $OBProperties = @{'computername'=$pc.ToUpper();
            'IPAddress'=$IPAddress;
            'SubnetMask'=$SubnetMask;
            'Gateway'=($DefaultGateway -join ",");
            'IsDHCPEnabled'=$IsDHCPEnabled;
            'DNSServers'=($DNSServers -join ",");
            'WINSServers'=($WINS -join ",");
            'MACAddress'=$MACAddress;
            'Error'=$NError}
        New-Object -TypeName PSObject -Prop $OBProperties
        
        <#
            $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $PC.ToUpper()
            $OutputObj | Add-Member -MemberType NoteProperty -Name IPAddress -Value $IPAddress
            $OutputObj | Add-Member -MemberType NoteProperty -Name SubnetMask -Value $SubnetMask
            $OutputObj | Add-Member -MemberType NoteProperty -Name Gateway -Value ($DefaultGateway -join ",")      
            $OutputObj | Add-Member -MemberType NoteProperty -Name IsDHCPEnabled -Value $IsDHCPEnabled
            $OutputObj | Add-Member -MemberType NoteProperty -Name DNSServers -Value ($DNSServers -join ",")     
            $OutputObj | Add-Member -MemberType NoteProperty -Name WINSServers -Value ($WINS -join ",")        
            $OutputObj | Add-Member -MemberType NoteProperty -Name MACAddress -Value $MACAddress
        #>
       }
      }
     }
    Write-Warning Sending results to: $ObCSVFile
    $OutputObj
    $OutputObj|Select computername, SubnetMask, DNSServers, Gateway, IPAddress, IsDHCPEnabled, MACAddress, WINSServers, Error |Export-Csv $ObCSVFile -NoClobber -NoTypeInformation -Append
}

function Run-ADCleanup ()
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


# ***********Return************

menu
