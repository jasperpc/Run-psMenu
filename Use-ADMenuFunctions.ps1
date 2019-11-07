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
    while ($admenuselect -lt 1 -or $admenuselect -gt 30)
    {
        Trap {"Error: $_"; Break;}        
        $MNum = 0;Clear-host |out-null
        Start-Sleep -m 50
        Write-host $t4 $Script:ADd4 -ForegroundColor Green
        # Exit to Main Menu
        $MNum ++;$adExit = $MNum; 
		    $p1 =" $MNum. `tSelect to ";
		    $p2 = "Exit ";
		    $p3 = "to Main Menu"; $NC = "Red";chcolor $p1 $p2 $p3 $NC
        # Prompt for credential with permissions across the network
        $MNum ++;$GCred=$MNum;$p1 =" $MNum.  `t Enter";
            $p2 = "Credentials ";$p3 = "for use in script."; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
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
		    $p2 = "users logged into, ";
		    $p3 = "selected computers in this domain"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Perform one of the following: Copy a file, fix links, view mapped drives to/on selected computers in this domain
        $MNum ++;$adManageADUFiles = $MNum;
		    $p1 =" $MNum. `tPerform one of the following: ";
		    $p2 = "Copy a file, fix links, `n$t4`t`tview mapped drives,`tor read data from a file ";
		    $p3 = "`n$t4`t`tto/on selected computers in this domain"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
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
		    $p3 = " (SamAccountName,Name,LastLogonDate,Created,mail,address)"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # AD Users List
        $MNum ++;$ADMenuFltrUserList = $MNum;
		    $p1 =" $MNum. `tFilter AD";
		    $p2 = "User List ";
		    $p3 = "and list workstation"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # New AD Objects List
        $MNum ++;$Get_NewADObjects = $MNum;
		    $p1 =" $MNum. `tFilter ";
		    $p2 = "New AD Objects ";
		    $p3 = "by date created."; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
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
            $p1 =" $MNum. `tBulk Remove by age";
		    $p2 = "Old AD Computer Objects ";
		    $p3 = "from Active Directory (with confirm/suspend)"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Decommission Old AD Computer object(s) from AD
        $MNum ++;$Decommission_ADPC = $MNum;
            $p1 =" $MNum. `tBulk ";
		    $p2 = "Decommission,Disable AD Computer ";
		    $p3 = "Objects in AD and prepend description with _D (with confirm/suspend)"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # TestFix-PCSecureChannel (trust relationship) 
        $MNum ++;$TestFix_PCSecureChannel = $MNum;
            $p1 =" $MNum. `tTest or ";
		    $p2 = "Fix the secure channel (trust relationship) ";
		    $p3 = "between computer and domain (with confirm/suspend)"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        Write-Host "$t4$Script:ADd4"  -ForegroundColor Green
        # Setting up the menu in this way allows you to move menu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $MNum ++;$p1 =" $MNum.  FIRST";$p2 = "HIGHLIGHTED ";$p3 = "END"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        [int]$admenuselect = Read-Host "`n`tPlease select the menu item you would like to do now"
    }
    switch($admenuselect)
    {
        $adExit{$admenuselect=$null;reloadmenu}
        $GCred{Get-Cred;reloadadmenu}
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
        $Get_NewADObjects{$admenuselect=$null;Get-NewADObjects;reloadadmenu}
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
        $Decommission_ADPC{$admenuselect=$null;Decommission-ADPC;reloadadmenu}
        $TestFix_PCSecureChannel{$admenuselect=$null;TestFix-PCSecureChannel;reloadadmenu}
        
        default
        {
            $admenuselect = $null
            reloadmenu
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
Function Manage-ADUserFiles()
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
        # Example: \\domain.local\dfs\Net Share\Folder
        # Folder where icons are pinned to the TaskBar
        #  ls "C:\Users\%username%\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
        $NAmedFNLocation = Read-Host "`nEnter the path to the new file to copy (exclude trailing \);
            `nExample: \\server\\share\folder"
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
    <#
    $p1 = "Type AD to use active computers listed in Active Directory or ";
    $p2 = "custom ";
    $p3 = "to type your own comma separated list";
    $NC = "Yellow";chcolor $p1 $p2 $p3 $NC;$ADorCSL = Read-Host "`t`t"
    #>
        
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
    $PingOutCSV = "C:\Users\%username%\Documents\WindowsPowerShell\"+$Global:CDate+"PingOut.csv"
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
Function EmptyGroups ()
{
    Write-Verbose "Getting AD Groups that have no members"
    Get-ADGroup -Filter * | where { -Not ($_ | Get-ADGroupMember)} |select name |FT -Autosize
}

Function OldDCPCs()
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
    <# Where-object Name -eq "pcname" | ft #> 
    Where-object Enabled -eq $True | ft
    # Add the following line if you want the result saved into a .CSV file
    #| Export-Csv OLD_Computer.csv -NoTypeInformation
}

Function GroupID()
{
    $ADGM = Read-Host "Enter name of group of which you wish to see members. (ex. Domain Admins)"
    Write-Host "`nRunning:`tGet-ADGroupMember -Identity $ADGM | ft"
    Get-ADGroupMember -Identity $ADGM | ft
}
Function Associate-UserAndPC
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
Function GroupMembership()
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
Function ADPcArch()
{
    Clear-Host
    Get-PCList
    <#
    $pclc=Get-ADComputer -Filter * |select -ExpandProperty Name
    $pclc
    #>
    # Display the command that will be run for each pc listed in Active directory Get OS information from pc list
    Write-Host '$sOS = Get-WmiObject -class Win32_OperatingSystem -Impersonation Impersonate -Credential $cred -computername $pc'
    foreach($Global:pcLine in $Global:pclist)
    {
        Identify-PCName
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
        # Get OS information from pc list
            $sOS = Get-WmiObject -class Win32_OperatingSystem -Impersonation Impersonate -Credential $cred -computername $Global:pc
            foreach($sProperty in $sOS)
            {
                $sProperty |select @{Name="System Name"; e=$sProperty.PSComputerName},PSComputerName,@{Name="OS"; Expression={__Class}},__Class, Caption, Description, @{Name="Architecture"; Expression={OSArchitecture}},OSArchitecture
                Write-Output "`nPC: $sProperty.CSNAme" #PSComputerName"
                write-host "`nOS: "$sProperty.Caption
                write-host "Architecture: "$sProperty.OSArchitecture
                write-host "Registered to: "$sProperty.Description
                write-host "Version #: " $sProperty.ServicePackMajorVersion
                Write-Output "`n$t4$Script:ADd4"
            }
        }
    }
}
# Get and display a variety of printer information for listed system
Function Get-PrinterInfo() # Called from pcinfo function
{
# param([Parameter(Mandatory=$False)])  # Troubleshoot problem with parenthesis if necessary to enable param here
# Using the Get-PCList from pcinfo function
    # foreach($pcLine in $Global:PCList){ # Required if not being called from pcinfo
    # Get and display printer information for listed system
        Write-Host 'Running: `nGet-WmiObject -Class win32_printer -Credential $cred -ComputerName $Global:PC| select name, drivername, portname, shared, sharename |ft -AutoSize'
        # $prin = Get-WmiObject -Class win32_printer -Credential $cred -ComputerName $Global:PC| select name, drivername, portname, shared, sharename |ft -AutoSize
        $Prin = Get-WmiObject -Class Win32_printer -Credential $cred -ComputerName $Global:PC| select Name, drivername, portname, shared, Location, PrinterState, PrinterStatus, ShareName, SystemName, DetectedErrorState, ExtendedDetectedErrorState
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
        Write-Host 'get-printerport -computer $Global:PC|select ComputerName,Name,Description,PrinterHostAddress |ft -A'
        $prin2 = get-printerport -computer $Global:PC|select ComputerName,Name,Description,PrinterHostAddress |ft -A
        # get-printerport -computer $Global:PC|select Caption, CommunicationStatus,ComputerName,Description, DetailedStatus, ElementName, HealthState, InstallDate, InstanceID, Name, OperatingStatus, PortMonitor, PrimaryStatus, PSComputerName, Status, StatusDescriptions |fl
        $prin2
    # Get and display Printer Driver information for listed system
        Write-Host 'get-printerDriver -computer $Global:PC|select * |ft -A # ComputerName,Name,Description,PrinterHostAddress,OEMUrl,Monitor,Path,provider |ft -A'
        $prin3 = get-printerDriver -computer $Global:PC|select ComputerName,Name|Sort Name|ft -A # Manufacturer
        $prin3
    # }  # Required if not running it from pcinfo
        if(!$NoMenu)
        {
            reloadmenu
        }
}
Function pcinfo()
{
    Clear-Host
    Get-PCList
    foreach($Global:PCLine in $Global:PCList)
    {
        Identify-PCName
    # Get the Dell Service Tag # from pc list
            $PCSN = (get-wmiobject -Credential $cred -computername $Global:PC -class win32_bios).SerialNumber
            Write-Output 'Running: `nget-wmiobject -Credential $cred -computername $Global:PC -class win32_bios |select "$Global:PC", SerialNumber |fl'
            Write-Host "`n$Global:PC : $PCSN" -ForegroundColor Yellow

    # Get OS information from pc list
        $sOS = Get-WmiObject -class Win32_OperatingSystem -Impersonation Impersonate -Credential $cred -computername $Global:PC
            Write-Host 'Running: `nGet-WmiObject -class Win32_OperatingSystem -Impersonation Impersonate -Credential $cred -computername $Global:PC'
        $ObjectOut = foreach($sProperty in $sOS){
                $OBProperties = @{'pscomputername'=$Global:PCList;
                'OS'=$sProperty.Caption;
                'Arch'=$sProperty.OSArchitecture;
                'Registration'=$sProperty.Description;
                'SPVersion'=$sProperty.ServicePackMajorVersion}
        New-Object -TypeName PSObject -Prop $OBProperties
        }
        $objectOut |FT Arch,OS,SPVersion,Registration
        
    # Get CPU Information for listed system
        $cpuinfo = get-wmiobject win32_processor -Impersonation Impersonate -Credential $cred -ComputerName $Global:PC |select-object name, numberofcores,numberoflogicalprocessors,
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
         }},role
    # @{Label = "CPU Status" ;Expression={$cpuinfo}}
        Write-Host "Using WMI -class win32_processor for CPU information" # ($_.capacity / 1048576).ToString('C0')
        $cpuinfo |FT
    # Get and display memory information for listed system
        Write-Host 'Running: `nGet-WmiObject cim_physicalmemory -Impersonation 3 -Credential $cred -ComputerName $Global:PC | select banklabel,capacity,caption,devicelocator,partnumber,speed |ft -AutoSize'
        $mem = Get-WmiObject cim_physicalmemory -Impersonation 3 -Credential $cred -ComputerName $Global:PC |
         select caption,devicelocator,speed,@{Label = "Gigs" ;Expression={($_.capacity / 1048576).ToString('N0')}},partnumber,banklabel |ft -AutoSize
        $mem
    # Get and display printer information for listed system
        Write-Host 'Running: `nGet-WmiObject -Class win32_printer -Credential $cred -ComputerName $Global:PC | select name, drivername, portname, shared, sharename |ft -AutoSize'
        $prin = Get-WmiObject -Class win32_printer -Credential $cred -ComputerName $Global:PC | select name, drivername, portname, shared, sharename |ft -AutoSize
        $prin
    # Testing another Printer Info command
    Write-Output 'Running: Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Network=True" | select Sharename,name | export-csv -path $Use_MMFPath\$env:computername.csv -NoTypeInformation'
    Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Network=True" | select Sharename,name | export-csv -path $Use_MMFPath\$env:computername.csv -NoTypeInformation
    Write-Warning ".csv saved to: $Use_MMFPath\$env:computername.csv"
    #Get-PrinterInfo via function
    $NoMenu = $true
    Get-PrinterInfo $Global:pc
    $NoMenu = $false
    # Get Disk information for system
        Write-Host '`nRunning: get-wmiobject win32_logicaldisk -Credential $cred -Computername $Global:PC | fl  systemname, Name, deviceid, description, Drivetype, mediatype, filesystem, @{n='MB freespace';e={[string]::Format("{0:0,000} MB",$_.FreeSpace/1MB)}}, 
@{n='MB Size';e={[string]::Format("{0:0,000} MB",$_.Size/1MB)}}, volumedirty, volumename, volumeserialnumber'
        get-wmiobject win32_logicaldisk -Credential $cred -Computername $Global:PC | fl  systemname, Name, deviceid, description, Drivetype, mediatype, filesystem, @{n='MB freespace';e={[string]::Format("{0:0,000} MB",$_.FreeSpace/1MB)}}, 
@{n='MB Size';e={[string]::Format("{0:0,000} MB",$_.Size/1MB)}}, volumedirty, volumename, volumeserialnumber
    #Get attached monitor information
        Write-Output "Attached Monitor Information:"
        gwmi WmiMonitorID -Namespace root\wmi -ComputerName $Global:pc | ForEach-Object {($_.UserFriendlyName -notmatch 0 | foreach {[char]$_}) -join ""; ($_.SerialNumberID -notmatch 0 | foreach {[char]$_}) -join ""}
    #Get process from pc
        Write-Host "Running: wmi search to see if the logon processes is running...`n"
        $logonInfo = get-wmiobject win32_process -computername $Global:PC -Credential $cred| select Name,sessionId,Creationdate  |Where-Object {($_.Name -like "Logon*")} 
        $logonInfo
    #Get Network information from pc
        Write-host "Checking for NIC status"    
        Get-WmiObject -Class win32_networkadapterconfiguration -ComputerName $Global:PC -Credential $cred|where {$_.IPAddress -ne $null}|select PSComputerName,MACAddress,IPSubnet,DNSServerSearchOrder, @{n="IP Address";e={($_ | select -ExpandProperty IPAddress)-join ','}},DHCPEnabled|ft
    Write-Host "`t`t********************************************************************" -ForegroundColor Green -BackgroundColor DarkRed
    }
}

Function ListDomPCs()
{
    Get-ADComputer -filter 'objectclass -eq "Computer"' | select -ExpandProperty dnshostname
}
Function Find-UserPCs
{
    # Where Users's first and last names are in the description of a Computer object
    # this will list the User and the associated computer
    [string[]] $Names = ((Read-Host "Enter a comma separated list of all or part of People's first or last names").split(",") | %{$_.trim()})
    foreach ($LName in $Names) 
    {
        get-adcomputer -Filter * -Properties *|where {$_.enabled -eq $true -and $_.Description -like "*$LName*"}|ft Description,Name
    }
}
Function UsrOnPC ()
{
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
 Function SysEvents()
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
Function UsrsOnPCs()
{
# Report the Username logged into any active computer listed in Active Directory
$Global:pcList = Get-ADComputer -Filter 'Enabled -eq "True"'
Write-Host "Logged In Computer`tUser" -ForegroundColor Yellow
Write-Host "IPAddress `tMachine Name `tUser Logged on"
foreach ($Global:PCLine in $Global:pcList)
	{
	Identify-PCName
    $CompIP = $null
    $CompName = $null
    $ErrorActionPreference = "silentlycontinue"
    $CompIP = ([system.net.dns]::GetHostAddresses($Global:pc)).IPAddressToString
    # Write-host "CompIP $CompIP pre-if" -ForegroundColor Red
    If ($CompIP)
        {
            $ping = test-connection $Global:pc -Count 1 -Quiet -ErrorAction SilentlyContinue
            # Verify Computername, hostname, with IP
            $CompName = [System.Net.Dns]::GetHostByAddress($CompIP).HostName
            If ($CompName -notlike "*$Global:pc*")
                {
                    Write-host "Computer Name $CompName is invalid for $Global:pc" -ForegroundColor Red
                }
        }
    Else{
            $ping = $false
            $CompIP = "___.___.___.___"
            # Write-host "No IP $CompIP found for: for $Global:pc" -foreground Red
        }
    # $ping = test-connection $Global:pc -Count 1 -Quiet -ErrorAction SilentlyContinue

  	if($ping -eq $true)
        {
            $pcinfo = @()
		    #Get explorer.exe processes
		    $proc = gwmi win32_process -computer $Global:pc -Filter "Name = 'explorer.exe'"
		    #Search collection of processes for username
		    ForEach ($p in $proc) {
	    	    $usr = ($p.GetOwner()).User
                $pscn = $p.pscomputername
                $pgo = $p.GetOwner().user
                Write-Output "$CompIP`t$pscn`t$pgo" |ft
           
		    }
        }
        else
        {
            Write-Host "$CompIP`t$Global:pc`tUNK/NONE" -ForegroundColor Red
        }
    }
        $ErrorActionPreference = "continue"
}
Function alladusers()
{
    # $ADUserList = get-aduser -Filter * |Select-Object SamAccountName, Name # |Export-Csv 20151015ADUsers.csv
    $ADUserList = Start-job -Name BuildADUserList -ScriptBlock {get-aduser -Filter * -Properties * |Select-Object SamAccountName, Name,LastLogonDate,Created,mail,address}|wait-job |Receive-Job # |Export-Csv 20151015ADUsers.csv
    # wait-job -Name BuildADUserList
    Start-Sleep 1
    $ADUserList |ft SamAccountName,Name,LastLogonDate,Created,mail,address -AutoSize
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

Function Get-ADPrivGrpMbrshp()
{
        Invoke-Expression ".\Get-ADPrivGrpMbrshp.ps1"
}
Function Disable-ADUser
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
Function SearchExchMbox()
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

Function GPReport()
{
    Clear-Host
    $PSScriptRoot
        $CalledFromFuction = $True
    $rptpath = "$PSScriptRoot\GPReport.htm"
    UsrOnPC
    # $Global:PC = Read-Host "Type the name of a PC"
    $User = Read-Host "Enter the username of the policy holder"
    Get-GPResultantSetOfPolicy -Computer $Global:PC -User $User -reporttype html -path $rptpath
    start chrome.exe $rptpath
    $CalledFromFuction = $False
}

Function Get-RDPStats()
{
Clear-Host
Write-host "Searching for active connections to servers"
$InvkCmd=$null
$Active=$false
$SN=$null
$servers = Get-ADComputer -Filter {OperatingSystem -like "*server*" -and Enabled -eq $true}
# $servers = $servers.trim("^\s+")
foreach ($server in $servers)
    {
        $SN = $Server.Name        
        if (Test-Connection $sn -Count 1 -ErrorAction SilentlyContinue )
            {
            $InvkCmd = invoke-command {query user /server:$sn}
            #$InvkCmd = invoke-command { qwinsta /server:$sn} 
            if ($InvkCmd -like "*Active*")
                {
                    Write-host "There is an active Session to $SN`n" -ForegroundColor Green
                    $InvkCmd
                    ""
                    $Active=$true
                    start-sleep -Milliseconds 500
                }
                elseif ($InvkCmd -like "*Disc*")
                {
                $sn + $InvkCmd
                    start-sleep -Milliseconds 500
                }
            }

            else 
            {
                "$SN is an AD server but cannot be reached "
                start-sleep -Milliseconds 500
            }
    }
If ($Active -eq $false)
    {
        Write-Host "No active RDP sessions found" -ForegroundColor Red
        start-sleep -Milliseconds 500
    }
}
Function Get-ADPcStarts()
{
    Clear-Host
    $CredUser = $Cred.UserName
    if ((Read-Host "Current credentials are for: `"$CredUser`" , type C to change or [Enter] to continue")-eq "C") {Get-Cred}
    Get-PCList
    # $Global:PCList=Get-ADComputer -Filter 'enabled -eq "true"' -Properties Enabled,Name,IPv4Address |Where {$_.ipv4address -ne $null} |select Name,IPv4Address
    $CSVFileNam = $CDate+"startups.csv"
    Write-host "File will be exported to $CSVFileNam" -ForegroundColor Cyan
    # Display the command that will be run for each pc listed in Active directory to get StartUp information from PCs in list
    Write-Host '$sStart = gwmi -Class win32_startupcommand -Impersonation Impersonate -Credential $cred -ComputerName $Global:pc |Select-Object PSComputerName,Command,Description,Location, Name,User' # ,UserSID
    foreach($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        # Only attempt to get information from computers that are reachable with a ping
        If (Test-Connection $Global:pc -Count 1 -Quiet -ErrorAction SilentlyContinue)
        {
            # Get Startup information from pc list
            $sStart = gwmi -Class win32_startupcommand -Impersonation Impersonate -Credential $cred -ComputerName $Global:pc |Select-Object PSComputerName,Command,Description,Location, Name,User # ,UserSID
            $CSVOutObj += $sStart
        }
    }
    $CSVOutObj |Select-Object @{n="System_Name";e="PSComputerName"},@{n="Relative_Path";e="__RelPath"},@{n="Local_Path";e="__Path"},Command,Description,Location,Name,User,UserSID| Export-Csv $CSVFileNam -NoClobber -NoTypeInformation
    Start-Sleep 1            

}
Function get-OpenFiles()
{
    Clear-Host
    Get-PCList
    foreach ($pcLine in $Global:PCList)
    {
        Identify-PCName # possibly need to bring this function into this script if it won't run from get-systemslist.ps1
        $srvobj = invoke-command {c:\Windows\System32\openfiles.exe /query /s $pc /fo csv /V} |convertfrom-csv |
        select Hostname,ID,Type,@{Name="UserName"; Expression="Accessed By"},@{Name="Path_FileName"; Expression="Open File (Path\executable)"} |
        Where {$_.Path_FileName -ne "D:\\"}|
        sort Path_FileName |ft
        $srvobj
    }
}

# Get-Shortcut called from Replace-Links function
Function Get-Shortcut {
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

      New-Object PSObject -Property $info
    }
  }
}
Function Set-Shortcut {
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
Function Replace-Links()
{
<#
.EXAMPLE
    Use the following example to assign F11 hotkey to PowerShell (uncomment line, make sure you have full admin privileges: 
    Get-Shortcut | Where-Object { $_.LinkPath -like '*Windows PowerShell.lnk' } | Set-Shortcut -Hotkey F11 
#>
param(
  [Parameter(ValueFromPipelineByPropertyName=$true)]
  $PathToLink,
  $value0 = (Read-Host "Enter new target Drive Letter in the following format L:\"),
  $DLValue = (Read-Host "Enter old target Drive Letter in the following format O:\")
  )

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
        $alldsktpLnkPropertiesOut = "C:\Users\Public\Documents\"+$Global:CDate+"AllDsktpLnkPropertiesOut.csv"
        $LSc |select * |export-csv -Path $alldsktpLnkPropertiesOut -Append -NoTypeInformation # -Verbose
        # $desktoplnkPropertiesOut = "C:\Users\Public\Documents\"+$Global:CDate+"DesktopLinkProperties.csv"
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
            <#This line workd, testing the line below with variable $DLValue#> #$Pattern1 = '^(S:\\)(.*?)$' #'^($DLValue\\)(.*?)$' (w/UNC: '^(\\\\server\\Net\sShare\\)(.*?) *$'
            $Pattern1 = "^($DLValue\)(.*?)$" # (w/UNC: '^(\\\\server\\Shared\sFiles\\)(.*?) *$'
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

Function iterate-Profiles()
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
Function Replace-DesktopFiles()
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
Function List-MappedDrives()
{
    # from calling function: $proc = gwmi win32_process -computer $pscn -Filter "Name = 'explorer.exe'"
    # make $userdir variable equal to current user SamAccountName
    $domain = $env:USERDOMAIN
    $userdir = $pgo
    $psdname = "$pscn"+"hku"
    $npsd = New-PSDrive -PSProvider Registry -Name $psdname -root HKEY_USERS
    # comment out this foreach and the next $userdir variable declaration to only apply the following action to the currently loged on user
    $ProfileList = gci "\\$pscn\C$\Users\" -Directory -Force
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
    $pc = $null

}
# list AD objects where "WhenCreated" date is greater than ...
Function Get-NewADObjects
{
$WCDate = Read-Host "Enter a start date using the format mm/dd/yyyy to list AD Objects created since this date."

(get-adobject -filter * -properties *|where {$_.WhenCreated -gt $WCDate}) |select Name,canonicalname,whencreated,modified |sort WhenCreated,modified
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
        # Move the system to the destination target OU
        $ADPCObj|Move-ADObject -TargetPath $DestTargetOU -Confirm
        # Pre-Pend D_ to computer object description
        $ADPCO_D = $ADPCObj.Description
        $ADPCO_D = "D_"+$ADPCO_D
        # Disable the computer object and change its description
        Set-ADComputer $ADPCObj -Description $ADPCO_D -Enabled $false -Confirm
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