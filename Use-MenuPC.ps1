<#
.NAME
    Use-MenuPC.ps1
.SYNOPSIS
    Provide menu system for identifying PC Information
.DESCRIPTION
 Menu design initially adopted from : https://quickclix.wordpress.com/2012/08/14/making-powershell-menus/
 2020-09-02

 Written by Jason Crockett - Laclede Electric Cooperative

.PARAMETERS
    Placeholder
.EXAMPLE
    "a command with a parameter would look like: (.\Use-MenuPC.ps1)"
        .\Use-MenuPC.ps1 []
.SYNTAX
    .\Use-MenuPC.ps1 []
.REMARKS
    To see the examples, type: help Use-MenuPC.ps1 -examples
    To see more information, type: help Use-MenuPC.ps1 -detailed"
.TODO
    Remove Windows Experience Index item
    Modify the Get-PCProcess function to query processes other than just the logon process
    Modify the Get-PCProcess function to create an object out of the response and clean up the output based on this
    Modify the Call-Get-SysEvents function to be a script called from the multi-threaded script
#>

Clear-Host

$Use_MPCPath = Split-Path -Parent $PSCommandPath
$Script:PCt4 = "`t`t`t`t"
$Script:PCd4 = "------------Use-PC Menu ($Global:PCCnt Systems currently selected)----------------"
$Script:PCNC = $null
$Script:PCp1 = $null
$Script:PCp2 = $null
$Script:PCp3 = $null
[int]$PCmenuselect = 0

Function chPCcolor($Script:PCp1,$Script:PCp2,$Script:PCp3,$Script:PCNC){
            
            Write-host $Script:PCt4 -NoNewline
            Write-host "$Script:PCp1 " -NoNewline
            Write-host $Script:PCp2 -ForegroundColor $Script:PCNC -NoNewline
            Write-host "$Script:PCp3" -ForegroundColor White
            $Script:PCNC = $null
            $script:p1 = $null
            $script:p2 = $null
            $script:p3 = $null
}

Function PCmenu()
{
    while ($PCmenuselect -lt 1 -or $PCmenuselect -gt 22)
    {
        Trap {"Error: $_"; Break;}        
        $PCMNum = 0;Clear-host |out-null
        Start-Sleep -m 50
        if ($Global:PCCnt -eq $null) {$Global:PCCnt = 0}Else {$Global:PCCnt = ($Global:PCList).Count}
        $Script:PCd4 = "------------Use-PC Menu ($Global:PCCnt Systems currently selected)----------------"
        Write-host $Script:PCt4 $Script:PCd4 -ForegroundColor Green
# Exit to Main Menu
        $PCMNum ++;$PCMExit=$PCMNum;
            $Script:PCp1 = " $PCMNum. `t";
            $Script:PCp2 = "Exit to Main Menu";
            $Script:PCp3 = " ";
            $Script:PCNC = "Red";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# Prompt for credential with permissions across the network
        $PCMNum ++;$GCred=$PCMNum;$Script:PCp1 =" $PCMNum.  `t Enter";
            $Script:PCp2 = "Credentials ";$Script:PCp3 = "for use in script."; $Script:PCNC = "Cyan";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# $Get_PCList
        $PCMNum ++;$Get_PCList=$PCMNum;$Script:PCp1 =" $PCMNum.  `t Select";
            $Script:PCp2 = "systems to use ";$Script:PCp3 = "with commands on menu ($Global:PCCnt Systems currently selected).";
            $Script:PCNC = "Cyan";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# PC Informational Menu
        $PCMNum ++;$Use_MenuPCInfo=$PCMNum;$Script:PCp1 = " $PCMNum. `t View ";
            $Script:PCp1 = " $PCMNum. `t Go to the";
            $Script:PCp2 = "PC Informational";
            $Script:PCp3 = " menu"
            $Script:PCNC = "DarkCyan";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# PC Applications Menu
        $PCMNum ++;$Use_MenuPCApp=$PCMNum;
            $Script:PCp1 = " $PCMNum. `t Go to the";
            $Script:PCp2 = "PC Applications";
            $Script:PCp3 = " menu"
            $Script:PCNC = "DarkCyan";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# PC Services Menu
        $PCMNum ++;$Use_MenuPCSvc=$PCMNum;
            $Script:PCp1 = " $PCMNum. `t Go to the";
            $Script:PCp2 = "PC Services";
            $Script:PCp3 = " menu"
            $Script:PCNC = "DarkCyan";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# Domain PCs that haven't logged in for X number of  days
        $PCMNum ++;$Get_OldDCPCs = $PCMNum;
		    $Script:PCp1 =" $PCMNum. `t Domain PCs:";
		    $Script:PCp2 = "##? days ";
		    $Script:PCp3 = "since previous logon"; $Script:PCNC = "yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# Wait for live port
        $PCMNum ++;$Test_CritSysLive=$PCMNum;
            $Script:PCp1 = " $PCMNum. `t Check ";
            $Script:PCp2 = "Active Status";
            $Script:PCp3 = " for critical systems";
            $Script:PCNC = "Yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# View PC Startups
        $PCMNum ++;$Get_ADPcStarts = $PCMNum;
            $Script:PCp1 =" $PCMNum. `t View ";
		    $Script:PCp2 = "Startups ";
		    $Script:PCp3 = "on PCs "; $Script:PCNC = "yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# View PC Startups
        $PCMNum ++;$Retrieve_BitlockerKey = $PCMNum;
            $Script:PCp1 =" $PCMNum. `t Retrieving ";
		    $Script:PCp2 = "BitLocker Recovery Key ";
		    $Script:PCp3 = "information from selected PCs "; $Script:PCNC = "yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# View Registry property value
        $PCMNum ++;$Manage_HKLM = $PCMNum;
            $Script:PCp1 =" $PCMNum. `t View ";
		    $Script:PCp2 = "Registry Property ";
		    $Script:PCp3 = "Value(s) "; $Script:PCNC = "yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# Reset Internet Explorer 
        $PCMNum ++;$Reset_IE = $PCMNum;
            $Script:PCp1 =" $PCMNum. `t Reset ";
		    $Script:PCp2 = "Internet Explorer (IE) ";
		    $Script:PCp3 = "settings "; $Script:PCNC = "yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# Event logs
        $PCMNum ++;$Call_Get_SysEvents = $PCMNum;
		    $Script:PCp1 =" $PCMNum. `t View";
		    $Script:PCp2 = "Event Logs ";
		    $Script:PCp3 = ""; $Script:PCNC = "yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# Check for Active RDP sessions on Domain servers
        $PCMNum ++;$Get_RDPStats=$PCMNum;
            $Script:PCp1 = " $PCMNum. `t Check ";
            $Script:PCp2 = "for Console or RDP";
            $Script:PCp3 = " sessions on Domain servers or select systems. (May require RunAs Admin...";
            $Script:PCNC = "Yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
    # Enable RDP on a system
        $PCMNum ++;$Enable_RDP=$PCMNum;
            $Script:PCp1 = " $PCMNum. `t";
            $Script:PCp2 = "Enable RDP";
            $Script:PCp3 = " on a Windows system"
            $Script:PCNC = "Yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
    # Process Information
        $PCMNum ++;$Get_PCProcess=$PCMNum;
            $Script:PCp1 = " $PCMNum. `t View ";
            $Script:PCp2 = "Logon Process ";
            $Script:PCp3 = "Information"
            $Script:PCNC = "Yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
#    "Discover boot times (from select systems)"
        $PCMNum ++;$Get_BootTime=$PCMNum;
            $Script:PCp1 = " $PCMNum. `t Discover ";
            $Script:PCp2 = "boot times -MultiThreaded ";
            $Script:PCp3 = "(Requires Run Script As Administrator)"
            $Script:PCNC = "Yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
        # List all domain member PCs
        $PCMNum ++;$List_DomPCs = $PCMNum;
		    $Script:PCp1 =" $PCMNum. `t List all";
		    $Script:PCp2 = "member PCs ";
		    $Script:PCp3 = "in this domain"; $Script:PCNC = "yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# Screen saver status on pc to see if a user is AFK
        $PCMNum ++;$Active_ScrnSav=$PCMNum;
            $Script:PCp1 = " $PCMNum. `t Check ";
            $Script:PCp2 = "screen saver status ";
            $Script:PCp3 = "on pc to see if a user is AFK"
            $Script:PCNC = "Yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# Get Scheduled Task Information
        $PCMNum ++;$Get_ScheduledTaskInformation=$PCMNum;
            $Script:PCp1 = " $PCMNum. `t Get a system's ";
            $Script:PCp2 = "Scheduled Task ";
            $Script:PCp3 = "Information"
            $Script:PCNC = "Yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# View-Certificates
        $PCMNum ++;$View_Certificates=$PCMNum;
            $Script:PCp1 = " $PCMNum. `t View ";
            $Script:PCp2 = "Certificates ";
            $Script:PCp3 = "on a system (possibly requires runas Administrator to include local system)"
            $Script:PCNC = "Yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
# View WinSAT / Windows Experience Index...
        $PCMNum ++;$Get_WinSAT = $PCMNum;
            $Script:PCp1 =" $PCMNum. `t View ";
		    $Script:PCp2 = "WinSAT ";
		    $Script:PCp3 = "/ Windows Experience Index... Function not found "; $Script:PCNC = "yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
        # Setting up the menu in this way allows you to move menu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $PCMNum ++;$Script:PCp1 =" $PCMNum.  FIRST";$Script:PCp2 = "HIGHLIGHTED ";$Script:PCp3 = "END"; $Script:PCNC = "yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
        Write-Host "$Script:PCt4 $Script:PCd4"  -ForegroundColor Green
        $Script:PCMNum = 0;[int]$PCmenuselect = Read-Host "`n`tPlease select the menu item you would like to do now"
    }
    Function Nullify-Prompts
    {
        $Global:BSN = $null # $null
        $Global:OSI = $null
        $Global:CPUI = $null
        $Global:MI = $null
        $Global:DI = "No"
        $Global:NICS = "No"
        $Global:PDI = $null
        $Global:PCP = $null
        $Get_AllPCInfo = "No"
    }
    Function Confirm-Prompts
    {
        $Global:BSN = "Yes"
        $Global:OSI = "Yes"
        $Global:CPUI = "Yes"
        $Global:MI = "Yes"
        $Global:DI = "Yes"
        $Global:NICS = "Yes"
        $Global:PDI = "Yes"
        $Global:PCP = "Yes"
    }
    # Update with text explanation of what command will be run
    $Global:BSNCMD = 'Running: `nget-wmiobject -Credential $cred -computername $Global:PC -class win32_bios |select "$Global:PC", SerialNumber |fl'
    $Global:OSICMD = ""
    $Global:CPUICMD = ""
    $Global:MICMD = ""
    $Global:DICMD = ""
    $Global:NICSCMD = ""
switch($PCmenuselect)
    {
    $PCMExit{$PCmenuselect=$null}
    $GCred{Clear-Host;Get-Cred;$PCmenuselect = $null;reload-PromptPCmenu} # Called from Run-PSMenu.ps1
    $Get_PCList{Clear-Host;Confirm-Prompts;Get-PCList;reload-PCmenu}
    $Use_MenuPCInfo{. .\Use-MenuPCInfo.ps1;PCIMenu}
    $Use_MenuPCApp{. .\Use-MenuPCApp.ps1;PCAppMenu}
    $Use_MenuPCSvc{. .\Use-MenuPCSvc.ps1;PCSvcMenu}
    $Get_OldDCPCs{$PCmenuselect = $null;Get-OldDCPCs;reload-PromptPCmenu}
    $Test_CritSysLive{Test-CritSysLive;reload-PromptPCmenu}
    $Get_ADPcStarts{$PCmenuselect = $null;Get-ADPcStarts;reload-PromptPCmenu}
    $Retrieve_BitlockerKey{$PCmenuselect = $null;Retrieve-BitlockerKey;reload-PromptPCmenu}
    $Manage_HKLM{Manage-HKLM;reload-PromptPCmenu}
    $Reset_IE{$PCmenuselect = $null;Reset-IE;reload-PromptPCmenu}
    $Call_Get_SysEvents{$PCmenuselect = $null;Call-Get-SysEvents $Global:PCList $Global:Cred;reload-PromptPCmenu}
    $Get_RDPStats{Get-RDPStats;reload-PromptPCmenu}
    $Enable_RDP{Enable-RDP;reloadmenu}
    $Get_PCProcess{Clear-Host;Nullify-Prompts;Write-Output $Global:PCPCMD;$Global:PCP = "Yes";Get-PCInfo $Global:PCP <#Get-PCProcess#>;reload-PromptPCmenu}
    $Get_BootTime{Clear-Host;.\Start-MultiThreads.ps1 .\Get-BootTime.ps1  -ScriptArgs $Global:PCList;reload-PromptPCmenu} # Called Externally
    $List_DomPCs{$PCmenuselect = $null;List-DomPCs;reload-PromptPCmenu}
    $Active_ScrnSav{Active-ScrnSav;reload-PromptPCmenu}
    $Get_ScheduledTaskInformation{Get-ScheduledTaskInformation;reload-PromptPCmenu}
    $View_Certificates{View-Certificates;reload-PromptPCmenu}
    $Get_WinSAT{$PCmenuselect = $null;get-winSAT;reload-PromptPCmenu}

        default
        {
            $PCmenuselect = $null
            reload-PromptPCmenu
        }
    }    
}

Function reload-PCmenu() # Reload with NO prompt (not needed until there's another menu down)
{
        $PCmenuselect = $null
        PCmenu
}
Function reload-PromptPCmenu() # Reload with prompt (not needed until there's another menu down)
{
        Read-Host "`nHit Enter to continue"
        $PCmenuselect = $null
        PCmenu
            $Script:PCp1 = " $PCMNum. `t View ";
            $Script:PCp2 = "Certificates ";
            $Script:PCp3 = "on a system (possibly requires runas Administrator to include local system)"
            $Script:PCNC = "Yellow";chPCcolor $Script:PCp1 $Script:PCp2 $Script:PCp3 $Script:PCNC
}

Function Call-Get-SysEvents()
{ 
    $WinLogName= Read-host "Enter Application, Security, Setup, or System"
    $EntryLevel_Type= Read-host " Enter Critical, Error, Warning, Information, FailureAudit, SuccessAudit"
    $cnt = Read-host "How many logs do you want to see of this type?" 
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    foreach ($Global:PCLine in $Global:PCLIst)
    {
    Identify-PCName
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
            if ($Global:pclist -eq "") {$Global:pclist = [environment]::machinename}
            if ($WinLogName -eq "") {$WinLogName = "System"} 
            if ($EntryLevel_Type -eq "") {$EntryLevel_Type = "Error"}
            if ($cnt -eq "") {$cnt = 50}
            Invoke-Expression ".\Get-SysEvents.ps1 -PCList $Global:PCList -WinLogName $WinLogName -EntryLevel_Type $EntryLevel_Type -cnt $cnt"
        }
    }
}

Function Get-ADPcStarts()
{
    Clear-Host
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    # $Global:PCList=Get-ADComputer -Filter 'enabled -eq "true"' -Properties Enabled,Name,IPv4Address |Where {$_.ipv4address -ne $null} |select Name,IPv4Address
    $CSVFileNam = $CDate+"startups.csv"
    Write-host "File will be opened in GridView and exported to $CSVFileNam" -ForegroundColor Cyan
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
    $CSVOutObj |Select-Object @{n="System_Name";e="PSComputerName"},@{n="Relative_Path";e="__RelPath"},@{n="Local_Path";e="__Path"},Command,Description,Location,Name,User,UserSID| Export-Csv $CSVFileNam -NoClobber -NoTypeInformation -Append
    $CSVOutObj |Select-Object @{n="System_Name";e="PSComputerName"},@{n="Relative_Path";e="__RelPath"},@{n="Local_Path";e="__Path"},Command,Description,Location,Name,User,UserSID|Out-GridView -Title $CSVFileNam
    Start-Sleep 1            

}

Function Enable-RDP()
{
#   $Cred = Get-Credential
    Write-Warning "Enter PC to enable RDP on"
    Get-PCList
    foreach ($Global:pcLine in $Global:PClist)
    {
        Identify-PCName
        (Get-WmiObject Win32_TerminalServiceSetting -ComputerName $Global:PC -Credential $Cred -Namespace root\cimv2\TerminalServices).SetAllowTsConnections(1,1)|Out-Null
        (Get-WmiObject -class "Win32_TSGeneralSetting" -ComputerName $Global:PC -Credential $Cred -Namespace root\cimv2\TerminalServices -Filter "TerminalName='RDP-tcp'").SetUserAuthenticationRequired(0)|Out-Null
    }
    Start-Sleep -Seconds 5

}

Function Get-pcinfo($Global:BSN, $Global:OSI, $Global:CPUI, $Global:MI, $PI, $PPI, $Global:DI, $Global:NICS, $Global:PDI, $Global:PCP)
{
    Write-Warning "Getting system information from the following systems:"
    foreach($Global:pcLine in $Global:PCList)
    {
        Identify-PCName # Called from Run-PSMenu.ps1
        if ($Global:BSN -eq "Yes"){Get-BIOSSerialNumber}
        if ($Global:OSI -eq "Yes"){Get-ArchOSInfo}
        if ($Global:CPUI -eq "Yes"){Get-CPUInfo;Start-Sleep -Seconds 2}
        if ($Global:MI -eq "Yes"){Get-MemInfo}
        if ($Global:DI -eq "Yes"){Get-DInfo}
        if ($Global:NICS -eq "Yes"){Get-NICStatus}
        if ($Global:PCP -eq "Yes"){Get-PCProcess}
    }
    Nullify-Prompts
}

Function Get-ProcessStatus()   # not ready yet!!!
{
    Clear-Host
    Write-Warning "From which computer(s) do you want to retrieve the status of the process?"
    Get-PCList
    foreach ($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        $ReadProc = Read-Host "`nWhat is the ID number of the process you wish to get information about? "
        if ($readproc -match "^[0-9]+$")
        {
            $gwmio32p = Get-WmiObject win32_process -ComputerName $Global:PC |Where-Object {$_.ProcessID -eq $ReadProc} |Select-Object Path
            # param($id=$pid)
            $gwmio32p
            Get-WmiObject -Class win32_process -Filter "ParentprocessId=$id"

        }
        else
        {
            $gwmio32p = Get-WmiObject win32_process -ComputerName $Global:PC -Credential $cred| Where-Object {$_.ProcessName -eq $ReadProc}
            $gwmio32p  | Select-Object PSComputerName, ProcessID, ProcessName, Path |ft
            # param($id=$pid)
            $gwmio32p
            Get-WmiObject -Class win32_process -Filter "ParentprocessId=$id"
            # $gwmichild = Get-WmiObject win32_process -ComputerName $Global:PC -Credential $cred| Where-Object {$_.ProcessName -eq $ReadProc} -filter "ParentProcessID=$ID"

        }
    }
}

# Get and display Printer Driver information for listed system
Function Get-PrinterDriverInfo
{
    Write-Host 'get-printerDriver -computer $Global:PC |select * |ft -A # ComputerName,Name,Description,PrinterHostAddress,OEMUrl,Monitor,Path,provider |ft -A'
    $prin3 = get-printerDriver -computer $Global:PC |select ComputerName,Name|ft -A
    $prin3
}
#Get process from pc
Function Get-PCProcess
{
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    Else
    {
    Write-Warning "Getting Logon Process information from the following list : $Global:PCList"
    }
    foreach ($Global:PCLine in $Global:PCList)
    {
        
        Identify-PCName # Called from Run-PSMenu.ps1
        If (Test-Connection $Global:pc -Count 1 -Quiet -ErrorAction SilentlyContinue) # Only attempt to get information from computers that are reachable with a ping
        {
            $logonInfo = get-wmiobject win32_process -computername $Global:PC -Credential $cred| select Name,sessionId,Creationdate  |Where-Object {($_.Name -like "PrintIsolationHost.exe*")} # Logon*")} 
            Write-Host "`nRunning: wmi search to see if the logon processes is running on $Global:PC ...`n"
            if ($logonInfo){$logonInfo} Else {Write-Output "Logon process not found for $Global:PC"}
        }
    }
}
Function Test-CritSysLive()
{
    Clear-Host
    $err=@()
    Write-Host "`t.\Critsys.txt contains a list of critical systems to check with this function." -ForegroundColor DarkYellow
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
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
}
Function Get-RDPStats()
{
# Compare with code at: https://superuser.com/questions/587737/powershell-get-a-value-from-query-user-and-not-the-headers-etc#702328
Clear-Host
$InvkCmd=$null
$Active=$false
$SN=$null
# Use the following if we want to retain a single PCList across many menu functions 
if ($Global:PCList -eq $null)
{
    Get-PCList
}
Write-host "Searching for Console or RDP connections to selected systems"
    $ObjectOut = foreach ($Global:PCLine in $Global:PCList)
        {
            Identify-PCName

            if (Test-Connection $Global:PC -Count 1 -ErrorAction SilentlyContinue )
                {
                $InvkCmd = invoke-command {query user /server:$Global:PC 2>$1} -ErrorAction SilentlyContinue
                if ($InvkCmd -like "*Active*")
                    {
                        Write-host "$Global:PC : Active" -ForegroundColor Green
                        $Active=$true
                        $IC = (($InvkCmd) -replace '\s{2,}', ','|convertfrom-csv)
                        $ICUserName = $IC.Username
                        $ICSessionName = $IC.sessionname
                        $ICID = $IC.id
                        $ICState = $IC.state
                        $ICIdleTime = $IC.'idle time'
                        $ICLogonTime = $IC.'logon time'
                        # $ICLine = $IC
                    $OBProperties = @{'pscomputername'=$Global:PC;
                        'UserName'="$ICUserName";
                        'SessionName'="$ICSessionName";
                        'ID'="$ICID";
                        'State'="$ICState";
                        'Idle_Time'="$ICIdleTime";
                        'Status'="Active ";
                        'Logon_Time'="$ICLogonTime"}

                        
                        start-sleep -Milliseconds 500
                    }
                elseif ($InvkCmd -like "*Disc*")
                    {
                        Write-host "$Global:PC : Disconnected" -ForegroundColor Yellow
                        # This Regex is not perfect since the header has SessionName in it but in a disconnected result
                        # there are only spaces in that area.
                        $IC = (($InvkCmd) -replace '\s{2,}', ','|convertfrom-csv)
                        $ICUserName = $IC.Username
                        $ICSessionName = ""
                        $ICID = $IC.sessionname # header is offset due to missing session name
                        $ICState = $IC.id # header is offset due to missing session name
                        $ICIdleTime = $IC.state # header is offset due to missing session name
                        $ICLogonTime = $IC.'idle time' # header is offset due to missing session name
                        $IC.'logon time' # header is offset due to missing session name
                        # $ICLine = $IC
                    $OBProperties = @{'pscomputername'=$Global:PC;
                        'UserName'="$ICUserName";
                        'SessionName'="$ICSessionName";
                        'ID'="$ICID";
                        'State'="$ICState";
                        'Idle_Time'="$ICIdleTime";
                        'Status'="Disconnected ";
                        'Logon_Time'="$ICLogonTime"}

                        start-sleep -Milliseconds 500
                    }
                else
                    {
                        Write-warning "$Global:PC : No session detected"

                    $OBProperties = @{'pscomputername'=$Global:PC;
                        'UserName'="";
                        'SessionName'="";
                        'ID'="";
                        'State'="";
                        'Idle_Time'="";
                        'Status'="No session detected ";
                        'Logon_Time'=""}

                        start-sleep -Milliseconds 500
                    }
                }

            else 
                {
                    Write-Warning "$Global:PC is a server but cannot be reached "

                $OBProperties = @{'pscomputername'=$Global:PC;
                    'UserName'="";
                    'SessionName'="";
                    'ID'="";
                    'State'="";
                    'Idle_Time'="";
                    'Status'="$Global:PC is a server but cannot be reached ";
                    'Logon_Time'=""}

                    start-sleep -Milliseconds 500
                }
            New-Object -TypeName PSObject -Prop $OBProperties
        }
    If ($Active -ne $true)
    {
        Write-Host "`nNo active RDP sessions found" -ForegroundColor Red
        start-sleep -Milliseconds 500
    }
    $ObjectOut |FT PSComputerName,ID,UserName,State,Idle_Time,Logon_Time,SessionName,Status -AutoSize -Wrap
}

Function Get-OldDCPCs() # View AD Computer last logon times
{
# adopted from module by Tilo 2013-08-27
# https://gallery.technet.microsoft.com/scriptcenter/Get-Inactive-Computer-in-54feafde
    Clear-Host
    Import-Module activedirectory
    $domain = $env:USERDOMAIN
    $DaysInactive = Read-host "`nInactive for how many days?"
    $time = (Get-Date).AddDays(-($DaysInactive))

# Search for all AD computers with lastLogonTimestamp less than our time
    $InactivePCs = Get-ADComputer -filter {lastlogontimestamp -lt $time} -Properties lastlogontimestamp |

# output hostname and lastlogonTimestamp into a CSV
    Select-Object Name,@{Name="Time Stamp"; Expression={[DateTime]::FromFileTime($_.lastlogonTimeStamp)}}, Enabled |
    Sort-object Enabled, "Time Stamp" |
    <# Where-object Name -eq "pcname" | ft #> 
    Where-object Enabled -eq $True
    $InactivePCs |ft
    Write-Warning "Captured the list to the $Global:PCList variable for use with other functions"
    # Add the following line if you want the result saved into a .CSV file
    #| Export-Csv OLD_Computer.csv -NoTypeInformation
    # Capture the list to the $Global:PCList variable for use with other functions
    $Global:PCList = $InactivePCs|select Name # Not really needed since the functions work only on live systems
}

Function List-DomPCs()
{
    Clear-Host
    Get-ADComputer -filter 'objectclass -eq "Computer"' -Properties dnshostname,Enabled,SamAccountName,UserPrincipalName,Name,DistinguishedName|Where {$_.Enabled -eq $true}| select @{n="FQDN";e={($_ | select -ExpandProperty dnshostname)-join ','}},Name,Enabled,SamAccountName,DistinguishedName|sort FQDN|ft -autosize
}

Function Active-ScrnSav()
{
Clear-Host
Write-Warning "An active screen saver means the user is AFK"
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    foreach ($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
            # Search for the screen saver process on a pc as a way of seeing if a user is active
             $scrpresent = Get-WmiObject Win32_process -ComputerName $Global:PC -credential $cred |Where-Object {($_.processname -like "scrn*" )} |Select-Object processname,@{n='CreatedTime';e={$_.ConvertToDateTime($_.CreationDate)}}
          
             if ($scrpresent.processname -like "scrn*") 
             {
                Write-Host "`nThe screen saver has been active on $Global:PC since "$scrpresent.createdtime "`n"
             }
             else
             {
                Write-Host "The screen saver is not active on $Global:PC"
             }

             $getwmiobject = Get-WmiObject -class Win32_computersystem -computername $Global:PC -Credential $Cred
             $Username = $Getwmiobject.username

             if($UserName -eq $NULL) {
             Write-host "`nThere is no Current user logged onto $Global:PC`n"
             }
                else {write-host "$Username is logged in`n"
             }
         }
    }
}

Function Get-ScheduledTaskInformation
{
    Clear-Host
    # http://powershell-guru.com/powershell-tip-56-convert-to-and-from-hex/
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    # https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-error-and-success-constants
    $ICSTPCObj = foreach ($Global:pcline in $Global:PCList)
    {
        Identify-PCName
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
            # $ICST = Invoke-command -computer $Global:PC {Get-ScheduledTask * | Get-ScheduledTaskInfo |Where {$_.LastTaskResult -gt 0}}|select @{n="LastTaskResult_Hex";e={[System.Convert]::ToString($_.LastTaskResult, 16)}},*
            $ICST = Invoke-command -computer $Global:PC {Get-ScheduledTask * | Get-ScheduledTaskInfo |Where {$_.LastRunTime -ne $null -and $_.LastTaskResult -gt 0}}|select @{n="LastTaskResult_Hex";e={[System.Convert]::ToString($_.LastTaskResult, 16)}},*
            $ICSTObj = foreach ($ICSTLine in $ICST)
            {
            $ICSTLine |select pscomputername, LastRunTime, LastTaskResult,LastTaskResult_Hex,
                @{Label = "LTRDesc" ;Expression={
                        switch ($ICSTLine.LastTaskResult_Hex)
                        {
                            41300 {"SCHED_S_TASK_READY" } #0x00041300 The task is ready to run at its next scheduled time.
                            41301 {"SCHED_S_TASK_RUNNING" } #0x00041301 The task is currently running.
                            41302 {"SCHED_S_TASK_DISABLED"} # 0x00041302 The task will not run at the scheduled times because it has been disabled.
                            41303 {"SCHED_S_TASK_HAS_NOT_RUN"} # 0x00041303 The task has not yet run.
                            41304 {"SCHED_S_TASK_NO_MORE_RUNS"} # 0x00041304 There are no more runs scheduled for this task.
                            41305 {"SCHED_S_TASK_NOT_SCHEDULED"} # 0x00041305 One or more of the properties that are needed to run this task on a schedule have not been set.
                            41306 {"SCHED_S_TASK_TERMINATED"} # 0x00041306 The last run of the task was terminated by the user.
                            41307 {"SCHED_S_TASK_NO_VALID_TRIGGERS"} # 0x00041307 Either the task has no triggers or the existing triggers are disabled or not set.
                            41308 {"SCHED_S_EVENT_TRIGGER"} # 0x00041308 Event triggers do not have set run times.
                            80041309 {"SCHED_E_TRIGGER_NOT_FOUND"} # 0x80041309 A task's trigger is not found.
                            8004130A {"SCHED_E_TASK_NOT_READY"} # 0x8004130A One or more of the properties required to run this task have not been set.
                            8004130B {"SCHED_E_TASK_NOT_RUNNING"} # 0x8004130B There is no running instance of the task.
                            8004130C {"SCHED_E_SERVICE_NOT_INSTALLED"} # 0x8004130C The Task Scheduler service is not installed on this computer.
                            8004130D {"SCHED_E_CANNOT_OPEN_TASK"} # 0x8004130D The task object could not be opened.
                            8004130E {"SCHED_E_INVALID_TASK"} # 0x8004130E The object is either an invalid task object or is not a task object.
                            8004130F {"SCHED_E_ACCOUNT_INFORMATION_NOT_SET"} # 0x8004130F No account information could be found in the Task Scheduler security database for the task indicated.
                            80041310 {"SCHED_E_ACCOUNT_NAME_NOT_FOUND"} # 0x80041310 Unable to establish existence of the account specified.
                            80041311 {"SCHED_E_ACCOUNT_DBASE_CORRUPT"} # 0x80041311 Corruption was detected in the Task Scheduler security database; the database has been reset.
                            80041312 {"SCHED_E_NO_SECURITY_SERVICES"} # 0x80041312 Task Scheduler security services are available only on Windows NT.
                            80041313 {"SCHED_E_UNKNOWN_OBJECT_VERSION"} # 0x80041313 The task object version is either unsupported or invalid.
                            80041314 {"SCHED_E_UNSUPPORTED_ACCOUNT_OPTION"} # 0x80041314 The task has been configured with an unsupported combination of account settings and run time options.
                            80041315 {"SCHED_E_SERVICE_NOT_RUNNING"} # 0x80041315 The Task Scheduler Service is not running.
                            80041316 {"SCHED_E_UNEXPECTEDNODE"} # 0x80041316 The task XML contains an unexpected node.
                            80041317 {"SCHED_E_NAMESPACE"} # 0x80041317 The task XML contains an element or attribute from an unexpected namespace.
                            80041318 {"SCHED_E_INVALIDVALUE"} # 0x80041318 The task XML contains a value which is incorrectly formatted or out of range.
                            80041319 {"SCHED_E_MISSINGNODE"} # 0x80041319 The task XML is missing a required element or attribute.
                            8004131A {"SCHED_E_MALFORMEDXML"} # 0x8004131A The task XML is malformed.
                            8004131B {"SCHED_S_SOME_TRIGGERS_FAILED"} # 0x0004131B The task is registered, but not all specified triggers will start the task.
                            8004131C {"SCHED_S_BATCH_LOGON_PROBLEM"} # 0x0004131C The task is registered, but may fail to start. Batch logon privilege needs to be enabled for the task principal.
                            8004131D {"SCHED_E_TOO_MANY_NODES"} # 0x8004131D The task XML contains too many nodes of the same type.
                            8004131E {"SCHED_E_PAST_END_BOUNDARY"} # 0x8004131E The task cannot be started after the trigger end boundary.
                            8004131F {"SCHED_E_ALREADY_RUNNING"} # 0x8004131F An instance of this task is already running.
                            80041320 {"SCHED_E_USER_NOT_LOGGED_ON"} # 0x80041320 The task will not run because the user is not logged on.
                            80041321 {"SCHED_E_INVALID_TASK_HASH"} # 0x80041321 The task image is corrupt or has been tampered with.
                            80041322 {"SCHED_E_SERVICE_NOT_AVAILABLE"} # 0x80041322 The Task Scheduler service is not available.
                            80041323 {"SCHED_E_SERVICE_TOO_BUSY"} #  0x80041323 The Task Scheduler service is too busy to handle your request. Please try again later.
                            80041324 {"SCHED_E_TASK_ATTEMPTED"} # 0x80041324 The Task Scheduler service attempted to run the task, but the task did not run due to one of the constraints in the task definition.
                            80041325 {"SCHED_S_TASK_QUEUED"} # 0x00041325 The Task Scheduler service has asked the task to run.
                            80041326 {"SCHED_E_TASK_DISABLED"} # 0x80041326 The task is disabled.
                            80041327 {"SCHED_E_TASK_NOT_V1_COMPAT"} # 0x80041327 The task has properties that are not compatible with earlier versions of Windows.
                            80041328 {"SCHED_E_START_ON_DEMAND"} # 0x80041328 The task settings do not allow the task to start on demand.
                     }
             }},TaskName
        }
        $ICSTObj
        }
    }
    $ICSTPCObj|sort pscomputername,Lasttaskresult,TaskName | Out-GridView -Title "Scheduled Tasks"
}

Function View-Certificates()
{
    Clear-Host
    $CertPath = "Cert:\LocalMachine"
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    $CertObjectOut = foreach ($Global:PCline in $Global:PCList)
    {
    Identify-PCName
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
        $certs = invoke-command {gci "Cert:\LocalMachine" -Recurse} -computername $Global:PC -Credential $cred
        # Set-Location Cert:\LocalMachine\My
        Write-Host = $Global:PC
        foreach ($Cert in $Certs)
            {
                $CertOBProperties = @{'pscomputername'=$Global:PC;
                            'Issuer'=($cert.Issuer);
                            'Subject'=($cert.Subject);
                            'NotBefore'=($cert.NotBefore);
                            'notafter'=($cert.notafter);
                            'DateChanged'=$CDAte}
                    New-Object -TypeName PSObject -Prop $CertOBProperties
            }
        #Get-ChildItem -Recurse cert: |select Subject, notafter
        }
    }
    $CertObjectOut |select PSComputerName, Issuer, Subject, NotBefore, notafter|sort NotAfter |Out-Gridview -Title $CertPath
}
Function Get-OSArch()
{
    Clear-Host
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
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
                Write-Output "`n$Script:PCt4$Script:ADd4"
            }
        }
    }
}

Function Reset-IE () 
{
    # Use Reset-IE to reset IE to its defaults and clear browsing history. Requires some user interaction
    # Not on menu yet
    RunDll32.exe InetCpl.cpl,ResetIEtoDefaults
    RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351
}
Function get-WinSAT()
{
    $wsatResults = @()
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    Write-Warning "Getting WinSAT information from the following list : $Global:PCList"
    foreach($Global:pcLine in $Global:PCList)
    {
        Identify-PCName # Called from Run-PSMenu.ps1
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
            $wsat = get-wmiobject win32_winsat -ComputerName $Global:PC -ErrorAction SilentlyContinue |select pscomputername,cpuscore,d3dscore,diskscore,graphicsscore,memoryscore,winsatassessmentstate,winsprlevel
            $wsatResults += $wsat
        }
    }
    $wsatResults |sort pscomputername,winsprlevel,cpuscore,diskscore,memoryscore,graphicsscore |ft
}
function Manage-HKLM
{
Clear-Host
$HKEY = $null
if (($HKPrompt = Read-Host "`nType a Key (like hklm,hkcu,hkcr,etc) or press [Enter] to use the default key: hklm") -eq "") {$HKPrompt = "hklm"}
if ($HKPrompt -eq "hklm"){$HKey="HKEY_LOCAL_MACHINE"}
elseif($HKPrompt -eq "hkcu"){$HKey="HKEY_CURRENT_USER"}
elseif($HKPrompt -eq "hkcr"){$HKey="HKEY_CLASSES_ROOT"}
$psdhkey = $null
$psdhkeyName = $null
$npsd = $null
$pathroot = $null
$path = $null
if ($Global:PCList -eq $null)
{
    Get-PCList
}
if (($PropertyName = Read-Host "`nType a Property Name to view or [Enter] to use default: RequireSecuritySignature") -eq "") {$PropertyName = "RequireSecuritySignature"}
$DefaultMidPath = "SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"
if (($MidPath = Read-Host "`nType the path exclude the root and the property [Enter] to use default:`n`t $DefaultMidPath") -eq "") {$MidPath = $DefaultMidPath}

    $HKObjectOut = foreach ($Global:PCline in $Global:PCList)
    {
        Identify-PCName
        if(Test-Connection -ComputerName $Global:PC -Count 1 -ea 0) {
            Write-host "." -NoNewline -ForegroundColor Green
            $psdhkey = "$Global:pc"+"\"+"$HKPrompt"
            $psdhkeyName = "$Global:pc"+"$HKPrompt"
            $npsess = New-PSSession -ComputerName $Global:PC -Credential $Cred -Name $psdhkeyName -ErrorVariable pssError -ErrorAction SilentlyContinue
            if (!$pssError)
            {
            $npsd = Invoke-Command -Session $npsess -ScriptBlock {New-PSDrive -PSProvider Registry -Name $using:psdhkeyName -root $using:HKey}  
                $pathroot  = Invoke-Command -Session $npsess -ScriptBlock {$using:npsd.Name + ":"}
                $path  = Invoke-Command -Session $npsess -ScriptBlock {$using:pathroot +$using:MidPath}
                $G_IP  = Invoke-Command -Session $npsess -ScriptBlock {Get-ItemProperty $using:path}
                $PropertyValue = ($G_IP).$PropertyName
                $srvComment = ($G_IP).srvComment
                Invoke-Command -Session $npsess -ScriptBlock {Remove-PSDrive $using:npsd}
                Remove-PSSession -Name $psdhkeyName
            }
            Else
            {
                $pathroot = ""
                $path = ""
                $propertyValue = ""
                $srvComment = "Error connecting to $Global:PC"
            }
                $HKOBProperties = @{'pscomputername'=$Global:PC;
                    'HKPrompt'=$HKPrompt;
                    'psdhkey'=$psdhkey;
                    'psdhkeyName'=$psdhkeyName;
                    'pathroot'=$pathroot;
                    'MidPath'=$MidPath;
                    'path'=$path;
                    'PropertyName'=$PropertyName;
                    'PropertyValue'=$PropertyValue;
                    'srvComment'=$srvComment}
                New-Object -TypeName PSObject -Prop $HKOBProperties
        }
        else
        {
            Write-host "." -NoNewline -ForegroundColor Red
        }
    $psdhkey = $null
    $psdhkeyName = $null
    $npsd = $null
    $pathroot = $null
    $path = $null
    $npsess = $null
    $G_IP = $null
    $pssError = $null
    }
    # $HKObjectOut |select pscomputername,HKPrompt,psdhkey,psdhkeyName,pathroot,MidPath,path,PropertyName,PropertyValue,srvComment |Out-GridView -Title "$MidPath $PropertyName"
    $HKObjectOut |select pscomputername,HKPrompt,MidPath,PropertyName,PropertyValue,srvComment |Out-GridView -Title "$MidPath $PropertyName"
}

Function Retrieve-BitlockerKey
{
    Clear-Host
    # Using the following to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    Write-Warning "Getting BitLocker Recovery information from the following list : $Global:PCList"
    $BitLOBjectOut = foreach($Global:pcLine in $Global:PCList)
    {
        Identify-PCName # Called from Run-PSMenu.ps1
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
            $ADPCObj = Get-ADComputer $Global:PC
            $Bitlocker_Object = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $ADPCObj.DistinguishedName -Properties 'msFVE-RecoveryPassword'
            
                $BitLOBProperties = @{'pscomputername'=$Global:PC;
                    'DistinguishedName'=$HKPrompt;
                    'msFVE-RecoveryPassword'=$Bitlocker_Object.'msFVE-RecoveryPassword';
                    'Name'=$Bitlocker_Object.Name;
                    'ObjectClass'=$Bitlocker_Object.ObjectClass;
                    'ObjectGUID'=$Bitlocker_Object.ObjectGUID}
                New-Object -TypeName PSObject -Prop $BitLOBProperties

        }
    }
    $BitLOBjectOut|where {$_.'msFVE-RecoveryPassword'}|sort pscomputername |FT pscomputername,msFVE-RecoveryPassword,ObjectClass -autosize -Wrap # ,ObjectGUID,Name,DistinguishedName
    [console]::beep(559,250);[console]::beep(559,100);[console]::beep(559,100);Start-sleep -milliseconds 250; # https://devblogs.microsoft.com/scripting/powertip-use-powershell-to-send-beep-to-console/
    [console]::beep(559,250);[console]::beep(559,250);[console]::beep(559,250);Start-sleep -milliseconds 250;
    [console]::beep(559,250);[console]::beep(559,100);Start-sleep -milliseconds 250;
    [console]::beep(559,100)

}
    <#
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    Identify-PCName
    If (Test-Connection $Global:pc -Count 1 -Quiet)
    {
    }
    #>

    