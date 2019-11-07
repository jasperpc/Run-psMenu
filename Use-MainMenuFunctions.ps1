<#
.NAME
<<<<<<< HEAD
    Use-MainMenuFunctions.ps1
=======
    Run-psMenu.ps1
>>>>>>> a1255d3f753d7d5f3e2386825dab4574a19aabcf
.SYNOPSIS
    Provide menu system for running Powershell scripts
.DESCRIPTION
 Menu design initially adopted from : https://quickclix.wordpress.com/2012/08/14/making-powershell-menus/
 2017-03-03

 Written by Jason Crockett - Laclede Electric Cooperative

 use Ctrl+J to bring up "snippets" that can be inserted into code
 Tips from MVA
    - Scripting ToolMaking
    - www.powershell.org
    - #PowerShell on Twitter & @JSnover
.PARAMETERS
    This parameter doesn't exist (Comps = list of computernames)
.EXAMPLE
    "a command with a parameter would look like: (.\Get-DiskInfo.ps1 -comps dc)"
<<<<<<< HEAD
        .\Use-MainMenuFunctions.ps1 []
.SYNTAX
    .\Use-MainMenuFunctions.ps1 []
.REMARKS
    To see the examples, type: help Use-MainMenuFunctions.ps1 -examples
    To see more information, type: help Use-MainMenuFunctions.ps1 -detailed
=======
        .\adutils.ps1 []
.SYNTAX
    .\adutils.ps1 []
.REMARKS
    To see the examples, type: help adutils.ps1 -examples
    To see more information, type: help adutils.ps1 -detailed
>>>>>>> a1255d3f753d7d5f3e2386825dab4574a19aabcf
.TODO
    Enable Script block logging
#>
# See Getting Started with PowerShell 3.0 Jump Start in MS Virtual Academy for loading modules from other systems

# Enable AD and PS Update modules
Import-Module ActiveDirectory

Function PSWindowsUpdate() # Possible ToDo here?
{
    #https://gallery.technet.microsoft.com/scriptcenter/2d191bcd-3308-4edd-9de2-88dff796b0bc#content
    #available functions are listed at the above website.
    # Not on menu yet
Import-Module PSWindowsUpdate

}
Function Reset-IE () 
{
    # Use Reset-IE to reset IE to its defaults and clear browsing history. Requires some user interaction
    # Not on menu yet
    RunDll32.exe InetCpl.cpl,ResetIEtoDefaults
    RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351
}

Clear-Host
$PSVersionTable.PSVersion
# set Global and Script variables
. .\Get-SystemsList.ps1
[BOOLEAN]$Global:ExitSession=$false

$Use_MMFPath = Split-Path -Parent $PSCommandPath

$Script:t4 = "`t`t`t`t"
$Script:d4 = "-----------------------------Use-MainMenuFunctions-------------------------------"
$Script:NC = $null
$Script:p1 = $null
$Script:p2 = $null
$Script:p3 = $null
$CDate = Get-Date -UFormat "%Y%m%d"
$Global:CDateTime = [datetime]::ParseExact($Global:CDate,'yyyymmdd',$null)
[int]$menuselect = 0

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

Function menu()
{
while ($menuselect -lt 1 -or $menuselect -gt 33)
{
    Clear-Host
        Trap {"Error: $_"; Break;}        
        $MNum = 0;Clear-host |out-null
        Start-Sleep -m 50
        Write-host $t4 $d4 -ForegroundColor Green
#    Write-Host "$t4  1. Exit Now" -ForegroundColor Red
        # Exit
        $MNum ++;
            $Exit=$MNum;
            $p1 = " $MNum. `t";
            $p2 = "Exit to Main Menu";
            $p3 = " ";
            $NC = "Red";chcolor $p1 $p2 $p3 $NC
#    Prompt for credential with permissions across the network
        $MNum ++;$GCred=$MNum;$p1 =" $MNum.  `t Enter";
            $p2 = "Credentials ";$p3 = "for use in script."; $NC = "Cyan";chcolor $p1 $p2 $p3 $NC
#    Write-Host "$t4  2. Active Directory Menu" -ForegroundColor Yellow
            # Go to the Active Directory Menu
        $MNum ++;$ADmenu=$MNum;
            $p1 = " $MNum. `t Go to ";
            $p2 = "Active Directory";
            $p3 = " menu"
            $NC = "Green";chcolor $p1 $p2 $p3 $NC
#    Write-Host "$t4  3. Net Utilities Menu" -ForegroundColor Yellow
            # Go to the Net Utilities Menu
        $MNum ++;$NetMenu=$MNum;
            $p1 = " $MNum. `t Go to ";
            $p2 = "Net Utilities";
            $p3 = " menu"
            $NC = "Magenta";chcolor $p1 $p2 $p3 $NC
# PC Info and Service Tag"
        $MNum ++;$Get_pcinfo=$MNum;
            $p1 = " $MNum. `t View ";
            $p2 = "PC Information ";
            $p3 = "and Dell Service Tag"
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
#    "View Java versions installed on a PC"
        $MNum ++;$ID_Java=$MNum;
            $p1 = " $MNum. `t View ";
            $p2 = "Java Versions ";
            $p3 = "installed on select systems"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
#    "Process status on a PC"
        $MNum ++;$Get_ProcessStatus=$MNum;
            $p1 = " $MNum. `t View ";
            $p2 = "process status ";
            $p3 = "on select systems"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# "Run `"update-help -Force`" to update PowerShell Help"
        $MNum ++;$update_help=$MNum;
            $p1 = " $MNum. `t Run ";
            $p2 = "Update on PowerShell Help ";
            $p3 = "information"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# List programs installed on a pc"
        $MNum ++;$get_apps=$MNum;
            $p1 = " $MNum. `t List ";
            $p2 = "installed programs ";
            $p3 = "on select systems"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# "$t4 14. Uninstall (a) program(s) on a pc"
        $MNum ++;$Uninst_App=$MNum;
            $p1 = " $MNum. `t";
            $p2 = "Uninstall program ";
            $p3 = "on select systems"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# Service Status Check"
        $MNum ++;$Get_Svc=$MNum;
            $p1 = " $MNum. `t Verify ";
            $p2 = "Service status ";
            $p3 = "on select systems"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
#    Write-Output "$t4  4. RDP - Enable on a workstation"
    # Enable RDP on a system
        $MNum ++;$EnableRDP=$MNum;
            $p1 = " $MNum. `t";
            $p2 = "Enable RDP";
            $p3 = " on a Windows system"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# MS KB update installation status" # Check for Most Recent Windows Update on servers
        $MNum ++;$get_KB=$MNum;
            $p1 = " $MNum. `t Verify installation status ";
            $p2 = "of MS KB or Windows updates ";
            $p3 = "on select systems"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# List most recent MS update for the servers
        $MNum ++;$List_RecentUpdate=$MNum;
            $p1 = " $MNum. `t List  ";
            $p2 = "most recent MS update ";
            $p3 = "on select systems"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# View associated logon names used to start a service
        $MNum ++;$Get_ServiceLogons=$MNum;
            $p1 = " $MNum. `t View ";
            $p2 = "service logon names ";
            $p3 = "on select systems"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# convert-Decimal to Binary to Hexadecimal
        $MNum ++;$convert_Numbers=$MNum;
            $p1 = " $MNum. `t convert ";
            $p2 = "Decimal to Binary to Hexadecimal ";
            $p3 = " "
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# Get text URL Web Page Info
        $MNum ++;$wGet_URL=$MNum;
            $p1 = " $MNum. `t Get ";
            $p2 = "Web Page in text ";
            $p3 = "from URL"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# Get a list of Files Modified in the last xx days
        $MNum ++;$Get_ModifiedFileList=$MNum;
            $p1 = " $MNum. `t Get a ";
            $p2 = "list of Files Modified ";
            $p3 = "in the last xx days"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# NTFS Folder - Change Ownership and Add FullControl permissions then prompt to delete.
            $MNum ++;
            $Manage_NTFS=$MNum;
            $p1 = " $MNum. `t NTFS Folder - ";
            $p2 = "Change Ownership and Add FullControl permissions ";
            $p3 = "then prompt to delete"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# Find Duplicate Files
        $MNum ++;$Compare_Files=$MNum;
            $p1 = " $MNum. `t Find ";
            $p2 = "Duplicate Files ";
            $p3 = "on a system"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# View a file hash
        $MNum ++;$run_get_filehash=$MNum;
            $p1 = " $MNum. `t View ";
            $p2 = "hash ";
            $p3 = "of a file"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# View-Certificates
        $MNum ++;$View_Certificates=$MNum;
            $p1 = " $MNum. `t View ";
            $p2 = "Certificates ";
            $p3 = "on a system"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# Screen saver status on pc to see if a user is AFK
        $MNum ++;$Active_ScrnSav=$MNum;
            $p1 = " $MNum. `t Check ";
            $p2 = "screen saver status ";
            $p3 = "on pc to see if a user is AFK"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# Get Scheduled Task Information
        $MNum ++;$Get_ScheduledTaskInformation=$MNum;
            $p1 = " $MNum. `t Get a system's ";
            $p2 = "Scheduled Task ";
            $p3 = "Information"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# $Create_DMARCReport
        $MNum ++;$Create_DMARCReport=$MNum;
            $p1 = " $MNum. `t Create a ";
            $p2 = "DMARC Report ";
            $p3 = "from xml files in zip files"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# Convert from ASCII to Base64 using UTF8 encoding then back
        $MNum ++;$Manage_Base64_ASCII_UTF8Encoding=$MNum;
            $p1 = " $MNum. `t Convert from  ";
            $p2 = "ASCII to Base64 and back ";
            $p3 = "using UTF8 encoding"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# Convert from ASCII to Base64 and then back
        $MNum ++;$Manage_Base64_ASCII_Encoding=$MNum;
            $p1 = " $MNum. `t Convert from  ";
            $p2 = "ASCII to Base64 and back ";
            $p3 = "using ASCII encoding"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# This Base64 to ASCII conversion combines the conversion and encoding into one step
        $MNum ++;$Manage_Base64_to_ASCII_Encoding=$MNum;
            $p1 = " $MNum. `t Convert from ";
            $p2 = "Base64 to ASCII only ";
            $p3 = "using ASCII encoding in one step"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
# This prompts to run a powershell script on selected remote system(s)
        $MNum ++;$Run_RemotePS=$MNum;
            $p1 = " $MNum. `t This prompts to  ";
            $p2 = "run a powershell script ";
            $p3 = "on selected remote system(s)"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
        # Setting up the menu in this way allows you to move menu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $MNum ++;$p1 =" $MNum.  FIRST";$p2 = "HIGHLIGHTED ";$p3 = "END"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        Write-Host "$t4$d4"  -ForegroundColor Green
        $Script:MNum = 0;[int]$MenuSelect = Read-Host "`n`tPlease select the menu item you would like to do now"
}

switch($menuselect)
    {
        # 1{}
        $Exit{$MenuSelect=$null}
        $GCred{Clear-Host;Get-Cred;$menuselect = $null;menu}
        $ADmenu{. .\Use-AdMenuFunctions;ADmenu}
        $NetMenu{Clear-Host;. .\Use-NetMenuFunctions;NetMenu}
        $Get_pcinfo{Get-pcinfo;reloadmenu}                  # Is this the same as 5 - NoDomPCInfo?
        $NonDomPCinfo{NoDomPCinfo;reloadmenu}
        $SysEventlogs{SysEvents;reloadmenu}
        $RmtPwr{New-PSRemote;reloadmenu}
        $BootTimes{Clear-Host;.\Start-MultiThreads.ps1 .\Get-BootTime.ps1;reloadmenu}
        $ID_Java{ID-Java;reloadmenu}
        $Get_ProcessStatus{Get-ProcessStatus;reloadmenu}
        $update_help{clear-host;write-Warning "Running `"update-help -Force`" to update PowerShell Help";update-help -Force;reloadmenu}
        $get_apps{get-apps;reloadmenu}
        $Uninst_App{Invoke-Expression ".\Uninstall-Application.ps1";;reloadmenu}
        $Get_Svc{Get-Svc;reloadmenu}
        $EnableRDP{Enable-RDP;reloadmenu}
        $get_KB{get-KB;reloadmenu}
        $List_RecentUpdate{List-RecentUpdate;reloadmenu}
        $Get_ServiceLogons{Get-ServiceLogons;reloadmenu}
        $convert_Numbers{convert-Numbers;reloadmenu}
        $wGet_URL{wGet-URL;reloadmenu}
        $Get_ModifiedFileList{Get-ModifiedFileList;reloadmenu}
        $Manage_NTFS{Clear-Host;Manage-NTFS;reloadmenu}
        $Compare_Files{Compare-Files;reloadmenu}
        $run_get_filehash{run-get-filehash;reloadmenu}
        $View_Certificates{View-Certificates;reloadmenu}
        $Active_ScrnSav{Active-ScrnSav;reloadmenu}
        $Get_ScheduledTaskInformation{Get-ScheduledTaskInformation;reloadmenu}
        $Create_DMARCReport{Invoke-Expression ".\Create-DMARCReport.ps1";reloadmenu}
        $Manage_Base64_ASCII_UTF8Encoding{Convert-ASCII_UTF8_To_Base64;Convert-Base64_ASCII_UTF8;reloadmenu}
        $Manage_Base64_ASCII_Encoding{Convert-ASCII_To_Base64;Convert-Base64_To_ASCII;reloadmenu}
        $Manage_Base64_to_ASCII_Encoding{Convert-Base64ToASCII;reloadmenu}
        $Run_RemotePS{Run-RemotePS;reloadmenu}
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

Function reloadmenu()
{
        Read-Host "Hit Enter to continue"
        # Start-Sleep -m 250
        $menuselect = $null
        menu
}
Function Get-Cred()
{
    # Get Credentials for later wmi requests
    $Global:cred = Get-Credential "$env:USERDOMAIN\Administrator"
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
Function NoDomPCinfo()
{
    Clear-Host
    Write-Warning "Enter a computername to see details"
    Get-PCList
    $DomName =  $env:USERDNSDOMAIN
    if (($DomName = Read-Host "Enter the domain name [Enter] to use default: $env:USERDNSDOMAIN") -eq "") {$DomName = $env:USERDNSDOMAIN}
    foreach($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        $sOS = Get-WmiObject -class Win32_OperatingSystem -Impersonation 3 -Credential "$DomName\administrator" -computername $Global:PC
        # Get the Dell Service Tag # from pc list
            get-wmiobject  -Impersonation 3 -Credential "$DomName\administrator" -computername $Global:PC -class win32_bios <#-erroraction SilentlyContinue#> |select "$Global:PC", SerialNumber |fl
        foreach($sProperty in $sOS){
        # Get OS information from pc list
            write-host "`nOS: "$sProperty.Caption
            write-host "Architecture: "$sProperty.OSArchitecture
            write-host "Registered to: "$sProperty.Description
            write-host "Version #: " $sProperty.ServicePackMajorVersion
        }
        # This is similar to the following from a DOS cli - wmic:root\cli>/node:"192.168.1.79" /user:workgroup\administrator memorychip get /all /format:list
        $mem = Get-WmiObject cim_physicalmemory -Impersonation 3 -Credential "$DomName\administrator" -ComputerName $Global:PC | select banklabel,capacity,caption,devicelocator,partnumber,speed |ft -AutoSize    
        $mem
    }
}
Function SysEvents()
{ 
    Get-PCList
    # $pclist = Read-host "Enter Computer list"
    $WinLogName= Read-host "Enter Application, Security, Setup, or System"
    $EntryLevel_Type= Read-host " Enter Critical, Error, Warning, Information, FailureAudit, SuccessAudit"
    $cnt = Read-host "How many logs do you want to see of this type?" 
        if ($Global:pclist -eq "") {$Global:pclist = [environment]::machinename}
        if ($WinLogName -eq "") {$WinLogName = "System"} 
        if ($EntryLevel_Type -eq "") {$EntryLevel_Type = "Error"}
        if ($cnt -eq "") {$cnt = 50}
        Invoke-Expression ".\Get-SysEvents.ps1 -PCList $Global:PCList -WinLogName $WinLogName -EntryLevel_Type $EntryLevel_Type -cnt $cnt"
}
Function New-PSRemote()
{
    Clear-Host
    # Connect to Server Remotely
    # Still needs some work for correct authorization / force credentials again to connect to remote machines
        # Prompt to decide to use the script's $Cred credential or another
        $CredUser = $Cred.UserName
        if ((Read-Host "Current credentials are for: `"$CredUser`" , type C to change or [Enter] to continue")-eq "C") {Get-Cred}
        Write-Warning "Enter the computername to which you want to open a PowerShell Remote Session"
        Get-PCList
        foreach ($Global:pcLine in $Global:PCList)
        {
            Identify-PCName
            # Connect to Exchange
            # Enter-PSSession -ComputerName ExchangeSvrName -Authentication Kerberos -Credential $userCredential
            # Add-PSSnapin *exchange*
            Enter-PSSession -ComputerName $Global:pc -Authentication Kerberos -Credential $Cred
        }
        Clear-host
break
}
Function ID-Java()
{
    Clear-Host
    Get-PCList
    foreach($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        $tcpc1 = test-connection -computername $Global:PC -Count 2 -Quiet #-ErrorAction SilentlyContinue
        if ($tcpc1 -eq $true)
        {
            #use WMI to identify Java versions on pc
            Write-Host "`nGet-WmiObject -Class win32_product -Filter "name like 'Java%%'" -ComputerName $Global:PC -Credential $cred |select -expand version |ft"
            Write-Host "`nDetermining (slowly) what versions of Java are installed on "$Global:PC" :`n"
            Get-WmiObject -Class win32_installedwin32program -Filter "name like 'Java%%'" -ComputerName $Global:PC -Credential $cred |select PSComputerName, Name, version, __Path |ft
            <# Get Java Uninstall String
            $javinstprop = get-itemproperty "HKLM:\SOFTWARE\microsoft\windows\currentversion\Installer\UserData\S-1-5-18\Products\4EA42A62D9304AC4784BF2381208370F\InstallProperties"
            $javinstprop |Select-Object uninstallstring
            $javinstprop = ""
            #>
        }
        else 
        {
        write-host "`n"$Global:PC " is not available right now.`n"
        }
    }      
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
Function Get-pcinfo()
{
    Clear-Host
    Write-Warning "Getting system information from the following systems:"
    Get-PCList
    foreach($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        $CU = $cred.UserName
        if ((Read-Host "Current credentials are: `"$CU`" , type C to change or [Enter] to continue")-eq "C") {Get-Cred}
        # Get the Dell Service Tag # from pc list
                get-wmiobject -Credential $cred -computername $Global:PC -class win32_bios |select "$Global:PC", SerialNumber |fl
                Write-Host 'Running: `nget-wmiobject -Credential $cred -computername $Global:PC -class win32_bios |select "$Global:PC", SerialNumber |fl'

        # Get OS information from pc list
            $sOS = Get-WmiObject -class Win32_OperatingSystem -Impersonation Impersonate -Credential $cred -computername $Global:PC
                Write-Host 'Running: `nGet-WmiObject -class Win32_OperatingSystem -Impersonation Impersonate -Credential $cred -computername $Global:PC'
            foreach($sProperty in $sOS){
                write-host "`nOS: "$sProperty.Caption
                write-host "Architecture: "$sProperty.OSArchitecture
                write-host "Registered to: "$sProperty.Description
                write-host "Version #: " $sProperty.ServicePackMajorVersion
            }
        
        # Get CPU Information for listed system
            $cpuinfo = get-wmiobject win32_processor -Impersonation Impersonate -Credential $cred -ComputerName $Global:PC |select-object role, name, numberofcores,numberoflogicalprocessors,
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
        # Get and display memory information for listed system
            Write-Host 'Running: `nGet-WmiObject cim_physicalmemory -Impersonation 3 -Credential $cred -ComputerName $Global:PC | select banklabel,capacity,caption,devicelocator,partnumber,speed |ft -AutoSize'
            $mem = Get-WmiObject cim_physicalmemory -Impersonation 3 -Credential $cred -ComputerName $Global:PC | select banklabel,capacity,caption,devicelocator,partnumber,speed |ft -AutoSize
            $mem
        # Get and display printer information for listed system
            Write-Host 'Running: `nGet-WmiObject -Class win32_printer -Credential $cred -ComputerName $Global:PC | select name, drivername, portname, shared, sharename |ft -AutoSize'
            $prin = Get-WmiObject -Class win32_printer -Credential $cred -ComputerName $Global:PC | select name, drivername, portname, shared, sharename |ft -AutoSize
            $prin
        # Testing another Printer Info command
        Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Network=True" | select Sharename,name | export-csv -path $printer\$env:computername.csv -NoTypeInformation
        # Get and display Printer Port information for listed system
            Write-Host 'get-printerport -computer $Global:PC |select ComputerName,Name,Description,PrinterHostAddress |ft -A'
            $prin2 = get-printerport -computer $Global:PC |select ComputerName,Name,Description,PrinterHostAddress |ft -A
            $prin2
        # Get and display Printer Driver information for listed system
            Write-Host 'get-printerDriver -computer $Global:PC |select * |ft -A # ComputerName,Name,Description,PrinterHostAddress,OEMUrl,Monitor,Path,provider |ft -A'
            $prin3 = get-printerDriver -computer $Global:PC |select ComputerName,Name|ft -A
            $prin3
        # Get Disk information for system
            Write-Host "`nRunning: get-wmiobject win32_logicaldisk -Credential $cred -Computername $Global:PC | fl  systemname, Name, deviceid, description, Drivetype, mediatype, filesystem, @{n=`'MB freespace`';e={[string]::Format(`"{0:0,000} MB`",$_.FreeSpace/1MB)}},@{n=`'MB Size`';e={[string]::Format(`"{0:0,000} MB`",$_.Size/1MB)}}, volumedirty, volumename, volumeserialnumber"
            get-wmiobject win32_logicaldisk -Credential $cred -Computername $Global:PC | fl  systemname, Name, deviceid, description, Drivetype, mediatype, filesystem, @{n='MB freespace';e={[string]::Format("{0:0,000} MB",$_.FreeSpace/1MB)}}, @{n='MB Size';e={[string]::Format("{0:0,000} MB",$_.Size/1MB)}}, volumedirty, volumename, volumeserialnumber
        #Get process from pc
            $logonInfo = get-wmiobject win32_process -computername $Global:PC -Credential $cred| select Name,sessionId,Creationdate  |Where-Object {($_.Name -like "Logon*")} 
            Write-Host "`nRunning: wmi search to see if the logon processes is running...`n"
            $logonInfo
        #Get Network information from pc
            Write-host "Checking for NIC status"    
            Get-WmiObject -Class win32_networkadapterconfiguration -ComputerName $Global:PC -Credential $cred|where {$_.IPAddress -ne $null}|select PSComputerName,MACAddress,IPSubnet,DNSServerSearchOrder, @{n="IP Address";e={($_ | select -ExpandProperty IPAddress)-join ','}},DHCPEnabled|ft
    }
}
Function get-apps()
{
    Get-PCList
    Write-Warning "Select systems on which to look for applications"
    $prglist = Read-Host "Enter a search term or nothing to find a program or all programs "
    foreach ($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        Get-WmiObject -class Win32_installedwin32program -ComputerName $Global:PC -Credential $cred |where {$_.name -ilike "*$prglist*"} |select pscomputername,Name,Version |Sort-Object name |ft #|ConvertTo-Csv -NoTypeInformation |select -Skip 1 |Out-File -Append ".\ProgsInPC.csv"
    }
         
}
Function Get-Svc()
{
    Clear-Host
    $CU = $cred.UserName
        if ((Read-Host "Current credentials are: `"$CU`" , type C to change or [Enter] to continue")-eq "C") {Get-Cred}
    Get-PCList
    $SvcName = Read-Host "Enter part of the service name (Ex: CRCPluginServer)"
    # foreach ($Global:PCLine in $Global:PCList){Identify-PCName;get-service remoteregistry -ComputerName $Global:PC |select MachineName,Status,StartType}
    # TO-Do: Move the action against the service into a different function and use this function just to list services with MachineName,Status, and StartType
    foreach($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        $WMIServ = Get-WmiObject win32_service -ComputerName $Global:PC -Credential $cred |select pscomputername,processid,name,state,pathname|Where-Object {$_.Name -like "*$SvcName*"}
        $WMIServ|ft -autosize -wrap
        if ($WMIServ.count -eq 1)
        {$1StatusChoice = Read-Host "Type yes to change the status of this service or any other key to continue."}
        elseif ($WMIServ.count -gt 1)
        {$StatusChoice = Read-Host "Type yes to change the status of one or more of these services or any other key to continue."}
        If ($1StatusChoice -eq "yes" -or $StatusChoice -eq "yes")
        {
            foreach ($WMIServLine in $WMIServ)
            {
                $svcName = $WMIServLine.name
                $gsout = Get-Service -name $svcName -ComputerName $Global:PC
                if ($StatusChoice -eq "yes") # process multiple possible services
                {
                    $1StatusChoice = Read-Host "Type yes to change the status of the `"$svcName`" service or any other key to continue."
                }
                If ($1StatusChoice -eq "yes")
                {
                    Write-Host -nonewline "`nEnter the new status for the service: " -ForegroundColor Green
                    Write-Host -NoNewline $gsout.ServiceName -ForegroundColor DarkYellow
                    Write-Host -nonewline " which is currently: "
                    Write-Host -nonewline $gsout.Status -ForegroundColor Red 
                    Write-Host " on"$gsout.MachineName"." -ForegroundColor Green
                    $NewSvcStat = Read-Host "Your options are Stopped, Paused, or Running"
                    If ($NewSvcStat -eq "Stopped")
                        {
                            $gsout.stop() # |Set-Service -Status Stopped -ComputerName $Global:PC
                        }
                    ElseIf ($NewSvcStat -eq "Running")
                        {
                            if (($gsout.StartType) -eq "disabled")
                                {
                                    $st = $gsout.starttype
                                    "The StartType for $SvcName is `"$st`" I`'ll try to start it. "
                                    $gsout |set-service -StartupType Automatic |Out-Null # This doesn't always work due to permissions, firewall, WMI, WinRM, Investigate...
                                    Invoke-Command -ComputerName $Global:PC -ScriptBlock {sc config $SvcName start= auto}
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
                    $1StatusChoice = ""
                    $NewSvcStat = ""
                }
            }
        }
        $gsout = $null
    }
    
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
Function List-RecentUpdate()
{
    Clear-Host
    $SN=$null
    $RecentUpdateList=$null
    Get-PCList
    # $servers = Get-ADComputer -Filter {OperatingSystem -like "*server*" -and Enabled -eq $true}
    # $servers = Get-ADComputer -Filter {Enabled -eq $true}
    $RecentUpdateList = foreach ($Global:pcLine in $Global:pclist)
    {
        Write-Host ". " -NoNewline
        Identify-PCName
        # Write-Warning $Global:PC
        $ping = test-connection $Global:PC -Count 1 -Quiet -ErrorAction SilentlyContinue
        if ($ping -eq $true){
            $RUL = Get-WMIObject -Class "Win32_QuickFixEngineering" -ComputerName $Global:PC -Credential $Cred |select PSComputerName,InstalledOn,HotFixID -Last 1

            $RULProperties = @{'pscomputername'=$Global:PC;
                    'Name'=$RUL.PSComputerName;
                    'InstalledOn'=$RUL.InstalledOn;
                    'HotFixID'=$RUL.HotFixID}
            New-Object -TypeName PSObject -Prop $RULProperties
        }
    }
    $RecentUpdateList |sort InstalledOn |ft
    
}
Function Get-ServiceLogons
{
    Clear-Host
    $uname = Read-Host "Enter all or part of a valid possible Username that a service may use to start"
    Get-PCList
    foreach ($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        Get-WmiObject -Class Win32_Service -ComputerName $Global:pc |Sort-Object StartName,PSComputerName |Where-Object {$_.startName -like "*$uname*"}|ft StartName,pscomputername,Name,State,Displayname -AutoSize
    }    
}
Function convert-Numbers()
{
    clear-host
    # add a section to input Hex number rather than just inputing a base 10 number
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
}
Function wGet-URL() # Initiate a web-request for a URL and return it in text form
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
}
Function Get-ModifiedFileList()
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
        $Dir2Scan = Read-Host "Hit [Enter] to scan \\FileServer\SharedFiles share on server or Enter an alternate path"
    # Set a default Dir2Scan value if nothing was entered in the prompt
        If($Dir2Scan -eq ""){$Dir2Scan = "\\FileServer\SharedFiles"}
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
    # Delete the CSV temp file when user exits out of menu
    # File will not be deleted if it is still open at this time
    Remove-Item $ExCSVFileName -Force
}
Function Manage-NTFS
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
Function run-get-filehash()
{
    # View the following for ideas
    # https://docs.microsoft.com/en-us/powershell/module/Microsoft.PowerShell.Utility/Get-FileHash?view=powershell-5.1
    Clear-Host
    $FullFileName = Read-Host "`nType the path and filename for the file for which you wish to view the hash`n Example: C:\Users\Public\Downloads\Software\Tools\Java\jdk-10.0.1_windows-x64_bin.exe`n"
    $HashAlgorithm = Read-Host "`nEnter the hashing algorithm you wish to use. `n`tChoose from SHA1 | SHA256 | SHA384 | SHA512 | MACTripleDES | MD5 | RIPEMD160"
    # Use of Get-FileHash: Get-FileHash [-Path] <String[]> [-Algorithm {SHA1 | SHA256 | SHA384 | SHA512 | MACTripleDES | MD5 | RIPEMD160}]
    Write-Warning "Running the following command:`n`nget-filehash -Path $FullFileName -Algorithm $HashAlgorithm -Verbose`n"
    get-filehash -Path $FullFileName -Algorithm $HashAlgorithm -Verbose
}
Function View-Certificates()
{
    Clear-Host
    Get-PCList
    foreach ($Global:PCline in $Global:PCList)
    {
    Identify-PCName
    $certs = invoke-command {gci Cert:\LocalMachine\My -Recurse} -computername $Global:PC
    # Set-Location Cert:\LocalMachine\My
    Write-Host = "Asset ID: "$Global:PC
    $certs |select PSComputerName, Issuer, Subject, NotBefore, notafter|sort NotAfter |FT PSComputerName,Issuer,Subject,NotBefore,NotAfter
    #Get-ChildItem -Recurse cert: |select Subject, notafter
    Write-host "`n"
    }
}
Function Active-ScrnSav()
{
Clear-Host
Write-Warning "An active screen saver means the user is AFK"
Get-PCList
    foreach ($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
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

Function Get-ScheduledTaskInformation
{
    Clear-Host
    # http://powershell-guru.com/powershell-tip-56-convert-to-and-from-hex/
    Get-pclist
    # https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-error-and-success-constants
    $ICSTPCObj = foreach ($Global:pcline in $Global:PCList)
    {
        Identify-PCName
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
    $ICSTPCObj|sort pscomputername,Lasttaskresult,TaskName | Out-GridView -Title "Scheduled Tasks"
}
# Get ASCII "plain text" to encode
Function Enter-ASCII_Text
{
    $Script:ASCIIText = Read-Host "Enter the text to encode here"
}
# Get Base64 to decode
Function Enter-Base64
{
    $Script:TextInBase64 = Read-Host "Enter the Base64 to decode here"
}
# Convert from ASCII to Base64 using UTF8 encoding
Function Convert-ASCII_UTF8_To_Base64
{
    Enter-ASCII_Text
    Write-Host "ASCIIText: $ASCIIText"
    $UTF8Bytes = [System.Text.Encoding]::UTF8.GetBytes($ASCIIText)
    Write-Host "UTF8Bytes: $UTF8Bytes" -ForegroundColor DarkGreen
    $TextInBase64 = [System.Convert]::ToBase64String($UTF8Bytes)
    Write-Host "TextInBase64: $TextInBase64`n" -ForegroundColor DarkYellow
}
# Convert from Base64 to ASCII using UTF8 encoding
Function Convert-Base64_ASCII_UTF8
{
    Enter-Base64
    $FromBase64Text = [System.Convert]::FromBase64String($TextInBase64)
    Write-Host "To Binary FromBase64Text: $FromBase64Text" -ForegroundColor DarkGreen
    $DecodedText = [System.Text.Encoding]::UTF8.GetString($FromBase64Text)
    Write-Host "DecodedText: $DecodedText`n" -ForegroundColor Yellow
}
# Convert from ASCII to Base64
Function Convert-ASCII_To_Base64
{
    Clear-Host
    Enter-ASCII_Text
    Write-Host "ASCIIText: $ASCIIText"
    $ASCIIBytes = [System.Text.Encoding]::ASCII.GetBytes($ASCIIText)
    Write-Host "`tASCIIBytes: $ASCIIBytes" -ForegroundColor Gray
    $ToBase64ASCIIText = [System.Convert]::ToBase64String($ASCIIBytes)
    Write-Host "`tToBase64ASCIIText: `n"
    Write-Host "$ToBase64ASCIIText`n" -ForegroundColor DarkGray
}
# Convert from Base64 to ASCII
Function Convert-Base64_To_ASCII
{
    Clear-Host
    Enter-Base64
    $FromBase64ASCIIText = [System.Convert]::FromBase64String($TextInBase64)
    $CodeWheelObject = foreach ($FB64At in $FromBase64ASCIIText)
    {
        $FB64At=$FB64At+2
        $FB64At
        $OBProperties = @{'number'=($FB64At).ToString()}
        New-Object -TypeName PSObject -Prop $OBProperties
    }
    Write-Host "`tCodeWheelObject: $CodeWheelObject" -ForegroundColor Cyan
    $DecodedCodeWheelObject = [System.Text.Encoding]::ASCII.GetString($CodeWheelObject)
    Write-Host "DecodedCodeWheelObject:"
    Write-Host "$DecodedCodeWheelObject"
    Write-Host "`tTo Binary FromBase64ASCIIText: $FromBase64ASCIIText" -ForegroundColor DarkGreen
    $DecodedASCIIText = [System.Text.Encoding]::ASCII.GetString($FromBase64ASCIIText)
    # UTF8.GetString($FromBase64ASCIIText)
    Write-Host "DecodedASCIIText:"
    Write-Host "`t$DecodedASCIIText`n" -ForegroundColor Yellow
}

# This Base64 to ASCII conversion combines the conversion and encoding into one step
Function Convert-Base64ToASCII
{
    Clear-Host
    Enter-Base64
    $DecodedText = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($TextInBase64))
    Write-Host "Base 64 `"$TextInBase64`" decodes to:`n" -ForegroundColor DarkGray
    Write-Host "$DecodedText" -ForegroundColor Yellow
}
Function Run-RemotePS
{
    Clear-Host
    $SBCmd = "hostname;"
    $SBCmd += Read-Host "Enter a Powershell command here"
    Get-PCList
    foreach ($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        # $PCs = Read-Host "Enter a comma separated list of systems here"
        Invoke-Command -ComputerName $Global:PC -ScriptBlock {param($p1) powershell $p1} -Credential $Cred -ArgumentList $SBCmd
    }
}

menu
