<#
.NAME
    Use-MenuPCInfo.ps1
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
    # Remove comment on pcinfo function(s) "Check version without parameters from Use-MenuAD.ps1 if this doen't work"
    Need to fix Get-NonDomPCinfo function Cred prompts *** Fix non-dom cred prompts ***
#>

Clear-Host
[BOOLEAN]$Global:ExitSession=$false
$Get_AllPCInfo = "No"
$Script:PCIt4 = "`t`t`t`t"
# $Script:PCId4 = "------------Use-PC Informational Menu ($Global:PCCnt Systems currently selected)----------------"
$Script:PCINC = $null
$Script:PCI1 = $null
$Script:PCI2 = $null
$Script:PCI3 = $null
[int]$PCmenuselect = 0

Function chPCIcolor($Script:PCI1,$Script:PCI2,$Script:PCI3,$Script:PCINC){
            
            Write-host $Script:PCIt4 -NoNewline
            Write-host "$Script:PCI1 " -NoNewline
            Write-host $Script:PCI2 -ForegroundColor $Script:PCINC -NoNewline
            Write-host "$Script:PCI3" -ForegroundColor White
            $Script:PCINC = $null
            $Script:PCI1 = $null
            $Script:PCI2 = $null
            $Script:PCI3 = $null
}

Function PCIMenu()
{
    while ($PCIMenuselect -lt 1 -or $PCIMenuselect -gt 17)
    {
        Trap {"Error: $_"; Break;}        
        $PCIMNum = 0;Clear-host |out-null
        Start-Sleep -m 50
        if ($Global:PCCnt -eq $null) {$Global:PCCnt = 0}Else {$Global:PCCnt = ($Global:PCList).Count}
        $Script:PCId4 = "------------Use-PC Informational Menu ($Global:PCCnt Systems currently selected)----------------"
        Write-host $Script:PCIt4 $Script:PCId4 -ForegroundColor Green
        # Exit to Main Menu
        $PCIMNum ++;$PCMExit=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t";
            $Script:PCI2 = "Exit to Main Menu";
            $Script:PCI3 = " ";
            $Script:PCINC = "Red";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    #    Prompt for credential with permissions across the network
        $PCIMNum ++;$GCred=$PCIMNum;$Script:PCI1 =" $PCIMNum.  `t Enter";
            $Script:PCI2 = "Credentials ";$Script:PCI3 = "for use in script."; $Script:PCINC = "Cyan";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    # Get PCList
        $PCIMNum ++;$Get_PCList=$PCIMNum;$Script:PCI1 =" $PCIMNum.  `t Get";
            $Script:PCI2 = "PC List ";$Script:PCI3 = "for use in script ($Global:PCCnt Systems currently selected)."; $Script:PCINC = "DarkCyan";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    # PC Info and Service Tag
        $PCIMNum ++;$Get_pcinfo=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t View ";
            $Script:PCI2 = "PC Information ";
            $Script:PCI3 = "and Dell Service Tag"
            $Script:PCINC = "Yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    # Service Tag and BIOS Serial number
        $PCIMNum ++;$Get_BIOSSerialNumber=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t View ";
            $Script:PCI2 = "Service Tag ";
            $Script:PCI3 = "and BIOS Serial number"
            $Script:PCINC = "Yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    # Operating System Information
        $PCIMNum ++;$Get_ArchOSInfo=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t View ";
            $Script:PCI2 = "OS and xArch of select Domain PCs (Compare with above OS information";
            $Script:PCI3 = "(OS) Information"
            $Script:PCINC = "Yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    # CPU Information
        $PCIMNum ++;$Get_CPUInfo=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t View ";
            $Script:PCI2 = "CPU ***FIX***";
            $Script:PCI3 = "Information"
            $Script:PCINC = "Yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    # Memory Information
        $PCIMNum ++;$Get_MemInfo=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t View ";
            $Script:PCI2 = "Memory (RAM) ";
            $Script:PCI3 = "Information"
            $Script:PCINC = "Yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    # Disk Information
        $PCIMNum ++;$Get_DiskInfo=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t View ";
            $Script:PCI2 = "Disk ";
            $Script:PCI3 = "Information"
            $Script:PCINC = "Yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    # Network (NIC) Information
        $PCIMNum ++;$Get_NICStatus=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t View ";
            $Script:PCI2 = "Network (NIC)  ";
            $Script:PCI3 = "Information"
            $Script:PCINC = "Yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    # Monitor Information
        $PCIMNum ++;$List_Monitors=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t View ";
            $Script:PCI2 = "Monitor Information ";
            $Script:PCI3 = "Information"
            $Script:PCINC = "Yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    # Screen Resolution Information
        $PCIMNum ++;$List_Resolutions=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t View ";
            $Script:PCI2 = "Screen Resolution ";
            $Script:PCI3 = "Information"
            $Script:PCINC = "Yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    #    Write-Output "$Script:PCIt4  5. View Non-Domain PC Information and Service Tag"
    # View PC Windows Product key
        $PCIMNum ++;$Get_WinProdKey=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t View ";
            $Script:PCI2 = "Windows Product key ";
            $Script:PCI3 = "for selected systems."
            $Script:PCINC = "Yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    # View Non-Domain PC Information and Service Tag
        $PCIMNum ++;$Get_NonDomPCinfo=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t View ";
            $Script:PCI2 = "non-domain PC Service Tag *** Fix non-dom cred prompts ***";
            $Script:PCI3 = "and generic pc information"
            $Script:PCINC = "Yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    # View System Logon Information
        $PCIMNum ++;$Get_ADPCLogonInfo=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t View ";
            $Script:PCI2 = "AD Logon ";
            $Script:PCI3 = "Information for select systems"
            $Script:PCINC = "Yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    # View System NTP Source Information (Time Server)
        $PCIMNum ++;$Get_NTPSourceInfo=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t View ";
            $Script:PCI2 = "NTP server ";
            $Script:PCI3 = "Information for select systems"
            $Script:PCINC = "Yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
    #    "Discover boot times (from select systems)"
        $PCIMNum ++;$BootTimes=$PCIMNum;
            $Script:PCI1 = " $PCIMNum. `t Discover ";
            $Script:PCI2 = "boot times ";
            $Script:PCI3 = "(Requires Run Script As Administrator)"
            $Script:PCINC = "Yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
        # Setting up the menu in this way allows you to move menu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $PCIMNum ++;$Script:PCI1 =" $PCIMNum.  FIRST";$Script:PCI2 = "HIGHLIGHTED ";$Script:PCI3 = "END"; $Script:PCINC = "yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
        Write-Host "$Script:PCIt4 $Script:PCId4"  -ForegroundColor Green
        $Script:MNum = 0;[int]$PCIMenuselect = Read-Host "`n`tPlease select the menu item you would like to do now"
    }

switch($PCIMenuselect)
    {
        $PCMExit{$PCIMenuselect=$null;reload-PCmenu}
        $GCred{Clear-Host;Get-Cred;$PCIMenuselect = $null;Reload-PCInfoMenu} # Called from Run-PSMenu.ps1
        $Get_PCList{Clear-Host;Get-PCList;Reload-PCInfoMenu}
        $Get_pcinfo{Clear-Host;Get-DInfo;Get-MemInfo;Get-CPUInfo;Get-BIOSSerialNumber;. ./Use-MenuNet.ps1;Call-Get-NICStatus;Get-ArchOSInfo;List-Resolutions;List-Monitors;Reload-PromptPCInfoMenu}
        $Get_BIOSSerialNumber{Clear-Host; Get-BIOSSerialNumber;Reload-PromptPCInfoMenu}
        $Get_ArchOSInfo{Clear-Host; Get-ArchOSInfo;Reload-PromptPCInfoMenu}
        $Get_CPUInfo{Clear-Host; Get-CPUInfo;Reload-PromptPCInfoMenu}
        $Get_MemInfo{Clear-Host; Get-MemInfo;Reload-PromptPCInfoMenu}
        $Get_DiskInfo{Clear-Host; Get-DInfo;Reload-PromptPCInfoMenu}
        $Get_NICStatus{Clear-Host;$Global:NICSCMD;$Get_AllPCInfo = "NICS";. ./Use-MenuNet.ps1;Call-Get-NICStatus;Reload-PromptPCInfoMenu}
        $List_Monitors{Clear-Host;List-Monitors;Reload-PromptPCInfoMenu}
        $List_Resolutions{Clear-Host;List-Resolutions;Reload-PromptPCInfoMenu}
        $Get_WinProdKey{Clear-Host;Get-WinProdKey;Reload-PromptPCInfoMenu}
        $Get_NonDomPCinfo{Clear-Host;Get-NonDomPCinfo;Reload-PCInfoMenu}
        $Get_ADPCLogonInfo{Clear-Host;Get-ADPCLogonInfo;Reload-PromptPCInfoMenu}
        $Get_NTPSourceInfo{Clear-Host;Get-NTPSourceInfo;Reload-PromptPCInfoMenu}
        $BootTimes{Clear-Host;.\Start-MultiThreads.ps1 .\Get-BootTime.ps1;Reload-PromptPCInfoMenu} # Called Externally
        default
        {
            # cls
            $PCIMenuselect = $null
            Reload-PCInfoMenu #clear-host;break;
        }
    }    
}

Function reload-PromptPCInfoMenu() # Reload with prompt (not needed until there's another menu down)
{
    Read-Host "Hit Enter to continue"
    $PCIMenuselect = $null
    PCIMenu
}
Function Reload-PCInfoMenu() # Reload with NO prompt
{
        $PCIMenuselect = $null
        PCIMenu
}

Function Get-NonDomPCinfo()
{
    Write-Warning "Enter a computername to see details"
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    $DomName =  $env:USERDNSDOMAIN
    if (($DomName = Read-Host "Enter the domain name [Enter] to use default: $env:USERDNSDOMAIN") -eq "") {$DomName = $env:USERDNSDOMAIN}
    foreach($Global:pcLine in $Global:PCList)
    {
        Identify-PCName # Called from Run-PSMenu.ps1
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
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
        $memResults += $mem
        # $mem
        }
    }
    $memResults
}

# Get the Dell Service Tag # from pc list
Function Get-BIOSSerialNumber
{
$BSNObjectOut = @()
$GBiosSN = @()
    Write-Host 'Running: `nget-wmiobject -Credential $cred -computername $Global:PC -class win32_bios |select "$Global:PC", SerialNumber |fl'
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
        $BSNObjectOut = Foreach($Global:PCLine in $Global:PCList)
        {
        Identify-PCName
            If (Test-Connection $Global:pc -Count 1 -Quiet)
            {
            Write-host "$Global:PC, " -NoNewline -ForegroundColor Green
            $GBiosSN = get-wmiobject -Credential $cred -computername $Global:PC -class win32_bios |Select SerialNumber
            New-Object PSObject -Property @{
                pscomputername = $Global:PC
                SerialNumber = $GBiosSN.SerialNumber
                }
            }
            else
            {
                Write-host "$Global:PC, " -NoNewline -ForegroundColor Red
            }
        }
    $BSNObjectOut|Sort pscomputername |FT pscomputername, SerialNumber
}
# Get OS information from pc list
Function Get-ArchOSInfo
{
$AOSObjectOut = @()

    Write-Host 'Running: `nGet-WmiObject -class Win32_OperatingSystem -Impersonation Impersonate -Credential $cred -computername $Global:PC  -Property * |Select Caption, OSArcheticture,Description, ServicePackMajorVersion, Version, BuildNumber'
    $AOSObjectOut = foreach($Global:pcLine in $Global:pclist)
    {
        $AOSOBProperties = @()
        $sOS = @()
        Identify-PCName
        If (Test-Connection $Global:pc -Count 1 -Quiet)
            {
        Write-host "$Global:PC, " -NoNewline -ForegroundColor Green
            $sOS = Get-WmiObject -class Win32_OperatingSystem -Impersonation Impersonate -Credential $cred -computername $Global:PC |Select PSComputerName, Caption, OSArchitecture, Version, BuildNumber, @{Label = "SP Ver" ;Expression={$_.ServicePackMajorVersion}}, Description
            #$sOS |Select PSComputerName, Caption, OSArchitecture, Version, BuildNumber, @{Label = "SP Ver" ;Expression={$_.ServicePackMajorVersion}}, Description
            $AOSOBProperties = @{'pscomputername'=$Global:PC; # $Global:PC
                'OS'= $sOS.Caption;
                'x-Arch'= $sOS.OSArchitecture;
                'Description'= $sOS.Description;
                'SP_Version'= $sOS.ServicePackMajorVersion;
                'Version'= $sOS.Version;
                'BuildNumber'= $sOS.BuildNumber}
            New-Object -TypeName PSObject -Prop $AOSOBProperties
            }
            else
            {
                Write-host "$Global:PC, " -NoNewline -ForegroundColor Red
            }
    }
    $AOSObjectOut |sort PSComputerName |Select PSComputerName, OS, x-Arch, Version, BuildNumber, Description, SP_Version |FT
}

# Get CPU Information for listed system
Function Get-CPUInfo
{
$cpuinfo = @()
$CPUIObj = @()
$CPUIObjectOut = @()
    foreach($Global:pcLine in $Global:pclist)
    {
    Identify-PCName
    # $CPUIObj = @()
    If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
        Write-host "$Global:PC, " -NoNewline -ForegroundColor Green
        $cpuinfo = get-wmiobject win32_processor -Impersonation Impersonate -Credential $cred -ComputerName $Global:PC |select-object pscomputername,role, name, numberofcores,numberoflogicalprocessors,
                @{Label = "CPUStats" ;Expression={
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
                $CPUIObj = foreach($CPUILine in $cpuinfo)
                {
                $CPUIOBProperties = @{'pscomputername'=$Global:PC;
                    CPUStatus = $CPUILine.cpustats;
                    role= $CPUILine.role;
                    name= $CPUILine.name;
                    cores= $CPUILine.numberofcores;
                    logicalprocessors= $CPUILine.numberoflogicalprocessors}
                New-Object -TypeName PSObject -Prop $CPUIOBProperties
                }
        }
        else
        {
            Write-host "$Global:PC, " -NoNewline -ForegroundColor Red
        }
    $CPUIObjectOut += $CPUIObj
    }
Write-Host "`nThe following is the information on the CPU"
$CPUIObjectOut |sort PSComputerName |FT PScomputerName,CPUStatus,Role,cores,logicalprocessors,Name
}


# Get and display memory information for listed system
Function Get-MemInfo
{
$MemObjOut = @()
$MemObj = @()
$mem = @()
Write-Host 'Running: `nGet-WmiObject cim_physicalmemory -Impersonation 3 -Credential $cred -ComputerName $Global:PC | select banklabel,capacity,caption,devicelocator,partnumber,speed |ft -AutoSize'
    foreach($Global:pcLine in $Global:pclist)
    {
        Write-host "$Global:PC, " -NoNewline -ForegroundColor Green
        Identify-PCName
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
            $mem = Get-WmiObject cim_physicalmemory -Impersonation 3 -Credential $cred -ComputerName $Global:PC |
                select pscomputername,caption,devicelocator,speed,@{Label = "Gigs" ;Expression={($_.capacity / 1048576).ToString('N0')}},partnumber,banklabel
            $MemObj = foreach($memline in $mem)
            {
                $MemObjProperties = @{
                    pscomputername = $Global:PC;
                    caption = $memline.caption;
                    devicelocator = $memline.devicelocator;
                    speed = $memline.speed;
                    Gigs = $memline.Gigs;
                    partnumber = $memline.partnumber;
                    banklabel = $memline.banklabel
                    }
            New-Object -TypeName PSObject -Prop $MemObjProperties
            }
        }
        else
        {
            Write-host "$Global:PC, " -NoNewline -ForegroundColor Red
        }
    $MemObjOut += $MemObj
    }
$MemObjOut |sort PSComputerName| Select pscomputername,Gigs,speed,devicelocator,caption,banklabel,partnumber | FT -AutoSize
}

# Get Disk information for system
Function Get-DInfo
{
$DIObjOut = @()
$DIObj = @()
$DiskInfo = @()
Write-Host "`nRunning: get-wmiobject win32_logicaldisk -Credential $cred -Computername $Global:PC | select systemname, Name, deviceid, description, Drivetype, mediatype, filesystem, @{n=`'MBfreespace`';e={[string]::Format(`"{0,8:#,##0.0} GB`",[math]::Round($_.FreeSpace/1GB,2))}}, @{n=`'MBSize`';e={[string]::Format(`"{0,8:#,##0.0} GB`",[math]::Round($_.Size/1GB))}}, volumedirty, volumename, volumeserialnumber"
    foreach($Global:pcLine in $Global:pclist)
    {
        Write-host "$Global:PC, " -NoNewline -ForegroundColor Green
        Identify-PCName
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
            $DiskInfo = get-wmiobject win32_logicaldisk -Credential $cred -Computername $Global:PC | select systemname, Name, deviceid, description, Drivetype, mediatype, filesystem, @{n='MBfreespace';e={[string]::Format("{0,8:#,##0.0} GB",[math]::Round($_.FreeSpace/1GB,2))}}, @{n='MBSize';e={[string]::Format("{0,8:#,##0.0} GB",[math]::Round($_.Size/1GB))}}, volumedirty, volumename, volumeserialnumber
            $DIObj = @()            
            $DIObj = foreach($DILine in $DiskInfo)
            {
                if ($DILine.MBSize -eq "0.0 GB" -and $DILine.MBfreespace -eq "0.0 GB"){$DILine.MBSize = "-";$DILine.MBfreespace = "-"}
                $DIObjProperties = @{
                    pscomputername = $Global:PC;
                    Name = $DILine.Name;
                    deviceid = $DILine.deviceid;
                    description = $DILine.description;
                    Drivetype = $DILine.Drivetype;
                    mediatype = $DILine.mediatype;
                    filesystem = $DILine.filesystem;
                    MBfreespace = $DILine.MBfreespace;
                    MBSize = $DILine.MBSize;
                    volumedirty = $DILine.volumedirty;
                    volumename = $DILine.volumename;
                    volumeserialnumber = $DILine.volumeserialnumber
                    }                
            New-Object -TypeName PSObject -Prop $DIObjProperties
            }
        }
        else
        {
            Write-host "$Global:PC, " -NoNewline -ForegroundColor Red
        }
    $DIObjOut += $DIObj
    }
$DIObjOut |sort PSComputerName| Select pscomputername, Name, MBfreespace, MBSize, volumename, description, filesystem |FT -AutoSize #, volumedirty, volumeserialnumber, Drivetype, mediatype, deviceid | Out-GridView 
""
}

Function Get-WinProdKey
{
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }   
    Write-Warning "`nRunning: Get-WmiObject -Query "select * from SoftwareLicensingService" -ComputerName $Global:PC |select __server,OA3xOriginalProductKey"
    $WPkObjectOut = foreach($Global:pcLine in $Global:pclist)
    {
        Write-host "$Global:PC, " -NoNewline -ForegroundColor Green
        Identify-PCName
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
            $WPkObWMI = Get-WmiObject -Query "select * from SoftwareLicensingService" -ComputerName $Global:PC|select __server,OA3xOriginalProductKey
        $WPkObProperties = @{'pscomputername'=$Global:PC;
            'OA3xOriginalProductKey'= $WPkObWMI.OA3xOriginalProductKey;
            'BuildNumber'= $sOS.BuildNumber}
        New-Object -TypeName PSObject -Prop $WPkObProperties
        }
        else
        {
            Write-host "$Global:PC, " -NoNewline -ForegroundColor Red
        }
    }
    $WPkObjectOut |FT pscomputername,OA3xOriginalProductKey
}
 
Function Get-ADPCLogonInfo()
{
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    #$adpcrslt = 
    foreach ($pcline in $Global:pclist)
    {
        Identify-PCName # Called from Run-PSMenu.ps1
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
        $adpcline = get-adcomputer $Global:pc -properties *|select Name,Enabled,Description,@{n="LastLogonyyyyMM";e={($_.lastlogondate).tostring("yyyyMMdd")}}
        <#
        $ADPCLineProperties = @{'pscomputername'=$Global:PC; # $Global:PC
                'Name'= $adpcline.Name;
                'Enabled'= $adpcline.Enabled;
                'Description'= $adpcline.Description;
                'LastLogonyyyyMM'= $adpcline.LastLogonyyyyMM}
            New-Object -TypeName PSObject -Prop $ADPCLineProperties
            #>
            $adpcResults += $adpcline

        }
    }
    $adpcrslt|sort LastLogonyyyyMM |ft
}
# View NTP Source information for selected systems
function Get-NTPSourceInfo{
# Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    
Write-Warning 'Running: w32tm /query /computer:$Global:PC /configuration | ?{$_ -match ''ntpserver:''} | %{($_ -split ":\s\b")[1]}'
$ntpsResult = foreach ($Global:PCLine in $Global:PCList)
    {
        Identify-PCName # Called from Run-PSMenu.ps1
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
            # $ntps = w32tm /query /computer:$Global:PC /SOURCE |new-object psobject -property @{Server = $Global:PC;NTPSource = $ntps}
            $ntps = w32tm /query /computer:$Global:PC /configuration | ?{$_ -match 'ntpserver:'} | %{($_ -split ":\s\b")[1]}
            new-object psobject -property @{Server = $Global:PC;NTPSource = $ntps}
            # w32tm /config /computer:$Global:PC /manualpeerlist:"2012DC.lec.lan LECv002.lec.lan tick.usno.navy.mil 0.us.pool.ntp.org" /syncfromflags:MANUAL
            # Get-Service -Name W32Time -ComputerName $Global:PC | Restart-Service -Confirm
        }
    }
$ntpsResult |ft -AutoSize
}
function List-Resolutions{
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    $LRObjectOut =""
    $LRObjectOut = foreach ($Global:PCLine in $Global:pclist)
    {
        Identify-PCName # Called from Run-PSMenu.ps1
        $i = 0
        $TC = Test-connection $Global:PC -Count 1 -ErrorAction SilentlyContinue -ErrorVariable TCError 
        if ($TC.StatusCode -eq 0)
        {
            Write-host "." -NoNewline -ForegroundColor Green
            $VidResList = GWMI win32_videocontroller -ComputerName $Global:PC -ErrorAction SilentlyContinue -ErrorVariable VidCError|where {$_.name -notlike "DameWare*"}|where {$_.currentHorizontalresolution -ne $null}|select * # |select PSComputerName,Name,Current*Resolution,PNPDeviceID
            foreach ($VidRes in $VidResList)
            {
                $HR = $VidRes.CurrentHorizontalResolution
                $VR = $VidRes.CurrentVerticalResolution
                $i++
                $LROBProperties = @{'pscomputername'=$Global:PC;
                    'HRes'=$HR;
                    'VRes'=$VR;
                    'Count'=$i;
                    'Error'=$VidCError}
                New-Object -TypeName PSObject -Prop $LROBProperties
            }
        }
    }
    #"$HR x $VR Resolution"
    $LRObjectOut |select pscomputername,HRes,VRes,Count,Error |FT
}
function List-Monitors{
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    $LDObjectOut =""
    $LDObjectOut = foreach ($Global:PCLine in $Global:pclist)
    {
        $i = 0
        Identify-PCName # Called from Run-PSMenu.ps1
        $TC = Test-connection $Global:PC -Count 1 -ErrorAction SilentlyContinue -ErrorVariable TCError 
        if ($TC.StatusCode -eq 0)
        {
            Write-host "." -NoNewline -ForegroundColor Green
            $MonitorID = gwmi WmiMonitorID -Namespace root\wmi -ComputerName $Global:pc -ErrorAction SilentlyContinue -ErrorVariable MonitorError |select PSComputerName,YearOfManufacture,InstanceName,SerialNumberID,UserFriendlyName,PNPDeviceID
            $DisplayParam = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams -ComputerName $Global:PC -Property * -ErrorAction SilentlyContinue -ErrorVariable DisplayError |select *
            foreach($Display in $DisplayParam)
            {
                $DisplayInstanceName = ($Display.CimInstanceProperties) |where {$_.Name -eq "InstanceName"}|select Value 
                $HSize = ($Display.CimInstanceProperties) |where {$_.Name -eq "MaxHorizontalImageSize"}|select Value # horizontal size in cm
                $VSize = ($Display.CimInstanceProperties) |where {$_.Name -eq "MaxVerticalImageSize"}|select Value # vertical size in cm
                $DiagSqrd = [math]::Pow($HSize.Value,2) + [math]::Pow($VSize.Value,2) # find the diagonal value from horizontal & vertical values in cm
                $DiagSize = ([math]::Sqrt($DiagSqrd))/2.54 # (2.54 cm / inch)
                $DiagSize = [math]::Round($DiagSize,2)
                $MonitorInfo = $MonitorID|where {$_.InstanceName -eq $DisplayInstanceName.Value} |select PSComputerName,YearOfManufacture,@{n="Model";e={[System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName -ne 00)}}# @{n="Serial Number";e={[System.Text.Encoding]::ASCII.GetString($_.SerialNumberID -ne 00)}}
                $Model = ($MonitorInfo.Model)
                $MonYear = ($MonitorInfo.YearOfManufacture)
                # "Monitor $i ($Model) on $Global:PC is $DiagSize inches (Made $MonYear)"
                $i++
                $LDOBProperties = @{'pscomputername'=$Global:PC;
                    'Model'=$Model;
                    'DiagSize'=$DiagSize;
                    'MonYear'=$MonYear;
                    'Count'=$i;
                    'Error'=$MonitorError}
                New-Object -TypeName PSObject -Prop $LDOBProperties
            }
        }
        Else
        {
            Write-host "." -NoNewline -ForegroundColor Red
            foreach ($TCErrorLine in $TCError)
            {
                $i++
                $TC=""
                $LDOBProperties = @{'pscomputername'=$Global:PC;
                    'Model'="Unknown";
                    'DiagSize'="Unknown";
                    'MonYear'="Unknown";
                    'Count'=$i;
                    'Error'=$TCErrorLine}
                New-Object -TypeName PSObject -Prop $LDOBProperties
            }
        }
    }
    # $LDObjectOut |gm
    $LDObjectOut |select PSComputerName,Model,DiagSize,MonYear,Count,Error |ft
}







