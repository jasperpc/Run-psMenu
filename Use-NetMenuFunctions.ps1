<#
.NAME
<<<<<<< HEAD
    Use-NetMenuFunctions.ps1
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
        .\Use-NetMenuFunctions.ps1 []
.SYNTAX
    .\Use-NetMenuFunctions.ps1 []
.REMARKS
    To see the examples, type: help Use-NetMenuFunctions.ps1 -examples
    To see more information, type: help Use-NetMenuFunctions.ps1 -detailed
=======
        .\adutils.ps1 []
.SYNTAX
    .\adutils.ps1 []
.REMARKS
    To see the examples, type: help adutils.ps1 -examples
    To see more information, type: help adutils.ps1 -detailed
>>>>>>> a1255d3f753d7d5f3e2386825dab4574a19aabcf
#>
[CmdletBinding()] #(allows the use of parameter attributes)
param(
        #[Parameter(Mandatory=$True)]
        [string[]]$comps = "$env:COMPUTERNAME" #was 'localhost' This can change it's just here to hold the place and illustrate how/where to insert parameters & for CmdletBinding to not fail
)


# Enable Exchange cmdlets (On an Exchange Server)
# add-pssnapin *exchange* -erroraction SilentlyContinue 

# Enable AD and PS Update modules
# See Getting Started with PowerShell 3.0 Jump Start in MS Virtual Academy for loading modules from other systems
# $ADPS=new-pssession -computername $ADSRV
Import-Module ActiveDirectory
Import-Module PSWindowsUpdate
# Import-Module AzureAutomationAuthoringToolkit

Clear-Host
# set script variables
[BOOLEAN]$Global:ExitSession=$false
# Get Credentials for later wmi requests
# $cred = Get-Credential "$env:USERDOMAIN\Administrator"

$Script:t4 = "`t`t`t`t"
$Script:Netd4 = "-------------------------------------------------------------------------------"
$Script:NC = $null
$Script:p1 = $null
$Script:p2 = $null
$Script:p3 = $null
$CDate = Get-Date -UFormat "%Y%m%d"
[int]$NetMenuSelect = 0

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

Function NetMenu()
{
while ($NetMenuSelect -lt 1 -or $NetMenuSelect -gt 18)
    {
        Trap {"Error: $_"; Break;}        
        $MNum = 0;Clear-host |out-null
        Start-Sleep -m 50
        Write-host $t4 $Script:Netd4 -ForegroundColor Green
        # Exit
        $MNum ++;$Exit=$MNum;
            $p1 = " $MNum. ";
            $p2 = "Exit to Main Menu";
            $p3 = " ";
            $NC = "Red";chcolor $p1 $p2 $p3 $NC
        # Test PC(s) for network connectivity
        $MNum ++;$test_Connections=$MNum;
            $p1 =" $MNum.  Run";
            $p2 = "Test-Connection ";
            $p3 = "on select PCs or all PCs in Domain";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Firewall Status
        $MNum ++;$FW_Status=$MNum;
            $p1 =" $MNum.  View";
            $p2 = "Firewall Status ";
            $p3 = "on select PCs or all PCs in Domain";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Test Ports
        $MNum ++;$test_ports=$MNum;
            $p1 =" $MNum.  Test";
            $p2 = "ports ";
            $p3 = "on a system or list of systems";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Ping
        $MNum ++;$Ping_Test=$MNum;
            $p1 =" $MNum.  Test";
            $p2 = "(Ping) Network ";
            $p3 = "Connection Persistence";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # IP from MAC
        $MNum ++;$Get_IP=$MNum;
            $p1 =" $MNum. ";
            $p2 = "MAC 2 IP ";
            $p3 = "- Use a known MAC address to find an IP address";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Status of NIC on PC
        $MNum ++;$Call_Get_NICStatus=$MNum;
            $p1 =" $MNum. Get ";
            $p2 = "NIC ";
            $p3 = "Status and information";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # View IPConfig-like information on select systems
        $MNum ++;$Get_IPConfig=$MNum;
            $p1 = " $MNum. View ";
            $p2 = "IPConfig";
            $p3 = " information for select systems"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC   
        # HostName from IP
        $MNum ++;$Get_hostname=$MNum;
            $p1 =" $MNum.  Find";
            $p2 = "hostname from IP ";
            $p3 = "address";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # Validate IP
        $MNum ++;$Test_ValidIPPrompt=$MNum;
            $p1 =" $MNum.  Identify";
            $p2 = "valid IP ";
            $p3 = "address from input";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # NetStat
        $MNum ++;$View_NetStat=$MNum;
            $p1 =" $MNum.  View";
            $p2 = "NetStat results ";
            $p3 = "for a system on the network";
            $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        # nlasvc restart
        $MNum ++;$Restart_NLASvc=$MNum;
            $p1 = " $MNum.  Restart";
            $p2 = "nlasvc";
            $p3 = ", the local Network Location Awareness Service";
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
        # Set NIC location
        $MNum ++;$Get_Set_ncp=$MNum;
            $p1 = " $MNum. Get ";
            $p2 = "Network Location ";
            $p3 = " and change it (with confirm) of a local NIC ";
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
        # Wait for live port
        $MNum ++;$Wait_Live=$MNum;
            $p1 = " $MNum. Wait for ";
            $p2 = "System port number";
            $p3 = " to come alive";
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
                # Wait for live port
        $MNum ++;$Test_CritSysLive=$MNum;
            $p1 = " $MNum. Check ";
            $p2 = "Active Status";
            $p3 = " for critical systems";
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
        # Check for Active RDP sessions on Domain servers
        $MNum ++;$Get_RDPStats=$MNum;
            $p1 = " $MNum. Check ";
            $p2 = "for Console or RDP";
            $p3 = " sessions on Domain servers or select systems.";
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
        # Check for Primary subnet of AD Site
        $MNum ++;$Get_ADSitSubnet=$MNum;
            $p1 = " $MNum. View ";
            $p2 = "AD Site and subnet";
            $p3 = " for the current Active Directory context"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC     
        # Browse-Share
        $MNum ++;$Browse_Share=$MNum;
            $p1 = " $MNum. Incomplete: Browse ";
            $p2 = "SMB Share ";
            $p3 = "on a system"
            $NC = "Yellow";chcolor $p1 $p2 $p3 $NC
            
        # Setting up the menu in this way allows you to move menu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $MNum ++;$p1 =" $MNum.  FIRST";$p2 = "HIGHLIGHTED ";$p3 = "END"; $NC = "yellow";chcolor $p1 $p2 $p3 $NC
        Write-Host "$t4$Script:Netd4"  -ForegroundColor Green
        $Script:MNum = 0;[int]$NetMenuSelect = Read-Host "`n`tPlease select the menu item you would like to do now"
    }


    switch($NetMenuSelect)
        {
            # Variables at the beginning of these lines in the switch are numbers assigned by their position in the above
            # configured menu layout. This allows them to be moved around and let the numbers automatically adjust
            $Exit{$NetMenuSelect=$null;reloadmenu}
            $test_Connections{test-connections;reloadnetmenu}
            $test_ports{test-ports;ReloadNetMenu}
            $FW_Status{FW-Status;ReloadNetMenu}
            $Restart_NLASvc{Restart-NLASvc;ReloadNetMenu}
            $Get_Set_ncp{Get-Set-ncp;ReloadNetMenu}
            $Ping_Test{Test-PCUp;ReloadNetMenu} #PingTest;reloadnetmenu}
            $Get_IP{Get-IP;ReloadNetMenu}
            $Get_IPConfig{Get-IPConfig;reloadnetmenu}
            $Call_Get_NICStatus{Call-Get-NICStatus;reloadnetmenu}
            $Get_hostname{Get-hostname;ReloadNetMenu}
            $Test_ValidIPPrompt{clear-host|Test-ValidIPPrompt;ReloadNetMenu}
            $View_NetStat{View-NetStat;ReloadNetMenu}
            $Wait_Live{Wait-Live;ReloadNetMenu}
            $Test_CritSysLive{Test-CritSysLive;reloadnetmenu}
            $Get_RDPStats{Get-RDPStats;reloadnetmenu}
            $Get_ADSitSubnet{Get-ADSiteSubnet;reloadnetmenu}
            $Browse_Share{Browse-Share;reloadmenu}

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
Function ReloadNetMenu()
{
        Read-Host "Hit Enter to continue"
        # Start-Sleep -m 250
        $NetMenuSelect = $null
        NetMenu
}
Function test-connections()
{
    Clear-Host
    $PP = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    Get-PCList
    <# Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    #>
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
Function test-ports()
{
    Clear-Host
    Get-PCList
    $port = Read-Host "`nEnter a port, CSV port list, or list..range "
    foreach ($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        Write-Host ". " -nonewline
        $IETPResult = Invoke-Expression ".\Test-Ports.ps1 -computer $Global:pc -port @($port)" # |FT
        $IETPRDetails += foreach ($IETPDetail in ($IETPResult |where {$_.port -ne $null}))
        {
            $IETPOBProperties = @{'pscomputername'=$Global:PC;    
                    'Port'=($IETPDetail.Port);
                    'TypePort'=($IETPDetail.TypePort);
                    'Open'=($IETPDetail.Open);
                    'Notes'=($IETPDetail.Notes)}
                New-Object -TypeName PSObject -Prop $IETPOBProperties
        }
    }
    $IETPRDetails|sort pscomputername,Open,Port |ft pscomputername,Port,Open,TypePort,Notes -Wrap -AutoSize
}
Function FW-Status()
{
    # If we cannot ping PC could be off or Firewall could be on
    # If we can ping but cannot get profile status RemoteRegistry Service could be off
    # Need to add a test for service being disabled and if it is disabled, a prompt to enable it and set its type to auto or demand, etc
    Clear-Host
    $RegErr = $false
    if (($FWprofile = Read-Host "Type firewall profile name to query or [Enter] to use default: DomainProfile") -eq "") {$FWprofile = "DomainProfile"}
    Get-pclist
foreach($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        $GSRemReg = Get-Service Remoteregistry -ComputerName $Global:PC -ErrorAction SilentlyContinue
        # Saving the status of the RemoteRegistry Service to the $RegStat variable
        $RegStat = $GSRemReg.Status
        if ($RegStat -eq $null)
            {
                Write-host "Unable to get RemoteRegistry Status for $Global:PC" -ForegroundColor Red
            }
        else
            {
                if ($RegStat -eq "stopped")
                    {
                        Write-host "$Global:PC RemoteRegistry Status is $RegStat " -foregroundcolor Yellow
                        # Attempt to start the registry service prior to continuing on (could add a verification of success here)
                        write-host "Trying to start Remote Registry Service on $Global:PC " -ForegroundColor DarkYellow
                        $GSRemReg |Start-Service |Out-Null
                    }
                 
        # Write-Host "RegStat" -ForegroundColor Green
        try
            {
                $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$Global:PC)
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
                        Write-Host "Error connecting to registry on $Global:PC" -ForegroundColor Red
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
                        $FWstatus = [bool]$firewallEnabled
                    }
            }
            If ($FWstatus -eq $false)
                {
                    
                    Write-Output "$Global:PC $FWprofile firewall is off "
                }
            Elseif($FWstatus -eq $true)
                {
                    Write-host "$Global:PC $FWprofile firewall is on " -ForegroundColor Green
                }
            Else
                {
                    "System is off or Firewall is blocking communication"
                }

            }
        #Reset FWStatus to true for next query
        $FWstatus = $null # $true
    }
}

Function Restart-NLASvc()
{
<#
In Windows 10, the nlasvc service must sometimes be restarted to force the network location to identify itself correctly
When it does not know where it is, Outlook cannot connect to the exchange server even if the pc has an IP address and can get to the internet.
The firewall on the default public location is much more restrictive
#>
    clear-host
    Write-Host "`nRestarting the Network Location Awareness service (nlasvc)" -ForegroundColor Green
    Restart-Service nlasvc -Force
    $getnlas = Get-Service nlasvc,netprofm
    $getnlas
    
}
Function Get-Set-ncp
{
    Clear-Host
    $II = ""
    $NetCategory = ""
    $GNcP = Get-NetConnectionProfile
    $GNcP |sort NetworkCategory -Descending|FT Name,NetworkCategory,InterfaceAlias,InterfaceIndex,IPv4Connectivity,IPv6connectivity # Caption,Description,ElementName,PSComputerName,InstanceID
    $II = Read-Host "Enter the InterfaceIndex number from the above network Interface to change that adapter or [ENTER] to return to the menu."
    If ($II -ne "")
    {
        write-host "You selected InterfaceIndex number `"$II`" to change.`n" -ForegroundColor Yellow
        $NetCategory = Read-Host "Which network category should this connection profile be configured as? Enter Public, Private, (or!!UNTESTED `"DomainAuthenticated`"!!)"
        if ($NetCategory -eq "Public" -or $NetCategory -eq "Private" -or $NetCategory -eq "DomainAuthenticated")
        {
            Set-NetConnectionProfile -InterfaceIndex $II -NetworkCategory $NetCategory -ErrorAction SilentlyContinue -Confirm
        }
        Else
        {
            Write-Warning "`"$NetCategory`" is not an option.`nYour options are Public, Private, or DomainAuthenticated."
            Start-sleep -Seconds 4
            Get-Set-ncp
        }
    }
    
}
# This function basically pings a system to see if it is alive. It logs the data and time and whether it is accessible or not
Function Test-PCUp()
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
            $key = [System.Console]::Readkey($true)
            Console.Read
            "ky is $ky"
        }
    }
    finally
    {
            Write-Host "ky is $ky"
    }
    Write-Host "ReloadNetMenu (function is done)"
}
Function Get-IP()
{
    Clear-Host
    Write-Host "`nYou may need to ping or communicate with a host prior to running this.`nTry pinging the broadcast address first."
    $macaddr = Read-Host "`nEnter some or all of a MAC address in the format AC-86-...."
    arp -a | select-string $macaddr #|% {$_.ToString().Trim().Split(" ")[0]}
    
}
Function Get-IPConfig()
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
       }
      }
     }
    Write-Warning Sending results to: $ObCSVFile
    $OutputObj
    $OutputObj|Select computername, SubnetMask, DNSServers, Gateway, IPAddress, IsDHCPEnabled, MACAddress, WINSServers, Error |Export-Csv $ObCSVFile -NoClobber -NoTypeInformation -Append
}
#Get Network information from pc
Function Get-NICStatus()
{
    Get-WmiObject -Class win32_networkadapterconfiguration -ComputerName $Global:pc -Credential $cred|where {$_.IPAddress -ne $null}|select PSComputerName,MACAddress,IPSubnet,DNSServerSearchOrder, @{n="IP Address";e={($_ | select -ExpandProperty IPAddress)-join ','}},DHCPEnabled # |ft #out-grid-view
}
Function Call-Get-NICStatus()
{
    Clear-Host
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
Function Get-hostname()
    {
        Clear-Host
        [array]$IPlc=(Read-Host "Enter a comma separated list of IP Addresses").split(",") | %{$_.trim()}
        foreach($IP in $IPlc)
        {
            $GHAddress = [System.Net.DNS]::GetHostByAddress($IP)
            $GHAddress | select * |ft
        }
        
    }
# Prompt for IP Address from User. Combine calls Test-ValidIP to test input for valid IP Address
Function Test-ValidIPPrompt
{
try {
        $gip = Read-host "`nType in a valid IP address to find information about"
        Test-ValidIP $gip -erroraction "stop"
            
    }
catch
    {
        $Another = Read-Host "`nInvalid IP Address. Type Y to enter another."
        while($Another -eq "Y")
            {
                Test-ValidIPPrompt
            }
    }
}
Function Test-ValidIP # Test input for valid IP Address
{
# Manually this can be called from Test-ValidIPPrompt
# This function illustrates parameter validation and using RegEx to validate IPv4 IP addresses including broadcast address
param ([ValidatePattern('^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$')]
[string]$gip
)
    Write-Output "$gip is a valid IP address. One moment please."
}
# View NetStat results for a system including the process name and active ports
Function View-NetStat
{
    Invoke-Expression .\Get-NetStat.ps1
        
}
Function Wait-Live()
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

    
}
Function Test-CritSysLive()
{
    Clear-Host
    $err=@()
    Write-Warning ".\Critsys.txt contains a list of critical systems to check with this function."
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
}
Function Get-RDPStats()
{
# Compare with code at: https://superuser.com/questions/587737/powershell-get-a-value-from-query-user-and-not-the-headers-etc#702328
Clear-Host
$InvkCmd=$null
$Active=$false
$SN=$null
Get-PCList
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
    If ($Active -eq $false)
    {
        Write-Host "No active RDP sessions found" -ForegroundColor Red
        start-sleep -Milliseconds 500
    }
    $ObjectOut |FT PSComputerName,ID,UserName,State,Idle_Time,Logon_Time,SessionName,Status -AutoSize -Wrap
}

Function Get-ADSiteSubnet()
{
    get-adreplicationsite -Filter * -Properties * |select Name,siteObjectBL,Subnets,Group |group siteObjectBL |select @{Name='Subnet';Expression={$_.Name.split(',')[0].Trim('{CN=')}},@{Name='Name';Expression={$_.Group.Name}} |ft
}

Function Browse-Share
{
param(
  [Parameter(ValueFromPipelineByPropertyName=$true)]
  [string[]] $ShareName = ($ShareName=if ($SN = Read-Host -Prompt "Please enter a UNC server and sharename ('\\$env:USERDNSDOMAIN\dfs\SharedFiles')") {$SN} else {"\\$env:USERDNSDOMAIN\dfs\SharedFiles"})
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
    
}