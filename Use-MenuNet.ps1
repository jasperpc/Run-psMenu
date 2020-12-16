<#
.NAME
    Use-MenuNet.ps1
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
        .\Use-MenuNet.ps1 []
.SYNTAX
    .\Use-MenuNet.ps1 []
.REMARKS
    To see the examples, type: help Use-MenuNet.ps1 -examples
    To see more information, type: help Use-MenuNet.ps1 -detailed
#>
# set script variables
# [BOOLEAN]$Global:ExitSession=$false
# Get Credentials for later wmi requests
# $cred = Get-Credential "$env:USERDOMAIN\Administrator"
# $CDate = Get-Date -UFormat "%Y%m%d"

$Script:Nt4 = "`t`t`t`t"
$Script:Nd4 = "------ Use-Net Menu ($Global:PCCnt Systems currently selected)-----"
$Script:NNC = $null
$Script:Np1 = $null
$Script:Np2 = $null
$Script:Np3 = $null
[int]$NetMenuSelect = 0

Function chNcolor($Script:Np1,$Script:Np2,$Script:Np3,$Script:NNC){
            
            Write-host $Script:Nt4 -NoNewline
            Write-host "$Script:Np1 " -NoNewline
            Write-host $Script:Np2 -ForegroundColor $Script:NNC -NoNewline
            Write-host "$Script:Np3" -ForegroundColor White
            $Script:NNC = $null
            $Script:Np1 = $null
            $Script:Np2 = $null
            $Script:Np3 = $null
}

Function NetMenu()
{
Clear-Host
while ($NetMenuSelect -lt 1 -or $NetMenuSelect -gt 12)
    {
        Trap {"Error: $_"; Break;}        
        $MNNum = 0;Clear-host |out-null
        Start-Sleep -m 50
        if ($Global:PCCnt -eq $null) {$Global:PCCnt = 0}Else {$Global:PCCnt = ($Global:PCList).Count}
        $Script:Nd4 = "------ Use-Net Menu ($Global:PCCnt Systems currently selected)-----"
        Write-host $Script:Nt4 $Script:Nd4 -ForegroundColor Green
        # Exit
        $MNNum ++;$NExit=$MNNum;
            $Script:Np1 = " $MNNum. `t";
            $Script:Np2 = "Exit to Main Menu";
            $Script:Np3 = " ";
            $Script:NNC = "Red";chNcolor $Script:Np1 $Script:Np2 $Script:Np3 $Script:NNC
        # $Get_PCList
        $MNNum ++;$Get_PCList=$MNNum;
            $Script:Np1 =" $MNNum.  `tSelect";
            $Script:Np2 = "systems to use ";
            $Script:Np3 = "with commands on menu ($Global:PCCnt Systems currently selected)";
            $Script:NNC = "Cyan";chNcolor $Script:Np1 $Script:Np2 $Script:Np3 $Script:NNC
        # Test PC(s) for network connectivity
        $MNNum ++;$test_Connections=$MNNum;
            $Script:Np1 =" $MNNum.  `tRun";
            $Script:Np2 = "Test-Connection ";
            $Script:Np3 = "on select PCs or all PCs in Domain";
            $Script:NNC = "yellow";chNcolor $Script:Np1 $Script:Np2 $Script:Np3 $Script:NNC
        # Test Ports
        $MNNum ++;$test_ports=$MNNum;
            $Script:Np1 =" $MNNum.  `tTest";
            $Script:Np2 = "ports ";
            $Script:Np3 = "on a system or list of systems";
            $Script:NNC = "yellow";chNcolor $Script:Np1 $Script:Np2 $Script:Np3 $Script:NNC
        # Status of NIC on PC
        $MNNum ++;$Call_Get_NICStatus=$MNNum;
            $Script:Np1 =" $MNNum.  `tGet ";
            $Script:Np2 = "NIC ";
            $Script:Np3 = "Status and information";
            $Script:NNC = "yellow";chNcolor $Script:Np1 $Script:Np2 $Script:Np3 $Script:NNC
        # View IPConfig-like information on select systems
        $MNNum ++;$Get_IPConfig=$MNNum;
            $Script:Np1 = " $MNNum.  `tView ";
            $Script:Np2 = "IPConfig";
            $Script:Np3 = " information for select systems"
            $Script:NNC = "Yellow";chNcolor $Script:Np1 $Script:Np2 $Script:Np3 $Script:NNC   
        # HostName from IP
        $MNNum ++;$Get_hostname=$MNNum;
            $Script:Np1 =" $MNNum.  `tFind";
            $Script:Np2 = "hostname from IP ";
            $Script:Np3 = "address(es)`n$Script:Nt4$Script:Nt4---";
            $Script:NNC = "yellow";chNcolor $Script:Np1 $Script:Np2 $Script:Np3 $Script:NNC
        # NetStat
        $MNNum ++;$View_NetStat=$MNNum;
            $Script:Np1 =" $MNNum.  `tView";
            $Script:Np2 = "NetStat results ";
            $Script:Np3 = "for a (single) system on the network (Requires elevation)";
            $Script:NNC = "yellow";chNcolor $Script:Np1 $Script:Np2 $Script:Np3 $Script:NNC
        # Ping
        $MNNum ++;$Test_PCUp=$MNNum;
            $Script:Np1 =" $MNNum.  `tTest";
            $Script:Np2 = "(Ping) Network ";
            $Script:Np3 = "Connection Persistence (unable to exit gracefully)";
            $Script:NNC = "yellow";chNcolor $Script:Np1 $Script:Np2 $Script:Np3 $Script:NNC
        # IP from MAC
        $MNNum ++;$Get_IP=$MNNum;
            $Script:Np1 =" $MNNum.`tIdentify ";
            $Script:Np2 = "MAC 2 IP ";
            $Script:Np3 = "- Use a known MAC address to find an IP address";
            $Script:NNC = "yellow";chNcolor $Script:Np1 $Script:Np2 $Script:Np3 $Script:NNC
        # Validate IP
        $MNNum ++;$Test_ValidIPPrompt=$MNNum;
            $Script:Np1 =" $MNNum. `tIdentify";
            $Script:Np2 = "valid IP ";
            $Script:Np3 = "address from input";
            $Script:NNC = "yellow";chNcolor $Script:Np1 $Script:Np2 $Script:Np3 $Script:NNC
        # Wait for live port
        $MNNum ++;$Wait_Live=$MNNum;
            $Script:Np1 = " $MNNum. `tWait for ";
            $Script:Np2 = "System port number";
            $Script:Np3 = " to come alive";
            $Script:NNC = "Yellow";chNcolor $Script:Np1 $Script:Np2 $Script:Np3 $Script:NNC
        # Setting up the menu in this way allows you to move menu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $MNNum ++;$Script:Np1 =" $MNNum.  FIRST";$Script:Np2 = "HIGHLIGHTED ";$Script:Np3 = "END"; $Script:NNC = "yellow";chNcolor $Script:Np1 $Script:Np2 $Script:Np3 $Script:NNC
        Write-Host "$Script:Nt4$Script:Nd4"  -ForegroundColor Green
        $Script:MNNum = 0;[int]$NetMenuSelect = Read-Host "`n`tPlease select the menu item you would like to do now"
    }
    switch($NetMenuSelect)
    {
        # Variables at the beginning of these lines in the switch are numbers assigned by their position in the above
        # configured menu layout. This allows them to be moved around and let the numbers automatically adjust
        $NExit{$NetMenuSelect=$null}
        $Get_PCList{Clear-Host;Get-PCList;reloadnetmenu}
        $test_Connections{Clear-Host;test-connections;reloadnetmenu}
        $test_ports{Clear-Host;test-ports;ReloadNetMenu}
        $Call_Get_NICStatus{Clear-Host;Call-Get-NICStatus;reloadnetmenu}
        $Get_IPConfig{Clear-Host;Get-IPConfig;reloadnetmenu}
        $Get_hostname{Clear-Host;Get-hostname;ReloadNetMenu}
        $View_NetStat{Invoke-Expression .\Get-NetStat.ps1;ReloadNetMenu}   # View NetStat results for a system including the process name and active ports
        $Test_PCUp{Clear-Host;Test-PCUp;ReloadNetMenu}
        $Get_IP{Clear-Host;Get-IP;ReloadNetMenu}
        $Test_ValidIPPrompt{clear-host|Test-ValidIPPrompt;ReloadNetMenu}
        $Wait_Live{Clear-Host;Wait-Live;ReloadNetMenu}
        default
        {
            $NetMenuSelect = $null
            ReloadNetMenu
        }
    }
}
Function ReloadNetMenu()
{
        Read-Host "Hit Enter to continue"
        $NetMenuSelect = $null
        NetMenu
}
# Run Test-Connection on select PCs or all PCs in Domain
Function test-connections()
{
    $PP = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    <# # Commented out because this is not desirable when called from another script
    Else
    {
    Write-Warning "The Current list is: $Global:PCList"
    }
    #>
    $Global:PCCnt = ($Global:PCList).Count
    Write-Warning "Running a PING test on $Global:PCCnt Systems"
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
    # Set Powershell's default $ProgressPreference back to what it started out as using $PP
    $ProgressPreference = $PP
}


# Test ports on a system or list of systems
Function test-ports()
{
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    Else
    {
        Write-Warning "The Current list is: $Global:PCList"
    }
    $IETPRDetails = @()
    $port = Read-Host "`nEnter a port, CSV port list, or list..range "
    if (($protocol = Read-Host "`nEnter UDP or hit [Enter] to use current default: TCP ") -eq "") {$protocol = "TCP"}

    foreach ($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
        Write-Host ". " -nonewline
        $IETPResult = Invoke-Expression ".\Test-Ports.ps1 -computer $Global:pc -port @($port) -$protocol" # |FT
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
    }
    $IETPRDetails|sort pscomputername,Open,Port |ft pscomputername,Port,Open,TypePort,Notes -Wrap -AutoSize
}
# This function basically pings a system to see if it is alive. It logs the data and time and whether it is accessible or not
Function Test-PCUp()
{
    $pcUp= Read-host "`nEnter a system name to test the connection"
        [bool]$hideKeyStrokes = $false #$true
    try
    {
    Write-host "`nPress x to quit" -foregroundcolor yellow
        while(($true) -or ($ky -ne "X"))
        {
            "$($ky = $key.key;Write-Output "$ky key As of ")$(Get-Date)$(Write-Output ", $pcUp ")$(if($(Test-connection -ComputerName $pcUp <#-Quiet#>)){Write-output "is accessible"}else{Write-Output "is inaccessible"})" |ft
            $key = [System.Console]::Readkey($true)
            Console.Read
            "ky is $ky"
        }
    }
    finally
    {
            Write-Host "ky is $ky"
    }
}

# Use a known MAC address to find an IP address
Function Get-IP()
{
    Write-Host "`nYou may need to ping or communicate with a host prior to running this.`nTry pinging the broadcast address first."
    $macaddr = Read-Host "`nEnter some or all of a MAC address in the format AC-86-...."
    arp -a | select-string $macaddr #|% {$_.ToString().Trim().Split(" ")[0]}
    
}

# Find hostname from IP address
# Prompt for IP Address from User. Combine calls Test-ValidIP to test input for valid IP Address
Function Get-hostname()
{
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    Else
    {
    Write-Warning "The Current list is: $Global:PCList"
    }
    foreach($Global:pcline in $Global:pclist)
    {
        if (Test-ValidIP $Global:pcline)
        {
            $GHAddress = [System.Net.DNS]::GetHostByAddress($Global:PCline)
            $GHAddress | select * |ft
        }
    }
        
}
# Identify valid IP address from input
Function Test-ValidIPPrompt # Depends on Test-ValidIP
{
<#
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
#>
try {
        $Global:pcline = Read-host "`nType in a valid IP address to find information about"
        Test-ValidIP $Global:pcline -erroraction "stop"
            
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
Function Test-ValidIP # Called from Test-ValidIPPrompt, Tests input for valid IP Address
{
# Manually this is called from Test-ValidIPPrompt
# This function illustrates parameter validation and using RegEx to validate IPv4 IP addresses including broadcast address
param ([ValidatePattern('^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$')]
[string]$Global:pcline
)
    Write-Output "$Global:pcline is a valid IP address. One moment please."
}
# Wait for  System port number to come alive
Function Wait-Live()
{
    $WaitPC = Read-host "Enter the name of a pc to wait for"
    $port = Read-host "Enter the port number (like 3389)"
    "`n"
    (Invoke-Expression ".\Test-Ports.ps1 -computer $WaitPC -port @($port)")
    while ((Invoke-Expression ".\Test-Ports.ps1 -computer $WaitPC -port @($port)").Open -eq "False")
    {
        Write-host "." -NoNewline
        start-sleep 5
    }
    "`n"
    (Invoke-Expression ".\Test-Ports.ps1 -computer $WaitPC -port @($port)")

    
}
#Get Network information from pc
Function Get-NICStatus() # Called from Call-Get-NICStatus
{
    Get-WmiObject -Class win32_networkadapterconfiguration -ComputerName $Global:pc -Credential $cred|where {$_.IPAddress -ne $null}|select PSComputerName,MACAddress,IPSubnet,DNSServerSearchOrder, @{n="IP Address";e={($_ | select -ExpandProperty IPAddress)-join ','}},DHCPEnabled # |ft #out-grid-view
}
# Get  NIC Status and information
Function Call-Get-NICStatus() # Depends on Get-NICStatus
{
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    Else
    {
    Write-Warning "The Current list is: $Global:PCList"
    }
    Write-host "Checking for NIC status"
    $PCNICStatus = foreach($Global:pcLine in $Global:pclist)
    {
        Identify-PCName
                If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
            Get-NICStatus
        }
    }
    $PCNICStatus |ft
}

# View  IPConfig information for select systems
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
 $Global:pclist (one or more computernames)

.ToDo
    Fix RegEx for matching IP addresses

#>
    $ObCSVFile = $PSScriptRoot+"\"+$Global:CDate+"_IPConfig.csv"
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    Else
    {
    Write-Warning "The Current list is: $Global:PCList"
    }    
    $OutputObj = foreach($Global:pcline in $Global:pclist) {
            # This if statement could potentially be moved to inside the Get-PCList function, but it might need to be modified to do so
            if ($Global:PCline.name -eq $null)
            {
                $Global:PC = $Global:PCline
            }
            else
            {
                $Global:PC = $Global:PCline.name
            }
      if(Test-Connection -ComputerName $Global:PC -Count 1 -ea 0) {
       try {
        $Networks = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Global:PC -EA Stop | ? {$_.IPEnabled}
       } catch {
            Write-Warning "Error occurred while querying $Global:PC."
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
        $OBProperties = @{'computername'=$Global:PC.ToUpper();
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
    Write-Warning "Sending results to a grid and: $ObCSVFile"
    $OutputObj|Select computername, SubnetMask, DNSServers, Gateway, IPAddress, IsDHCPEnabled, MACAddress, WINSServers, Error |Out-GridView
    $OutputObj|Select computername, SubnetMask, DNSServers, Gateway, IPAddress, IsDHCPEnabled, MACAddress, WINSServers, Error |Export-Csv $ObCSVFile -NoClobber -NoTypeInformation -Append
}
