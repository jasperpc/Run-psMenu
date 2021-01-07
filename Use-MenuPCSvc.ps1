<#
.NAME
    Use-MenuPCSvc.ps1
.SYNOPSIS
   Use PC Services functions
.DESCRIPTION
    Menu design patterned after: https://quickclix.wordpress.com/2012/08/14/making-powershell-menus/
    2020-11-13
    Written by Jason Crockett - Laclede Electric Cooperative
.PARAMETERS
    No command-line parameters needed
.EXAMPLE
    .\Use-MenuPCSvc.ps1 []
.SYNTAX
    .\Use-MenuPCSvc.ps1 []
.REMARKS

.TODO

#>
# set script variables
# [BOOLEAN]$Global:ExitSession=$false
$Script:PCSt4 = "`t`t`t`t"
# $Script:PCSd4 = "------------Use-PC Svc Menu ($Global:PCCnt Systems currently selected)----------------"
$Script:PCSNC = $null
$Script:PCS1 = $null
$Script:PCS2 = $null
$Script:PCS3 = $null
$Script:PCSSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
[int]$PCSvcMenuSelect = 0
function chPCScolor($Script:PCS1,$Script:PCS2,$Script:PCS3,$Script:PCSNC)
{
    Write-host $Script:PCSt4 -NoNewline
    Write-host "$Script:PCS1 " -NoNewline
    Write-host "$Script:PCS2" -ForegroundColor $Script:PCSNC -NoNewline
    Write-host "$Script:PCS3" -ForegroundColor White
    $Script:PCSNC = $null
    $Script:PCS1 = $null
    $Script:PCS2 = $null
    $Script:PCS3 = $null
}
function PCSvcMenu()
{
while ($PCSvcMenuSelect -lt 1 -or $PCSvcMenuSelect -gt 8)
    {
        Clear-Host
        Trap {"Error: $_"; Break;}        
        $Script:PCSMNum = 0;Clear-host |out-null
        Start-Sleep -m 250
        if ($Global:PCCnt -eq $null) {$Global:PCCnt = 0}Else {$Global:PCCnt = ($Global:PCList).Count}
        $Script:PCSd4 = "------------Use-PC Svc Menu ($Global:PCCnt Systems currently selected)----------------"
        Write-host $Script:PCSt4 $Script:PCSd4 -ForegroundColor Green
#    Write-Host "$Script:PCSt4  1. Exit Now" -ForegroundColor Red
        # Exit
        $Script:PCSMNum ++;$AExit=$Script:PCSMNum;
            $Script:PCS1 = " $Script:PCSMNum. `t";
            $Script:PCS2 = "Exit Main PCSvcMenu";
            $Script:PCS3 = " ";
            $Script:PCSNC = "Red";chPCScolor $Script:PCS1 $Script:PCS2 $Script:PCS3 $Script:PCSNC
# Prompt for credential with permissions across the network
        $Script:PCSMNum ++;$GCred=$Script:PCSMNum;$Script:PCS1 =" $Script:PCSMNum.  `t Enter";
            $Script:PCS2 = "Credentials ";$Script:PCS3 = "for use in script."; $Script:PCSNC = "Cyan";chPCScolor $Script:PCS1 $Script:PCS2 $Script:PCS3 $Script:PCSNC
# $Get_PCList
        $Script:PCSMNum ++;$Get_PCList=$Script:PCSMNum;$Script:PCS1 =" $Script:PCSMNum.  `t Select";
            $Script:PCS2 = "systems to use ";$Script:PCS3 = "with commands on menu ($Global:PCCnt Systems currently selected).";
            $Script:PCSNC = "Cyan";chPCScolor $Script:PCS1 $Script:PCS2 $Script:PCS3 $Script:PCSNC
# Get services using specific accounts to run
        $Script:PCSMNUM ++;$Get_ServiceLogons=$Script:PCSMNUM;
            $Script:PCSp1 =" $Script:PCSMNUM. `t View ";
            $p2 = "RunAs accounts ";
            $p3 = "for Services";
            $NC = "yellow";chPCScolor $Script:PCSp1 $p2 $p3 $NC
# Firewall Status
        $Script:PCSMNUM ++;$FW_Status=$Script:PCSMNUM;
            $Script:PCSp1 =" $Script:PCSMNUM. `t View ";
            $p2 = "Firewall Status ";
            $p3 = "on select PCs or all PCs in Domain";
            $NC = "yellow";chPCScolor $Script:PCSp1 $p2 $p3 $NC
# Set NIC location
        $Script:PCSMNUM ++;$Get_Set_ncp=$Script:PCSMNUM;
            $Script:PCSp1 = " $Script:PCSMNUM. `t Get ";
            $p2 = "Network Location ";
            $p3 = " and change it (with confirm) of a local NIC ";
            $NC = "Yellow";chPCScolor $Script:PCSp1 $p2 $p3 $NC
# servce restart
        $Script:PCSMNUM ++;$Restart_SelectSvc=$Script:PCSMNUM;
            $Script:PCSp1 = " $Script:PCSMNUM. `t Restart ";
            $p2 = "Service";
            $p3 = " on selected system";
            $NC = "Yellow";chPCScolor $Script:PCSp1 $p2 $p3 $NC
# servce query
        $Script:PCSMNUM ++;$Query_SelectSvc=$Script:PCSMNUM;
            $Script:PCSp1 = " $Script:PCSMNUM. `t Query ";
            $p2 = "Service";
            $p3 = " on selected system";
            $NC = "Yellow";chPCScolor $Script:PCSp1 $p2 $p3 $NC
        # Setting up the menu in this way allows you to move menu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $PCIMNum ++;$Script:PCI1 =" $PCIMNum.  FIRST";$Script:PCI2 = "HIGHLIGHTED ";$Script:PCI3 = "END"; $Script:PCINC = "yellow";chPCIcolor $Script:PCI1 $Script:PCI2 $Script:PCI3 $Script:PCINC
        Write-Host "$Script:PCIt4 $Script:PCId4"  -ForegroundColor Green
        $Script:MNum = 0;[int]$PCSvcMenuSelect = Read-Host "`n`tPlease select the menu item you would like to do now"
    }

switch($PCSvcMenuSelect)
    {
        $PCMExit{$PCIMenuselect=$null;reload-PCmenu}
        $GCred{Clear-Host;Get-Cred;$PCIMenuselect = $null;Reload-PCSvcMenu} # Called from Run-PSMenu.ps1
        $Get_PCList{Clear-Host;Get-PCList;Reload-PCSvcMenu}
        $Get_ServiceLogons{Clear-Host;Get-ServiceLogons;Reload-PromptPCSvcMenu}
        $FW_Status{Clear-Host;FW-Status;Reload-PCSvcMenu}
        $Get_Set_ncp{Clear-Host;Get-Set-ncp;Reload-PCSvcMenu}
        $Restart_SelectSvc{Clear-Host;Restart-SelectSvc;Reload-PromptPCSvcMenu}
        $Query_SelectSvc{Clear-Host;Query-SelectSvc;Reload-PromptPCSvcMenu}
default
        {
            # cls
            $PCSvcMenuSelect = $null
            Reload-PCSvcMenu #clear-host;break;
        }
    }    
}

Function reload-PromptPCSvcMenu() # Reload with prompt (not needed until there's another menu down)
{
    Read-Host "Hit Enter to continue"
    $PCSvcMenuSelect = $null
    PCSvcMenu
}
Function Reload-PCSvcMenu() # Reload with NO prompt
{
        $PCSvcMenuSelect = $null
        PCSvcMenu
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
Function Get-ServiceLogons
{
    $uname = Read-Host "Enter all or part of a valid possible Username that a service may use to start"
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    foreach ($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        Get-WmiObject -Class Win32_Service -ComputerName $Global:pc |Sort-Object StartName,PSComputerName |Where-Object {$_.startName -like "*$uname*"}|ft StartName,pscomputername,Name,State,Displayname -AutoSize
    }    
}
Function Get-Set-ncp
{
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
Function FW-Status()
{
    # If we cannot ping PC could be off or Firewall could be on
    # If we can ping but cannot get profile status RemoteRegistry Service could be off
    # Need to add a test for service being disabled and if it is disabled, a prompt to enable it and set its type to auto or demand, etc
    $RegErr = $false
    if (($FWprofile = Read-Host "Type firewall profile name to query or [Enter] to use default: DomainProfile") -eq "") {$FWprofile = "DomainProfile"}
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
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
Function Restart-SelectSvc()
{
$SvcPrompt = $null
if (($ServiceNamePrompt = Read-Host "Enter a service name or [Enter] to use default: NLASvc") -eq "") {$ServiceNamePrompt = "NLASvc"}
<#
In Windows 10, the nlasvc service must sometimes be restarted to force the network location to identify itself correctly
When it does not know where it is, Outlook cannot connect to the exchange server even if the pc has an IP address and can get to the internet.
The firewall on the default public location is much more restrictive
#>
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    $Global:PCCnt = ($Global:PCList).Count
    if ($Global:PCCnt -gt 1)
    {
        $SvcPrompt =Read-Host "You have $Global:PCCnt systems selected. Type Y to restart the $ServiceNamePrompt on all systems"
    }
    else{$SvcPrompt = "Y"}
    if ($SvcPrompt -eq "Y")
    { 
        $SvcObjResult = foreach($Global:pcLine in $Global:PCList)
        {
            Identify-PCName # Called from Run-PSMenu.ps1
            If (Test-Connection $Global:pc -Count 1 -Quiet)
            {
                $getSvc = Get-Service $ServiceNamePrompt -ComputerName $Global:PC # |Select MachineName,Name,Status,StartType 
                # $getSvc |FT
                If ($getSvc.Status -eq "Stopped")
                {
                    $getSvcPrompt = Read-Host "Would you like to restart the $ServiceNamePrompt service? (Y if Yes)"
                    if ($getSvcPrompt -eq "Y")
                    {
                        Write-Host "`nAttempting to restart the $ServiceNamePrompt service" -ForegroundColor Green
                        $getSvc |Restart-Service -Force
                        $getSvc = Get-Service $ServiceNamePrompt -ComputerName $Global:PC |Select MachineName,Name,Status,StartType 
                    }
                }
                New-Object PSObject -Property @{
                    pscomputername = $Global:PC;
                    MachineName = $getSvc.MachineName;
                    Name = $getSvc.Name;
                    Status = $getSvc.Status;
                    StartType  = $getSvc.StartType
                    }
            }
            $SvcObjResult += $SvcObjResult
        }
    }
    else
    {Write-Warning "You may use the menu to select only a single system if you wish"}
    $SvcObjResult |FT
}

Query-SelectSvc
Function Query-SelectSvc()
{
$SvcPrompt = $null
if (($ServiceNamePrompt = Read-Host "Enter a service name or [Enter] to use default: NLASvc") -eq "") {$ServiceNamePrompt = "NLASvc"}
<#
In Windows 10, the nlasvc service must sometimes be restarted to force the network location to identify itself correctly
When it does not know where it is, Outlook cannot connect to the exchange server even if the pc has an IP address and can get to the internet.
The firewall on the default public location is much more restrictive
#>
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    $Global:PCCnt = ($Global:PCList).Count
    if ($Global:PCCnt -gt 1)
    {
        $SvcPrompt =Read-Host "You have $Global:PCCnt systems selected. Type Y to view the NLASvc on all systems"
    }
    else{$SvcPrompt = "Y"}
    if ($SvcPrompt -eq "Y")
    { 
        $SvcObjResult = foreach($Global:pcLine in $Global:PCList)
        {
            Identify-PCName # Called from Run-PSMenu.ps1
            If (Test-Connection $Global:pc -Count 1 -Quiet)
            {
                $getSvc = Get-Service $ServiceNamePrompt -ComputerName $Global:PC |Select MachineName,Name,Status,StartType 
                # $getSvc |FT
                <#
                If ($getSvc.Status -eq "Stopped")
                {
                    $getSvcPrompt = Read-Host "Would you like to restart the service: $ServiceNamePrompt? (Y if Yes)"
                    if ($getSvcPrompt -eq "Y")
                    {
                        Write-Host "`nAttempting to restart the Network Location Awareness service (nlasvc)" -ForegroundColor Green
                        Restart-Service $ServiceNamePrompt -Force
                        $getSvc = Get-Service $ServiceNamePrompt -ComputerName $Global:PC |Select MachineName,Name,Status,StartType 
                    }
                }
                #>
                New-Object PSObject -Property @{
                    pscomputername = $Global:PC;
                    MachineName = $getSvc.MachineName;
                    Name = $getSvc.Name;
                    Status = $getSvc.Status;
                    StartType  = $getSvc.StartType
                    }
            }
            $SvcObjResult += $SvcObjResult
        }
    }
    else
    {Write-Warning "You may use the menu to select only a single system if you wish"}
    $SvcObjResult |FT
}
