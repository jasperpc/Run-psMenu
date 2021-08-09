<#
.NAME
    Run-PSMenu.ps1
.SYNOPSIS
    Provide menu system for running Powershell scripts
.DESCRIPTION
 Menu design initially adopted from : https://quickclix.wordpress.com/2012/08/14/making-powershell-menus/
 2017-03-03

 Written by Jason Crockett - Laclede Electric Cooperative

 use Ctrl+J to bring up "snippets" that can be inserted into code
 Tips from MVA
    - Scripting ToolMaking
    - www.powershell.odmarc
    rg
    - #PowerShell on Twitter & @JSnover
.PARAMETERS
    This parameter doesn't exist (Comps = list of computernames)
.EXAMPLE
    "a command with a parameter would look like: (.\Run-PSMenu.ps1 -comps dc)"
        .\Run-PSMenu.ps1 []
.SYNTAX
    .\Run-PSMenu.ps1 []
.REMARKS
    To see the examples, type: help Run-PSMenu.ps1 -examples
    To see more information, type: help Run-PSMenu.ps1 -detailed"
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

# Clear-Host
$PSVersionTable.PSVersion
# set Global and Script variables
. .\Get-SystemsList.ps1
[BOOLEAN]$Global:ExitSession=$false

# Enable if needed
# $Use_MMFPath = Split-Path -Parent $PSCommandPath
$Script:t4 = "`t`t`t`t"
$Script:NC = $null
$Script:p1 = $null
$Script:p2 = $null
$Script:p3 = $null
$Global:CDate = Get-Date -UFormat "%Y%m%d"
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
while ($menuselect -lt 1 -or $menuselect -gt 12)
{
    Clear-Host
        Trap {"Error: $_"; Break;}        
        $MNum = 0;Clear-host |out-null
        Start-Sleep -m 50
        if ($Global:PCCnt -eq $null) 
        {$Global:PCCnt = 0;$Script:d4 = "-------------------------------- Run-PSMenu --------------------------------"}
        Else 
        {$Global:PCCnt = ($Global:PCList).Count;$Script:d4 = "---------------- Run-PSMenu ($Global:PCCnt Systems currently selected) ----------------"}
        Write-host $t4 $d4 -ForegroundColor Green
#    Exit from the Main Menu
        $MNum ++;
            $Exit=$MNum;
            $p1 = " $MNum. `t";
            $p2 = "Exit to Main Menu";
            $p3 = " ";
            $NC = "Red";chcolor $p1 $p2 $p3 $NC
#    Prompt for credential with permissions across the network
            $MNum ++;$Get_Cred=$MNum;
            $p1 =" $MNum.  `t Enter";
            $p2 = "Credentials ";
            $p3 = "for use in script."; $NC = "Cyan";chcolor $p1 $p2 $p3 $NC
#    List Systems currently in the $Global:PCList variable
            $MNum ++;$List_PCList=$MNum;
            $p1 =" $MNum.  `t List";
            $p2 = "Global PCList variable ";
            $p3 = "contents."; $NC = "Cyan";chcolor $p1 $p2 $p3 $NC
#    Net Utilities Menu
            # Go to the Net Utilities Menu
        $MNum ++;$Use_MenuNet=$MNum;
            $p1 = " $MNum. `t Go to";
            $p2 = "Net Utilities";
            $p3 = " menu"
            $NC = "DarkCyan";chcolor $p1 $p2 $p3 $NC
#    Users and Groups Menu
        $MNum ++;$Use_MenuUserNGroup=$MNum;
            $p1 = " $MNum. `t Go to the";
            $p2 = "manage Users and groups";
            $p3 = " menu"
            $NC = "DarkCyan";chcolor $p1 $p2 $p3 $NC
# PC-Related Menu"
        $MNum ++;$Use_MenuPC=$MNum; #Previously $Get_pcinfo
            $p1 = " $MNum. `t Go to ";
            $p2 = "PC";
            $p3 = " Menu"
            $NC = "DarkCyan";chcolor $p1 $p2 $p3 $NC
# File-Related Menu"
        $MNum ++;$Use_MenuFile=$MNum;
            $p1 = " $MNum. `t Go to ";
            $p2 = "File-Related";
            $p3 = " Menu"
            $NC = "DarkCyan";chcolor $p1 $p2 $p3 $NC
# Printer-Related Menu"
        $MNum ++;$Use_MenuPrint=$MNum;
            $p1 = " $MNum. `t Go to ";
            $p2 = "Printer-Related";
            $p3 = " Menu"
            $NC = "DarkCyan";chcolor $p1 $p2 $p3 $NC
# Utility Menu"
        $MNum ++;$Use_MenuUtility=$MNum;
            $p1 = " $MNum. `t Go to ";
            $p2 = "Utility";
            $p3 = " Menu"
            $NC = "DarkCyan";chcolor $p1 $p2 $p3 $NC
# "Run `"update-help -Force`" to update PowerShell Help"
        $MNum ++;$update_help=$MNum;
            $p1 = " $MNum. `t -";
            $p2 = "Update PowerShell Help ";
            $p3 = "information"
            $NC = "Magenta";chcolor $p1 $p2 $p3 $NC
# Create a new Remote PowerShell connection"
        $MNum ++;$New_PSRemote=$MNum;
            $p1 = " $MNum. `t Create a new ";
            $p2 = "Remote PowerShell connection ";
            $p3 = "(disconnect from this menu)"
            $NC = "Magenta";chcolor $p1 $p2 $p3 $NC
# This prompts to run a powershell script on selected remote system(s)
        $MNum ++;$Run_RemotePS=$MNum;
            $p1 = " $MNum. `t This prompts to ";
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
        $Get_Cred{Clear-Host;Get-Cred;$menuselect = $null;menu}
        $List_PCList{Clear-Host;"PCName";foreach ($Global:PCLIne in $Global:PCList){Identify-PCName;$Global:PC};reload-PromptMenu}
        $Use_MenuNet{. .\Use-MenuNet.ps1;NetMenu;reload-NoPromptMenu}
        $Use_MenuUserNGroup{. .\Use-MenuUserNGroup.ps1;UnGmenu;reload-NoPromptMenu}
        $Use_MenuPC{. .\Use-MenuPC.ps1;PCMenu;reload-NoPromptMenu}
        $Use_MenuFile{. .\Use-MenuFile.ps1;FileMenu;reload-NoPromptMenu}
        $Use_MenuPrint{. .\Use-MenuPrint.ps1;PrintMenu;reload-NoPromptMenu}
        $Use_MenuUtility{. .\Use-MenuUtility.ps1;UtilityMenu;reload-NoPromptMenu}
        $update_help{clear-host;write-Warning "Running `"update-help -Force`" to update PowerShell Help";update-help -Force;reload-PromptMenu}
        $New_PSRemote{New-PSRemote;reload-PromptMenu}
        $Run_RemotePS{$p1=$null;$p2=$null;$p3=$null;Run-RemotePS;reload-PromptMenu}
        default
        {
            cls
            $MenuSelect = $null
            $global:NC = " "
            $global:p1 = " "
            $global:p2 = " "
            $global:p3 = " "
            $Script:MNum = 0
            $Global:ExitSession=$true;break
        }

    }    
}
Function reload-NoPromptMenu()
{
        $menuselect = $null
        menu
}
Function reload-PromptMenu()
{
        Read-Host "Hit Enter to continue"
        # Start-Sleep -m 250
        $menuselect = $null
        menu
}
Function Get-Cred()
{
    # Get Credentials for later wmi requests
    # Prompt to decide to use the script's $Cred credential or another
    $CredUser = $Cred.UserName
    if ((Read-Host "Current credentials are for: `"$CredUser`" , type C to change or [Enter] to continue")-eq "C") 
        {$Global:cred = Get-Credential "$env:USERDOMAIN\Administrator"}
}
Function New-PSRemote()
{
    Clear-Host
    # Connect to Server Remotely
    # Still needs some work for correct authorization / force credentials again to connect to remote machines
        # Prompt to decide to use the script's $Cred credential or another
        Get-Cred
        # The following two lines are now bundled into the Get-Cred function
        # $CredUser = $Cred.UserName
        # if ((Read-Host "Current credentials are for: `"$CredUser`" , type C to change or [Enter] to continue")-eq "C") {Get-Cred}
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
Function Run-RemotePS
{
    Clear-Host
    $SBCmd = "hostname;"
    Write-Warning $SBCmd
    $SBCmd = Read-Host "Enter a Powershell command here"
    Write-Warning $SBCmd
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    foreach ($Global:pcLine in $Global:PCList)
    {
        $ICResult = ""
        Identify-PCName
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
            $s = New-PSSession -ComputerName $Global:PC -Credential $Global:Cred
            #Working but Commented 2020/10/28 for testing purposes: Invoke-Command -ScriptBlock {param($p1) powershell $p1} -Session $s -ArgumentList $SBCmd
            $ICResult = Invoke-Command -ScriptBlock {param($p1) powershell $p1} -Session $s -ArgumentList $SBCmd -ErrorAction SilentlyContinue
            Write-output "Written output: $ICResult"
            $s |Remove-PSSession
            # Invoke-Command -ComputerName $Global:PC -ScriptBlock {param($p1) powershell $p1} -Credential $Cred -ArgumentList $SBCmd
        }
    }
}

menu
