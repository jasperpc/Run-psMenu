<#
.NAME
    Use-MenuPrint.ps1
.SYNOPSIS
   Use Printer-related functions
.DESCRIPTION
    PrintMenu design patterned after: https://quickclix.wordpress.com/2012/08/14/making-powershell-PrintMenus/
    2020-11-13
    Written by Jason Crockett - Laclede Electric Cooperative
.PARAMETERS
    No command-line parameters needed
.EXAMPLE
    .\Use-MenuPrint.ps1 []
.SYNTAX
    .\Use-MenuPrint.ps1 []
.REMARKS

.TODO
    Modify Get-PrintInfo function to list the name of the computer along with the discovered information
#>
# set script variables
[BOOLEAN]$Global:ExitSession=$false
$script:Prt4 = "`t`t`t`t"
$Script:Prd4 = "------ Use-Print Menu ($Global:PCCnt Systems currently selected)-----"
$script:PrNC = $null
$Script:PrP1 = $null
$Script:PrP2 = $null
$Script:PrP3 = $null
$Script:PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
[int]$PrintMenuselect = 0
function chPRcolor($Script:PrP1,$Script:PrP2,$Script:PrP3,$NC)
{
    Write-host $script:Prt4 -NoNewline
    Write-host "$Script:PrP1 " -NoNewline
    Write-host "$Script:PrP2" -ForegroundColor $NC -NoNewline
    Write-host "$Script:PrP3" -ForegroundColor White
    $script:PrNC = $null
    $Script:PrP1 = $null
    $Script:PrP2 = $null
    $Script:PrP3 = $null
}
function PrintMenu()
{
while ($PrintMenuSelect -lt 1 -or $PrintMenuSelect -gt 6)
    {
        Clear-Host
        Trap {"Error: $_"; Break;}        
        $Script:PRMNum = 0;Clear-host |out-null
        Start-Sleep -m 250
        if ($Global:PCCnt -eq $null) {$Global:PCCnt = 0}Else {$Global:PCCnt = ($Global:PCList).Count}
        $Script:Prd4 = "------ Use-Print Menu ($Global:PCCnt Systems currently selected)-----"
        Write-host $script:Prt4 $script:Prd4 -ForegroundColor Green
# 1. Exit Print Menu" -ForegroundColor Red
        # Exit
        $Script:PRMNum ++;$PRExit=$Script:PRMNum;
            $Script:PrP1 = " $Script:PRMNum. `t";
            $Script:PrP2 = "Exit Main PrintMenu";
            $Script:PrP3 = " ";
            $NC = "Red";chPRcolor $Script:PrP1 $Script:PrP2 $Script:PrP3 $NC
# Get PCList
        $Script:PRMNum ++;$Get_PrPCList=$Script:PRMNum;$Script:PrP1 =" $PrMNum.  `t Select";
            $Script:PrP2 = "systems to use ";$Script:PrP3 = "with commands on menu ($Global:PCCnt Systems currently selected)";
            $NC = "Cyan";chPRcolor $Script:PrP1 $Script:PrP2 $Script:PrP3 $NC
# Printer, Port, and Driver Info
        $Script:PRMNum ++;$Get_PrintInfo=$Script:PRMNum;
            $Script:PrP1 = " $Script:PRMNum. `t View ";
            $Script:PrP2 = "Printer, Port, and Driver ";
            $Script:PrP3 = "Information"
            $NC = "Yellow";chPRcolor $Script:PrP1 $Script:PrP2 $Script:PrP3 $NC
# Printer Information
        $Script:PRMNum ++;$Get_PrinterInfo=$Script:PRMNum;
            $Script:PrP1 = " $Script:PRMNum. `t View ";
            $Script:PrP2 = "Printer ";
            $Script:PrP3 = "Information"
            $NC = "Yellow";chPRcolor $Script:PrP1 $Script:PrP2 $Script:PrP3 $NC
# Printer Port Information
        $Script:PRMNum ++;$Get_PrinterPortInfo=$Script:PRMNum;
            $Script:PrP1 = " $Script:PRMNum. `t View ";
            $Script:PrP2 = "Printer Port ";
            $Script:PrP3 = "Information"
            $NC = "Yellow";chPRcolor $Script:PrP1 $Script:PrP2 $Script:PrP3 $NC
# Printer Driver Information
        $Script:PRMNum ++;$Get_PrinterDriverInfo=$Script:PRMNum;
            $Script:PrP1 = " $Script:PRMNum. `t View ";
            $Script:PrP2 = "Printer Driver ";
            $Script:PrP3 = "Information"
            $NC = "Yellow";chPRcolor $Script:PrP1 $Script:PrP2 $Script:PrP3 $NC
                       
        # Setting up the PrintMenu in this way allows you to move PrintMenu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $Script:PRMNum ++;$Script:PrP1 =" $Script:PRMNum.  FIRST";$Script:PrP2 = "HIGHLIGHTED ";$Script:PrP3 = "END"; $NC = "yellow";chPRcolor $Script:PrP1 $script:p2 $Script:PrP3 $NC
        Write-Host "$script:Prt4$script:Prd4"  -ForegroundColor Green
        $Script:PRMNum = 0;[int]$PrintMenuSelect = Read-Host "`n`tPlease select the PrintMenu item you would like to do now"
    }
    Function Nullify-PrintPrompts
    {
        $Script:PI = $null
        $Script:PPI = $null
        $Script:PDI = $null
    }
    Function Confirm-PrPrompts
    {
        $Script:PI = "Yes"
        $Script:PPI = "Yes"
        $Script:PDI = "Yes"
    }
    $PICMD = ""
    $PPICMD = ""
    $PDICMD = ""
    switch($PrintMenuSelect)
    {
        # Variables at the beginning of these lines in the switch are numbers assigned by their position in the above
        # configured PrintMenu layout. This allows them to be moved around and let the numbers automatically adjust
        $PRExit{$PrintMenuSelect=$null}
        $Get_PrPCList{Clear-Host;Confirm-PrPrompts;Get-PCList;reloadPrintMenu}
        $Get_PrintInfo{Clear-Host;Confirm-PrPrompts;Get-PrintInfo $Script:PI $Script:PPI $Script:PDI;reloadPrintMenu}
        $Get_PrinterInfo{Clear-Host;Write-Output $PICMD;$Script:PI = "Yes";Get-PrinterInfo;ReloadPrintMenu}
        $Get_PrinterPortInfo{Clear-Host;Write-Output $PPICMD;$Script:PPI = "Yes";Get-PrinterPortInfo;ReloadPrintMenu}
        $Get_PrinterDriverInfo{Clear-Host;Write-Output $PDICMD;$Script:PDI = "Yes";Get-PrinterDriverInfo;ReloadPrintMenu}
        default
        {
            $PrintMenuSelect = $null
            $global:NC = " "
            $global:p1 = " "
            $global:p2 = " "
            $global:p3 = " "
            $Script:PRMNum = 0
            ReloadPrintMenu
        }
    }
}

function ReloadPrintMenu()
{
        Read-Host "Hit Enter to continue"
        $PrintMenuSelect = $null
        $Script:PRMNum = 0
        PrintMenu
}

Function Get-PrintInfo($Script:PI, $Script:PPI, $Script:PDI)
{
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    Else
    {
    Write-Warning "Getting printer information from the following list : $Global:PCList"
    }
    foreach($Global:pcLine in $Global:PCList)
    {
        Identify-PCName # Called from Run-PSMenu.ps1
        if ($Script:PI -eq "Yes"){Get-PrinterInfo}
        if ($Script:PPI -eq "Yes"){Get-PrinterPortInfo}
        if ($Script:PDI -eq "Yes"){Get-PrinterDriverInfo}
    }
    Nullify-PrintPrompts
}
# Get and display printer information for listed system
Function Get-PrinterInfo
{
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    Else
    {
    Write-Warning "Getting printer information from the following list : $Global:PCList"
    }
    $Global:PCCnt = ($Global:PCList).Count
    foreach($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        # Only attempt to get information from computers that are reachable with a ping
        If (Test-Connection $Global:pc -Count 1 -Quiet -ErrorAction SilentlyContinue)
        {
            $Script:PrP1 = "`nRunning: Get-WmiObject -Class win32_printer -Credential $cred -ComputerName"
            $Script:PrP2 = "$Global:PC"
            $Script:PrP3 = " | select name, drivername, portname, shared, sharename |ft -AutoSize"
            $NC = "Yellow";
            chPRcolor $Script:PrP1 $Script:PrP2 $Script:PrP3 $NC
            $prin = Get-WmiObject -Class win32_printer -Credential $cred -ComputerName $Global:PC | select name, drivername, portname, shared, sharename |ft -AutoSize
            $prin
            # Testing another Printer Info command
            # Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Network=True" | select Sharename,name | export-csv -path .\$env:computername.csv -NoTypeInformation

       }
    }
}
# Get and display Printer Port information for listed system
Function Get-PrinterPortInfo
{
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    Else
    {
    Write-Warning "Getting printer information from the following list : $Global:PCList"
    }
    $Global:PCCnt = ($Global:PCList).Count
    foreach($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        # Only attempt to get information from computers that are reachable with a ping
        If (Test-Connection $Global:pc -Count 1 -Quiet -ErrorAction SilentlyContinue)
        {
    $Script:PrP1 = "`nRunning: get-printerport -computer "
    $Script:PrP2 = "$Global:PC"
    $Script:PrP3 = " |select ComputerName,Description,PrinterHostAddress,Name |ft -Wrap -AutoSize"
    $NC = "Yellow";
    chPRcolor $Script:PrP1 $Script:PrP2 $Script:PrP3 $NC
    $prin2 = Get-PrinterPort -ComputerName $Global:PC |select ComputerName,Description,PrinterHostAddress,Name |ft -Wrap -AutoSize
    $prin2
        }
    }
}

# Get and display Printer Driver information for listed system
Function Get-PrinterDriverInfo
{
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    Else
    {
    Write-Warning "Getting printer information from the following list : $Global:PCList"
    }
    $Global:PCCnt = ($Global:PCList).Count
    foreach($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        # Only attempt to get information from computers that are reachable with a ping
        If (Test-Connection $Global:pc -Count 1 -Quiet -ErrorAction SilentlyContinue)
        {
    $Script:PrP1 = "`nRunning: get-printerDriver -computer"
    $Script:PrP2 = "$Global:PC"
    $Script:PrP3 = " |select * |ft -wrap -Autosize"
    $NC = "Yellow";
    chPRcolor $Script:PrP1 $Script:PrP2 $Script:PrP3 $NC
    $prin3 = get-printerDriver -computer $Global:PC |select ComputerName,Name|ft -wrap -Autosize
    $prin3
        }
    }
}