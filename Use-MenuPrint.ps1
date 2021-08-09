<#
.NAME
    Use-MenuPrint.ps1
.SYNOPSIS
   Use Printer-related functions
.DESCRIPTION
    PrintMenu design patterned after: https://quickclix.wordpress.com/2012/08/14/making-powershell-PrintMenus/
    2020-11-17
    Written by Jason Crockett - Laclede Electric Cooperative
.PARAMETERS
    No command-line parameters needed
.EXAMPLE
    .\Use-MenuPrint.ps1 []
.SYNTAX
    .\Use-MenuPrint.ps1 []
.REMARKS

.TODO
    Modify Get-PrintInfo function to list the name of the computer along with the discovered information and cleanup output
#>
# set script variables
[BOOLEAN]$Global:ExitSession=$false
$Global:Prt4 = "`t`t`t`t"
$Global:Prd4 = "------ Use-Print Menu ($Global:PCCnt Systems currently selected)-----"
$Global:PrNC = $null
$Global:PrP1 = $null
$Global:PrP2 = $null
$Global:PrP3 = $null
$Global:PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
[int]$PrintMenuselect = 0
function chPRcolor($Global:PrP1,$Global:PrP2,$Global:PrP3,$global:PrNC)
{
    Write-host $Global:Prt4 -NoNewline
    Write-host "$Global:PrP1 " -NoNewline
    Write-host "$Global:PrP2" -ForegroundColor $global:PrNC -NoNewline
    Write-host "$Global:PrP3" -ForegroundColor White
    $Global:PrNC = $null
    $Global:PrP1 = $null
    $Global:PrP2 = $null
    $Global:PrP3 = $null
}
function PrintMenu()
{
while ($PrintMenuSelect -lt 1 -or $PrintMenuSelect -gt 7)
    {
        Clear-Host
        Trap {"Error: $_"; Break;}        
        $Global:PRMNum = 0;Clear-host |out-null
        Start-Sleep -m 250
        if ($Global:PCCnt -eq $null) {$Global:PCCnt = 0}Else {$Global:PCCnt = ($Global:PCList).Count}
        $Global:Prd4 = "------ Use-Print Menu ($Global:PCCnt Systems currently selected)-----"
        Write-host $Global:Prt4 $Global:Prd4 -ForegroundColor Green
    # Exit Print Menu
        $Global:PRMNum ++;$PRExit=$Global:PRMNum;
            $Global:PrP1 = " $Global:PRMNum. `t";
            $Global:PrP2 = "Exit Main PrintMenu";
            $Global:PrP3 = " ";
            $global:PrNC = "Red";chPRcolor $Global:PrP1 $Global:PrP2 $Global:PrP3 $global:PrNC
    # Get a globally available list of systems
        $Global:PRMNum ++;$Get_PrPCList=$Global:PRMNum;$Global:PrP1 =" $PrMNum.  `t Select";
            $Global:PrP2 = "systems to use ";$Global:PrP3 = "with commands on menu ($Global:PCCnt Systems currently selected)";
            $global:PrNC = "Cyan";chPRcolor $Global:PrP1 $Global:PrP2 $Global:PrP3 $global:PrNC
    # Printer, Port, and Driver Info
        $Global:PRMNum ++;$Get_PrintInfo=$Global:PRMNum;
            $Global:PrP1 = " $Global:PRMNum. `t View ";
            $Global:PrP2 = "Printer, Port, and Driver ";
            $Global:PrP3 = "Information"
            $global:PrNC = "Yellow";chPRcolor $Global:PrP1 $Global:PrP2 $Global:PrP3 $global:PrNC
    # Printer Information (name, drivername, portname, shared, sharename)
        $Global:PRMNum ++;$Get_PrinterInfo=$Global:PRMNum;
            $Global:PrP1 = " $Global:PRMNum. `t View ";
            $Global:PrP2 = "Printer ";
            $Global:PrP3 = "Information"
            $global:PrNC = "Yellow";chPRcolor $Global:PrP1 $Global:PrP2 $Global:PrP3 $global:PrNC
    # Printer Port Information
        $Global:PRMNum ++;$Get_PrinterPortInfo=$Global:PRMNum;
            $Global:PrP1 = " $Global:PRMNum. `t View ";
            $Global:PrP2 = "Printer Port ";
            $Global:PrP3 = "Information"
            $global:PrNC = "Yellow";chPRcolor $Global:PrP1 $Global:PrP2 $Global:PrP3 $global:PrNC
    # Printer Driver Information
        $Global:PRMNum ++;$Get_PrinterDriverInfo=$Global:PRMNum;
            $Global:PrP1 = " $Global:PRMNum. `t View ";
            $Global:PrP2 = "Printer Driver ";
            $Global:PrP3 = "Information"
            $global:PrNC = "Yellow";chPRcolor $Global:PrP1 $Global:PrP2 $Global:PrP3 $global:PrNC
    # Printer logs
        $Global:PRMNum ++;$Get_PrinterLogs=$Global:PRMNum;
            $Global:PrP1 = " $Global:PRMNum. `t View ";
            $Global:PrP2 = "Printer Logs ";
            $Global:PrP3 = "on selected system(s)"
            $global:PrNC = "Yellow";chPRcolor $Global:PrP1 $Global:PrP2 $Global:PrP3 $global:PrNC
                       
        # Setting up the PrintMenu in this way allows you to move PrintMenu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $Global:PRMNum ++;$Global:PrP1 =" $Global:PRMNum.  FIRST";$Global:PrP2 = "HIGHLIGHTED ";$Global:PrP3 = "END"; $global:PrNC = "yellow";chPRcolor $Global:PrP1 $Global:p2 $Global:PrP3 $global:PrNC
        Write-Host "$Global:Prt4$Global:Prd4"  -ForegroundColor Green
        $Global:PRMNum = 0;[int]$PrintMenuSelect = Read-Host "`n`tPlease select the PrintMenu item you would like to do now"
    }
    switch($PrintMenuSelect)
    {
        # Variables at the beginning of these lines in the switch are numbers assigned by their position in the above
        # configured PrintMenu layout. This allows them to be moved around and let the numbers automatically adjust
        $PRExit{$PrintMenuSelect=$null}
        $Get_PrPCList{Clear-Host;Get-PCList;Reload-PromptPrintMenu}
        $Get_PrintInfo{Clear-Host;Get-PrinterInfo;Get-PrinterPortInfo;Get-PrinterDriverInfo;Reload-PromptPrintMenu}
        $Get_PrinterInfo{Clear-Host;Get-PrinterInfo;Reload-PromptPrintMenu}
        $Get_PrinterPortInfo{Clear-Host;Get-PrinterPortInfo;Reload-PromptPrintMenu}
        $Get_PrinterDriverInfo{Clear-Host;Get-PrinterDriverInfo;Reload-PromptPrintMenu}
        $Get_PrinterLogs{Clear-Host;Get-PrinterLogs;Reload-PromptPrintMenu}
        default
        {
            $PrintMenuSelect = $null
            $Global:PrNC = " "
            $Global:PrP1 = " "
            $Global:p2 = " "
            $Global:p3 = " "
            $Global:PRMNum = 0
            Reload-PromptPrintMenu
        }
    }
}
function Reload-PromptPrintMenu()
{
        Read-Host "Hit Enter to continue"
        $PrintMenuSelect = $null
        $Global:PRMNum = 0
        PrintMenu
}

function Reload-PrintMenu()
{
        $PrintMenuSelect = $null
        $Global:PRMNum = 0
        PrintMenu
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
            $Global:PrP1 = "`nRunning: Get-WmiObject -Class win32_printer -Credential $cred -ComputerName"
            $Global:PrP2 = "$Global:PC"
            $Global:PrP3 = " | select name, drivername, portname, shared, sharename |ft -AutoSize"
            $Global:PrNC = "Yellow";
            chPRcolor $Global:PrP1 $Global:PrP2 $Global:PrP3 $Global:PrNC
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
    $Global:PrP1 = "`nRunning: get-printerport -computer "
    $Global:PrP2 = "$Global:PC"
    $Global:PrP3 = " |select ComputerName,Description,PrinterHostAddress,Name |ft -Wrap -AutoSize"
    $Global:PrNC = "Yellow";
    chPRcolor $Global:PrP1 $Global:PrP2 $Global:PrP3 $Global:PrNC
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
    $Global:PrP1 = "`nRunning: get-printerDriver -computer"
    $Global:PrP2 = "$Global:PC"
    $Global:PrP3 = " |select * |ft -wrap -Autosize"
    $Global:PrNC = "Yellow";
    chPRcolor $Global:PrP1 $Global:PrP2 $Global:PrP3 $Global:PrNC
    $prin3 = get-printerDriver -computer $Global:PC |select ComputerName,Name|ft -wrap -Autosize
    $prin3
        }
    }
}
# Get Printer logs
Function Get-PrinterLogs
{
<#
    .NAME
        Get-PrintLogs.ps1
    .SYNOPSIS
        Used to gather Print logs from the print server
    .Description
        # Laclede Electric Cooperative
        # Jason Crockett
         As a prerequisite, First do the following: "Event Viewer -> Applications and Service Logs -> Microsoft -> Windows -> Print Service -> Operational -> Enable log"

         This code was verified and modified by Jason Crockett - Laclede Electric Cooperative to include a parameter for computername and to export to a CSV file.
         Most of the code is from BSOD2600 in a post on Tuesday, May 03, 2011 on
         https://social.technet.microsoft.com/Forums/scriptcenter/en-US/007be664-1d8d-461c-9e0b-d8177106d4f8/vb-script-or-powershell-script-for-auditing-win2k8-print-server?forum=ITCG
    .Parameter 
        Comps = "CSV list of computernames"
        Will default to local computername if parameter is not given in CLI
    .EXAMPLE
        .\Get-PrintLogs CompName
#>
# Use the following if we want to retain a single PCList across many menu functions 
if ($Global:PCList -eq $null)
{
    Get-PCList
}
<# Used when this was a separate script file
param(
    [parameter(position=1)]
    [string]$comps= $env:COMPUTERNAME
)
#>
$Global:PCCnt = ($Global:PCList).Count
foreach($Global:pcLine in $Global:PCList)
{
    Identify-PCName
    # Only attempt to get information from computers that are reachable with a ping
    If (Test-Connection $Global:pc -Count 1 -Quiet -ErrorAction SilentlyContinue)
    {
        Write-Warning "Collecting event logs from $Global:pc..."
        $Date = (get-date) - (new-timespan -day 8) # modify to get a different timespan of logs
        $PrintEntries = Get-WinEvent -ea SilentlyContinue -ComputerName $Global:pc -FilterHashTable @{ProviderName="Microsoft-Windows-PrintService"; StartTime=$Date; ID=307}
        $strOutput = ""
        $File = "Printing Audit - " + (Get-Date).ToString("yyyy-MM-dd") + ".csv"
        write-output "Date,Username,Full Name,Client,Printer Name,Print Size,Pages,Document" | Out-File $File -Encoding ascii
        write-Host "Parsing event log entries..."

        ForEach ($PrintEntry in $PrintEntries)
         { 
         #Get date and time of printjob from TimeCreated
         $Date_Time = $PrintEntry.TimeCreated
         $entry = [xml]$PrintEntry.ToXml() 
         $docName = $entry.Event.UserData.DocumentPrinted.Param2
         $Username = $entry.Event.UserData.DocumentPrinted.Param3
         $Client = $entry.Event.UserData.DocumentPrinted.Param4
         $PrinterName = $entry.Event.UserData.DocumentPrinted.Param5
         $PrinterPort = $entry.Event.UserData.DocumentPrinted.Param6
         $PrintSize = $entry.Event.UserData.DocumentPrinted.Param7
         $PrintPages = $entry.Event.UserData.DocumentPrinted.Param8

         #Get full name from AD
         if ($UserName -gt "")
         {
          $DirectorySearcher = New-Object System.DirectoryServices.DirectorySearcher
          $LdapFilter = "(&(objectClass=user)(samAccountName=${UserName}))"
          $DirectorySearcher.Filter = $LdapFilter
          $UserEntry = [adsi]"$($DirectorySearcher.FindOne().Path)"
          $ADName = $UserEntry.displayName
         }
 
         #$rawMessage
         $strOutput = $Date_Time.ToString()+ "," +$UserName+ "," +$ADName+ "," +$Client+ "," +$PrinterName+ "," +$PrintSize+ "," +$PrintPages+ "," +$docName
         write-output $strOutput | Out-File $File -append  
 
         }

         Write-Host "The file name is $File"
         }
    }
}