<#
.NAME
    Use-MenuPCApp.ps1
.SYNOPSIS
   Use PC Application functions
.DESCRIPTION
    Menu design patterned after: https://quickclix.wordpress.com/2012/08/14/making-powershell-menus/
    2020-11-13
    Written by Jason Crockett - Laclede Electric Cooperative
.PARAMETERS
    No command-line parameters needed
.EXAMPLE
    .\Use-MenuPCApp.ps1 []
.SYNTAX
    .\Use-MenuPCApp.ps1 []
.REMARKS

.TODO
    Re-work the program search and other high-volume / time consuming functions to utilize multi-threading
#>
# set script variables
# [BOOLEAN]$Global:ExitSession=$false
$Script:PCAt4 = "`t`t`t`t"
# $Script:PCAd4 = "------------Use-PC App Menu ($Global:PCCnt Systems currently selected)----------------"
$Script:PCANC = $null
$Script:PCAp1 = $null
$Script:PCAp2 = $null
$Script:PCAp3 = $null
$Script:PCAPSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
[int]$menuselect = 0
function chAcolor($Script:PCAp1,$Script:PCAp2,$Script:PCAp3,$Script:PCANC)
{
    Write-host $Script:PCAt4 -NoNewline
    Write-host "$Script:PCAp1 " -NoNewline
    Write-host "$Script:PCAp2" -ForegroundColor $Script:PCANC -NoNewline
    Write-host "$Script:PCAp3" -ForegroundColor White
    $Script:PCANC = $null
    $Script:PCAp1 = $null
    $Script:PCAp2 = $null
    $Script:PCAp3 = $null
}
function PCAppMenu()
{
while ($PCAppMenuSelect -lt 1 -or $PCAppMenuSelect -gt 8)
    {
        Clear-Host
        Trap {"Error: $_"; Break;}        
        $Script:PCAMNum = 0;Clear-host |out-null
        Start-Sleep -m 250
        if ($Global:PCCnt -eq $null) {$Global:PCCnt = 0}Else {$Global:PCCnt = ($Global:PCList).Count}
        $Script:PCAd4 = "------------Use-PC App Menu ($Global:PCCnt Systems currently selected)----------------"
        Write-host $Script:PCAt4 $Script:PCAd4 -ForegroundColor Green
#    Write-Host "$Script:PCAt4  1. Exit Now" -ForegroundColor Red
        # Exit
        $Script:PCAMNum ++;$AExit=$Script:PCAMNum;
            $Script:PCAp1 = " $Script:PCAMNum. `t";
            $Script:PCAp2 = "Exit Main PCAppMenu";
            $Script:PCAp3 = " ";
            $Script:PCANC = "Red";chAcolor $Script:PCAp1 $Script:PCAp2 $Script:PCAp3 $Script:PCANC
# Prompt for credential with permissions across the network
        $Script:PCAMNum ++;$GCred=$Script:PCAMNum;$Script:PCAp1 =" $Script:PCAMNum.  `t Enter";
            $Script:PCAp2 = "Credentials ";$Script:PCAp3 = "for use in script."; $Script:PCANC = "Cyan";chPCcolor $Script:PCAp1 $Script:PCAp2 $Script:PCAp3 $Script:PCANC
# $Get_PCList
        $Script:PCAMNum ++;$Get_PCList=$Script:PCAMNum;$Script:PCAp1 =" $Script:PCAMNum.  `t Select";
            $Script:PCAp2 = "systems to use ";$Script:PCAp3 = "with commands on menu ($Global:PCCnt Systems currently selected).";
            $Script:PCANC = "Cyan";chPCcolor $Script:PCAp1 $Script:PCAp2 $Script:PCAp3 $Script:PCANC
# List most recent MS update for the servers/systems
        $Script:PCAMNum ++;$List_RecentUpdate=$Script:PCAMNum;
            $Script:PCAp1 = " $Script:PCAMNum. `t List ";
            $Script:PCAp2 = "most recent MS update ";
            $Script:PCAp3 = "on select systems (Requires Elevation)"
            $Script:PCANC = "Yellow";chPCcolor $Script:PCAp1 $Script:PCAp2 $Script:PCAp3 $Script:PCANC
# MS KB update installation status" # Check for Most Recent Windows Update on servers
        $Script:PCAMNum ++;$get_KB=$Script:PCAMNum;
            $Script:PCAp1 = " $Script:PCAMNum. `t Verify installation status ";
            $Script:PCAp2 = "of MS KB or Windows updates ";
            $Script:PCAp3 = "on select systems (may need to run as Administrator)"
            $Script:PCANC = "Yellow";chAcolor $Script:PCAp1 $Script:PCAp2 $Script:PCAp3 $Script:PCANC
# List programs installed on a pc"
        $Script:PCAMNum ++;$get_apps=$Script:PCAMNum;
            $Script:PCAp1 = " $Script:PCAMNum. `t List ";
            $Script:PCAp2 = "installed programs ";
            $Script:PCAp3 = "on select systems. Requires Admin Credentials"
            $Script:PCANC = "Yellow";chAcolor $Script:PCAp1 $Script:PCAp2 $Script:PCAp3 $Script:PCANC
# "$Script:PCAt4 14. Uninstall (a) program(s) on a pc"
        $Script:PCAMNum ++;$Uninst_App=$Script:PCAMNum;
            $Script:PCAp1 = " $Script:PCAMNum. `t";
            $Script:PCAp2 = "Uninstall program ";
            $Script:PCAp3 = "on select systems (Requires Run Script as Administrator"
            $Script:PCANC = "Yellow";chAcolor $Script:PCAp1 $Script:PCAp2 $Script:PCAp3 $Script:PCANC
#    "View Java versions installed on a PC"
        $Script:PCAMNum ++;$ID_Java=$Script:PCAMNum;
            $Script:PCAp1 = " $Script:PCAMNum. `t View ";
            $Script:PCAp2 = "Java Versions ";
            $Script:PCAp3 = "installed on select systems"
            $Script:PCANC = "Yellow";chAcolor $Script:PCAp1 $Script:PCAp2 $Script:PCAp3 $Script:PCANC
                       
        # Setting up the PCAppMenu in this way allows you to move PCAppMenu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $Script:PCAMNum ++;$Script:PCAp1 =" $Script:PCAMNum.  FIRST";$Script:PCAp2 = "HIGHLIGHTED ";$Script:PCAp3 = "END"; $Script:PCANC = "yellow";chAcolor $Script:PCAp1 $Script:PCAp2 $Script:PCAp3 $Script:PCANC
        Write-Host "$Script:PCAt4$Script:PCAd4"  -ForegroundColor Green
        $Script:PCAMNum = 0;[int]$PCAppMenuSelect = Read-Host "`n`tPlease select the PCAppMenu item you would like to do now"
    }
    switch($PCAppMenuSelect)
    {
        # Variables at the beginning of these lines in the switch are numbers assigned by their position in the above
        # configured PCAppMenu layout. This allows them to be moved around and let the numbers automatically adjust
        $AExit{$PCAppMenuSelect=$null;reload-PCmenu}
        $GCred{Clear-Host;Get-Cred;$PCAppmenuselect = $null;Reload-PCAppMenu} # Called from Run-PSMenu.ps1
        $Get_PCList{Clear-Host;Get-PCList;Reload-PCAppMenu}
        $List_RecentUpdate{List-RecentUpdate;Reload-PromptPCAppMenu}
        $get_KB{get-KB;Reload-PromptPCAppMenu}
        $get_apps{Clear-Host;get-apps;Reload-PromptPCAppMenu}
        $Uninst_App{Invoke-Expression ".\Uninstall-Application.ps1";Reload-PromptPCAppMenu}
        $ID_Java{ID-Java;Reload-PromptPCAppMenu}
        default
        {
            $PCAppMenuSelect = $null
            $global:NC = " "
            $global:p1 = " "
            $global:p2 = " "
            $global:p3 = " "
            $Script:PCAMNum = 0
            $Global:ExitSession=$true;clear-host;break
            Reload-PCAppMenu
        }
    }
}

function Reload-PromptPCAppMenu()
{
        Read-Host "Hit Enter to continue"
        $PCAppMenuSelect = $null
        PCAppMenu
}

function Reload-PCAppMenu()
{
        $PCAppMenuSelect = $null
        PCAppMenu
}

Function List-RecentUpdate()
{
    Clear-Host
    $SN=$null
    $RecentUpdateList=$null
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    # $servers = Get-ADComputer -Filter {OperatingSystem -like "*server*" -and Enabled -eq $true}
    # $servers = Get-ADComputer -Filter {Enabled -eq $true}
    Write-Host "Searching for the most recent updates for $Global:PCCnt Systems"
    $RecentUpdateList = foreach ($Global:pcLine in $Global:pclist)
    {
        Write-Host ". " -NoNewline
        Identify-PCName
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
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
Function get-KB()
{
    clear-host
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    $ghf = $null
    $HotFixErr=$null
    $PCCount = ($Global:pclist).count
    $kb = Read-host "`nEnter a KB number or a partial number with wildcards not including the KB"
    $HFFileName = ".\$CDate" + "_KB$kb.csv"
    $ObjectOut = foreach($Global:pcline in $Global:pclist)
    {
        Identify-PCName
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
            get-hotfix -computername $Global:pc |where {$_.HotFixID -like "KB$kb"}|sort InstalledOn|Select PSComputerName,Source,Description,HotFixID,InstalledBy,InstalledOn|Export-Csv $HFFileName -Append -NoClobber -NoTypeInformation #|out-gridview -Title "Hotfix: KB$KB"
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
    }
    Read-Host "Hit [Enter] to view the following most recent security updates applied to these systems."
    $ObjectOut|sort InstalledOn,PSComputerName -Descending|select pscomputername,HotFixID,Description,InstalledOn,installedBy,HotFixErr|ft -AutoSize -Wrap
    Write-Warning "HotFixes appended to $HFFileName"
}

Function get-apps()
{
    $W32Error = $null
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    $prglist = Read-Host "Enter a search term or nothing to find a program or all programs (use * for wildcards)"
    Write-Warning "Selecting systems on which to look for applications"
    $ObjInstW32Out = foreach ($Global:pcLine in $Global:PCList)
    {
            Identify-PCName
            if (Test-Connection $Global:PC -Count 1 -ErrorAction SilentlyContinue -ErrorVariable W32Error)
            {
            Write-Host "." -nonewline -ForegroundColor Green
                $InstW32 = Get-WmiObject -class Win32_installedwin32program -ComputerName $Global:PC -Credential $cred -ErrorAction SilentlyContinue -ErrorVariable W32Error |where {$_.name -ilike "$prglist"} |select pscomputername,Name,Version
                if ($InstW32)
                {
                    foreach ($InstW32Line in $InstW32)
                    {
                        if ($InstW32Line -eq $null)
                        {if ($W32Error){$W3get2Names = "Error"}elseif ($InstW32Line.Name){$W32Names = "UNK"}} 
                            $OBInstW32Properties = @{'pscomputername'=$Global:PC;
                                'Name'=($InstW32Line.Name);
                                'Version'=($InstW32Line.Version);
                                'Date'=$CDAte;
                                'WMIError' = $W32Error}
                        New-Object -TypeName PSObject -Prop $OBInstW32Properties 
                        $W32Error = $null
                        $InstW32 = $null
                        $InstW32Line = $null
                    }
                }
                Else
                {
                    $OBInstW32Properties = @{'pscomputername'=$Global:PC;
                        'Name'= "NONE";
                        'Version'="N/A";
                        'Date'=$CDAte;
                        'WMIError' = "$prglist not found"}
                    New-Object -TypeName PSObject -Prop $OBInstW32Properties 
                    $W32Error = $null
                }
            }
            Else
            {
            Write-Host "." -nonewline -ForegroundColor Red
                    $OBInstW32Properties = @{'pscomputername'=$Global:PC;
                        'Name'= "Unknown";
                        'Version'="";
                        'Date'=$CDAte;
                        'WMIError' = $W32Error}
                    New-Object -TypeName PSObject -Prop $OBInstW32Properties 
                    $W32Error = $null
            }
    }
    "`n"
    $ObjInstW32Out |Sort-Object WMIError,PSComputerName,Name |select pscomputername,Name,Version,Date,WMIError | FT -AutoSize -Wrap #Out-GridView # |ConvertTo-Csv -NoTypeInformation |select -Skip 1 |Out-File -Append ".\ProgsInPC.csv"
}

Function ID-Java()
{
    Clear-Host
    $AppVar = Read-Host "Enter all or part of the name of an application"
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    $PrgObjResult = foreach($Global:pcLine in $Global:PCList)
    {
            Identify-PCName
        If (Test-Connection $Global:pc -Count 1 -Quiet)
        {
            $tcpc1 = test-connection -computername $Global:PC -Count 2 -Quiet #-ErrorAction SilentlyContinue
            if ($tcpc1 -eq $true)
            {
                <#
                #use WMI to identify Java versions on pc
                Write-Host "`nGet-WmiObject -Class win32_product -Filter "name like 'Java%%'" -ComputerName $Global:PC -Credential $cred |select -expand version |ft"
                Write-Host "`nDetermining (slowly) what versions of Java are installed on "$Global:PC" :`n"
                Get-WmiObject -Class win32_installedwin32program -Filter "name like 'Java%%'" -ComputerName $Global:PC -Credential $cred |select PSComputerName, Name, version, __Path |ft
                #>
                #use WMI to identify Java versions on pc
                Write-Host "`nGet-WmiObject -Class win32_product -Filter "name like 'Cylance%%'" -ComputerName $Global:PC -Credential $cred |select -expand version |ft"
                Write-Host "`nDetermining (slowly) what versions of Cylance are installed on "$Global:PC" :`n"
                # $InstPrg = 
                Get-WmiObject -Class win32_installedwin32program -Filter "name like 'Cylance%%'" -ComputerName $Global:PC -Credential $cred |select PSComputerName, Name, version, __Path
                start-sleep 2
            }
            else 
            {
            write-host "`n"$Global:PC " is not available right now.`n"
            }        
                New-Object PSObject -Property @{
                    pscomputername = $Global:PC;
                    Name = $InstPrg.Name;
                    version = $InstPrg.version;
                    __Path = $InstPrg.__Path
                    }
        }
            $PrgObjResult += $PrgObjResult
    }      
}
