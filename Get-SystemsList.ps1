<#
.NAME
    Get-SystemsList.ps1
.SYNOPSIS
 Return a list of systems to the script that calls it
.DESCRIPTION
 2019-06-14
 Written by Jason Crockett - Laclede Electric Cooperative
 Function Get-PCList returns a list of systems to the script that calls it
 Function Identify-PCName takes the list of pcnames and makes sure it can be 
 handled correctly
.PARAMETERS
 N/A
.EXAMPLE
 .\Get-SystemsList.ps1 []
.SYNTAX
 .\Get-SystemsList.ps1 []
.REMARKS
 To see the examples, type: help Get-SystemsList.ps1 -examples
 To see more information, type: help Get-SystemsList.ps1 -detailed
.TODO
#>
function chcolor($p1,$p2,$p3,$NC){
            Write-host $t4 -NoNewline
            Write-host "$p1 " -NoNewline
            Write-host "$p2 " -ForegroundColor $NC -NoNewline
            Write-host "$p3" -ForegroundColor White
            $script:NC = $null
            $script:p1 = $null
            $script:p2 = $null
            $script:p3 = $null
}

function Get-PCList()
{$ADorCSL = $null
    $p1 = "Type AD to use active computers listed in Active Directory or";
    $p2 = " C or S or F";
    $p3 = " to type your own comma separated list -OR- for Servers only - OR File";
    $NC = "Yellow";
    chcolor $p1 $p2 $p3 $NC;$ADorCSL = Read-Host "`t`t"
if ($ADorCSL -eq "AD")
    {
        # Get the list from computers listed in Active Directory
        # $Global:PCList = Get-ADComputer -Filter 'Enabled -eq "True"' # (use this line if you don't care what OU the systems are in)
        # Filter out AD for Enabled systems in the Domain Computers and Domain Controllers OUs
        # Add more properties or * if you need more returned with the query
        $Global:PCList = Get-ADComputer -Filter 'Enabled -eq "True"'-Properties * |where {$_.DistinguishedName -like "*Domain Computers*" -or $_.DistinguishedName -like "*Domain Controllers*" <# -and $_.Name -like "LEC104"#>} |sort-object Name # Enabled,Name,IPv4Address,DistinguishedName,Location
        $Global:PCLocation = $Global:PCList.Location
    }
elseif ($ADorCSL -eq "C")
    {
        # Get the list from computers listed in comma-separated list typed by the user
        $PCListPrompt = Read-Host "Type in a comma separated list of computers"
        [array]$Global:PCList = $PCListPrompt.split(",") |%{$_.trim()}
        # need to standardize on how this list is presented throughout the script then remove this line: Use just $PCList throughout the script.
        $computers = $Global:PCList |sort-object
        # https://poshoholic.com/2009/01/19/powershell-quick-tip-how-to-retrieve-the-current-line-number-and-file-name-in-your-powershell-script/
        $CLineNum = $MyInvocation.ScriptLineNumber
        Write-Warning "Called from current line in Script: $CLineNum"

    }
elseif ($ADorCSL -eq "S")
    {
        # Get the list from Servers listed in Active Directory
        # Filter out AD for Enabled systems in the Domain Computers and Domain Controllers OUs
        $Global:PCList = Get-ADComputer -Filter {OperatingSystem -like "*server*" -and Enabled -eq $true} |where {$_.DistinguishedName -like "*Domain Computers*" -or $_.DistinguishedName -like "*Domain Controllers*"}|sort-object Name

    }
elseif ($ADorCSL -eq "F")
{
    if (($PCListFile = Read-Host "Enter a file name containing a list of systems - CritSys.txt/pl.txt, etc. [Enter] to use default: .\pl.txt") -eq "") {$PCListFile = ".\pl.txt"}
    # $Global:PCList use a file as an optional data source
    $Global:PCList = Get-Content $PCListFile |sort # .\pl.txt # Get list from this file located in the current directory (single item per line)
    Write-Warning "Getting systems from $PCListFile"
}
Else
    {
    $opt1 = $null
        $p1 = "That doesn't match the options type ";
        $p2 = "yes ";
        $p3 = "to run against this system only";
        $NC = "Yellow";chcolor $p1 $p2 $p3 $NC;$opt1 = Read-Host "`t`t"
    If ($opt1 -eq "yes")
        {
        # Use the local computername as the list of computers 
        $pscn = $env:COMPUTERNAME #$proc.pscomputername
        # Get the username for the currently logged-in user
            # $proc = gwmi win32_process -Filter "Name = 'explorer.exe'"
            # $proc.GetOwner().user
        $pgo = (gwmi win32_process -Filter "Name = 'explorer.exe'" -ComputerName $pscn).GetOwner().user 
        $Global:PCList = $pscn
        # need to standardize on how this list is presented throughout the script then remove this line: Use just $PCList throughout the script.
        $computers = $Global:PCList
        }
    Else
        {
                $ErrorActionPreference = "continue"
        }
    }
$Global:PCCnt = ($Global:PCList).Count
}

function Identify-PCName
{
    # This if statement could potentially be moved to inside the Get-PCList function, but it might need to be modified to do so
    if ($Global:pcline.name -eq $null)
    {
        $Global:pc = $Global:pcline
    }
    else
    {
        $Global:pc = $Global:pcline.name
    }
}
