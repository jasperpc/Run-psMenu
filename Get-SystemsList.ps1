
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
    $p2 = " C or S ";
    $p3 = "to type your own comma separated list -OR- for Servers only";
    $NC = "Yellow";
    chcolor $p1 $p2 $p3 $NC;$ADorCSL = Read-Host "`t`t"
<# 
    # Option to add a pclist from a file as a source
    $pclist use a file as an optional data source
    $pclist = Get-Content .\pl.txt # Get CSV list from this file listed in the current directory
    $pclist
    $pclist.GetType()
#>
<#
            # Run in a "foreach ($pcline in $PCList){}" statement
            # This if statement could potentially be moved to inside the Get-PCList function, but it might need to be modified to do so
            if ($pcline.name -eq $null)
            {
                $pc = $pcline
            }
            else
            {
                $pc = $pcline.name
            }
#>
if ($ADorCSL -eq "AD")
    {
        # Get the list from computers listed in Active Directory
        # $Global:PCList = Get-ADComputer -Filter 'Enabled -eq "True"'
        # Filter out AD for Enabled systems in the Domain Computers and Domain Controllers OUs
        $Global:PCList = Get-ADComputer -Filter 'Enabled -eq "True"' |where {$_.DistinguishedName -like "*Domain Computers*" -or $_.DistinguishedName -like "*Domain Controllers*"}
    }
elseif ($ADorCSL -eq "C")
    {
        # Get the list from computers listed in comma-separated list typed by the user
        $PCListPrompt = Read-Host "Type in a comma separated list of computers"
        [array]$Global:PCList = $PCListPrompt.split(",") |%{$_.trim()}
        # need to standardize on how this list is presented throughout the script then remove this line: Use just $PCList throughout the script.
        $computers = $Global:PCList
        # https://poshoholic.com/2009/01/19/powershell-quick-tip-how-to-retrieve-the-current-line-number-and-file-name-in-your-powershell-script/
        $CLineNum = $MyInvocation.ScriptLineNumber
        Write-Warning "Called from current line in Script: $CLineNum"

    }
elseif ($ADorCSL -eq "S")
    {
        # Get the list from Servers listed in Active Directory
        # $Global:PCList = Get-ADComputer -Filter 'Enabled -eq "True"'
        # Filter out AD for Enabled systems in the Domain Computers and Domain Controllers OUs
        # $Global:PCList = Get-ADComputer -Filter '-and Enabled -eq "True"' |where {$_.DistinguishedName -like "*Domain Computers*" -or $_.DistinguishedName -like "*Domain Controllers*"}
        $Global:PCList = Get-ADComputer -Filter {OperatingSystem -like "*server*" -and Enabled -eq $true} |where {$_.DistinguishedName -like "*Domain Computers*" -or $_.DistinguishedName -like "*Domain Controllers*"}

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
        $pgo = (gwmi win32_process -Filter "Name = 'explorer.exe'" -ComputerName lec142).GetOwner().user 
        [script]$PCList = $pscn
        # need to standardize on how this list is presented throughout the script then remove this line: Use just $PCList throughout the script.
        $computers = $PCList
        }
    Else
        {
                $ErrorActionPreference = "continue"
        }
    }

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
