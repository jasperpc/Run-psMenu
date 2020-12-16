<#
.NAME
    Get-BootTime.ps1
.SYNOPSIS
 Discover boot times (from select systems)
.DESCRIPTION
 (c) 2019
 Written by Jason Crockett - Laclede Electric Cooperative
 Used to discover boot times (from select systems)
.PARAMETERS
    [array]$Global:PCList
.EXAMPLE
    .\Get-BootTime.ps1 []
    .\Start-MultiThreads.ps1 .\Get-BootTime.ps1
.SYNTAX
    .\Get-BootTime.ps1 []
.REMARKS
    To see the examples, type: help Get-BootTime.ps1 -examples
    To see more information, type: type: help Get-BootTime.ps1 -detailed
.TODO
    When this script is called from Start-MultiThreads.ps1 the $Global:Cred variable is not passed correctly. 
    This may affect the ability of the script to connect to a higher level server.
    This script runs correctly when run by itself after $Global:Cred with a high enough permission set is in memory.
#>

Param([array]$Global:PCList)


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

function Discover-BootTime
{
    if ($Global:PCList -eq $null)
        {
            Clear-Host
        . .\Get-SystemsList.ps1 # load functions in .ps1
            Get-PCList
        }
$ObjectOut = foreach($Global:pcline in $Global:pclist)
    {
            $GWMOResult=$null
            $DateCol=$null
            $TimeCol=$null
        Identify-PCName
        $GWMOResult = Get-WmiObject win32_operatingsystem -ComputerName $Global:pc -Credential $Global:Cred |select csname, @{Label='LastBootUpDate';expression={$_.converttodatetime($_.lastbootuptime).tostring("yyyy/MM/dd")}}, @{Label='LastBootUpTime';expression={$_.converttodatetime($_.lastbootuptime).tostring("hh:mm")}}
        if($GWMOResult -eq $null)
        {
            # Run Command by connecting to PC using Invoke-Command if WMI is blocked
            $data = Invoke-Command -ComputerName $Global:pc -ScriptBlock {systeminfo |find "Boot Time"} -ErrorAction SilentlyContinue
            # Regex pattern to group and strip boot date and time
            $Pattern1 = '^(\w*\s*\w*\s*\w*)(:\s*)(\d*\/\d*\/\d*)(,\s)(\d*:\d*:\d*\s[A-Z]*)'
            if ($data -ne $null)
            {
                $DateCol = ([regex]::Matches($data, $Pattern1).Groups[3].Value)
                $TimeCol = ([regex]::Matches($data, $Pattern1).Groups[5].Value)
            }
        }
        Else
        {
            $DateCol = $GWMOResult.LastBootUpDate
            $TimeCol = $GWMOResult.LastBootUpTime
        }
        # Build Object
        $OBProperties = @{'__Server'=$Global:pc;
            'BootDate'=$DateCol;
            'BootTime'=$TimeCol}
        New-Object -TypeName PSObject -Prop $OBProperties
    }
$ObjectOut |sort BootDate,__Server |select * # ComputerName
}
Discover-BootTime # Comment out this line if this function will only be run when called from elsewhere



