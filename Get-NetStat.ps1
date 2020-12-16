<#
.Synopsis
 Runs netstat -ano and get-process from powershell against a Win system and combines output with process name.
.Description
 Run netstat -ano from powershell and output with process name.

 Written by Jason Crockett - Laclede Electric Cooperative
 2016-11-03

.Parameter

.Example 
.\Get-NetStat.ps1

.ToDo
    Fix RegEx for matching IP addresses

#>
Clear-Host
$ErrorActionPreference = "SilentlyContinue"
$error.Clear()
# Set script variables
$properties=@()
$tasks=@()
$ps=$null
$script:data=$null
. .\Get-SystemsList.ps1 # load functions in .ps1
Get-PCList
# Prompt for Result filtering specifics
# $Global:PC=Read-Host "`nEnter the name of a system to diagnose"
$ForeignIP = Read-Host "`nEnter all or start of the foreign IP address to limit results on foreign IP"
$ForeignCondition = Read-Host "`nEnter -notlike or -like to see results that donot or do match the foreign IP"
$LocalIP = Read-Host "`nEnter all or start of the local IP address to limit results on local IP"
$port = Read-Host "`nEnter a port number to limit results to a specific local or remote port number"

foreach ($Global:pcLine in $Global:PCList)
{
    Identify-PCName
    function Get-FromLive()
    {
        $script:data = Invoke-Expression "& `".\PSexec.exe`" \\$Global:PC netstat -ano"
        if ($LASTEXITCODE -eq 0)
            {
                Read-Host "PSExec completed Successfully"
            }
        $script:ps = Get-WmiObject -class Win32_process -ComputerName $Global:PC |Select-Object -Property @{N='ProcessName';E={$_.Name}},@{N='Id';E={$_.ProcessID}}, pscomputername|Sort-Object ProcessName
        foreach ($line in $ps[0..$ps.Count])
        {
            # Remove beginning and ending whitespace characters
            $line=$line -replace "^\s+", "" #.trim()
            # Split on whitespace characters
            $line=$line -split '\s+'
        
            # Identify protocol and build object (UDP has no "State")
            $Script:tasks+= @{     # This should possibly be $Script:properties.... like the other ones...?
                ProcessName = $line[0]
                ID = $line[1]
                }
        }
    }
}
function Get-FromLocal()
{
    #store output in $Script:data variable
    $script:data = netstat -ano
    $Script:ps = Get-Process -ComputerName $Global:PC |select Id, ProcessName |Sort-Object ProcessName
}

function NetProcLines()
{
    foreach ($line in $Script:data[4..$Script:data.Count])
    {
    # Remove beginning and ending whitespace characters
    $line=$line.trim()
    $line=$line -replace '\[::(?<a>1)?\]', '[-${a}]'
    # Split on whitespace characters
    $line=$line -split '\s+'
        
    # Identify protocol and build object (UDP has no "State")
    # if ($line -match '^DNS Servers.*') # At one time I had this instead - Not sure why if it doesn't come back I'll delete it.
    if ($line -match '^[T][C][P].*')
        {
            $Script:properties+= @{
                Protocol = $line[0]
                LocalAddressIP = ($line[1] -split ":")[0]
                LocalAddressPort = ($line[1] -split ":")[1]
                ForeignAddressIP = ($line[2] -split ":")[0]
                ForeignAddressPort = ($line[2] -split ":")[1]
                State = $line[3]
                ProcName = ( $ps | Where-Object {$_.ID -eq $line[4]} ).ProcessName
                ProcNum = $line[4]
                FAIPName = (Test-ValidIP ($line[2] -split":")[0])
                }
            # $Script:properties |select faipName
            }
    Elseif ($line -match '^[U][D][P].*')
        {
            $Script:properties+= @{
                Protocol = $line[0]
                LocalAddressIP = ($line[1] -split ":")[0]
                LocalAddressPort = ($line[1] -split ":")[1]
                ForeignAddressIP = ($line[2] -split ":")[0]
                ForeignAddressPort = ($line[2] -split ":")[1]
                State = $null
                ProcName = ( $ps | Where-Object {$_.ID -eq $line[3]} ).ProcessName
                ProcNum = $line[3]
            }
        }
    }
}
function Test-Foreign
{
    #Build output with filters and sorting
    if ($ForeignCondition -eq "-like")
        {
            $Script:properties | % {New-Object psobject -Property $_} |
            Where-Object {($_.ForeignAddressIP -like "$ForeignIP*") `
                -and ($_.LocalAddressIP -like "$LocalIP*") `
                -and (($_.ForeignAddressPort -like "$port*") `
                -or ($_.LocalAddressPort -like "$port*")) } |
                    Sort-Object ProcName,CPU |
                    ft ProcName,Protocol,LocalAddressIP, LocalAddressPort, ForeignAddressIP,ForeignAddressPort, ProcNum,CPU,State -AutoSize
        }
    elseif ($ForeignCondition -eq "-notlike")
        {
            $Script:properties | % {New-Object psobject -Property $_} |
            Where-Object {($_.ForeignAddressIP -notlike "$ForeignIP*") `
                -and ($_.LocalAddressIP -like "$LocalIP*") `
                -and (($_.ForeignAddressPort -like "$port*") `
                -or ($_.LocalAddressPort -like "$port*")) } |
                    Sort-Object ProcName,CPU |
                    ft FAIPName,ProcName,Protocol,LocalAddressIP, LocalAddressPort, ForeignAddressIP,ForeignAddressPort, ProcNum,CPU, State,FAIPName -AutoSize
        }
    else
        {
        $Script:properties | % {New-Object psobject -Property $_} |
        Where-Object {($_.ForeignAddressIP -like "$ForeignIP*") `
            -and ($_.LocalAddressIP -like "$LocalIP*") `
            -and (($_.ForeignAddressPort -like "$port*") `
            -or ($_.LocalAddressPort -like "$port*")) } |
                Sort-Object ProcName,CPU |
                ft ProcName,Protocol,LocalAddressIP, LocalAddressPort, ForeignAddressIP,ForeignAddressPort, ProcNum,CPU, State # -AutoSize
        }
}


function Test-ValidIP
{
    param(
            [string]$Script:gip
            )
            if($Script:gip -eq "[-]")
                {
                    Write-Output "- - - - $Script:gip"
                }
            elseif ($Script:gip -notmatch '^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$')
                {
                    throw "Value $Script:gip does not match a valid IP"
                }
            elseif($Script:gip  -ne "0.0.0.0")
                {
                    throw "Value $Script:gip does not match a valid IP"
                }
            else
                {
                    $FAIP = $Script:gip
                    if ($FAIP -ne "0.0.0.0")
                    {
                        $ADDC = Get-ADDomainController
                        $FNameHost = Resolve-DnsName $FAIP -Server $ADDC |select NameHost
                        $FNameHost.NameHost
                        # Write-Output "$Script:gip is a valid IP address. One moment please."
                    }
                }
}

function Get-ForeignHosts()
{
    $properties.ForeignAddressIP
}
foreach ($Global:PCLine in $Global:PCList)
{
    Identify-PCName
    if ("$Global:PC" -eq $env:COMPUTERNAME)
        {
            Get-FromLocal
        }
    Else 
        {
            Get-FromLive
        }
}
# Count lines of output
# $Script:data.count

NetProcLines
Test-Foreign
$ShowErrorPrompt = Read-Host "There were " $error.Count " foreign addresses that are not valid IPv4 addresses. Do you wish to see the errors? (Y or N)"
if ($ShowErrorPrompt -eq "Y")
{
    $error
}
