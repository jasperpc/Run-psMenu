<#  
.SYNOPSIS  
    Tests for active port(s) on IP connected systems.
.DESCRIPTION
    Tests for one or more active port(s) on one or more IP connected systems.
.PARAMETER $pc
    Name of system(s) where you want to test the port connection(s).
.PARAMETER $port
    Port(s) to test for activity
.NOTES  
    Name: Get-ActivePorts.ps1
    Author: Jason Crockett
    DateCreated: 2016-03-22
    DateUpdated: 2017-03-02
    List of Ports: http://www.iana.org/assignments/port-numbers
     
    To Do:
        Add ability to input files of CSV system names
        Add ability to input an array for a range of ports
.LINK  
    NONE
.EXAMPLE  
    Get-ActivePorts.ps1 -pc server,workstation -port 445,339,90,80,3389
    Checks ports 445,339,90,80,3389 on systems 'server' and 'workstation' to see if they are listening
#>  

    Param(
        [Parameter(
            Mandatory = $True,
            Position = 0,
            ParameterSetName = '',
            ValueFromPipeline = $True)]
            [array]$pc,
        [Parameter(
            Position = 1,
            Mandatory = $True,
            ParameterSetName = '')]
            [array]$port)

if ($pc -eq $null)
{
    # Need to replace this with a call to the function that handles this in Run-psMenu.ps1
    [array]$pc = (Read-host "`nEnter a comma separated list of systems to scan.").Split(",") |%{$_.trim()}
}
if ($port -eq $null)
{
    [array]$port = (Read-host "`nEnter a comma separated list of port number(s) to scan? ")# .Split(",") |%{$_.trim()}
}
foreach ($system in $pc)
{
        $i=0
        $j=0
        foreach ($prt in $port)
        {  
            $mysock = New-Object net.sockets.tcpclient
            $IAsyncResult = [IAsyncResult] $mysock.BeginConnect($system, $prt,$null, $null)
            $mcts = Measure-Command {$succ = $IAsyncResult.AsyncWaitHandle.WaitOne(3000,$true)}| % totalseconds
            If ($mysock.Connected -eq $true)
                {
                    Write-Host "`nPort $prt is open on $system"
                    $j=1
                }
            Else
                {$i=$i+1 }
            $mysock.Close()
            $mysock.Dispose()    
        }
    if ($j -ne 0)
    {
        if ($i -ne 0)
        {
            Write-Host <#-Object#> "`nThe other "$i " port(s) are closed on $system"
        }        
    }
    else
    {
        if ($i -ne 0)
        {
            Write-Host "`nAll "$i " port(s) are closed on $system"
        }
    }
    
    

}
