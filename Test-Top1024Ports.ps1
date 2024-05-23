<#
.NAME
    Test-Top1024Ports.ps1
.SYNOPSIS
    Test egress communication from current host through firewalls using open ports on "portquiz.net" by default
.DESCRIPTION
    Written by Jason Crockett - Sho-Me Electric Power Cooperative
    Configured the following based on https://www.blackhillsinfosec.com/poking-holes-in-the-firewall-egress-testing-with-allports-exposed/
    ** allports.exposed was no longer live so using portquiz.net for this functionality
    Modify $TestDom if you wish to test open ports on a different domain rather than egress capabilities
    Results are saved to a CSV file
.PARAMETERS
    Parameters not yet incorporated
.EXAMPLE
    "Run command as follows: (.\Test-Top1024Ports.ps1 [])"
        .\Test-Top1024Ports.ps1 []
.SYNTAX
    .\Test-Top1024Ports.ps1 []
.REMARKS
    To see the examples, type: help Test-Top1024Ports.ps1 -examples
    To see more information, type: help Test-Top1024Ports.ps1 -detailed
.TODO
    Add Read-host or parameters to select ports and domain
#>
$CDate = Get-Date -UFormat "%Y%m%d"
$DeviceName = $env:COMPUTERNAME
$TestDom = "portquiz.net" 
$MyIP = (Resolve-DnsName -Name $env:COMPUTERNAME |where {$_.IPAddress -notlike "fe80*"}).IPAddress
$CSVPath = ".\"+$CDate+"_"+$DeviceName+"_"+$MyIP+"_EgressTest.csv"
# Regular expression pattern to match a single number or a range with two numbers
$pattern = '^\d+$|^(\d+)\.\.(\d+)$'
$i=0
Function Test-EgressPorts()
{
if (($DefaultTestPorts = Read-Host "`nType a port range like the following: 80..88 or press [Enter] to use the default: 1..1024") -eq "") 
    {$DefaultTestPorts = 1..1024} 
    else 
    {
    # Test if the input matches the pattern
        if ($DefaultTestPorts -match $pattern) {
            # Convert the string to an actual range
            $DefaultTestPorts = Invoke-Expression $DefaultTestPorts
            #Write-Host "Captured number range:" $numberRange
        } else {
            Write-Host "Invalid input. Please enter a single number or a range with two numbers separated by '..'."
            exit
        }
    }
[array]$ObjectOut = foreach ($port in $DefaultTestPorts)
    {
        $test = new-object system.Net.Sockets.TcpClient
        $wait = $test.beginConnect($TestDom,$port,$null,$null)
        ($wait.asyncwaithandle.waitone(250,$false))
            if($test.Connected)
            {
                Write-host "$port " -NoNewline -ForegroundColor Red
                $PortStat = "open"
            if ($i -gt 30){$i=0;Write-Host "`n"}
            }else
            {
                Write-host "." -NoNewline -ForegroundColor Green
                $PortStat = "closed"
                if ($i -gt 30){$i=0;Write-Host "`n"}
            } 
        $OBProperties = @{'pscomputername'=$DeviceName;
                'Destination'=$TestDom;
                'Port'=$port;
                'Date'=$CDate;
                'Status'=$PortStat}
         New-Object -TypeName PSObject -Prop $OBProperties
         $test.Close()
         $i++
    }
    $ObjectOut | where {$_.Port -ne $null} |select Date, pscomputername, Destination, Port, Status |Export-Csv -NoClobber -NoTypeInformation -Path $CSVPath #| FT # |Out-GridView
    Write-Warning "Results of test saved to $CSVPath"
}
Test-EgressPorts
