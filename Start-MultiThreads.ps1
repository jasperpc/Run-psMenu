<#
.NAME
    Start-MultiThreads
.SYNOPSIS
 This start multiple threads of a particular script to run as jobs
.DESCRIPTION
 2010-12-01
 http://www.get-blog.com/?p=22     # (using jobs)
 http://www.get-blog.com/?p=189    # (Using runspaces) Not incorporated in this script yet.
 Original Author:  Ryan Witschger

 2019-06-11
 Modified by: Jason Crockett
 Laclede Electric Cooperative
.PARAMETERS
    Position one is a file name
    Position two is a comma-separated value list of computer names
.EXAMPLE
    Use this command with parameters as follows:
    .\start-multithreads.ps1 Used-scriptname.ps1 pc1,pc2
    .\start-multithreads.ps1 Used-scriptname.ps1 (answer Read-Host with pc1,pc2,etc)
    .\start-multithreads.ps1 (answer Read-Host prompts with .\scriptname.ps1 and with pc1,pc2,etc)
    .\start-multithreads.ps1 Used-scriptname.ps1 domain

.REMARKS
    Start-MultiThreads.ps1 [[-ScriptFile] string] [[-ComputerList] string] [-MaxThreads string[]] [-Sleeptimer string[]] [-MaxWaitAtEnd string[]] [-OutputType string[]]
    Default start variables are as follows:
    $MaxThreads = 20,
    $Sleeptimer = 500,
    $MaxWaitAtEnd = 600,
    $OutputType = "Text" (other option is "GridView")

    To see the examples, type: help Start-MultiThreads.ps1 -examples
    To see more information, type: help Start-MultiThreads.ps1 -detailed
#>
Param([string]$ScriptFile = $(Read-Host "Enter the script path and filename"),
    [int]$MaxThreads = 20,
    [int]$Sleeptimer = 500,
    [int]$MaxWaitAtEnd = 600,
    [string]$OutputType = "Text")
Clear-Host
. .\Get-SystemsList.ps1 # load functions in .ps1
Get-PCList
# Start All Jobs
$i = 0
# Capture a list of jobs so we can close only jobs created through this script
[Array]$JobList = $null
foreach($Global:PCLine in $Global:PCList)
{
    # Identify-PCName isolates the system name regardless of input type
    Identify-PCName
    # Uncomment the following line to view the Name of the systems as they are processed
    # "The computer is $Global:PC"
    # Check for number of open threads
    # Wait for some threads to close if necessary prior to opening more
    While($(Get-job -State Running).count -ge $MaxThreads)
    {
        Write-Progress -Activity "Creating Server List"`
                        -Status "Waiting for Threads to Close"`
                        -CurrentOperation "$i threads created - $($(get-job -state running).count) threads open"`
                        -PercentComplete ($i / $Global:PCList.Count *100)
        Start-Sleep -Milliseconds $Sleeptimer
    }
    $Gdt = Get-Date -Format yyyyMMddHHmmfffff
    $JbName = "$Gdt"+"$Global:pc"
    $JobList+=$JbName #Get-Variable -Name $JbName ([char]$_) -ValueOnly
    # Starting job - $Global:pc"
    $i++
    Start-job -filepath $ScriptFile -ArgumentList $Global:pc -name $JbName |Out-Null
    Write-Progress -Activity "Creating Server List"`
            -Status "Waiting for Threads to Close"`
            -CurrentOperation "$i threads created - $($(get-job -state running).count) threads open"`
            -PercentComplete ($i / $Global:PCList.Count *100)
}

# Wait for all jobs to finish
While($(get-job -State Running).count -gt 0)
{
    $ComputersStillRunning = ""
    ForEach($System in $(Get-job -state Running))
    {
        $ComputersStillRunning += ",$($Sytem.name)"
        
    }
    $ComputersStillRunning =$ComputersStillRunning.Substring(1)
    Write-Progress -Activity "Creating Server List"`
                -Status "$($(Get-Job -State Running).count) threads remaining"`
                -CurrentOperation "$ComputersStillRunning"`
                -PercentComplete ($(Get-job -State Completed).count / $(Get-job).count * 100)
    Start-Sleep -Milliseconds $Sleeptimer
}

Write-Warning "Reading all jobs"
If($OutputType -eq "Text")
{
    Get-job |receive-job |Select-Object @{Name="ComputerName"; Expression="__Server"} ,* -ExcludeProperty RunspaceId|
        select * -ExcludeProperty __Server|Sort-Object|ft -AutoSize -Wrap
}
ElseIf($OutputType -eq "GridView")
{
    Get-job |receive-job |Select-Object @{Name="ComputerName"; Expression="__Server"} ,* -ExcludeProperty RunspaceId|
        select * -ExcludeProperty __Server|Sort-Object |out-gridview
}

if ((Get-job).count -gt 0)
{
    # Kill jobs from this script only
    "Killing existing jobs"
    foreach ($Jb in $JobList)
    {
        if ($Jb -ne $null)
        {
            Get-job $Jb |remove-job -Force
        }
    }
}

