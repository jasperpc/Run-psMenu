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
.TODO
    Combine $ScriptFile and $ScriptArgs if possible since the script currently works with both entered in position 1 anyways
    See if there is a better way to wait for jobs to return before clearing out the progress bar - currently the progress
       bar gets removed after all the jobs have started, not when they are completed. This was changed because the progress bar 
       was refusing to go away in a timely manner.
    Clean up comments
#>
Param([string]$ScriptFile = $(Read-Host "Enter the script path and filename"), # OK to include list of PCs after filename where applicable
    # [string]$ScriptArgs = $(. .\Get-SystemsList.ps1;Get-PCList), #$null, # Not working yet - have conflicts with pcname arg
    [string]$ScriptArgs = $(if ($Global:PCList -eq $null){. .\Get-SystemsList.ps1;Get-PCList}), #$null, # Not working yet - have conflicts with pcname arg
    [int]$MaxThreads = 20,
    [int]$Sleeptimer = 500,
    [int]$MaxWaitAtEnd = 600,
    [string]$OutputType = "Text")
Clear-Host
# $Global:PCList = $null
# . .\Get-SystemsList.ps1 # load functions in .ps1
#Get-PCList
# Start All Jobs
# $Global:PCList
$i = 0
# Capture a list of jobs so we can close only jobs created through this script
[Array]$JobList = $null
foreach($Global:PCLine in $Global:PCList)
{
    $Complete = Get-Date
    # Identify-PCName isolates the system name regardless of input type
    Identify-PCName
    $Global:PC
    # Uncomment the following line to view the Name of the systems as they are processed
    #Write-Warning "The computer is $Global:PC"
    # Check for number of open threads
    # Wait for some threads to close if necessary prior to opening more
    While($(Get-job -State Running).count -ge $MaxThreads)
    {
        Write-Progress -Activity "Creating Server List"`
                        -Status "Waiting for Threads to Close"`
                        -CurrentOperation "$i threads created - $($(get-job -state running).count) threads open"`
                        -PercentComplete ($i / $Global:PCList.Count *100)`
                        -Completed
        Start-Sleep -Milliseconds $Sleeptimer
    }
    $Gdt = Get-Date -Format yyyyMMddHHmmfffff
    $JbName = "$Gdt"+"$Global:pc"
    $JobList+=$JbName #Get-Variable -Name $JbName ([char]$_) -ValueOnly
    # Starting job - $Global:pc"
    $i++
    # $ArgList = $ScriptFile +" " + $Global:ScriptArgs # " + $Global:pc + " 
    Start-job -filepath $ScriptFile -ArgumentList $Global:pc -Credential $Global:Cred -name $JbName |Out-Null
    <#
    $ScriptBlock = {($ScriptFile)} #  $ScriptArgs
    Start-job -ScriptBlock $ScriptBlock -ArgumentList 'no', 'test' -name $JbName |Out-Null
    #>
    # Invoke-Command -ComputerName ($Global:pc) -ScriptBlock {$ScriptFile} -JobName $JbName -ThrottleLimit 16 -AsJob
    # Start-job -ScriptBlock {invoke-expression "& `"$SCriptFile`" $Global:pc $ScriptArgs"} -Name $JbName |out-null
        Write-Progress -Activity "Creating Server List"`
            -Status "Waiting for Threads to Close"`
            -CurrentOperation "$i threads created - $($(get-job -state running).count) threads open"`
            -PercentComplete ($i / $Global:PCList.Count *100)
        <#
        If ($(New-TimeSpan $Complete $(Get-Date)).totalseconds -ge $MaxWaitAtEnd)
        {
            "Killing all jobs still running . . ."
            Get-Job -State Running | Remove-Job -Force # maybe just do a stop-job here rather than removing the job here...?
        }
        #>
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
                -PercentComplete ($(Get-job -State Completed).count / $(Get-job).count * 100)`
                -Completed
    Start-Sleep -Milliseconds $Sleeptimer
}

Write-Warning "Reading all jobs"
If($OutputType -eq "Text")
{
    #foreach ($JBLine in $JobList)
    #{
        #$GTxtJb = 
        Get-job |receive-job -Keep |Select-Object @{Name="ComputerName"; Expression="__Server"},* -ExcludeProperty RunspaceId |select * -ExcludeProperty __Server, RunspaceId |Sort-Object |ft -AutoSize -Wrap
    #}
    # $GTxtJb |gm
    # Get-job |receive-job |Select-Object @{Name="ComputerName"; Expression="__Server"} ,* -ExcludeProperty RunspaceId|select * -ExcludeProperty __Server|Sort-Object |ft -AutoSize -Wrap
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

# $GTxtJb |select * |FT -AutoSize -Wrap
# $GTxtJb |gm #FT -AutoSize -Wrap

