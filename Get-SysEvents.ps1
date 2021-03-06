<#
.NAME
    Get-SysEvents.ps1
.SYNOPSIS
    Used to gather Service TAG numbers and event logs from multiple systems
.DESCRIPTION
    (c) Copyright: Laclede Electric Cooperative; Jason Crockett
    2015-10-08
.PARAMETERS
     -pclist
     -WinLogName='System'
     -EntryLevel_Type='Error'
     -MaxEvents = 1000
     -cnt = 10
     -id
     -Source
.EXAMPLE
    .\Get-SysEvents.ps1 -comps Server1, PC1 -evnts Application -types Warning -cnt 5 -id 4101 -Source Wininit
    Winitit Source shows Chkdsk results. Increase the -cnt until it is shown
.REMARKS
    Excellent write-up on Windows Event logs at:
    https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4776
    https://www.ultimatewindowssecurity.com/securitylog/book/page.aspx?spid=chapter4
#>

# Prep to accept incomming parameter values
param(
[Parameter(Mandatory=$False)]
[string[]] $PCList= (if ($Global:PCList -eq $null){. .\Get-SystemsList.ps1 Get-PCList}), # $env:COMPUTERNAME
[string[]] $WinLogName='System',
[string] $EntryLevel_Type='Error',
[string] $MaxEvents = 1000,
[string] $cnt = 10,
[string] $id,
[string] $Source,
[securestring] $Cred
)
Clear-host
# Declare Variables
$usr = [environment]::username
$pchost = [environment]::machinename
$i = 0
# set Global and Script variables
. .\Get-SystemsList.ps1

if ($Global:Cred -eq $null)
{
    $Global:cred = Get-Credential "$env:USERDOMAIN\Administrator"
}
# Change value in $cnt to increase the number of events in the log to view

# Processes
Write-host "The current user name is: " $usr `n
Write-host "The current computer name is: " $pchost `n

<# Write-host " the environmental UserDomain is: "$env:USERDOMAIN `n #>
Write-host "You requested $EntryLevel_Type logs on $($Global:pclist.Count) System(s)"
Start-Sleep -Seconds 2
# Get the Dell Service Tag # from a list of pcs

function pcinfo($Global:pclist, $Global:Cred)
{
    get-wmiobject -computername $Global:pclist -class win32_bios -erroraction SilentlyContinue |select SerialNumber, pscomputername
}
pcinfo $Global:pclist $Global:Cred
# Get-EventLog -LogName System -Newest 50 -EntryType Error -ComputerName pcname
function Erretrieve([string[]]$Global:pclist, [string]$WinLogName,[string]$EntryLevel_Type)
{

    If ($WinLogName -eq "Security") 
        {
        Write-host "`nPlease wait. It could take a while for me to gather the correct information"
        Get-EventLog -LogName $WinLogName -Newest $cnt -Entrytype $EntryLevel_Type -ComputerName $Global:pclist | 
            where {$_.EventID -eq "4625" -or $_.EventID -eq "4648" -or $_.EventID -eq "4776" -or $_.EventID -eq "4676" -or $_.EventID -eq "4674"} |
        ft -Property MachineName, EventID, Index, TimeGenerated, Source, Message, MaximumKilobytes, Log -AutoSize -Wrap #InstanceID EntryType
        }
    else 
        {
            Get-EventLog -LogName $WinLogName -Newest $cnt -Entrytype $EntryLevel_Type -ComputerName $Global:pclist |
                ft -Property MachineName, EventID, Index, TimeGenerated, Source, Message, MaximumKilobytes, Log -AutoSize -Wrap #InstanceID EntryType
        }
}

<# Enable progress bar if needed sometime 
for ($i = 1; $i -le 100; $i++)
    {Write-Progress -Activity "Please wait..." -status "$i% Complete:" -PercentComplete $i;}
#>


Foreach ($Global:pcLine in $Global:pclist){
    Identify-PCName
Write-Warning "`nUsing Get-WinEvent rather than Get-EventLog (this supports credentials)`n"
    If (Test-Connection $Global:pc -Count 1 -Quiet)
    {
        If ($id) # $id exists with a value
        {
            If ($Source) # $Source exists with a value
            {
                Write-Warning $Source
                Get-WinEvent -LogName $WinLogName -MaxEvents $MaxEvents -ComputerName $Global:pc -Credential $Global:cred |where {$_.ID -like "*$id*" -and $_.providerName -like "*$Source*"}|select TimeCreated,MachineName,Id,LogName,Message,ProcessId,ProviderId -Last $cnt|ft -AutoSize -Wrap
            }
            Else
            {
                Get-WinEvent -LogName $WinLogName -MaxEvents $MaxEvents -ComputerName $Global:pc -Credential $Global:cred |where {$_.ID -like "*$id*"}|select TimeCreated,MachineName,Id,LogName,Message,ProcessId,ProviderId -Last $cnt |ft -AutoSize -Wrap
            }
        }
        <#
        ElseIf ($Source) # $Source -ne $null
        {
            Write-Warning $Source
            Get-WinEvent -LogName $WinLogName -MaxEvents $MaxEvents -ComputerName $Global:pc -Credential $Global:cred |where {$_.providerName -like "$Source"}|select ProviderName,MachineName,Id,LogName,Message,ProcessId,ProviderId -Last $cnt |ft -AutoSize -Wrap
            # no need for an if or ElseIf because the if for "$id -ne $null" already
        }
        #>
        Else
        {
            # Get-WinEvent -LogName $WinLogName -MaxEvents $MaxEvents -ComputerName $Global:pc -Credential $Global:cred |where {$_.ID -eq "4634" -or "4624" -or "4625" -or "4740"} |select ProviderName,MachineName,Id,LogName,Message,ProcessId,ProviderId,Source -Last $cnt |fl -AutoSize -Wrap
            Get-WinEvent -LogName $WinLogName -MaxEvents $MaxEvents -ComputerName $Global:pc -Credential $Global:cred |select ProviderName,MachineName,Id,LogName,Message,ProcessId,ProviderId,Source -Last $cnt |ft -AutoSize -Wrap
            Get-WinEvent -ListLog $WinLogName -ComputerName $Global:pc -Credential $Global:cred -Verbose    
        }
    }
    else 
    {
        Write-Warning "$Global:PC cannot be reached "
    }
}

$continueEventLog = Read-host "Type Y to also run a non-filtered Get-EventLog or Security EventLog. Hit any other key to ignore"
If ($continueEventLog)
{
    Erretrieve $Global:pclist $WinLogName $EntryLevel_Type  -ErrorAction Suspend
}
Write-Host "done"