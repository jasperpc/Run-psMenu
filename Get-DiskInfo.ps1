


<#
.Synopsis
Short explanation
.Description
Long Description
.Parameter Comps
list of computernames
.Example 
.\Get-DiskInfo.ps1 -comps dc
#>
[CmdletBinding()] #(allows the use of parameter attributes)
param(
        [Parameter(Mandatory=$True)]
        [string[]]$comps = 'localhost'
)
# Collect Information about the volumes on referenced systems
# Using Win32_Volume instead of Win32_LogicalDisk allows us to easily exclude the optical drive and include drives without drive letters
Get-WmiObject win32_volume -filter 'drivetype=3' -ComputerName $comps | select * |Ft pscomputername ,label,driveletter,@{n='freeGb';e={"{0:N1}" -f ($_.freespace / 1gb -as [int])};align="right"},@{n='TotalGb';e={"{0:N1}" -f ($_.Capacity /1gb -as [int])};align="right"}

<#
# If we were running it from/as a function, it would be as follows:
function Get-DiskInfo{
    # Collect Information about the volumes on referenced systems
    # Using Win32_Volume instead of Win32_LogicalDisk allows us to easily exclude the optical drive and include drives without drive letters
    Get-WmiObject win32_volume -filter 'drivetype=3' -ComputerName $comps | select * |Ft pscomputername ,label,driveletter,@{n='freeGb';e={"{0:N1}" -f ($_.freespace / 1gb -as [int])};align="right"},@{n='TotalGb';e={"{0:N1}" -f ($_.Capacity /1gb -as [int])};align="right"}
}
Get-DiskInfo($comps)
#>