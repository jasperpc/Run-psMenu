<#  
.SYNOPSIS  
    Uninstall an application from a local or remote computer
.DESCRIPTION
    Uninstall an application from a local or remote computer
.VARIABLE $pc
    Name of computer from which to uninstall the application.
.VARIABLE $PartialAppName
    Portion of the name of the application to uninstall 
.NOTES  
    Name: Uninstall-Application.ps1
    Author: Jason Crockett
    DateCreated: 2019-08-07
    Thanks to : https://blogs.technet.microsoft.com/heyscriptingguy/2011/12/14/use-powershell-to-find-and-uninstall-software/
    
    To Do:
        Consider a link to Installer to install current (new) version of application removed.
        Insert parameters to fully automate this with a specific application and a list of computers.
        Create a CSV or HTML log showing which uninstallations were successfull.
.LINK  
    https://blogs.technet.microsoft.com/heyscriptingguy/2011/12/14/use-powershell-to-find-and-uninstall-software/
.EXAMPLE  
    .\Uninstall-Application.ps1
        
    Prompts for an application name and a system name. 
    You are only required to enter a portion of the application name. You will be given an opportunity to bypass 
    any applications you do not wish to install
#>

# Prep to accept incomming parameter values
<#
param(
[Parameter(Mandatory=$False)]
[array[]] $Global:pcList= Get-PCList
# [securestring] $Cred
)
#>
Clear-Host
# . .\Get-SystemsList.ps1;
# Use the following if we want to retain a single PCList across many menu functions 
if ($Global:PCList -eq $null)
{
    Get-PCList
}
$PartialAppName = Read-host "What application do you wish to uninstall? (short form) " # "java" # 
foreach ($Global:pcLine in $Global:pclist)
{
    If (Test-Connection $Global:pc -Count 1 -Quiet)
    {
    Identify-PCName
    $i = 1
    $pcname = $Global:PC
    # Get list of Application versions installed on a system
    Write-Host "I'm searching for $PartialAppName on $Global:PC. Please wait..."
    $AppVerInfo = gwmi win32_product -ComputerName $Global:PC -filter "Name like '%$PartialAppName%'" 
    $appverInfo |select PSComputerName, Name, Version, IdentifyingNumber|sort-object Name
        if ($appverinfo -eq $null)
        {
            Write-host "Sorry, $PartialAppName was not found on $Global:PC"
        }
        Else
        {
            foreach ($AVIline in $AppVerInfo)
            {
            # Break out the fields in the $AppVerInfo result into usable variables
            # $IdentNum
                $IdentNum = $AVIline |Select -Expandproperty IdentifyingNumber
                $IdentNum = $IdentNum.Trim('{}')
            # $IdentName / IdentStringName (this can be consolidated)
                $IdentStringName = $AVIline|Select -Expandproperty Name
            # $IdentVersion
                $IdentVersion = $AVIline |select -expandproperty Version
                $classKey="IdentifyingNumber=`"`{$Identnum`}`",Name=`"$IDentStringName`",Version=`"$IdentVersion`""
            # Assign a variable so that it can be uninstalled
                $WMICK = ([wmi]"\\$Global:PC\root\cimv2:Win32_product.$classkey")
            # Write-host "..wmick.SystemProperties.."
                $WMICK |select PSComputerName,Name,Version,IdentifyingNumber |ft
                $Uninst = "N"
                $Uninst = Read-Host "`nEnter `"Y`" if you want to Uninstall the above application. "
                if ($Uninst -eq "Y")
                {
                    Write-Host "I'm uninstalling  $IDentStringName - $IDentVersion `n"
                    # Uninstall the product
                    if ($WMICK.Uninstall().returnvalue -eq 0) 
                    {
                        write-host "Successfully uninstalled $IDentStringName from $($Global:PC)" 
                    }
                    else 
                    {
                        write-warning "Failed to uninstall $IDStringName from $($Global:PC)." 
                    }
                }
            # Increment to the next record found in the query (used for testing and reference only)
                $i ++
            }

        }
    }
}