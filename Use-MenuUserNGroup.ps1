<#
.NAME
    Use-MenuUserNGroup.ps1
.SYNOPSIS
    Run Active Directory / Network Administrative tasks
.DESCRIPTION
    Called from Run-psMenu to load functions for Active Directory / Network Administrative tasks
 2018-06-04

 Written by Jason Crockett - Laclede Electric Cooperative

 .PARAMETERS
    This parameter doesn't exist (Comps = list of computernames)
.EXAMPLE
    "a command with a parameter would look like: (.\Get-DiskInfo.ps1 -comps dc)"
    .\Use-MenuUserNGroup.ps1 []
.SYNTAX
    .\Use-MenuUserNGroup.ps1 []
.REMARKS
    
.TODo
    In functions like Find-PCsUser test for IP address and get the system name from the IP if necessary prior to querying against it.
        See the following functions in Use-MenuNet.ps1: Get-hostname, Test-ValidIP, and Test-ValidIPPrompt
    In Function Get-GroupMembership we need to change the above configuration to also pull in Contacts it is in the works
#>

<#
[CmdletBinding()] #(allows the use of parameter attributes)
param(
        #[Parameter(Mandatory=$True)]
            #[string[]]$comps = "$env:COMPUTERNAME",
)
#>

# Enable AD modules
Import-Module ActiveDirectory

# Enable Exchange cmdlets (On an Exchange Server)
# add-pssnapin *exchange* -erroraction SilentlyContinue 

Clear-Host
# set script variables
# . .\Get-SystemsList.ps1
[BOOLEAN]$Global:ExitSession=$false
# Get Credentials for later wmi requests
# $cred = Get-Credential "$env:USERDOMAIN\Administrator"

$Script:UnGt4 = "`t`t`t`t"
$Script:UnGd4 = "------------Use-Users and Groups Menu ($Global:PCCnt Systems currently selected)----------------"
$Script:UnGNC = $null
$Script:UnGp1 = $null
$Script:UnGp2 = $null
$Script:UnGp3 = $null
$CDate = Get-Date -UFormat "%Y%m%d"
[int]$UnGMenuselect = 0

Function chcolor($Script:UnGp1,$Script:UnGp2,$Script:UnGp3,$Script:UnGNC){
            
            Write-host $Script:UnGt4 -NoNewline
            Write-host "$Script:UnGp1 " -NoNewline
            Write-host $Script:UnGp2 -ForegroundColor $Script:UnGNC -NoNewline
            Write-host "$Script:UnGp3" -ForegroundColor White
            $script:UnGNC = $null
            $script:UnGp1 = $null
            $script:UnGp2 = $null
            $script:UnGp3 = $null
}
# The following is an example of combining -Property and -ExpandProperty
# (Get-ADUser -filter 'SamAccountName -like $ADUser1' -Properties * |Select-Object -property @{n="ADUserInfo";e={$_.SamAccountName,$_.memberof}}|select -ExpandProperty ADUserInfo)
Function UnGMenu()
{
    while ($UnGMenuselect -lt 1 -or $UnGMenuselect -gt 17)
    {
        Trap {"Error: $_"; Break;}        
        $UnGMNum = 0;Clear-host |out-null
        Start-Sleep -m 50
        if ($Global:PCCnt -eq $null) {$Global:PCCnt = 0}Else {$Global:PCCnt = ($Global:PCList).Count}
        $Script:UnGd4 = "------------Use-Users and Groups Menu ($Global:PCCnt Systems currently selected)----------------"
        Write-host $Script:UnGt4 $Script:UnGd4 -ForegroundColor Green
        # Exit to Main Menu
        $UnGMNum ++;$adExit = $UnGMNum; 
		    $Script:UnGp1 =" $UnGMNum. `tSelect to ";
		    $Script:UnGp2 = "Exit ";
		    $Script:UnGp3 = "to Main Menu"; $Script:UnGNC = "Red";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
        # Prompt for credential with permissions across the network
        $UnGMNum ++;$GCred=$UnGMNum;$Script:UnGp1 =" $UnGMNum.  `tEnter";
            $Script:UnGp2 = "Credentials ";$Script:UnGp3 = "for use in script."; $Script:UnGNC = "Cyan";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
    # Get PCList
        $UnGMNum ++;$Get_PCList=$UnGMNum;$Script:UnG1 =" $UnGMNum.  `tGet";
            $Script:UnG2 = "PC List ";$Script:UnG3 = "for use in script ($Global:PCCnt Systems currently selected)."; $Script:UnGNC = "DarkCyan";chcolor $Script:UnG1 $Script:UnG2 $Script:UnG3 $Script:UnGNC
# List Users assigned to PC
        $UnGMNum ++;$Find_PCsUser = $UnGMNum;
		    $Script:UnGp1 =" $UnGMNum. `tUse PC/IP List to ";
		    $Script:UnGp2 = "view Users ";
		    $Script:UnGp3 = "assigned to PCs."; $Script:UnGNC = "yellow";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
# List which user(s) are logged onto selected computer(s) listed in the domain
        $UnGMNum ++;$Find_UsrsOnPCs = $UnGMNum;
		    $Script:UnGp1 =" $UnGMNum. `tUse PC/IP list to ";
		    $Script:UnGp2 = "find IPs and users ";
		    $Script:UnGp3 = "logged into selected systems. (non-local)`n$Script:UnGt4$Script:UnGt4$Script:UnGt4---"; $Script:UnGNC = "yellow";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
# AD Manage AD Users menu
        $UnGMNum ++;$Inv_Manage_ADUsers = $UnGMNum;
		    $Script:UnGp1 =" $UnGMNum. `tOpen";
		    $Script:UnGp2 = "Manage-ADUsers.ps1 ";
		    $Script:UnGp3 = "menu "; $Script:UnGNC = "green";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
# Group Memberships assigned to a user
        $UnGMNum ++;$Get_UserGroupMembership = $UnGMNum;
		    $Script:UnGp1 =" $UnGMNum. `tEnter user to ";
		    $Script:UnGp2 = "View Groups ";
		    $Script:UnGp3 = "assigned to users."; $Script:UnGNC = "yellow";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
# List Primary user PC assignments
        $UnGMNum ++;$Find_UserPCs = $UnGMNum;
		    $Script:UnGp1 =" $UnGMNum. `tEnter user to ";
		    $Script:UnGp2 = "view PCs ";
		    $Script:UnGp3 = "assigned to users."; $Script:UnGNC = "yellow";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
# User Members in a group
        $UnGMNum ++;$Get_GroupMembership = $UnGMNum;
		    $Script:UnGp1 =" $UnGMNum. `tEnter Group ";
		    $Script:UnGp2 = "to view Members ";
		    $Script:UnGp3 = "assigned to group."; $Script:UnGNC = "yellow";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
# User Members in a localgroup
        $UnGMNum ++;$Get_LocalGroupMembership = $UnGMNum;
		    $Script:UnGp1 =" $UnGMNum. `tEnter name of ";
		    $Script:UnGp2 = "LocalGroup ";
		    $Script:UnGp3 = "to view Members."; $Script:UnGNC = "yellow";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
# List Empty Active Directory Groups
        $UnGMNum ++;$Find_EmptyGroups = $UnGMNum;
		    $Script:UnGp1 =" $UnGMNum. `tList all";
		    $Script:UnGp2 = "Empty Groups ";
		    $Script:UnGp3 = "in the active directory"; $Script:UnGNC = "yellow";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
# Search Groups for Object Membership
        $UnGMNum ++;$Find_SystemGroupMembership = $UnGMNum;
		    $Script:UnGp1 =" $UnGMNum. `tSearch multiple";
		    $Script:UnGp2 = "AD Groups for Computer Object ";
		    $Script:UnGp3 = "Membership - Slow with large group - need to speed this up."; $Script:UnGNC = "yellow";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
# Active Directory Groups and Members -> .CSV
        $UnGMNum ++;$Get_ADGrps_n_Members = $UnGMNum;
		    $Script:UnGp1 =" $UnGMNum. `tAD";
		    $Script:UnGp2 = "Groups and Members ";
		    $Script:UnGp3 = "-> .CSV"; $Script:UnGNC = "yellow";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
# AD Users List
        $UnGMNum ++;$Filter_adusers = $UnGMNum;
		    $Script:UnGp1 =" $UnGMNum. `tEnter User ";
		    $Script:UnGp2 = "to Filter AD User List ";
		    $Script:UnGp3 = "and display PW and logon date."; $Script:UnGNC = "yellow";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
# AD accounts that will never expire
        $UnGMNum ++;$Search_ADNonExpiring = $UnGMNum;
		    $Script:UnGp1 =" $UnGMNum. `tAD accounts";
		    $Script:UnGp2 = "never to expire ";
		    $Script:UnGp3 = ""; $Script:UnGNC = "yellow";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
# AD Privileged Group Membership and password age
        $UnGMNum ++;$Get_ADPrivGrpMbrshp = $UnGMNum;
		    $Script:UnGp1 =" $UnGMNum. `tAD Privileged";
		    $Script:UnGp2 = "Group Membership and Password ";
		    $Script:UnGp3 = "age"; $Script:UnGNC = "yellow";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
# Group Memberships by individual"
        $UnGMNum ++;$get_GrpMbrXInd = $UnGMNum;
            $Script:UnGp1 =" $UnGMNum. `tView ";
		    $Script:UnGp2 = "Group Memberships ";
		    $Script:UnGp3 = "by individual (Requires import of adutils)"; $Script:UnGNC = "red";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
        Write-Host "$Script:UnGt4$Script:UnGd4"  -ForegroundColor Green
        # Setting up the menu in this way allows you to move menu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $UnGMNum ++;$Script:UnGp1 =" $UnGMNum.  FIRST";$Script:UnGp2 = "HIGHLIGHTED ";$Script:UnGp3 = "END"; $Script:UnGNC = "yellow";chcolor $Script:UnGp1 $Script:UnGp2 $Script:UnGp3 $Script:UnGNC
        [int]$UnGMenuselect = Read-Host "`n`tPlease select the menu item you would like to do now"
    }
    switch($UnGMenuselect)
    {
        $adExit{$UnGMenuselect=$null}
        $GCred{Clear-Host;Get-Cred;$UnGMenuselect=$null;reload-PromptUnGMenu}
        $Get_PCList{Clear-Host;Get-PCList;reload-NoPromptUnGMenu}
        $Find_PCsUser{$UnGMenuselect=$null;Find-PCsUser;reload-PromptUnGMenu}
        $Find_UsrsOnPCs{$UnGMenuselect=$null;Find-UsrsOnPCs;reload-PromptUnGMenu} # Good one to multi-thread
        $Inv_Manage_ADUsers{$UnGMenuselect=$null;Invoke-Expression ".\Manage-ADUsers.ps1";reload-NoPromptUnGMenu}
        $Get_GroupMembership{$UnGMenuselect=$null;Clear-Host;Get-GroupMembership;reload-PromptUnGMenu}
        $Get_LocalGroupMembership{$UnGMenuselect=$null;Clear-Host;Get-LocalGroupMembership;reload-PromptUnGMenu}
        $Find_UserPCs{$UnGMenuselect=$null;Find-UserPCs;reload-PromptUnGMenu}
        $Get_UserGroupMembership{$UnGMenuselect=$null;Get-UserGroupMembership;reload-PromptUnGMenu}
        $Find_EmptyGroups{$UnGMenuselect=$null;Find-EmptyGroups;reload-PromptUnGMenu}
        $Find_SystemGroupMembership{$UnGMenuselect=$null;Find-SystemGroupMembership;reload-PromptUnGMenu}
        $Get_ADGrps_n_Members{$UnGMenuselect=$null;Get-ADGrps-n-Members;reload-PromptUnGMenu}
        $Filter_adusers{$UnGMenuselect=$null;Filter-adusers;reload-PromptUnGMenu}
        $Search_ADNonExpiring{$UnGMenuselect=$null;Search-ADNonExpiring;reload-PromptUnGMenu}
        $Get_ADPrivGrpMbrshp{$UnGMenuselect=$null;Get-ADPrivGrpMbrshp;reload-PromptUnGMenu}
        $get_GrpMbrXInd{Write-Warning "get-GrpMbrXInd Requires import of adutils";reload-PromptUnGMenu}
        default
        {
            $UnGMenuselect = $null
            reloUnGMenu
        }
    }
}
Function reload-NoPromptUnGMenu() # Reload with NO prompt (not needed until there's another menu down)
{
        $UnGMenuselect = $null
        UnGMenu
}

Function reload-PromptUnGMenu() # Reload with prompt
{
        Read-Host "Hit Enter to continue"
        $UnGMenuselect = $null
        UnGMenu
}

Function Find-EmptyGroups ()
{
    Clear-Host
    Write-Verbose "Getting AD Groups that have no members"
    Get-ADGroup -Filter * | where { -Not ($_ | Get-ADGroupMember)} |select name |FT -Autosize
}

Function Find-SystemGroupMembership() # Search for system members of AD Groups containing computer objects
{
    Clear-Host
    $GrpSearchString = Read-Host "Type all or part of a group for which you wish to see members (Wildcards accepted) Ex: *Sec*"
    Get-PClist
    $FEPCResult = foreach ($Global:PCLine in $Global:PCList)
    {
        Write-host "." -NoNewline -ForegroundColor Green
        Identify-PCName
        $ADGrps = Get-ADGRoup -filter {Name -like $GrpSearchString}
        $FEResult = foreach ($ADGrp in $ADGrps)
            { 
                $ADGrp |Get-ADGroupMember |Where {$_.name -eq $Global:PC} |sort Name |Select Name,@{n="Group";e={(((($ADGrp.DistinguishedName).split(",") |%{$_.trim()})[0]).split("=")|%{$_.trim()})[1]}}
            }
        if ($FEResult){$FEResult} Else{"$Global:PC"}
    
    }
    $FEPCResult |FT
}


Function Associate-UserAndPC
{
    Write-Output "The following are the workstations found for each user"
    $AUnPResult = $AUnP|Where {$_.Name}|sort Name -Unique
    $OBClass1 = $AUnP.OBClass
    # "OBClass is $OBClass"
    Start-sleep 1
    $UsrPCAssn = foreach ($line in $AUnPResult)
    {
        $LName = ($line.Name)
        $LNameWild = ($line.Name) + "*"
        If ($ObClass1 -eq "User")
        {
            $A1 = (get-adcomputer -Properties * -filter 'description -like $LNameWild' -ErrorAction SilentlyContinue |Where {$_.enabled -eq $true}|select Name,Description -ErrorAction SilentlyContinue)
            $A1SAN = $A1.SamAccountName
            $A1Name = $A1.Name
            Write-Host "$A1Name`t" -NoNewline
            Write-Host "$Lname`t" -NoNewline
            Write-Host "$A1SAN`t"
        }
        Elseif ($ObClass1 -eq "Computer")
        {
            $A1 = (Get-aduser -Filter *  -ErrorAction SilentlyContinue |where {$_.Name -like $line.Description}|select Name, SamAccountName -ErrorAction SilentlyContinue)
            $A1SAN = $A1.SamAccountName
            $A1Name = $A1.Name
            Write-Host "$Lname`t" -NoNewline
            Write-Host "$A1SAN`t`t" -NoNewline
            Write-Host "$A1Name`t"

            # "$LName `t $A1"
        }
    }
    $UsrPCAssn|ft -autosize # Name, SamAccountName 
    $AUnP=$null
}
Function Get-GroupMembership() # Include PC Name where User name is listed in PC obj Description
{
param(
[Parameter(Mandatory=$False)]

$ADGPrompt = (Read-Host "Enter name of one group for which you wish to see members. (ex. Domain Admins)")
)
    $ADG = Get-ADGroup $ADGPrompt -Properties * | Select-Object -property @{n="ADGroupInfo";e={$_.distinguishedname,$_.members}} |select -ExpandProperty ADGroupInfo
    $adgm = $null
    Write-Warning "`nRunning:`tGet-ADGroupMember -Identity $ADG"
    # 
    $adgm = Get-ADGroupMember -Identity $ADGPrompt -Recursive |sort ObjectClass,Name,Description # |select * # Name,SamAccountname,distinguishedName,ObjectClass

    # Need to change the above configuration to also pull in Contacts 
    # $adgm = Get-ADObject -filter * -Properties MemberOf,Mail |where {$ADGPrompt -contains $_.memberof}
    $AUnP = foreach ($adgmRec in $adgm)
    {
        If ($adgmRec.ObjectClass -eq "user" -or $adgmRec.ObjectClass -eq "computer")
        {
            $adgmpcname = $adgmRec.Name
            $adgmRecObClass = $adgmRec.ObjectClass
            $ADPCDescription = Get-ADcomputer -Filter 'Name -eq $adgmpcname' -Properties * -ErrorAction SilentlyContinue |select Name,Description
            $SAN = ($adgmRec |select SamAccountName).SamAccountName
            $OBName =($adgmRec |select Name).Name
            $OBDNAme = ($adgmRec |select DistinguishedName).DistinguishedName
            $OBDescription =($ADPCDescription |select Description).Description
            $ObClass = $adgmRecObClass #).Class
            If ($adgmRec.ObjectClass -eq "user")
            {
                $OBProxyAddress = Get-ADUser -Filter 'SamAccountName -eq $SAN' -Properties * |select UserPrincipalName #ProxyAddresses
                $OBProxyAd = $OBProxyAddress.ProxyAddresses
                Write-Host "($SAN)`t$OBName`t$OBProxyAd" -ForegroundColor Green
            }
            If ($adgmRec.ObjectClass -eq "computer")
            {
                Write-Host "`t$OBName`t$OBDescription" -ForegroundColor Green
            }
            $OBProperties = @{'SamAccountName'=$SAN;
                        'Name'=$OBName;
                        'DistinguishedName'=$OBDNAme;
                        'Description' = $OBDescription;
                        'OBClass' = $ObClass;
                        'OBProxyAd' = $OBProxyAd}
        }
        New-Object -TypeName PSObject -Prop $OBProperties
        $ADPCDescription = $null
    }
        Associate-UserAndPC
        $ADG = $null
}
Function Get-LocalGroupMembership()# Include PC Name where User name is listed in PC obj Description
{
param(
[Parameter(Mandatory=$False)]

$GPrompt = (Read-Host "Enter name of one group for which you wish to see members. (ex. Domain Admins)")
)
 # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        #$Global:PCList = (Get-ADComputer -Filter * |select Name)
        Get-PCList
    }
    foreach ($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        $Global:PC;Invoke-Command -ScriptBlock { Get-LocalGroupMember administrators } -ComputerName ($Global:PC).Name | select PSComputerName, Name, SID, PrincipalSource |FT -AutoSize -Wrap
    }

<#
    $ADG = Get-ADGroup $GPrompt -Properties * | Select-Object -property @{n="ADGroupInfo";e={$_.distinguishedname,$_.members}} |select -ExpandProperty ADGroupInfo
    $adgm = $null
    Write-Warning "`nRunning:`tGet-ADGroupMember -Identity $ADG"
    # 
    $adgm = Get-ADGroupMember -Identity $GPrompt -Recursive |sort ObjectClass,Name,Description # |select * # Name,SamAccountname,distinguishedName,ObjectClass

    # Need to change the above configuration to also pull in Contacts 
    # $adgm = Get-ADObject -filter * -Properties MemberOf,Mail |where {$GPrompt -contains $_.memberof}
    $AUnP = foreach ($adgmRec in $adgm)
    {
        If ($adgmRec.ObjectClass -eq "user" -or $adgmRec.ObjectClass -eq "computer")
        {
            $adgmpcname = $adgmRec.Name
            $adgmRecObClass = $adgmRec.ObjectClass
            $ADPCDescription = Get-ADcomputer -Filter 'Name -eq $adgmpcname' -Properties * -ErrorAction SilentlyContinue |select Name,Description
            $SAN = ($adgmRec |select SamAccountName).SamAccountName
            $OBName =($adgmRec |select Name).Name
            $OBDNAme = ($adgmRec |select DistinguishedName).DistinguishedName
            $OBDescription =($ADPCDescription |select Description).Description
            $ObClass = $adgmRecObClass #).Class
            If ($adgmRec.ObjectClass -eq "user")
            {
                $OBProxyAddress = Get-ADUser -Filter 'SamAccountName -eq $SAN' -Properties * |select UserPrincipalName #ProxyAddresses
                $OBProxyAd = $OBProxyAddress.ProxyAddresses
                Write-Host "($SAN)`t$OBName`t$OBProxyAd" -ForegroundColor Green
            }
            If ($adgmRec.ObjectClass -eq "computer")
            {
                Write-Host "`t$OBName`t$OBDescription" -ForegroundColor Green
            }
            $OBProperties = @{'SamAccountName'=$SAN;
                        'Name'=$OBName;
                        'DistinguishedName'=$OBDNAme;
                        'Description' = $OBDescription;
                        'OBClass' = $ObClass;
                        'OBProxyAd' = $OBProxyAd}
        }
        New-Object -TypeName PSObject -Prop $OBProperties
        $ADPCDescription = $null
    }
        Associate-UserAndPC
        $ADG = $null
#>
}
Function Get-UserGroupMembership()
{
 Clear-Host
 [array]$ADUsers=(Read-Host "Enter a sam account name or * to see what groups the user(s) belong to").split(",") | %{$_.trim()}
    foreach ($ADUser in $ADUsers)
    {
        Write-Host $ADUser -ForegroundColor Yellow
        $GADUList1 = (Get-ADUser -filter 'SamAccountName -like $ADUser'|select -expandproperty SamAccountName)
        foreach ($ADUser1 in $GADUList1)
        {
            $GADUList = (Get-ADUser -filter 'SamAccountName -like $ADUser1' -Properties * |Select-Object select -ExpandProperty memberof)
            Write-Warning "$ADUser1 is a member of the following group(s)"
            foreach ($GADU in $GADUList)
            {
                ($GADU.split(",")[0].substring(3))
            }
        }
        
    }
   # Invoke-Expression ./Review-ADGroupMembership.ps1
}

Function Find-UserPCs
{
    # Where Users's first and last names are in the description of a Computer object
    # this will list the User and the associated computer
    Clear-Host
    [string[]] $Names = ((Read-Host "Enter a comma separated list of all or part of People's first or last names").split(",") | %{$_.trim()})
    foreach ($LName in $Names) 
    {
        get-adcomputer -Filter * -Properties *|where {$_.enabled -eq $true -and $_.Description -like "*$LName*"}|ft Name, Description -autosize -wrap
    }
}
Function Find-PCsUser
{
    # Where Users's first and last names are in the description of a Computer object
    # this will list the User and the associated computer
    # Use this format for the Computer Object Description: "First Last, Optional Extended Description"
    Clear-Host
    $Global:PCLocation = $null
    # Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
    $PCsUserRslt = foreach ($Global:pcLine in $Global:PCList)
    {
        Identify-PCName
        $PCDescription = (Get-ADComputer $Global:PC -Properties * |select Description).Description
        if ($Global:PCLocation -eq $null){$Global:PCLocation = "Unk"}
        $PCDescLikeQuery = ($PCDescription.split(","))[0]
        $GADUserRslt = Get-ADUser -Filter * -Properties * |where {$_.Name -like $PCDescLikeQuery}|select SAMAccountName, Name, Mail, EmailAddress, Enabled
if ($GADUserRslt){
        $PCsUserProperties = @{'SystemName'=$Global:PC;
                        'SamAccountName'=$GADUserRslt.SAMAccountName;
                        'Name'=$GADUserRslt.Name;
                        'Mail'=$GADUserRslt.Mail;
                        'Description'=$PCDescription;
                        'EmailAddress' = $GADUserRslt.EmailAddress;
                        'Enabled' = $GADUserRslt.Enabled
                        'Location' = $Global:PCLocation}
        }
        Else{
        $PCsUserProperties = @{'SystemName'=$Global:PC;
                'SamAccountName'=$GADUserRslt.SAMAccountName;
                'Name'=$PCDescription;
                'Mail'=$GADUserRslt.Mail;
                'Description'=$PCDescription;
                'EmailAddress' = $GADUserRslt.EmailAddress;
                'Enabled' = $GADUserRslt.Enabled
                'Location' = $Global:PCLocation}
        }
        New-Object -TypeName PSObject -Prop $PCsUserProperties
    $PCDescription = $null
    $PCDescLikeQuery = $null
    $GADUserRslt = $null
    }

    $PCsUserRslt |ft SystemName, SAMAccountName, Description, Mail, EmailAddress, Name, Enabled -Wrap -AutoSize
}

Function Find-UsrsOnPCs()
{
# Report the IP Address, ComputerName and Username logged into any active computer listed in Active Directory
Clear-Host
$Global:PCLocation = $null
# Use the following if we want to retain a single PCList across many menu functions 
    if ($Global:PCList -eq $null)
    {
        Get-PCList
    }
Write-Host "Machine Name `tIPAddress `tUser Logged on `tStatus"
foreach ($Global:PCLine in $Global:pcList)
	{
    Identify-PCName
    $CompIP = $null
    $CompName = $null
    $ErrorActionPreference = "silentlycontinue"
    # Get IP from HostName if possible
    $CompIP = ([system.net.dns]::GetHostAddresses($Global:pc)).IPAddressToString
    If ($CompIP)
        {
            $Stat = "Ping Failed"
            $ping = test-connection $Global:pc -Count 1 -Quiet -ErrorAction SilentlyContinue
            # Verify Computername, hostname, with IP
            $CompName = [System.Net.Dns]::GetHostByAddress($CompIP).HostName
            If ($CompName -notlike "*$Global:pc*")
            # If ($Global:pc -notlike "*$CompName*")
                {
                    Write-host "Cannot use $Global:pc as a valid computer name " -ForegroundColor Red
                }
        }
    Else{
            $ping = $false
            $CompIP = "   .   .   .   "
            $Stat = "No Entry in DNS"
        }
    if($ping -eq $true)
        {
            # $pcinfo = @()
		    #Get explorer.exe process
		    $proc = gwmi win32_process -computer $Global:pc -Filter "Name = 'explorer.exe'" -Credential $cred -ErrorAction SilentlyContinue
		    #Search collection of processes for process owner username
            if ($proc -ne $null)
            {
                $Stat = "is logged in."
		        ForEach ($p in $proc) {
                    $pscn = $p.pscomputername
                    $pgo = $p.GetOwner().user
                    Write-Host "$pscn :`t$CompIP`t$pgo`t$Stat`t$Global:PCLocation" -ForegroundColor Green
                }
                $proc = $null
            }
            else
            {
                $Stat = "gwmi unable to get process."
                Write-Host "$Global:PC :`t$CompIP`tUNK`t$Stat`t$Global:PCLocation" -ForegroundColor Yellow
                # Write-host "$Global:PC :`tUnable to get process" -ForegroundColor Red # $script:computer
            }
        }
        else
        {
            Write-Host "$Global:pc :`t$CompIP`tUNK`t$Stat`t$Global:PCLocation" -ForegroundColor Red
            
        }
    }
    $ErrorActionPreference = "continue"
    Write-Output "`n"
    if ($CalledFromFuction -eq $True)
        {
            return                
        }
}

Function Filter-adusers()
{
    Clear-Host
    $Fltr = Read-Host "Type All or part of a users name to filter the list. `nUse * for all users and as a wild-card at the beginning or end of your search criteria."
    $ADUserList = get-aduser -filter 'Name -like $Fltr' -Properties Samaccountname, ObjectGUID, passwordlastset, Name, EmailAddress,DisplayName,LastLogonDate,Created,Mail,Enabled |where {$_.enabled -eq $true} |select Samaccountname, passwordlastset, Name, EmailAddress,DisplayName,LastLogonDate, ObjectGUID,Created,Mail,Enabled |sort SamAccountName, ObjectGUID # |Out-GridView
    $ADUserList.count
    $ADUserList |FT -AutoSize -Wrap
    $adgm = $ADUserList|select Name,DisplayName, Enabled
    $adgm |FT -AutoSize 
}
Function Get-ADGrps-n-Members()
{
$groupsusers=get-adgroup -filter * | 
    ForEach-Object{
                $settings=@{Group=$_.DistinguishedName;Member=$null}
        $_ | get-adgroupmember |
		ForEach-Object
                {
                    $settings.Member=$_.DistinguishedName
                    New-Object PsObject -Property $settings
                }
    }
    
$ScriptDIR = Split-Path $Script:MyInvocation.MyCommand.Path
$GrpUsrcsv = "\GroupsUsers.csv"
$groupsusers | Export-Csv $scriptdir$GrpUsrcsv -NoTypeInformation
Write-host "The information has been saved to: " $ScriptDIR$GrpUsrcsv 
}

Function Search-ADNonExpiring()
{
    # adapted from Netwrix
    Search-ADAccount -PasswordNeverExpires |Select-Object ObjectClass, Name, SAMAccountName, PasswordNeverExpires |ft # | Export-Csv c:\temp\users_password_expiration_false.csv
}

Function Get-ADPrivGrpMbrshp()
{
        Invoke-Expression ".\Get-ADPrivGrpMbrshp.ps1"
}
Function ADUser-Logons()
{
    $nowlessnum =((get-date).AddDays(-30))
    $usecnt = Get-ADUser -Filter {lastlogondate -lt $nowlessnum} |ft
    $usecnt

    
}
