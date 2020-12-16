<#
.NAME
  Review-ADGroupMembership.ps1
.SYNOPSIS
  Search AD/Exchange for Group Membership of specified Users
.DESCRIPTION
  2018-07-20
  Written by Jason Crockett - Laclede Electric Cooperative
  Search Active Directory / Exchange for Group Membership of specified Users
.PARAMETERS
  None yet, but could add parameters for Group, ADUsers, and formatting
.EXAMPLE
  .\Review-ADGroupMembership []
.SYNTAX
  .\Review-ADGroupMembership []
.REMARKS
  Change output to format-list or to out-gridview
#>
Import-Module ActiveDirectory
# Add-PSSnapin *Exchange* -ErrorAction SilentlyContinue
$GDG = Get-ADGroup -Filter 'Name -like "*"'|select SamAccountName, Name, ObjectClass, GroupCategory
$FormatEnumerationLimit=-1 # Default was 4, but -1 doesn't truncate
# $ADUser = Read-Host "Type a search string to find user(s)"
[array]$ADUsers=(Read-Host "Type a search string of a user to find their group(s)").split(",") | %{$_.trim()}

$ObjectOutResult = foreach ($line in $GDG)
{
        $ObjectOut = foreach ($ADUser in ($ADUsers))
            {
            $GDGM = get-ADgroupmember $line.name -ErrorAction SilentlyContinue |select SamAccountName,DisplayName,AddressListMembership |where {$_.SamAccountName -like "$ADUser"} -ErrorAction SilentlyContinue
                if ($GDGM) {
                    $Membership = @{'SamAccountName' = $GDGM.SamAccountName;
                    # 'DisplayName' = $GDGM.DisplayName;
                    # 'AddressListMembership' = $GDGM.AddressListMembership;
                    'GroupName' = $line.Name
                    }
                New-Object -type PSObject -Property $Membership
                }
            }
        if($ObjectOut)
        {
        $Result = @{'SamAccountName' = $ObjectOut.SamAccountName;
            # 'DisplayName' = $ObjectOut.DisplayName;
            # 'AddressListMembership' = $ObjectOut.AddressListMembership;
            'GroupName' = $line.Name # $ObjectOut.GroupName
            }
        New-Object -type PSObject -Property $Result
        }
}
$ObjectOutResult |Select GroupName,SamAccountName <#,DisplayName,AddressListMembership#> |sort GroupName,SamAccountName|FT -autosize # Out-GridView # FL