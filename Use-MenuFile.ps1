<#
.NAME
    Use-MenuFile.ps1
.SYNOPSIS
   Use File-related functions
.DESCRIPTION
    Menu design patterned after: https://quickclix.wordpress.com/2012/08/14/making-powershell-menus/
    2020-11-13
    Written by Jason Crockett - Laclede Electric Cooperative
.PARAMETERS
    No command-line parameters needed
.EXAMPLE
    .\Use-MenuFile.ps1 []
.SYNTAX
    .\Use-MenuFile.ps1 []
.REMARKS

.TODo
    Break out funciton Iterate-Profiles into an isolated script since this is used by multiple menu items
    Ex: Copy a file / Delete a file / fix links / view mapped drives / read data from a file
    Ex: Don't copy to all profiles but to just the public profile, etc
    Break/rename out Manage-ADUserFiles into multiple functions or scripts
#>

Clear-Host
# set script variables
# . .\Get-SystemsList.ps1
[BOOLEAN]$Global:ExitSession=$false
# Get Credentials for later wmi requests
# $cred = Get-Credential "$env:USERDOMAIN\Administrator"
# $CDate = Get-Date -UFormat "%Y%m%d"

$Script:Ft4 = "`t`t`t`t"
$Script:Fd4 = "-----------------Use-File Menu ($Global:PCCnt Systems currently selected)----------------"
$Script:FNC = $null
$Script:Fp1 = $null
$Script:Fp2 = $null
$Script:Fp3 = $null
[int]$ad = 0

Function chFCcolor($Script:Fp1,$Script:Fp2,$Script:Fp3,$Script:FNC){
            
            Write-host $Script:Ft4 -NoNewline
            Write-host "$Script:Fp1 " -NoNewline
            Write-host $Script:Fp2 -ForegroundColor $Script:FNC -NoNewline
            Write-host "$Script:Fp3" -ForegroundColor White
            $Script:FNC = $null
            $Script:Fp1 = $null
            $Script:Fp2 = $null
            $Script:Fp3 = $null
}
# The following is an example of combining -Property and -ExpandProperty
# (Get-ADUser -filter 'SamAccountName -like $ADUser1' -Properties * |Select-Object -property @{n="ADUserInfo";e={$_.SamAccountName,$_.memberof}}|select -ExpandProperty ADUserInfo)
Function FileMenu()
{
    while ($FileMenuSelect -lt 1 -or $FileMenuSelect -gt 32)
    {
        Trap {"Error: $_"; Break;}        
        $FMNum = 0;Clear-host |out-null
        Start-Sleep -m 50
        if ($Global:PCCnt -eq $null) {$Global:PCCnt = 0}Else {$Global:PCCnt = ($Global:PCList).Count}
        $Script:Fd4 = "------------Use-PC Menu ($Global:PCCnt Systems currently selected)----------------"
        Write-host $Script:Ft4 $Script:Fd4 -ForegroundColor Green
# Exit to Main Menu
        $FMNum ++;$adExit = $FMNum; 
		    $Script:Fp1 =" $FMNum. `tSelect to ";
		    $Script:Fp2 = "Exit ";
		    $Script:Fp3 = "to Main Menu"; $Script:FNC = "Red";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC
# $Get_PCList
        $FMNum ++;$Get_PCList=$FMNum;$Script:Fp1 =" $FMNum.  `t Select";
            $Script:Fp2 = "systems to use ";$Script:Fp3 = "with commands on menu ($Global:PCCnt Systems currently selected).";
            $Script:FNC = "Cyan";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC
        # Perform one of the following: Copy a file, fix links, view mapped drives to/on selected computers in this domain
        # **Requires permission to view C$ shares on all selected systems **
        $FMNum ++;$adManageADUFiles = $FMNum;
		    $Script:Fp1 =" $FMNum. `tPerform one of the following: ";
		    $Script:Fp2 = "Copy a file, fix links, `n$Script:Ft4`t`tview mapped drives,`tor read data from a file ";
		    $Script:Fp3 = "`n$Script:Ft4`t`tto/on selected computers in this domain`t**Elevated Permissions Required**"; $Script:FNC = "yellow";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC
        # List OpenFiles on Servers
        $FMNum ++;$adOpFiles = $FMNum;
            $Script:Fp1 =" $FMNum. `tList ";
		    $Script:Fp2 = "open files ";
		    $Script:Fp3 = "on AD file servers... "; $Script:FNC = "yellow";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC
# Find Duplicate Files
        $FMNum ++;$Compare_Files=$FMNum;
            $Script:Fp1 = " $FMNum. `tFind ";
            $Script:Fp2 = "Duplicate Files ";
            $Script:Fp3 = "on a system"
            $Script:FNC = "Yellow";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC
# View a file hash
        $FMNum ++;$run_get_filehash=$FMNum;
            $Script:Fp1 = " $FMNum. `tView ";
            $Script:Fp2 = "hash ";
            $Script:Fp3 = "of a file"
            $Script:FNC = "Yellow";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC
# Get a list of Files Modified in the last xx days
        $FMNum ++;$Get_ModifiedFileList=$FMNum;
            $Script:Fp1 = " $FMNum. `tGet a ";
            $Script:Fp2 = "list of Files Modified ";
            $Script:Fp3 = "in the last xx days"
            $Script:FNC = "Yellow";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC
# NTFS Folder - Change Ownership and Add FullControl permissions then prompt to delete.
            $FMNum ++;
            $Manage_NTFS=$FMNum;
            $Script:Fp1 = " $FMNum. `tNTFS Folder - ";
            $Script:Fp2 = "Change Ownership and Add FullControl permissions ";
            $Script:Fp3 = "then prompt to delete"
            $Script:FNC = "Yellow";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC
        # Browse-Share
        $FMNum ++;$Browse_Share=$FMNum;
            $Script:Fp1 = " $FMNum. `tIncomplete: Browse ";
            $Script:Fp2 = "SMB Share ";
            $Script:Fp3 = "on a system"
            $Script:FNC = "Yellow";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC

        Write-Host "$Script:Ft4$Script:Fd4"  -ForegroundColor Green
        # Setting up the menu in this way allows you to move menu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $MNum ++;$Script:Fp1 =" $MNum.  FIRST";$Script:Fp2 = "HIGHLIGHTED ";$Script:Fp3 = "END"; $Script:FNC = "yellow";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC
        [int]$FileMenuSelect = Read-Host "`n`tPlease select the menu item you would like to do now"
    }
    switch($FileMenuSelect)
    {
        $adExit{$FileMenuSelect=$null}
        $Get_PCList{Clear-Host;Confirm-Prompts;Get-PCList;Reload-PromptFileMenu}
        $adManageADUFiles{$FileMenuSelect=$null;Manage-ADUserFiles;Reload-PromptFileMenu} # Requires permission to view C$ shares on all selected systems
        $adOpFiles{$FileMenuSelect=$null;get-OpenFiles;Reload-PromptFileMenu}
        $Compare_Files{Compare-Files;Reload-PromptFileMenu}
        $run_get_filehash{run-get-filehash;Reload-PromptFileMenu}
        $Get_ModifiedFileList{Get-ModifiedFileList;Reload-PromptFileMenu}
        $Manage_NTFS{Clear-Host;Manage-NTFS;Reload-PromptFileMenu}
        $Browse_Share{Browse-Share;Reload-PromptFileMenu}
        
        default
        {
            $FileMenuSelect = $null
            Reload-PromptFileMenu
        }
    }
}


Function Reload-PromptFileMenu()
{
        Read-Host "Hit Enter to continue"
        # Start-Sleep -m 250
        $FileMenuSelect = $null
        FileMenu
}



# This can be designed to be called from different areas.
# It is currently called from function UsrsOnPCs to run against all active PCs
Function List-MappedDrives()
{
    # from calling function: $proc = gwmi win32_process -computer $pscn -Filter "Name = 'explorer.exe'"
    # make $userdir variable equal to current user SamAccountName
    $domain = $env:USERDOMAIN
    $userdir = $pgo
    $psdname = "$pscn"+"hku"
    $npsd = New-PSDrive -PSProvider Registry -Name $psdname -root HKEY_USERS
    # comment out this foreach and the next $userdir variable declaration to only apply the following action to the currently loged on user
    $ProfileList = gci "\\$pscn\C$\Users\" -Directory -Force
    foreach ($PL in $ProfileList)
    {
        $Name = $PL.Name
        if ($Name -ne "All Users" -and $Name -notlike "Default*" -and $Name -ne "Public")
        {
            $strw32acct = Get-wmiobject -Class win32_useraccount -Filter "Domain = '$domain' AND Name = '$Name'" |select Name,FullName,SID
            $strSID = $strw32acct.SID
            Write-Host $strw32acct.Name,$strSID,"`n" -NoNewline
            $path = $psdname + ":\" + $strSID + "\Network"
            if (test-path -Path $path) {Write-host "`t$Name,TRUE,$strSID,network,$pscn`t" -ForegroundColor Green} else {Write-Host "`t$Name,FALSE,$strSID,network,$pscn`t" -ForegroundColor Red}
            if (Test-Path -Path $path"\s") {Write-host "`t$Name,TRUE,$strSID,network\s,$pscn`t" -ForegroundColor Green} else {Write-host "`t$Name,FALSE,$strSID,network\s,$pscn`t" -ForegroundColor Red}
            # "Path:$path"
            # gci $path |select Name,@{n="Remote Path";e={($_.Property | select -ExpandProperty )-join ','}}
            gci $path |select @{n="DriveLetter";e={$_.pschildname}},@{n="RegistryKey";e={$_.Name}}|ft
            $strSID=$null
        }
    }
    Remove-PSDrive -Name $psdname
    $pc = $null

}
Function Manage-ADUserFiles()
{
    $Choose2CopyFile = $null
        $Script:Fp1 = "Type the word - ";
        $Script:Fp2 = "copy ";
        $Script:Fp3 = " - if you wish to copy a file to each user's desktop";
        $Script:FNC = "Yellow";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC;$Choose2CopyFile = Read-Host "`t`t"
    $ReplaceLink = $null
        $Script:Fp1 = "Type ";
        $Script:Fp2 = "dlink or plink ";
        $Script:Fp3 = "to fix desktop links - OR - Links in a path";
        $Script:FNC = "Yellow";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC;$ReplaceLink = Read-Host "`t`t"
    $Opt2 = $null
        $Script:Fp1 = "Type the word - ";
        $Script:Fp2 = "map ";
        $Script:Fp3 = "- if you wish to see mapped drives on each users's system";
        $Script:FNC = "Yellow";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC;$Opt2 = Read-Host "`t`t"
    $CollectFile = $null
        $Script:Fp1 = "Type the word - ";
        $Script:Fp2 = "collect ";
        $Script:Fp3 = " - if you wish to collect the contents of a pre-determined file from each system's public download folder";
        $Script:FNC = "Yellow";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC;$CollectFile = Read-Host "`t`t"
    # Name of old file to replace
    # FNExample1: Shared Files (LEC).lnk
    # FNExample2:  WorkStudio INFOCENTER.url
    # Set the value of $NPconfirm to "N" as default for "Shall I continue to fix all links at $Computer Y/N?"
    $global:NPconfirm = "N"
    if ($Choose2CopyFile -eq "Copy")
        {
        $NamedFiles = Read-Host "Enter a file name to remove. Ex: SharedFiles (S).lnk or *"
        # Location of new file to copy
        # Example: \\domain.local\dfs\Net Share\Folder
        # Folder where icons are pinned to the TaskBar
        #  ls "C:\Users\%username%\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
        $NAmedFNLocation = Read-Host "`nEnter the path to the new file to copy (exclude trailing \);
            `nExample: \\server\\share\folder"
        # Name of Source file
        $SourceCopyFN = Read-Host "Enter the source for the new file name to copy `nExample: SharedFiles (L).lnk" # "SharedFiles (L).lnk"
        if ((Read-Host "Current source Filename is $SourceCopyFN . Type C to enter a new name for the destination file or [Enter] to keep it the same ")-eq "C") 
            {$ReplaceWithFN = Read-Host "Type a filename to replace $SourceCopyFN"}
        ELSE
            {$ReplaceWithFN = $SourceCopyFN}
        # Concatonate source filename and location
        $copyFile = $NAmedFNLocation + "\" + $SourceCopyFN
        }
    if ($CollectFile)
        {
            $NamedFiles = Read-Host "Enter a file name to read. Ex: printers-"
            $CFdesk = Read-Host "Is the file on the desktop? Yes or No."
            # need to de-personalize this -OR- get the .\ or script folder to work
            $OutFile = "$PSScriptRoot\$NamedFiles"+"Results.csv"
            $AppendOverwrite = Read-Host "Type O to overwrite $OutFile`nAll other options will Append"
            Write-Warning "Collect Function Results will be saved in: $OutFile"
        }
    $ADorCSL = $null
    Get-PCList
        If (!$Global:pclist)
            {
                $opt1 = $null
                $Script:Fp1 = "That doesn't match the options type ";
                $Script:Fp2 = "yes ";
                $Script:Fp3 = "to run against this system only";
                $Script:FNC = "Yellow";chFCcolor $Script:Fp1 $Script:Fp2 $Script:Fp3 $Script:FNC;$opt1 = Read-Host "`t`t"
        
            If ($opt1 -eq "yes")
                {
                $proc = gwmi win32_process -Filter "Name = 'explorer.exe'"
                $pscn = $proc.pscomputername
                $pgo = $proc.GetOwner().user
                "pscn is $pscn"
                iterate-Profiles
                }
            Else
                {
                        $ErrorActionPreference = "continue"
                }
             }

    Write-Host "Logged In Computer`tUser" -ForegroundColor Yellow
    Write-Host "IPAddress `tMachine Name `tUser Logged on"
    foreach ($Global:PCline in $Global:PCList)
        {
            Identify-PCName
    $CompIP = $null
    $CompName = $null
    $PingOutCSV = "C:\Users\%username%\Documents\WindowsPowerShell\"+$Global:CDate+"PingOut.csv"
    $ErrorActionPreference = "silentlycontinue"
    $CompIP = ([system.net.dns]::GetHostAddresses($Global:PC)).IPAddressToString |where  {$_ -like "*.*.*.*"}
    If ($CompIP)
        {
            # Get ping results into object
            $ping = test-connection $Global:PC -Count 1 -Quiet -ErrorAction SilentlyContinue
            $PingRslt = New-Object -type PSObject -Property ([Ordered]@{
                Hostname = $Global:PC
                Online = [bool]($ping)
                PingDate = $Global:CDate
                })
            $PingRslt|select * | Export-Csv $PingOutCSV -NoTypeInformation -Append
            # Verify Computername, hostname, with IP
            # "$Computer IP is $CompIP"
            $CompName = [System.Net.Dns]::GetHostByAddress($CompIP).HostName
            $CN_GHBA = if($Global:PC){$Global:PC}Else{"nothing"}
            # Write-host "Computer is $Computer , CompName is $CompName" -BackgroundColor DarkYellow
            If ($Global:PC -notlike "*$Global:PC*")
                {
                    Write-host "Testing $Global:PC, but reverse lookup returns: $CN_GHBA" -ForegroundColor Red
                }
        }
        Else
        {
            $ping = test-connection $Global:PC -Count 1 -Quiet -ErrorAction SilentlyContinue
            $PingRslt = New-Object -type PSObject -Property ([Ordered]@{
                Hostname = $Global:PC
                Online = [bool]($ping)
                PingDate = $Global:CDate
                })
            $PingRslt|select * | Export-Csv $PingOutCSV -NoTypeInformation -Append
            $ping = $false
            $CompIP = "___.___.___.___"
            # Write-host "No IP $CompIP found for: for $Computer" -foreground Red
        }

  	if($ping -eq $true)
        {
        $pcinfo = @()
		#Get explorer.exe processes
		$proc = gwmi win32_process -computer $Global:PC -Filter "Name = 'explorer.exe'"

		#Search collection of processes for username
		ForEach ($p in $proc) 
            {
	    	    # $usr = ($p.GetOwner()).User
                $pscn = $p.pscomputername
                $pgo = $p.GetOwner().user
                Write-Output "$CompIP`t$pscn`t$pgo`t" |ft
            
                if ($Choose2CopyFile -eq "copy" -or $ReplaceLink -eq "dlink" -or $ReplaceLink -eq "plink" -or $CollectFile -eq "collect") # If the user previously confirmed he wants to copy by typing the word "copy" then run the following function
                    {
                        iterate-Profiles
                    }
                if ($Opt2 -eq "map") # If the user previously confirmed he wants to see mapped drives by typing the word "map" then run the following function
                    {
                        List-MappedDrives
                    }            
            # ***********Return************
           
		    }
        }
    else
        {
            Write-Host "$CompIP`t$Global:PC`tUNK/NONE" -ForegroundColor Red
        }
    }
        $ErrorActionPreference = "continue"
    if ($CollectFile){Write-Warning "Collect Function Results has been saved in: $OutFile"}
}

Function get-OpenFiles()
{
    Clear-Host
    Get-PCList
    foreach ($Global:PCLine in $Global:PCList)
    {
        Identify-PCName # possibly need to bring this function into this script if it won't run from get-systemslist.ps1
        $srvobj = invoke-command {c:\Windows\System32\openfiles.exe /query /s $Global:PC /fo csv /V} |convertfrom-csv |
        select Hostname,ID,Type,@{Name="UserName"; Expression="Accessed By"},@{Name="Path_FileName"; Expression="Open File (Path\executable)"} |
        Where {$_.Path_FileName -ne "D:\\"}|
        sort Path_FileName | Out-GridView # ft -AutoSize -Wrap
        $srvobj
    }
}

# Get-Shortcut called from Replace-DriveLetterLinks function
Function Get-Shortcut {
  param(
    $path = $null
  )
  # https://stackoverflow.com/questions/484560/editing-shortcut-lnk-properties-with-powershell
  $obj = New-Object -ComObject WScript.Shell

  if ($path -eq $null) {
    $pathUser = [System.Environment]::GetFolderPath('StartMenu')
    $pathCommon = $obj.SpecialFolders.Item('AllUsersStartMenu')
    $path = dir $pathUser, $pathCommon -Filter *.lnk -Recurse 
  }
  if ($path -is [string]) {
    $path = dir $path -Filter *.lnk
  }
  $path | ForEach-Object { 
    if ($_ -is [string]) {
      $_ = dir $_ -Filter *.lnk
    }
    if ($_) {
      $link = $obj.CreateShortcut($_.FullName)
      $link |select *
      $info = @{}
      $info.Hotkey = $link.Hotkey
      $info.TargetPath = $link.TargetPath
      $info.LinkPath = $link.FullName
      $info.Arguments = $link.Arguments
      $info.Target = try {Split-Path $info.TargetPath -Leaf } catch { 'n/a'}
      $info.Link = try { Split-Path $info.LinkPath -Leaf } catch { 'n/a'}
      $info.WindowStyle = $link.WindowStyle
      $info.IconLocation = $link.IconLocation
      $info.WorkingDirectory = try {Split-Path $info.WorkingDirectory -Leaf } catch { 'n/a'}

      New-Object PSObject -Property $info
    }
  }
}
Function Set-Shortcut {
  param(
  [Parameter(ValueFromPipelineByPropertyName=$true)]
  $LinkPath,
  $Hotkey,
  $IconLocation,
  $Arguments,
  $TargetPath,
  $WorkingDirectory
  )
    # https://stackoverflow.com/questions/484560/editing-shortcut-lnk-properties-with-powershell
  begin {
    $shell = New-Object -ComObject WScript.Shell
  }

  process {
    $link = $shell.CreateShortcut($LinkPath)

    $PSCmdlet.MyInvocation.BoundParameters.GetEnumerator() |
      Where-Object { $_.key -ne 'LinkPath' } |
      ForEach-Object { $link.$($_.key) = $_.value }
    $link.Save()
  }
}
Function Replace-DriveLetterLinks()
{
<#
.EXAMPLE
    Use the following example to assign F11 hotkey to PowerShell (uncomment line, make sure you have full admin privileges: 
    Get-Shortcut | Where-Object { $_.LinkPath -like '*Windows PowerShell.lnk' } | Set-Shortcut -Hotkey F11 
#>
param(
  [Parameter(ValueFromPipelineByPropertyName=$true)]
  $PathToLink,
  $value0 = (Read-Host "Enter new target Drive Letter in the following format L:\"),
  $DLValue = (Read-Host "Enter old target Drive Letter in the following format O:\")
  )

        # Set the WorkingDirectory driveletter to the same driveletter as the new target drive letter
        $WDvalue0 = $value0
    $DLCSVOut = "C:\Users\%username%\Documents\WindowsPowerShell\"+$Global:CDate+"DesktopLinks.csv"
    $LinksFound = ls $PathToLink *.lnk -Recurse # \\PCName\C$\Users\ProfileName\Desktop
    # Write-warning "Writing CSV to $DLCSVOut"
    $LinksFound |select DirectoryName, Name |export-csv -Path $DLCSVOut -Append -NoTypeInformation # select Directory, DirectoryName, FullName, Name, BaseName 
    foreach ($lf in $LinksFound)
    {
        $source = $lf.fullname
        $TP_GSc_qry = $DLValue+"*"
        $GSc = Get-shortcut $source |select FullName, RelativePath,TargetPath,WorkingDirectory |where {$_.TargetPath -like "$TP_GSc_qry" -and $_.WorkingDirectory -ne "n/a"}
        $LSc = Get-shortcut $source |select FullName, RelativePath,TargetPath,WorkingDirectory |where {$_.WorkingDirectory -ne "n/a"}
        $alldsktpLnkPropertiesOut = "C:\Users\Public\Documents\"+$Global:CDate+"AllDsktpLnkPropertiesOut.csv"
        $LSc |select * |export-csv -Path $alldsktpLnkPropertiesOut -Append -NoTypeInformation # -Verbose
        # $desktoplnkPropertiesOut = "C:\Users\Public\Documents\"+$Global:CDate+"DesktopLinkProperties.csv"
        $desktoplnkPropertiesOut = $Script:PSScriptRoot+"\"+$Global:CDate+"DesktopLinkProperties.csv"
        $GSc |select * |export-csv -Path $desktoplnkPropertiesOut -Append -NoTypeInformation # -Verbose
        # Zero out variables
        $GSclp = $null
        $GSctp = $null
        $GScwd = $null
        $NewTargetPath = $null
        $NewWDPath = $null
        if ($GSc)
        {
            # Get LinkPath to variable $GSclp
            $GSclp = $GSc.FullName
            # Get TargetPath to variable $GSctp
            $GSctp = $GSC.TargetPath
            # Get WorkingDirectory to variable $GScwd
            $GScwd = $GSC.WorkingDirectory
            # Create a regex pattern to break the TargetPath into two groups: Drive and Path This is currently case sensitive
            <#This line workd, testing the line below with variable $DLValue#> #$Pattern1 = '^(S:\\)(.*?)$' #'^($DLValue\\)(.*?)$' (w/UNC: '^(\\\\server\\Net\sShare\\)(.*?) *$'
            $Pattern1 = "^($DLValue\)(.*?)$" # (w/UNC: '^(\\\\server\\Shared\sFiles\\)(.*?) *$'
            # $Pattern1 = '^(\\\\server\\share\sname\\)(.*?) *$'
            # $Pattern1 = '^(\\\\SERVER\\Share\sName\\)(.*?) *$'
            # Match the drive letter and path into two separate variables
            
            $value1 = [regex]::Matches($GSctp, $Pattern1).Groups[1].Value # This is case sensitive using -match is case insensitive
            $value2 = [regex]::Matches($GSctp, $Pattern1).Groups[2].Value # This is case sensitive using -match is case insensitive
            
            <#
            # If I can tweak this style of regex, it would be case insensitive
            $value1 = ($GSctp -match $Pattern1).Groups[1].Value
            $value2 = ($GSctp -match $Pattern1).Groups[2].Value
            #>
            # Display the values showing the results of the Regex matches
            Write-warning "Re: $source"
            Write-warning "0(NewDrive TargetPath) = $value0"
            Write-warning "1(Matches Group1) = $value1"
            Write-warning "2(Matches Group2)= $value2"
            
            if ($GScwd)
            {
                $WDvalue1 = [regex]::Matches($GScwd, $Pattern1).Groups[1].Value # This is case sensitive using -match is case insensitive
                $WDvalue2 = [regex]::Matches($GScwd, $Pattern1).Groups[2].Value # This is case sensitive using -match is case insensitive
                <#
                # Display the values showing the results of the Regex matches
                $WDvalue0
                $WDvalue1
                $WDvalue2
                #>
            }
            if ($value1 -eq $DLValue) #"S:\")
            {
                $NewTargetPath = $value0+$value2
                $NewTargetPath
                if ($GScwd)
                {
                    $NewWDPath = $WDvalue0+$WDvalue2
                    $NewWDPath
                }
                Write-host "The new path on $PSCN is:`n`t $NewTargetPath `nand the WorkingDirectory is `n`t$NewWDPath" -ForegroundColor Green
                $lnkPropertiesNewOld = $Script:PSScriptRoot+"\"+$Global:CDate+"lnkPropertiesNewOld.csv"
                $GSc |select * |export-csv -Path $lnkPropertiesNewOld -Append -NoTypeInformation # -Verbose
                If ($Global:NPconfirm -eq "Y" -or "A") # NPconfirm is not N or None
                {
                    Set-Shortcut -TargetPath $NewTargetPath -LinkPath $GSclp -WorkingDirectory $NewWDPath
                }
                else # NPconfirm is N or None so DO NOTHING
                {
                    Write-Warning "I am not replacing $GSctp with $NewTargetPath"
                }
            }
        }
    }
}

Function iterate-Profiles() # Requires permission to view C$ shares on all selected systems
{
    # from calling function: $proc = gwmi win32_process -computer $pscn -Filter "Name = 'explorer.exe'"
    if ($Choose2CopyFile -eq "copy" -or $ReplaceLink -eq "dlink")
    {
        # make $userdir variable equal to current user SamAccountName
        $userdir = $pgo
        # Set $ProfileList to contain all contents of a systems Users folder (Usually just a list of Profile folder names)
        $ProfileList = gci "\\$pscn\C$\Users\" -Directory -Force
    }
    if ($CollectFile -eq "collect")
    {
        # make $userdir variable equal to current user SamAccountName
        $userdir = $pgo
        # Set $ProfileList to contain all contents of a systems Users folder (Usually just a list of Profile folder names)
        $ProfileList = gci "\\$pscn\C$\Users\" -Directory -Force
        if ($AppendOverwrite -eq "O")
            {
                $sw = New-Object -TypeName System.IO.StreamWriter -ArgumentList $OutFile
                # Change value of $AppendOverwrite it so that it would overwrite initially and then append from then on.
                $AppendOverwrite = "NoMoreO"
            }
        else
            {
                $sw = New-Object -TypeName System.IO.StreamWriter -ArgumentList $OutFile ,'Append' #$CollectFile Out
            }
        if ($CFdesk -eq "Yes"){$sw.WriteLine("Workstation,FileFound")} # $CollectFile Out header **Incomplete**
        elseif ($NamedFiles -eq "Printers-") {$sw.WriteLine("Workstation,Name,Location,Printer,PrinterState,PrinterStatus,ShareName,SystemName")}#$CollectFile Out header
    }
    if ($ReplaceLink -eq "dlink" -or $ReplaceLink -eq "plink")
    {
        $value0 = (Read-Host "Enter new target Drive Letter in the following format L:\") # "L:\" # Change this into a = Read-Host "Enter new target Drive Letter in the following format L:\"
        $DLValue = (Read-Host "Enter old target Drive Letter in the following format O:\") # "S:\" # Change this into a = Read-Host "Enter old target Drive Letter in the following format O:\"
    }
    # comment out this foreach and the next $userdir variable declaration to only apply the following action to the currently loged on user
    if ($ReplaceLink -eq "dlink")
    {
        # Should create a do-dlink function with the rest of this if statement and loop back to it on the NPconfirm -ne "A"
        if ($Global:NPconfirm -eq "Y" -or $Global:NPconfirm -eq "N")
        {
            $Global:NPconfirm = Read-Host "Type A to fix all, Y to fix next, N to NOT fix next, or None to NOT fix any."
        }
        Elseif ($Global:NPconfirm -eq "None")
        {
            # No Need to do anything here
            # Is there another place that we need to test for None rather than just N?
        }
        Elseif ($Global:NPconfirm -ne "A")
        {
            Write-Host "Your options are A,Y, or N. Please try again"
        }
    }
    if ($ReplaceLink -eq "plink")
    {
        $PathToLink = "\\$pscn\C$\Users" # $PathToLink = Read-Host "Type the root of the path to search. Do not end in \.`nUse a local drive letter or UNC path."
        Replace-DriveLetterLinks $PathToLink $value0 $DLValue
    }
    foreach ($PL in $ProfileList)
    {
    $userdir = $PL.Name
    $PathToLink = "\\$pscn\C$\Users\$userdir\Desktop"
        if ($ReplaceLink -eq "dlink")
            {
                 Replace-DriveLetterLinks $PathToLink $value0 $DLValue
            }
        if ($Choose2CopyFile -eq "copy") # If the user previously confirmed he wants to copy by typing the word "copy" then run the following function
            {
                # Don't copy to the following profiles
                if ($userdir -ne "Public" -and $userdir -ne "All Users" -and $userdir -ne "Default User" -and $userdir -ne "Default" -and $userdir -ne "Default.Migrated")
                {
                    replace-DesktopFiles
                }
            }
        # Process the contents of a text file on a system. Output results to a text file.
        if ($CollectFile -eq "collect")
            {
                # Currently set to a static location. This could be automated
                if ($CFdesk -eq "Yes")
                {
                    $P2File = "\\$pscn\C$\Users\$userdir\Desktop\" + $NamedFiles
                    if (Test-Path $P2File)
                    {
                        Write-Output "Found $P2File"
                        Write-Warning "AppendOverwrite $AppendOverwrite"
                        $content = New-Object -TypeName System.IO.StreamReader -ArgumentList $P2File
                        #$FL = GC $P2File -first 1
                        $FL = GC $P2File
                        # Write-Warning "FL"
                        # $FL
                        $FLObjectOut = foreach ($FLLine in $FL)
                        {
                            # "$FLLine"
                            $FSplit0 = ($FLLine.Split("="))[0]
                            $FSplit1 = ($FLLine.Split("="))[1]
                            $FLOBProperties = @{'pscomputername'=$pscn;
                                    'FSplit0'=$FSplit0;
                                    'FSplit1'=$FSplit1;
                                    'UserDir'=$userdir;
                                    'DateChanged'=$CDAte}
                            New-Object -TypeName PSObject -Prop $FLOBProperties
                        }
                    }
                }
                if ($NamedFiles -eq "printers-"){$P2File = "\\$pscn\C$\Users\Public\Downloads\" + $NamedFiles + $userdir + ".txt"
                    if (Test-Path $P2File)
                    {
                        Write-Output "Found $P2File"
                        Write-Warning "AppendOverwrite $AppendOverwrite"
                        $content = New-Object -TypeName System.IO.StreamReader -ArgumentList $P2File
                        $FL = GC $P2File -first 1
                        # Comment out the belowpattern you do not want otherwise, the second pattern will succeed
                            $Pattern1 = '^([A-Z][a-z]*\s*)([A-Z][a-z]*\s*)([A-Z][a-z]*\s*)([A-Z][a-z]*\s*)([A-Z][a-z]*\s*)([A-Z][a-z]*\s*)'
                            $Pattern1 = '^(Location*\s*)(Name*\s*)(PrinterState*\s*)(PrinterStatus*\s*)(ShareName*\s*)(SystemName*\s*)'
                        # Comment out the above pattern you do not want otherwise, the second pattern will succeed
                        $C1Width = ($ColLocationWidth = [regex]::Matches($FL, $Pattern1).Groups[1].Value).length
                        $C2Width = ($ColNameWidth = [regex]::Matches($FL, $Pattern1).Groups[2].Value).length
                        $C3Width = ($ColPrinterStateWidth = [regex]::Matches($FL, $Pattern1).Groups[3].Value).length
                        $C4Width = ($ColPrinterStatusWidth = [regex]::Matches($FL, $Pattern1).Groups[4].Value).length
                        $C5Width = ($ColShareNameWidth = [regex]::Matches($FL, $Pattern1).Groups[5].Value).length
                        $C6Width = ($ColSystemNameWidth = [regex]::Matches($FL, $Pattern1).Groups[6].Value).length
                        $0to1 = $C1Width
                        $0to2 = $0to1 + $C2Width
                        $0to3 = $0to2 + $C3Width
                        $0to4 = $0to3 + $C4Width
                        $0to5 = $0to4 + $C5Width
                        $0to6 = $0to5 + $C6Width
                        while($null -ne ($line = $content.ReadLine())) {
                            $line
                            $Location = ($line.SubString(0,$C1Width).trim())
                            $Location
                            $sw.Write($pscn);$sw.Write(",")
                            $sw.Write($userdir);$sw.Write(",")
                            if ($Location){$sw.Write($line.SubString(0,$C1Width).trim())}
                            else {$sw.Write("Local")};$sw.Write(",")
                                $sw.Write($line.SubString($0to1,$C2Width).trim());$sw.Write(",")
                                $sw.Write($line.SubString($0to2,$C3Width).trim());$sw.Write(",")
                                $sw.Write($line.SubString($0to3,$C4Width).trim());$sw.Write(",")
                                $sw.Write($line.SubString($0to4,$C5Width).trim());$sw.Write(",")
                                $sw.Write($line.SubString($0to5,$C6Width).trim());$sw.Write(",")
                                $sw.Write($line.SubString($0to6,$C7Width).trim());$sw.Write("`n")
        
                        }
                    }
                }
            }
    }
Write-output "GC OutFile: $OutFile"
Write-Warning "`nStart Contents of OutFile`n"
Get-content $OutFile # Collect OutFile
Write-Warning "`End Contents of OutFile`n"
$sw.close() # Close Collect OutFile
$FLObjectOut|where {$_.FSplit0 -eq "URL"} |Select pscomputername,UserDir,FSplit0,FSplit1 |Export-Csv -Append $OutFile -NoTypeInformation
Write-Warning "Start Contents of OutFile"
Get-content $OutFile # Collect OutFile
Write-Warning "End Contents of OutFile"
}


# Function replaces / copies a file to a folder on specified systems
# This can be designed to be called from different areas.
# It is currently called from function iterate-Profiles via UsrsOnPCs to run against certain active PCs
Function Replace-DesktopFiles()
{
    $ciResult = gci $PathToLink |where {$_.Name -like $NamedFiles} |select Name, FullName, LastWriteTime, @{Name="ComputerName"; Expression={$pscn}} # look for deleting shared files link "SharedFiles (S)"
    # Write-warning "CiResult is: $ciResult"
    # $ciResult = gci $PathToLink |where {$_.Name -like "*share*.lnk"} |select Name,LastWriteTime, @{Name="ComputerName"; Expression={$pscn}} # look for deleting shared files link
    # $ciResult = gci $PathToLink -ErrorAction SilentlyContinue |where {$_.Name -like "*INFOCENTER*.url"} |select Name,LastWriteTime, @{Name="ComputerName"; Expression={$pscn}} # look for infocenter.url link
    if ($ciResult -ne $null)
    {
        Write-Output "$CiResult" # -WarningVariable a ;$a
        foreach($cr in $ciResult)
        {
            # Determine Path and Filename to delete
            $fn = $cr.name 
            $rmfile = $PathToLink +"\"+$fn 
            # uncomment to delete the old file found first prior to copying its replacement.
            del $rmfile # -Confirm
            # Export a list of all shortcuts marked for replacement
            $CSVName = ".\_"+$Global:CDate+"DesktopLinks.csv"
            $cr |Export-Csv $CSVName -NoTypeInformation -Append
        }
    }
    # copy $copyFile to new location on $desktop with destination filename: $ReplaceWithFN

    Write-Warning "Copying $copyFile to $PathToLink\$ReplaceWithFN"
    copy $copyFile "$PathToLink\$ReplaceWithFN" -ErrorAction Inquire -ErrorVariable ER # -ErrorAction SilentlyContinue 
    # $cr = $error[0].exception.message
    $CSVName = ".\_"+$Global:CDate+"DesktopLinks.csv"
    # if ($cr -ne "")
    Write-Warning "ER $ER"
    if ($ER -ne "")
    {
        $ER|Out-File $CSVName -Append -NoClobber
        # Write-Output "Was trying to copy $copyFile to $PathToLink\$ReplaceWithFN" |Out-File $CSVName -Append -NoClobber
        $ER = ""    
    }
    Else
    {
        Write-Output "Copied $copyFile to $PathToLink\$ReplaceWithFN" |Out-File $CSVName -Append -NoClobber
    }
    Write-host "$CSVName" -ForegroundColor Green


}

Function Compare-Files
{
$BaseDirectory = Read-Host "Enter the root of the path to search for duplicate files"
$wpsdir = gci $BaseDirectory -Recurse
$ObjectOut = foreach ($wpsdir1 in $wpsdir.fullname)
{
        $GFH = Get-FileHash $wpsdir1 -Algorithm SHA1 -Verbose
        if ($GFH.Hash -ne $null)
        {
            $GFH |Select *
            $FilePath = $GFH.Path
            $FileHash = $GFH.Hash
            $FileAlgorithm = $GFH.Algorithm
            $OBProperties = @{'FilePath'=$FilePath;
                        'FileHash'=$FileHash;
                        'Algorithm'=$FileAlgorithm}
                New-Object -TypeName PSObject -Prop $OBProperties
        }
}

$ObjectOutNonDups = $ObjectOut.FileHash |select -Unique
$ObjectOut|Where {$_.FileHash -ne $null} |sort FileHash |Group-Object FileHash |? {$_.count -gt 1} | %{$_.Group} |ft

$ObCSVFile = $PSScriptRoot+"\"+$Global:CDate+"_Compare_Files.csv"
$ObjectOut|Where {$_.FileHash -ne $null} |sort FileHash |Group-Object FileHash |? {$_.count -gt 1} | %{$_.Group} |Export-Csv $ObCSVFile -NoClobber -NoTypeInformation -Append
Write-Host "All lines: " -NoNewline
    $ObjectOut.Count
Write-Host "Non-Duplicate lines: " -NoNewline
    $ObjectOutNonDups.Count

# $ObjectOut |select FileHash, Algorithm, FilePath
}

Function run-get-filehash()
{
    # View the following for ideas
    # https://docs.microsoft.com/en-us/powershell/module/Microsoft.PowerShell.Utility/Get-FileHash?view=powershell-5.1
    Clear-Host
    $ComparePrompt = $null
    $FullFileName = Read-Host "`nType the path and filename for the file for which you wish to view the hash`n Example: C:\Users\Public\Downloads\Software\Tools\Java\jdk-10.0.1_windows-x64_bin.exe`n"
    $HashAlgorithm = Read-Host "`nEnter the hashing algorithm you wish to use. `n`tChoose from SHA1 | SHA256 | SHA384 | SHA512 | MACTripleDES | MD5 | RIPEMD160"
    $ComparePrompt = Read-Host "Type or Paste a hash/checksum here to compare it with what is found"
    # Use of Get-FileHash: Get-FileHash [-Path] <String[]> [-Algorithm {SHA1 | SHA256 | SHA384 | SHA512 | MACTripleDES | MD5 | RIPEMD160}]
    Write-Warning "Running the following command:`n`nget-filehash -Path $FullFileName -Algorithm $HashAlgorithm -Verbose`n"
    $FFNHash = get-filehash -Path $FullFileName -Algorithm $HashAlgorithm -Verbose
    if ($ComparePrompt -eq $FFNHash.Hash)
        {Write-Host "The hashes match" -ForegroundColor Green
        $FFNHash |select * |FT -autosize -wrap
        }
    else
        {
        $FFNHash |select * |FT -autosize -wrap
        }
}


Function Get-ModifiedFileList()
{
    Clear-Host
    # Set up Variables
        $Dir2Scan = "" # Dir2Scan is the folder that will be the root folder for the search$DaySpan = $null # DaySpan is the number of days since a file was modified
        $Ext2Scan = $null # Ext2Scan is a specific file extension to filter by
        $FFilter = $null # FFilter uses a typical filter syntax to find a file
        $tdir = "$env:TMP"+"\" # This is the user's temporary folder - change it to ".\" or PWD for current folder
    # Prompt for optional search criteria
        $DaySpan = Read-Host "Enter the number of days of modified files to list"
        # Set number of days to 100 years if nothing was entered
        If($DaySpan -eq ""){$DaySpan = "36500"}
        $Dir2Scan = Read-Host "Hit [Enter] to scan \\FileServer\SharedFiles share on server or Enter an alternate path"
    # Set a default Dir2Scan value if nothing was entered in the prompt
        If($Dir2Scan -eq ""){$Dir2Scan = "\\FileServer\SharedFiles"}
        $Ext2Scan = Read-Host "`nType an extension to restrict search"
        $FFilter = Read-Host "`nEnter a file search filter if necessary"
    # Inform user of selections made
        Write-Warning "`nSearching $Dir2Scan for files up to $DaySpan day(s) old."
    # $ExCSVFileName = ((Get-date -Format "yyyy_M_d_hms") + "_Files.csv") #Use this option if you want the file saved in the current script folder
        $ExCSVFileName = "$tdir" + ((Get-date -Format "yyyy_M_d_hms") + "_Files.csv")
        "`n"
    # Run different searches depending on whether a file age in days was declared
    Get-ChildItem -Filter $FFilter -Recurse $Dir2Scan -ErrorAction SilentlyContinue |Where-Object {$_.Extension -like "*$Ext2Scan"}|
        Where-Object {$_.LastWriteTime -gt (get-date).AddDays(-$DaySpan)}|
        ForEach-Object {$_ |add-member -name "Owner" -MemberType NoteProperty -Value (get-acl $_.FullName).Owner -PassThru}|
        Sort-Object fullname |select Fullname,Name,psdrive,Directory,Extension,CreationTime,LastWriteTime,Length,Owner|
        Export-Csv $ExCSVFileName -Force -NoTypeInformation 
    $YNPrompt = Read-Host "Output at $ExCSVFileName is temporary. `nBe certain to save it elsewhere if you want a permanent copy.`nDo you wish to open it now in Excel? - Y or N"
    if($YNPrompt -eq "Y"){Start-Process -wait $ExCSVFileName}else{Import-Csv $ExCSVFileName|Out-GridView -Title $ExCSVFileName}# Using Start-Process rather than Invoke-Item because it has a -wait feature
    Start-Sleep -m 250
    # Delete the CSV temp file when user exits out of menu
    # File will not be deleted if it is still open at this time
    Remove-Item $ExCSVFileName -Force
}
Function Manage-NTFS
{
<#
.NAME
    Manage-NTFS
.SYNOPSIS
    Change NTFS permissions and perform actions to target file/folder
.DESCRIPTION
     Change NTFS permissions so that the user will have permissions to 
     delete a target file/folder. This could be modified to prompt for and perform other actions
.PARAMETERS
    $NewOwn is the name of the target owner 
    $NewPermUser is the name of the target user to receive full permissions
    $Targetfolders is the name of the Folder(s) (wild-card accepted) to which the permission change will be applied
.EXAMPLE
    Manage-NTFS -NewOwn domain\NewOwnerName -NewPermUser domain\NewPermUserName -TargetFolders C:\TEMP - Copy*
.SYNTAX
    Manage-NTFS [-NewOwn] <string[]> [-NewPermUser] <string[]> [-TargetFolders] <string[]>
.REMARKS
    NTFSSecurity module can be downloaded from the following site:
    https://gallery.technet.microsoft.com/scriptcenter/1abd77a5-9c0b-4a2b-acef-90dbb2b84e85
    Instructions for the module's use are found at:
    https://blogs.technet.microsoft.com/fieldcoding/2014/12/05/ntfssecurity-tutorial-1-getting-adding-and-removing-permissions/
    See >help *ntfs* for other cmdlets in the ntfssecurity module
#>
[CmdletBinding()] #(allows the use of parameter attributes)
    param
    (
            #[Parameter(Mandatory=$True)]
            $NewOwn =(Read-Host "Enter a valid new Owner dom\UserName"), 
            $NewPermUser = (Read-Host "Enter a valid new User for full permissions dom\UserName"),
            $TargetFolders = (Read-Host "Enter a valid Target Folder like: C:\TEMP - Copy*" )
    )
Import-Module ntfssecurity
$folders = ls $Targetfolders # $Targetfolders # List direcotry to variable If user doesn't have list permissions this ls may fail
$FStart = $folders.Count
Write-Warning
    foreach ($folder in $folders)
    {
        # Get-NTFSOwner $folder
        $folder | Get-NTFSOwner
        # Change $folder ownership
        # Set-NTFSOwner -path $folder -Account $NewOwn
        $folder | Set-NTFSOwner -Account $NewOwn
        # Add FullControl to $folder for the $NewPermUser user
        # Get-Item $folder|Add-NTFSAccess -Account $NewPermUser -AccessRights FullControl
        $folder |Get-Item|Add-NTFSAccess -Account $NewPermUser -AccessRights FullControl 
        # Remove the $folder directory and all contents
        rmdir $folder -Recurse -Confirm
    }
$FEnd = ls $Targetfolders 
$dif = $FEnd.Count
Write-Warning "Out of $FStart Objects, only $dif remain."
}

Function Browse-Share
{
param(
  [Parameter(ValueFromPipelineByPropertyName=$true)]
  [string[]] $ShareName = ($ShareName=if ($SN = Read-Host -Prompt "Please enter a UNC server and sharename ('\\$env:USERDNSDOMAIN\dfs\SharedFiles')") {$SN} else {"\\$env:USERDNSDOMAIN\dfs\SharedFiles"})
  )
    $gPd = Get-PSDrive|where {$_.Root -like "?:\"} |select Name
    $gpdcsvList = $gpd.Name -join ", "
    $NewDL = Read-Host "Enter a drive letter to map to $ShareName (other than $gpdcsvList)"
    $NewPath = Read-Host "The path to the location you would like to list starting with \"
    New-PSDrive -Name "$NewDL" -PSProvider "FileSystem" -Root "$ShareName" -Credential $cred
    Get-PSDrive |ft -AutoSize
    $NewDLPath = "$NewDL"+":"
    $DL = Get-ChildItem $NewDLPath
    $DL = Get-ChildItem ("$NewDLPath"+"$NewPath")
    $DL |select PSDrive,CreationTime,FullName,LastAccessTime,Mode <#| Sort Mode -Descending -property @{E="FullName";Descending=$false} @{Expression="Mode";Descending=$true}|sort FullName #>|ft -AutoSize -Wrap # Root,Parent,PSChildName,
    Remove-PSDrive -Name "$NewDL"
    
}
