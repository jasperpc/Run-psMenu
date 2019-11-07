<#
.NAME
    Create-DMARCReport.ps1
.SYNOPSIS
    Read a DMARC .zip or .gz file extract the .xml and produce a human-readable report
.DESCRIPTION
 2019-11-07
 Zip handling written by https://techibee.com/powershell/reading-zip-file-contents-without-extraction-using-powershell/2152
 Written by Jason Crockett - Laclede Electric Cooperative
.PARAMETERS
    $ExportCSVFileName
.EXAMPLE
    .\Create-DMARCReport.ps1 []
.SYNTAX
    .\Create-DMARCReport.ps1 []
.REMARKS
    To see the examples, type: help Create-DMARCReport.ps1 -examples
    To see more information, type: help Create-DMARCReport.ps1 -detailed
.TODO
    Change the line determining $PCReadFolder if necessary to change the default folder
    Change the line determining $PCReadFileEXT if not reading .XML files in the compressed file
    The .zip files may not close correctly and the process may need to be changed.
#>
[cmdletbinding()]            
param(            
 [Parameter(Mandatory=$false)] # $true
 # [string[]]$FileName,
 [String]$ExportCSVFileName            
)
[datetime]$epoch = '1970-01-01 00:00:00'
#Exit if the shell is using lower version of dotnet            
$dotnetversion = [Environment]::Version            
if(!($dotnetversion.Major -ge 4 -and $dotnetversion.Build -ge 30319)) {            
 write-error "You do not have Microsoft DotNet Framework 4.5 installed. Script exiting"            
 exit(1)            
}            


function Read-FileInZip($ZipFilePath, $FilePathInZip) {
    # https://stackoverflow.com/questions/37561465/how-to-read-contents-of-a-csv-file-inside-zip-file-using-powershell
    Add-Type -assembly "System.IO.Compression.FileSystem"
    try {
        if (![System.IO.File]::Exists($ZipFilePath)) {
            throw "Zip file ""$ZipFilePath"" not found."
        }

        $Zip = [System.IO.Compression.ZipFile]::OpenRead($ZipFilePath)
        $ZipEntries = [array]($Zip.Entries | where-object {
                return $_.FullName -eq $FilePathInZip
            });
        if (!$ZipEntries -or $ZipEntries.Length -lt 1) {
            throw "File ""$FilePathInZip"" couldn't be found in zip ""$ZipFilePath""."
        }
        if (!$ZipEntries -or $ZipEntries.Length -gt 1) {
            throw "More than one file ""$FilePathInZip"" found in zip ""$ZipFilePath""."
        }

        $ZipStream = $ZipEntries[0].Open()

        $Reader = [System.IO.StreamReader]::new($ZipStream)
        return $Reader.ReadToEnd()
    }
    finally {
        if ($Reader) { $Reader.Close() }
        # These dispose lines are supposed to close the file but I still can't move the .zip file afterwards
        if ($Reader) { $Reader.Dispose() }
        if ($Zip) {$Zip.Dispose()}
    }
}
    
Clear-Host
# Import dotnet libraries
[Void][Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')
$ObjArray = @()
$TR = $null
$HR = $null
$R1 = $null
$CDate = Get-Date -UFormat "%Y%m%d"
$MoveFilePrompt = "No"
$Global:CDateTime = [datetime]::ParseExact($Global:CDate,'yyyymmdd',$null)
$Script:PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
# if (($PCReadFolder = Read-Host "Enter a folder name containing the list of target files. [Enter] to use default: C:\Users\Public\Downloads\Security\Web_Email\New\") -eq "") {$PCReadFolder = "C:\Users\Public\Downloads\Security\Web_Email\New\"}
$PCReadFolder = "C:\Users\Public\Downloads\Security\Web_Email\New\" # Comment this line and anable the if (($PCReadFolder line to allow for user prompt
# if (($PCWriteFolder = Read-Host "Enter the folder name Where you wish to move the target files after processing. [Enter] to use default: C:\Users\Public\Downloads\Security\Web_Email\Old\") -eq "") {$PCReadFolder = "C:\Users\Public\Downloads\Security\Web_Email\Old\"}
$PCWriteFolder = "C:\Users\Public\Downloads\Security\Web_Email\Old\" # Comment this line and anable the if (($PCReadFolder line to allow for user prompt
# if (($PCReadFileEXT = Read-Host "Enter a file extension to select the target files. [Enter] to use default: .xml") -eq "") {$PCReadFileEXT = ".xml"}
$PCReadFileEXT = ".xml" # Comment this line and enable the if (($PCReadFileEXT line to allow for user prompt
Write-Output "`n"
Write-Warning "Folder is: $PCReadFolder and final extension to parse is: $PCReadFileEXT"
$Folder = $PCReadFolder |gci |sort LastWriteTime # GCI "C:\Users\Public\Downloads\Security\Web_Email\New*"
$WorkingFiles = ($Folder |where {$_.name -like "*.zip" -or $_.name -like "*.xml" -or $_.name -like "*.gz"})
if (($MoveFilePrompt = Read-Host "Type [Enter] move files from New to Old folder") -eq "") {$MoveFilePrompt = "Yes"}
foreach($WorkingFileLine in $WorkingFiles) 
{            
    $WorkingFileLineLastWriteTime = $WorkingFileLine.LastWriteTime
    $WFName = $WorkingFileLine.Name
    $WorkingFileLine = $WorkingFileLine.FullName
    # $WorkingFileLine |gm
    $WFL_LWT = $WorkingFileLine.LastWriteTime
    if(Test-Path $WorkingFileLine) 
    {            
        if ((ls $WorkingFileLine).extension -eq ".zip")
        {
            $RawFiles = [IO.Compression.ZipFile]::OpenRead($WorkingFileLine).Entries
            $Fileextn = [IO.Path]::GetExtension($WorkingFileLine)
        }
        elseif ((ls $WorkingFileLine).extension -eq ".gz")
        {
            $RawFiles =$WorkingFileLine
            $Fileextn = [IO.Path]::GetExtension($WorkingFileLine)
        }
        elseif ((ls $WorkingFileLine).extension -eq ".xml")
        {
            $RawFiles = $WorkingFileLine
            $Fileextn = [IO.Path]::GetExtension($WorkingFileLine)
        }
        # Get timestamps from zip files and process into human-readable format
        $TC = try{
            $Pattern1 = '.*\!(\d{10}\!\d{10}).*' # Select the coded timestamps with the pattern !, 10 digits, !, 10 digits
            $DateCode = [regex]::Matches($WorkingFileLine, $Pattern1).Groups[1].Value
            [Array]$DateCode = $DateCode.split("!") |%{$_.trim()}
            [int]$Date1 = $DateCode[0]
                [datetime]$Date1 = $epoch.AddSeconds($date1)
            [int]$Date2 = $DateCode[1]
                [datetime]$Date2 = $epoch.AddSeconds($date2)
            } catch{
                Write-Warning "Error with $Pattern in $WorkingFileLine"
                [int]$Date1 = $null# [datetime]$epoch
                [string]$Date2 = ([datetime]$WorkingFileLineLastWriteTime).ToString("MM/dd/yyyy")
            }
        if ($TC){"Unable to get dates from filename: $TC"}
        $DateCode = $null
        foreach($RawFile in $RawFiles) 
        {
            if ($Fileextn -eq $PCReadFileEXT)
            {
                [xml]$PCReadXMLFile =  Get-Content $WorkingFileLIne
            }
            Elseif ($Fileextn -eq ".zip")
            {
                [xml]$PCReadXMLFile = Read-FileInZip $WorkingFileLine $RawFile.FullName
            }
            Elseif ($Fileextn -eq ".gz")
            {
                # Need to do a test-path for 7z.exe and warning if not found
                [xml]$PCReadXMLFile = invoke-command -ScriptBlock { &'C:\Program Files\7-Zip\7z.exe' e $WorkingFileLine -so}
            }
            $F_mtorg_name = $PCReadXMLFile.DocumentElement.report_metadata.org_name
            $F_mtdtBegin = $PCReadXMLFile.DocumentElement.report_metadata.date_range.begin
            $F_mtdtEnd = $PCReadXMLFile.DocumentElement.report_metadata.date_range.end
            $F_mtEmail = $PCReadXMLFile.DocumentElement.report_metadata.email
            $F_mtContact = $PCReadXMLFile.DocumentElement.report_metadata.extra_contact_info
            $F_mtRptID = $PCReadXMLFile.DocumentElement.report_metadata.report_id
            $F_ppDomain = $PCReadXMLFile.DocumentElement.policy_published.domain
            $F_ppadkim = $PCReadXMLFile.DocumentElement.policy_published.adkim
            if ($F_ppadkim -eq "r"){$F_ppadkim = "relaxed"} elseif ($F_ppadkim -eq "s"){$F_ppadkim = "strict"}
            $F_ppaspf = $PCReadXMLFile.DocumentElement.policy_published.aspf
            if ($F_ppaspf -eq "r"){$F_ppaspf = "relaxed"} elseif ($F_ppaspf -eq "s"){$F_ppaspf = "strict"}
            $F_ppp = $PCReadXMLFile.DocumentElement.policy_published.p
            $F_ppsp = $PCReadXMLFile.DocumentElement.policy_published.sp
            $F_pppct = $PCReadXMLFile.DocumentElement.policy_published.pct
            $HR += "<tr>
                <td>$F_mtorg_name</td>
                <td>$F_mtEmail</td>
                <td>$F_mtContact</td>
                <td>$F_mtRptID</td>
                <td>$F_mtdtBegin</td>
                <td>$F_mtdtEnd</td>
                <td>$F_ppdomain</td>
                <td>$F_ppadkim</td>
                <td>$F_ppaspf</td>
                <td>$F_ppp</td>
                <td>$F_ppsp</td>
                <td>$F_pppct</td>
            </tr>"
            $CNodeRec = foreach ($Record in $PCReadXMLFile.ChildNodes.record)
            {
                if (($Record.count) -ne 0)
                {
                    $R1 = $Record.row.policy_evaluated
                    if ($R1)
                    {
                        $F_riheader_from = $Record.identifiers.header_from
                        $F_rrSource_ip = $Record.row.source_ip
                        $RDN = Resolve-DnsName $F_rrSource_ip -ErrorAction SilentlyContinue -ErrorVariable RDNSN
                            $IP_PTR_Domain = if(($RDN).NameHost){$RDN.NameHost}else{"UNK"}
                        $F_rrCount = $Record.row.count
                        $F_rrDisposition = $R1.disposition # (Policy Applied)
                        $F_rrdkim = $R1.dkim
                        if ($F_rrdkim -eq "fail"){$F_rrdkim = "<font color=`"Orange`">$F_rrdkim</font>"}
                        $F_rrspf = $R1.spf
                        if ($F_rrspf -eq "fail"){$F_rrspf = "<font color=`"Red`">$F_rrspf</font>"}
                        $F_rrReasonType = $R1.reason.type
                        $F_rrReasonComment = $R1.reason.comment
                        $F_riRank = $Record.identifiers.Rank
                        $F_riSyncRoot = $Record.auth_results.SyncRoot
                        $F_rardkim = $Record.auth_results.dkim
                        $F_rardkimDomain = $Record.auth_results.dkim.domain
                        $F_rardkimResult = $Record.auth_results.dkim.result
                        $F_rardkimSelector = $Record.auth_results.dkim.selector
                        $F_rarspfDomain = $Record.auth_results.spf.domain
                        $F_rarspfResult = $Record.auth_results.spf.result
                        $F_rChildNodes = $Record.ChildNodes # Unused
                        $R1 |Add-Member -MemberType NoteProperty -Name F_mtRptID $F_mtRptID
                        $R1 |Add-Member -MemberType NoteProperty -Name F_StartDate $Date1
                        $R1 |Add-Member -MemberType NoteProperty -Name F_EndDate $Date2
                        $R1 |Add-Member -MemberType NoteProperty -Name F_riheader_from -Value $F_riheader_from
                        $R1 |Add-Member -MemberType NoteProperty -Name F_rrSource_ip -Value $F_rrSource_ip
                        $R1 |Add-Member -MemberType NoteProperty -Name IP_PTR_Domain -Value $IP_PTR_Domain
                        $R1 |Add-Member -MemberType NoteProperty -Name F_rrCount -Value $F_rrCount
                        $R1 |Add-Member -MemberType NoteProperty -Name F_rrDisposition -Value $F_rrDisposition
                        $R1 |Add-Member -MemberType NoteProperty -Name F_rrdkim -Value $F_rrdkim
                        $R1 |Add-Member -MemberType NoteProperty -Name F_rrspf -Value $F_rrspf
                        $R1 |Add-Member -MemberType NoteProperty -Name F_rrReasonType -Value $F_rrReasonType
                        $R1 |Add-Member -MemberType NoteProperty -Name F_rrReasonComment -Value $F_rrReasonComment
                        $R1 |Add-Member -MemberType NoteProperty -Name F_riRank -Value $F_riRank
                        $R1 |Add-Member -MemberType NoteProperty -Name F_riSyncRoot -Value $F_riSyncRoot
                        $R1 |Add-Member -MemberType NoteProperty -Name F_rardkim -Value $F_rardkim
                        $R1 |Add-Member -MemberType NoteProperty -Name F_rardkimDomain -Value $F_rardkimDomain
                        $R1 |Add-Member -MemberType NoteProperty -Name F_rardkimResult -Value $F_rardkimResult
                        $R1 |Add-Member -MemberType NoteProperty -Name F_rardkimSelector -Value $F_rardkimSelector
                        $R1 |Add-Member -MemberType NoteProperty -Name F_rarspfDomain -Value $F_rarspfDomain
                        $R1 |Add-Member -MemberType NoteProperty -Name F_rarspfResult -Value $F_rarspfResult
                        $R1 |Add-Member -MemberType NoteProperty -Name F_rChildNodes -Value $F_rChildNodes # Unused
                    }
                    $TRF_mtRptID = $R1.F_mtRptID
                    $TRF_StartDate = $R1.F_StartDate
                    $TRF_EndDate = $R1.F_EndDate
                    $TRF_riheader_from = $R1.F_riheader_from
                    $TRF_rrSource_ip = $R1.F_rrSource_ip
                    $TRF_IP_PTR_Domain = $R1.IP_PTR_Domain
                    $TRF_rrCount =$R1.F_rrCount
                    $TRF_rrDisposition = $R1.F_rrDisposition
                    $TRF_rrdkim = $R1.F_rrdkim
                    $TRF_rrspf = $R1.F_rrspf
                    $TRF_rrReasonType = $R1.F_rrReasonType
                    $TRF_rarspfResult = $R1.F_rarspfResult
                    $TRF_rardkimResult = $R1.F_rardkimResult
                    $TRF_rardkimSelector = $R1.F_rardkimSelector
                    $TRF_rrReasonComment = $R1.reason.comment
                    $TRF_riRank = $R1.F_riRank
                    $TRF_riSyncRoot = $R1.F_riSyncRoot
                    $TRF_rarspfDomain = $R1.F_rarspfDomain
                    $TRF_rardkim = $R1.TRF_rardkim
                    $TRF_rardkimDomain = $R1.F_rardkimDomain

                    $TR+="<tr>
                        <td>$TRF_mtRptID</td>
                        <td>$TRF_StartDate</td>
                        <td>$TRF_EndDate</td>
                        <td>$TRF_riheader_from</td>
                        <td>$TRF_rrSource_ip</td>
                        <td>$TRF_IP_PTR_Domain</td>
                        <td>$TRF_rrCount</td>
                        <td>$TRF_rrDisposition</td>
                        <td>$TRF_rrdkim</td>
                        <td>$TRF_rrspf</td>
                        <td>$TRF_rrReasonType</td>
                        <td>$TRF_rrReasonComment</td>
                        <td>$TRF_riRank</td>
                        <td>$TRF_riSyncRoot</td>
                        <td>$TRF_rardkim</td>
                        <td>$TRF_rardkimDomain</td>
                        <td>$TRF_rardkimResult</td>
                        <td>$TRF_rardkimSelector</td>
                        <td>$TRF_rarspfDomain</td>
                        <td>$TRF_rarspfResult</td>
                    </tr>"
                }
            }
        }
    }
    "Processing: $WorkingFileLine"
}

$FLC = $WorkingFiles.count
$HTMLTitle = "$CDate DMARC Report on $FLC files"
$HRHeader = "
<tr>
    <td colspan=12><CENTER><font color=`"Black`">$HTMLTitle</font></CENTER></td>
</tr>
<tr>
    <td>org_name</td>
    <td>email</td>
    <td>extra_contact_info</td>
    <td>Report_ID</td>
    <td>begin</td>
    <td>end</td>
    <td>domain</td>
    <td>adkim</td>
    <td>aspf</td>
    <td>p</td>
    <td>sp</td>
    <td>pct</td>
</tr>
"
 $TRHeader = "
    <tr style=`"color: black`">
        <td>Report_ID</td>
        <td>Start Date</td>
        <td>End Date</td>
        <td>From Domain:</td>
        <td>Source IP</td>
        <td>Domain PTR</td>
        <td>Msg Count</td>
        <td>Policy Applied</td>
        <td>dkim</td>
        <td>spf</td>
        <td>Reason Type</td>
        <td>Reason Comment</td>
        <td>Rank</td>
        <td>Sync Root</td>
        <td>dkim</td>
        <td>dkim Domain</td>
        <td>dkim result</td>
        <td>dkim selector</td>
        <td>spf Domain</td>
        <td>spf result</td>
    </tr>

"
$HTMLBody = @"
    <table>
    $HRHeader
    $HR
</table>
<p>
<table>
    $TRHeader
    $TR
</table>
"@
$Header = @"
    <style>
    TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;color: blue;}
    TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
    </style>
"@
ConvertTo-Html -Property Name,Used,Provider,root,currentlocation -Title $HTMLTitle -Body $HTMLBody -Head $Header |Out-File -FilePath .\DMARCReport.html
Start-Process "chrome" -Argument ".\DMARCReport.html"
if ($ExportCSVFileName){            
    try {            
        $Global:ObjArray  | Export-CSV -Path $ExportCSVFileName -NotypeInformation            
    } catch {            
        Write-Error "Failed to export the output to CSV. Details : $_"            
    }            
}
# Clean up compressed files 
if ($MoveFilePrompt -eq "Yes")
{
    foreach($WorkingFileLine in $WorkingFiles) 
    {
        "WorkingFileLine: $WorkingFileLine"
        Move-Item "$PCReadFolder\$WorkingFileLine" "$PCWriteFolder\$WFName" -Force -ErrorAction SilentlyContinue -ErrorVariable MoveErr
    }
    "$MoveErr"
}