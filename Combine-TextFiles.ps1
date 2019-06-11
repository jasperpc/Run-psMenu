<#
 Combine text files in a folder
 Originally used to combine Toshiba text (.trf) traffic reports
#>
[CmdletBinding()] #(allows the use of parameter attributes)
param(
        #[Parameter(Mandatory=$True)]
        [string[]]$ext = ".trf",
        [string[]]$Path = "C:\Users\Public\Downloads\_Temp Clear Phone\TRAFFIC"
)
"Path is $Path"
$SourceFolder = (gi -Path $Path -Verbose).Name
"SourceFolder is $SourceFolder"
"ext is $ext"
$newFileName = "_"+$SourceFolder+"_combinedfile.txt"
"newFileName is $newFileName"
$flist = ls $Path|where {$_.Extension -eq $ext}|select -ExpandProperty Name
foreach ($fl in $flist){gc $fl|out-file $newFileName -Append}
notepad "$newFileName"
<#
 Combine text files in a folder
 Originally used to combine Toshiba text (.trf) traffic reports
#>
[CmdletBinding()] #(allows the use of parameter attributes)
param(
        #[Parameter(Mandatory=$True)]
        [string[]]$ext = ".trf",
        [string[]]$Path = "C:\Users\Public\Downloads\_Temp Clear Phone\TRAFFIC"
)
"Path is $Path"
$SourceFolder = (gi -Path $Path -Verbose).Name
"SourceFolder is $SourceFolder"
"ext is $ext"
$newFileName = "_"+$SourceFolder+"_combinedfile.txt"
"newFileName is $newFileName"
$flist = ls $Path|where {$_.Extension -eq $ext}|select -ExpandProperty Name
foreach ($fl in $flist){gc $fl|out-file $newFileName -Append}
notepad "$newFileName"
