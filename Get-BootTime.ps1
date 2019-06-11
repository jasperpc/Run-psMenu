#    "Discover boot times (from select systems)"

Param([array]$Global:PCList)


if ($Global:PCList -eq $null)
{
    Clear-Host
    . .\Get-SystemsList.ps1 # load functions in .ps1
    Get-PCList
}
function Identify-PCName
{
    # This if statement could potentially be moved to inside the Get-PCList function, but it might need to be modified to do so
    if ($Global:pcline.name -eq $null)
    {
        $Global:pc = $Global:pcline
    }
    else
    {
        $Global:pc = $Global:pcline.name
    }
}
function Discover-BootTime
{
$ObjectOut = foreach($Global:pcline in $Global:pclist)
    {
            $DateCol=$null
            $TimeCol=$null
        Identify-PCName
        <# Run Command by connecting to PC using Invoke-Expression combined with psexec
           $script:data = Invoke-Expression "& `".\PSexec.exe`" \\$Global:pc systeminfo |find `"Boot Time`""
        #>
        # Run Command by connecting to PC using Invoke-Command
        $data = Invoke-Command -ComputerName $Global:pc -ScriptBlock {systeminfo |find "Boot Time"} -ErrorAction SilentlyContinue
        # Regex pattern to group and strip boot date and time
        $Pattern1 = '^(\w*\s*\w*\s*\w*)(:\s*)(\d*\/\d*\/\d*)(,\s)(\d*:\d*:\d*\s[A-Z]*)'
        if ($data -ne $null)
        {
            $DateCol = ([regex]::Matches($data, $Pattern1).Groups[3].Value)
            $TimeCol = ([regex]::Matches($data, $Pattern1).Groups[5].Value)
        }
        # Build Object
        $OBProperties = @{'__Server'=$Global:pc;
            'BootDate'=$DateCol;
            'BootTime'=$TimeCol}
        New-Object -TypeName PSObject -Prop $OBProperties
        Write-Host "." -NoNewline
    }
$ObjectOut |select *
}
Discover-BootTime # Comment out this line if this function will only be run when called from elsewhere



