<#
.NAME
    Use-MenuUtility.ps1
.SYNOPSIS
    Menu for calling Utility functions
.DESCRIPTION
 Menu design initially adopted from : https://quickclix.wordpress.com/2012/08/14/making-powershell-menus/
 2020-11-12
 Written by Jason Crockett - Laclede Electric Cooperative
.EXAMPLE
    "a command with a parameter would look like: (.\Use-MenuUtility.ps1)"
        .\Use-MenuUtility.ps1 []
.SYNTAX
    .\Use-MenuUtility.ps1 []
.REMARKS
    To see the examples, type: help Use-MenuUtility.ps1 -examples
    To see more information, type: help Use-MenuUtility.ps1 -detailed"
.TODO

    *****   Need to fix the menus repeating even after trying to exit the menu     *****

#>
Clear-Host
# set Global and Script variables
# . .\Get-SystemsList.ps1
# [BOOLEAN]$Global:ExitSession=$false
# [BOOLEAN]$Global:ExitMenuUtilitySession=$false


$Script:Ut4 = "`t`t`t`t"
$Script:Ud4 = "------ Use-Utility Menu ($Global:PCCnt Systems currently selected)-----"
$Script:UNC = $null
$Script:Up1 = $null
$Script:Up2 = $null
$Script:Up3 = $null
[int]$UtilityMenuSelect = 0

Function Utilitychcolor($Script:Up1,$Script:Up2,$Script:Up3,$Script:UNC){
            
            Write-host $Script:Ut4 -NoNewline
            Write-host "$Script:Up1 " -NoNewline
            Write-host $Script:Up2 -ForegroundColor $Script:UNC -NoNewline
            Write-host "$Script:Up3" -ForegroundColor White
            $Script:UNC = $null
            $Script:Up1 = $null
            $Script:Up2 = $null
            $Script:Up3 = $null
}

Function Utilitymenu()
{
while ($UtilityMenuSelect -lt 1 -or $UtilityMenuSelect -gt 10)
{
        Trap {"Error: $_"; Break;}        
        $MUNum = 0;Clear-host |out-null
        Start-Sleep -m 50
        if ($Global:PCCnt -eq $null) {$Global:PCCnt = 0}Else {$Global:PCCnt = ($Global:PCList).Count}
        $Script:Ud4 = "------ Use-Utility Menu ($Global:PCCnt Systems currently selected)-----"
        Write-host $Script:Ut4 $Script:Ud4 -ForegroundColor Green
# Exit
        $MUNum ++;$UMExit=$MUNum;
            $Script:Up1 = " $MUNum. `t";
            $Script:Up2 = "Exit to Main Menu";
            $Script:Up3 = " ";
            $Script:UNC = "Red";Utilitychcolor $Script:Up1 $Script:Up2 $Script:Up3 $Script:UNC
# $Create_DMARCReport
        $MUNum ++;$Create_DMARCReport=$MUNum;
            $Script:Up1 = " $MUNum. `t Create a ";
            $Script:Up2 = "DMARC Report ";
            $Script:Up3 = "from xml files in zip files"
            $Script:UNC = "Yellow";Utilitychcolor $Script:Up1 $Script:Up2 $Script:Up3 $Script:UNC
# Get text URL Web Page Info
        $MUNum ++;$wGet_URL=$MUNum;
            $Script:Up1 = " $MUNum. `t Get ";
            $Script:Up2 = "Web Page in text ";
            $Script:Up3 = "from URL"
            $Script:UNC = "Yellow";Utilitychcolor $Script:Up1 $Script:Up2 $Script:Up3 $Script:UNC
# convert-Decimal to Binary to Hexadecimal
        $MUNum ++;$convert_Numbers=$MUNum;
            $Script:Up1 = " $MUNum. `t convert ";
            $Script:Up2 = "Decimal to Binary to Hexadecimal ";
            $Script:Up3 = " "
            $Script:UNC = "Yellow";Utilitychcolor $Script:Up1 $Script:Up2 $Script:Up3 $Script:UNC
# Convert text to Base64 using legacy ASCII encoding 
        $MUNum ++;$Convert_ASCII_To_Base64=$MUNum;
            $Script:Up1 = " $MUNum. `t Convert text from ";
            $Script:Up2 = "ASCII to Base64 ";
            $Script:Up3 = "using Legacy ASCII encoding"
            $Script:UNC = "Yellow";Utilitychcolor $Script:Up1 $Script:Up2 $Script:Up3 $Script:UNC
# Convert text to Base64 using UTF8 ASCII encoding 
        $MUNum ++;$Convert_ASCII_UTF8_To_Base64=$MUNum;
            $Script:Up1 = " $MUNum. `t Convert text from ";
            $Script:Up2 = "UTF8 to Base64 ";
            $Script:Up3 = "using UTF8 encoding"
            $Script:UNC = "Yellow";Utilitychcolor $Script:Up1 $Script:Up2 $Script:Up3 $Script:UNC
# This Base64 to legacy ASCII conversion combines the conversion and encoding into one step
        $MUNum ++;$Convert_Base64_To_ASCII=$MUNum;
            $Script:Up1 = " $MUNum. `t Convert from ";
            $Script:Up2 = "Base64 to ASCII text ";
            $Script:Up3 = "using Legacy ASCII encoding"
            $Script:UNC = "Yellow";Utilitychcolor $Script:Up1 $Script:Up2 $Script:Up3 $Script:UNC
# Convert Base64 to UTF8 using UTF8 encoding
        $MUNum ++;$Convert_Base64_ASCII_UTF8=$MUNum;
            $Script:Up1 = " $MUNum. `t Convert from ";
            $Script:Up2 = "Base64 to UTF8 text ";
            $Script:Up3 = "using UTF8 encoding"
            $Script:UNC = "Yellow";Utilitychcolor $Script:Up1 $Script:Up2 $Script:Up3 $Script:UNC
# Convert from ASCII to Base64 using UTF8 encoding then back
        $MUNum ++;$Manage_Base64_ASCII_UTF8Encoding=$MUNum;
            $Script:Up1 = " $MUNum. `t Convert text from  ";
            $Script:Up2 = "UTF8 to Base64 and back ";
            $Script:Up3 = "using UTF8 encoding"
            $Script:UNC = "Yellow";Utilitychcolor $Script:Up1 $Script:Up2 $Script:Up3 $Script:UNC
# Convert from ASCII to Base64 and then back
        $MUNum ++;$Manage_Base64_ASCII_Encoding=$MUNum;
            $Script:Up1 = " $MUNum. `t Convert text from  ";
            $Script:Up2 = "ASCII to Base64 and back ";
            $Script:Up3 = "using legacy ASCII encoding"
            $Script:UNC = "Yellow";Utilitychcolor $Script:Up1 $Script:Up2 $Script:Up3 $Script:UNC
        # Setting up the menu in this way allows you to move menu items around and have the numbering change automatically
        # It also allows you to put a word in the middle of the line and have it change color to be easier to see
        # $MUNum ++;$Script:Up1 =" $MUNum.  FIRST";$Script:Up2 = "HIGHLIGHTED ";$Script:Up3 = "END"; $Script:UNC = "yellow";Utilitychcolor $Script:Up1 $Script:Up2 $Script:Up3 $Script:UNC
        Write-Host "$Script:Ut4$Script:Ud4"  -ForegroundColor Green
        $Script:MUNum = 0;[int]$UtilityMenuSelect = Read-Host "`n`tPlease select the menu item you would like to do now"
}

switch($UtilityMenuSelect)
    {
        $UMExit{$UtilityMenuSelect=$null}
        $Create_DMARCReport{Invoke-Expression ".\Create-DMARCReport.ps1";ReloadUtilityMenu}
        $wGet_URL{Get-wURL;ReloadUtilityMenu}
        $convert_Numbers{convert-Numbers;ReloadUtilityMenu}
        $Convert_ASCII_To_Base64{Convert-ASCII_To_Base64;ReloadUtilityMenu}
        $Convert_ASCII_UTF8_To_Base64{Convert-ASCII_UTF8_To_Base64;ReloadUtilityMenu}
        $Convert_Base64_To_ASCII{Clear-Host;Convert-Base64_To_ASCII;ReloadUtilityMenu}
        $Convert_Base64_ASCII_UTF8{Clear-Host;Convert-Base64_ASCII_UTF8;ReloadUtilityMenu}
        $Manage_Base64_ASCII_UTF8Encoding{Convert-ASCII_UTF8_To_Base64;Convert-Base64_ASCII_UTF8;ReloadUtilityMenu}
        $Manage_Base64_ASCII_Encoding{Convert-ASCII_To_Base64;Convert-Base64_To_ASCII;ReloadUtilityMenu}
        default
        {
            $UtilityMenuSelect = $null
            ReloadUtilityMenu
        }

    }    
}

Function ReloadUtilityMenu()
{
        Read-Host "Hit Enter to continue"
        $UtilityMenuSelect = $null
        UtilityMenu
}
Function Get-wURL() # Initiate a web-request for a URL and return it in text form
{
    Clear-Host
    $URL = Read-Host "Please type a URL to investigate"
    $R = Invoke-WebRequest $URL -UseBasicParsing # (wget)
    # Include Images
    $I = $R.Images |ft src # |select -ExpandProperty images |ft
    # Include links
    $L = ($R.Links).href |ft href
    # Show Raw html content
    # $RC = $R.RawContent
    # Show html headers
    # $RH = $R.Headers
    $I+$L |ft
}

Function convert-Numbers()
{
    clear-host
    # add a section to input Hex number rather than just inputing a base 10 number
    [int64]$Num = Read-host "Enter a base 10 number to convert)"
    $BinNum = [convert]::ToString($Num,2)
    "Binary: $BinNum"
    $DecNum = [convert]::ToInt64($BinNum,2)
    "Decimal: $DecNum"
    $hexNum = $Num |Format-Hex
    $hexNum = $hexNum -replace "^0{8}\s+","0x"
    "Hexadecimal: $hexNum"
    # Number with 2 places after the decimal
    $Numd2 = "{0:N2}" -f $Num
    $Numd2
    # Formatted with Currency
    $Cur = "{0:C2}" -f $Num
    $Cur
    # Formatted in Binary (8 places)
    $BinNum = "{0:D8}" -f $BinNum
    $BinNum
    # Formatted as a percent
    $Percent = "{0:P0}" -f ($DecNum/100)
    $Percent
    # Formatted in Traditional 0x Hex
    $Hex = "{0:X0}" -f $Num # $hexNum
    "Hex: $Hex"
    $hexNum
}
# Get ASCII "plain text" to encode
Function Enter-ASCII_Text # Called from multiple other functions
{
    $Script:ASCIIText = Read-Host "Enter the text to encode here"
}
# Get Base64 to decode
Function Enter-Base64 # Called from multiple other functions
{
    $Script:TextInBase64 = Read-Host "Enter the Base64 to decode here"
}
# Convert from ASCII to Base64 using UTF8 encoding
Function Convert-ASCII_UTF8_To_Base64 # Depends on Enter-ASCII_Text function
{
    Clear-Host
    Enter-ASCII_Text
    Write-Host "ASCIIText: $ASCIIText"
    $UTF8Bytes = [System.Text.Encoding]::UTF8.GetBytes($ASCIIText)
    Write-Host "UTF8Bytes: $UTF8Bytes" -ForegroundColor DarkGreen
    $TextInBase64 = [System.Convert]::ToBase64String($UTF8Bytes)
    Write-Host "TextInBase64: $TextInBase64`n" -ForegroundColor DarkYellow
}
# Convert from Base64 to ASCII using UTF8 encoding
Function Convert-Base64_ASCII_UTF8 # Depends on Enter-Base64 function
{
    Enter-Base64
    $FromBase64Text = [System.Convert]::FromBase64String($TextInBase64)
    Write-Host "To Binary FromBase64Text: $FromBase64Text" -ForegroundColor DarkGreen
    $DecodedText = [System.Text.Encoding]::UTF8.GetString($FromBase64Text)
    Write-Host "DecodedText: $DecodedText`n" -ForegroundColor Yellow
}
# Convert from ASCII to Base64
Function Convert-ASCII_To_Base64  # Depends on Enter-ASCII_Text function
{
    Clear-Host
    Enter-ASCII_Text
    Write-Host "ASCIIText: $ASCIIText"
    $ASCIIBytes = [System.Text.Encoding]::ASCII.GetBytes($ASCIIText)
    Write-Host "`tASCIIBytes: $ASCIIBytes" -ForegroundColor Gray
    $ToBase64ASCIIText = [System.Convert]::ToBase64String($ASCIIBytes)
    Write-Host "`tToBase64ASCIIText: `n"
    Write-Host "`t$ToBase64ASCIIText`n" -ForegroundColor DarkGray
}
# Convert from Base64 to ASCII
Function Convert-Base64_To_ASCII  # Depends on Enter-Base64 function
{
    Enter-Base64
    $FromBase64ASCIIText = [System.Convert]::FromBase64String($TextInBase64)
    $CodeWheelObject = foreach ($FB64At in $FromBase64ASCIIText)
    {
        $FB64At=$FB64At+2
        $FB64At
        $OBProperties = @{'number'=($FB64At).ToString()}
        New-Object -TypeName PSObject -Prop $OBProperties
    }
    <#
    Write-Host "`tCodeWheelObject: $CodeWheelObject" -ForegroundColor Cyan
    $DecodedCodeWheelObject = [System.Text.Encoding]::ASCII.GetString([system.convert]::FromBase64String($CodeWheelObject))
    Write-Host "DecodedCodeWheelObject:"
    Write-Host "$DecodedCodeWheelObject"
    #>
    Write-Host "`tThe Binary of FromBase64ASCIIText: `n" -ForegroundColor DarkGreen
    Write-Host "`t$FromBase64ASCIIText`n" -ForegroundColor Yellow
    $DecodedASCIIText = [System.Text.Encoding]::ASCII.GetString($FromBase64ASCIIText)
    Write-Host "DecodedASCIIText: `n"
    Write-Host "`t$DecodedASCIIText`n" -ForegroundColor Yellow
}

# This Base64 to ASCII conversion combines the conversion and encoding into one step
Function Convert-Base64ToASCII # Depends on Enter-Base64 function
{
    Clear-Host
    Enter-Base64
    $DecodedText = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String("LEC136"<#$Script:TextInBase64#>))
    Write-Host "Base 64 `"$TextInBase64`" decodes to:`n" -ForegroundColor DarkGray
    Write-Host "$DecodedText" -ForegroundColor Yellow
}