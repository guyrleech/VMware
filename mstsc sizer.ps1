#requires -version 3

<#
.SYNOPSIS
    Show monitor resolutions and allow to pick one for mstsc launch where it uses most of the screen or other dimensions from parameters

.NOTES
    @guyrleech 2022/08/21

    Modification History:

    2022/08/24 @guyrleech  Added msrdc support
    2022/08/25 @guyrleech  Added GUI
    2022/09/23 @guyrleech  Fixed bug where : in .rdp file
    2022/10/13 @guyrleech  Added Other RDP options capability & persist to registry. Added username text box
    2022/10/14 @guyrleech  Added editable Comment column to display view
    2022/11/09 @guyrleech  Detect if window maximised and undo that before sizing/positioning.
                           Added -useOtherOptions otherwise tries to login locally with no password if AZ options in rdp file
                           Added logic to find existing windows as msrdc re-uses same process
                           Added Launch button on other tabs
    2022/11/14 @guyrleech  Fix for when UserFriendlyName is empty/null. Added BOE manufacturer   
    2022/11/15 @guyrleech  Change method of getting width and height as was relative to DPI scaling which made size too small
    2022/11/16 @guyrleech  Added VMware tab
    2022/12/14 @guyrleech  Changed -percentage to support x:y
    2022/12/19 @guyrleech  Added code to work for profile install of msrdc
    2023/02/21 @guyrleech  Persist computers list to HKCU
    2023/02/24 @guyrleech  Add VMware VMs to main computers list if connected to

    ## TODO persist the "comment" column in memory so that it is available when undocked and redocked
    ## TODO check for existing process and offer to bring that to front, kill or launch another although if msrdc will not work
#>

<#
Copyright © 2022 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding()]

Param
(
    [string]$address ,
    [string]$displayModel ,
    [string]$displayManufacturer ,
    [string]$displayManufacturerCode ,
    [string]$username ,
    [string]$remoteDesktopName ,
    [switch]$primary ,
    [string]$percentage ,
    [switch]$usemsrdc ,
    [switch]$noResize , ## use mstsc with no width/height parameters
    [string]$widthHeight , ## colon delimited
    [string]$xy , ## colon delimited
    [string]$drivesToRedirect = '*' ,
    [switch]$usegridviewpicker ,
    [switch]$fullScreen ,
    [switch]$showDisplays ,
    [switch]$showManufacturerCodes ,
    [switch]$useOtherOptions ,
    [switch]$noMove ,
    [int]$windowWaitTimeSeconds = 60 ,
    [int]$pollForWindowEveryMilliseconds = 333 ,
    [string]$exe = 'mstsc.exe' ,
    [string]$configKey = 'HKCU:\SOFTWARE\Guy Leech\mstsc wrapper'
)

#region data

[array]$script:vms = $null
$script:vmwareConnection = $null

# keep user added comments so can set when displays change
##$script:itemscopy = New-Object -TypeName System.Collections.Generic.List[object]

## https://docs.microsoft.com/en-us/windows-server/remote/remote-desktop-services/clients/rdp-files
[string]$rdpTemplate = @'
desktopwidth:i:$width
desktopheight:i:$height
full address:s:$address
use multimon:i:0
screen mode id:i:$screenmode
dynamic resolution:i:2
smart sizing:i:0
drivestoredirect:s:$drivesToRedirect
'@

[string]$mainwindowXAML = @'
<Window x:Class="mstsc_msrdc_wrapper.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:mstsc_msrdc_wrapper"
        mc:Ignorable="d"
        Title="Guy's mstsc/msrdc Wrapper Script" Height="500" Width="809">
    <Grid HorizontalAlignment="Left" VerticalAlignment="Top">
        <TabControl HorizontalAlignment="Center" Height="432" VerticalAlignment="Center" Width="768">
            <TabItem Header="Main">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="23*"/>
                        <ColumnDefinition Width="50*"/>
                        <ColumnDefinition Width="103*"/>
                        <ColumnDefinition Width="68*"/>
                        <ColumnDefinition Width="518*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.ColumnSpan="5" Height="110" Margin="15,10,-160,0" VerticalAlignment="Top" Width="907">
                        <DataGrid x:Name="datagridDisplays" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" SelectionMode="Single" />
                    </StackPanel>
                    <Label Content="Computer" HorizontalAlignment="Left" Height="38" Margin="14,132,0,0" VerticalAlignment="Top" Width="71" Grid.ColumnSpan="3"/>
                    <Button x:Name="btnLaunch" Content="_Launch" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="25" Margin="2,0,0,10" VerticalAlignment="Bottom" Width="96" Grid.Column="1"/>
                    <CheckBox x:Name="chkboxmsrdc" Grid.Column="4" Content="Use msrdc instead of mstsc" HorizontalAlignment="Left" Height="21" Margin="145,189,0,0" VerticalAlignment="Top" Width="292"/>
                    <ComboBox x:Name="comboboxComputer" Grid.Column="2" HorizontalAlignment="Left" Height="27" Margin="14,137,0,0" VerticalAlignment="Top" Width="254" IsEditable="True" IsDropDownOpen="False" Grid.ColumnSpan="3"/>
                    <CheckBox x:Name="chkboxPrimary" Grid.Column="4" Content="Use primary monitor" HorizontalAlignment="Left" Height="21" Margin="145,215,0,0" VerticalAlignment="Top" Width="292"/>
                    <TextBox x:Name="txtboxDrivesToRedirect" Grid.Column="2" HorizontalAlignment="Left" Height="26" Margin="14,230,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="254" Text="*" Grid.ColumnSpan="3"/>
                    <Label Content="Drive&#xA;Redirection" HorizontalAlignment="Left" Height="46" Margin="14,220,0,0" VerticalAlignment="Top" Width="71" Grid.ColumnSpan="3"/>
                    <CheckBox x:Name="chkboxNoMove" Grid.Column="4" Content="Do not move window" HorizontalAlignment="Left" Height="21" Margin="145,245,0,0" VerticalAlignment="Top" Width="292"/>
                    <RadioButton x:Name="radioFullScreen" Grid.Column="4" Content="Fullscreen" HorizontalAlignment="Left" Height="24" Margin="145,296,0,0" VerticalAlignment="Top" Width="206" GroupName="WindowSize"/>
                    <RadioButton x:Name="radioPercentage" Grid.Column="4" Content="Screen Percentage (X:Y)" HorizontalAlignment="Left" Height="24" Margin="145,272,0,0" VerticalAlignment="Top" Width="206" GroupName="WindowSize"/>
                    <RadioButton x:Name="radioWidthHeight" Grid.Column="4" Content="Width &amp; Height" HorizontalAlignment="Left" Height="24" Margin="145,324,0,0" VerticalAlignment="Top" Width="206" GroupName="WindowSize" />
                    <TextBox x:Name="txtboxWindowPosition" Grid.Column="2" HorizontalAlignment="Left" Height="26" Margin="14,283,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="254" Text="0,0" Grid.ColumnSpan="3"/>
                    <Label Content="Window&#xA;Position" HorizontalAlignment="Left" Height="46" Margin="14,273,0,0" VerticalAlignment="Top" Width="71" Grid.ColumnSpan="3"/>
                    <TextBox x:Name="txtboxScreenPercentage" Grid.Column="4" HorizontalAlignment="Left" Height="23" Margin="314,273,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="158"/>
                    <TextBox x:Name="txtboxWidthHeight" Grid.Column="4" HorizontalAlignment="Left" Height="23" Margin="314,319,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="158">
                        <TextBox.InputBindings>
                            <MouseBinding Gesture="LeftDoubleClick" />
                        </TextBox.InputBindings>
                    </TextBox>
                    <RadioButton x:Name="radioFillScreen" Grid.Column="4" Content="Fill Screen" HorizontalAlignment="Left" Height="24" Margin="145,348,0,0" VerticalAlignment="Top" Width="206" GroupName="WindowSize"/>
                    <Button x:Name="btnRefresh" Content="_Refresh" HorizontalAlignment="Left" Height="25" Margin="71,0,0,10" VerticalAlignment="Bottom" Width="96" Grid.Column="2" Grid.ColumnSpan="2"/>
                    <Button x:Name="btnCreateShortcut" Content="_Create Shortcut" HorizontalAlignment="Left" Height="25" Margin="14,0,0,10" VerticalAlignment="Bottom" Width="96" Grid.Column="4"/>
                    <Label Content="User" HorizontalAlignment="Left" Height="38" Margin="14,181,0,0" VerticalAlignment="Top" Width="71" Grid.ColumnSpan="3"/>
                    <TextBox x:Name="textboxUsername" Grid.Column="2" HorizontalAlignment="Left" Height="27" Margin="14,186,0,0" VerticalAlignment="Top" Width="254" Grid.ColumnSpan="3"/>

                </Grid>

            </TabItem>
            <TabItem Header="Mstsc Options">
                <Grid Margin="0,0,100,100   " Grid.Column="1" Height="200">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="214*"/>
                        <ColumnDefinition Width="53*"/>
                    </Grid.ColumnDefinitions>
                    <CheckBox x:Name="chkboxMultimon" Content="_Multi Monitor" HorizontalAlignment="Left" Height="28" VerticalAlignment="Top" Width="267" Grid.ColumnSpan="2"/>
                    <CheckBox x:Name="chkboxSpan" Content="_Span" HorizontalAlignment="Left" Height="28" Margin="0,29,0,0" VerticalAlignment="Top" Width="267" Grid.ColumnSpan="2"/>
                    <CheckBox x:Name="chkboxAdmin" Content="_Admin" HorizontalAlignment="Left" Height="28" Margin="0,57,0,0" VerticalAlignment="Top" Width="267" Grid.ColumnSpan="2"/>
                    <CheckBox x:Name="chkboxPublic" Content="_Public" HorizontalAlignment="Left" Height="28" Margin="0,85,0,0" VerticalAlignment="Top" Width="267" Grid.ColumnSpan="2"/>
                    <CheckBox x:Name="chkboxRemoteGuard" Content="Remote _Guard" HorizontalAlignment="Left" Height="28" Margin="0,113,0,0" VerticalAlignment="Top" Width="267" Grid.ColumnSpan="2"/>
                    <CheckBox x:Name="chkboxRestrictedAdmin" Content="_Restricted Admin" HorizontalAlignment="Left" Height="28" Margin="0,141,0,0" VerticalAlignment="Top" Width="267" Grid.ColumnSpan="2"/>
                    <Button x:Name="btnLaunchMstscOptions" Content="_Launch" HorizontalAlignment="Left" Height="25" VerticalAlignment="Bottom" Width="96" Margin="10,0,0,-128" IsDefault="True"/>

                </Grid>
            </TabItem>
            <TabItem Header="Other Options">

                <Grid x:Name="OtherRDPOptions" Margin="55,0,528,0" Height="309">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="146*"/>
                        <ColumnDefinition Width="33*"/>
                    </Grid.ColumnDefinitions>
                    <CheckBox x:Name="chkboxDoNotSave" Content="Do Not Save" Width="196" Grid.Column="1" Margin="46,133,-209,154"/>
                    <CheckBox x:Name="chkboxDoNotApply" Content="Do Not Apply" Width="196" Grid.Column="1" Margin="46,106,-209,181"/>
                    <Label Content="Other RDP File Options:" HorizontalAlignment="Center" Height="49" Margin="0,19,0,0" VerticalAlignment="Top" Width="144"/>
                    <TextBox x:Name="txtBoxOtherOptions" HorizontalAlignment="Left" Margin="10,57,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="230" Width="168" ForceCursor="True" IsManipulationEnabled="True" AcceptsReturn="True"  VerticalScrollBarVisibility="Visible" Grid.ColumnSpan="2"/>
                    <Button x:Name="btnLaunchOtherOptions" Content="_Launch" HorizontalAlignment="Left" Height="25" VerticalAlignment="Bottom" Width="96" Margin="10,0,0,-19" IsDefault="True"/>
                </Grid>
            </TabItem>
            <TabItem Header="VMware">

                <Grid x:Name="VMwareOptions" Margin="55,0,409,0" Height="342">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="129*"/>
                        <ColumnDefinition Width="140*"/>
                    </Grid.ColumnDefinitions>
                    <Button x:Name="btnLaunchVMwareOptions" Content="_Launch" HorizontalAlignment="Left" Height="25" VerticalAlignment="Bottom" Width="96" Margin="10,0,0,-24" IsDefault="True"/>
                    <ListView x:Name="listViewVMwareVMs" Grid.ColumnSpan="2" Height="293" Margin="10,39,70,0" VerticalAlignment="Top" SelectionMode="Single" >
                        <ListView.View>
                            <GridView>
                                <GridViewColumn/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                    <Label x:Name="labelVMwareVMs" Content="VMs" HorizontalAlignment="Center" Height="29" Margin="0,5,0,0" VerticalAlignment="Top" Width="122"/>
                    <Label Content="Filter" HorizontalAlignment="Left" Height="29" Margin="94,2,0,0" VerticalAlignment="Top" Width="123" Grid.Column="1"/>
                    <CheckBox x:Name="checkBoxVMwareRegEx" Content="RegEx" Height="29" Width="93" Grid.Column="1" Margin="304,41,-255,272" IsChecked="True"/>
                    <TextBox x:Name="textBoxVMwareFilter" TextWrapping="Wrap" Grid.Column="1" Margin="94,39,-143,275" />
                    <Button x:Name="buttonApplyFilter" Content="Apply _Filter" Height="31" Width="117" Grid.Column="1" Margin="94,86,-69,225"/>
                    <Label Content="vCenter" HorizontalAlignment="Left" Height="29" VerticalAlignment="Top" Width="124" Grid.Column="1" Margin="100,226,0,0" />
                    <TextBox x:Name="textBoxVMwareRDPPort" TextWrapping="Wrap" Height="28" Width="189" Grid.Column="1" Margin="102,156,-136,158" />
                    <Button x:Name="buttonVMwareConnect" Content="_Connect" Height="31" Width="117" Grid.Column="1" Margin="102,301,-64,10"/>
                    <RadioButton x:Name="radioButtonConnectByIP" Content="Connect by _IP" Margin="97,202,-127,121" Grid.Column="1" GroupName="GroupBy"/>
                    <RadioButton x:Name="radioButtonConnectByName"   Content="Connect by _Name" Margin="216,202,-202,121" Grid.Column="1" GroupName="GroupBy" IsChecked="True"/>
                    <Label Content="RDP Port" HorizontalAlignment="Left" Height="29" VerticalAlignment="Top" Width="124" Grid.Column="1" Margin="97,122,0,0" />
                    <TextBox x:Name="textBoxVMwarevCenter" TextWrapping="Wrap" Height="28" Grid.Column="1" Margin="100,255,-202,59" />
                    <Button x:Name="buttonVMwareDisconnect" Content="_Disconnect" Height="31" Width="117" Grid.Column="1" Margin="240,301,-202,10"/>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
'@

#endregion data
<#
    desktopscalefactor:i:200
    compression:i:0
#>

#region pre-main

## https://sirconfigmgr.de/display-inventory/
[hashtable]  $ManufacturerHash = @{ 
    'AAC' =	'AcerView'
    'ACR' = 'Acer'
    'AOC' = 'AOC'
    'AIC' = 'AG Neovo'
    'APP' = 'Apple Computer'
    'AST' = 'AST Research'
    'AUO' = 'Asus'
    'BNQ' = 'BenQ'
    'CMO' = 'Acer'
    'CPL' = 'Compal'
    'CPQ' = 'Compaq'
    'CPT' = 'Chunghwa Picture Tubes, Ltd.'
    'CTX' = 'CTX'
    'DEC' = 'DEC'
    'DEL' = 'Dell'
    'DPC' = 'Delta'
    'DWE' = 'Daewoo'
    'EIZ' = 'EIZO'
    'ELS' = 'ELSA'
    'ENC' = 'EIZO'
    'EPI' = 'Envision'
    'FCM' = 'Funai'
    'FUJ' = 'Fujitsu'
    'FUS' = 'Fujitsu-Siemens'
    'GSM' = 'LG Electronics'
    'GWY' = 'Gateway 2000'
    'HEI' = 'Hyundai'
    'HIT' = 'Hyundai'
    'HSL' = 'Hansol'
    'HTC' = 'Hitachi/Nissei'
    'HWP' = 'HP'
    'IBM' = 'IBM'
    'ICL' = 'Fujitsu ICL'
    'IVM' = 'Iiyama'
    'KDS' = 'Korea Data Systems'
    'LEN' = 'Lenovo'
    'LGD' = 'Asus'
    'LPL' = 'Fujitsu'
    'MAX' = 'Belinea' 
    'MEI' = 'Panasonic'
    'MEL' = 'Mitsubishi Electronics'
    'MS_' = 'Panasonic'
    'NAN' = 'Nanao'
    'NEC' = 'NEC'
    'NOK' = 'Nokia Data'
    'NVD' = 'Fujitsu'
    'OPT' = 'Optoma'
    'PHL' = 'Philips'
    'REL' = 'Relisys'
    'SAN' = 'Samsung'
    'SAM' = 'Samsung'
    'SBI' = 'Smarttech'
    'SGI' = 'SGI'
    'SNY' = 'Sony'
    'SRC' = 'Shamrock'
    'SUN' = 'Sun Microsystems'
    'SEC' = 'Hewlett-Packard'
    'TAT' = 'Tatung'
    'TOS' = 'Toshiba'
    'TSB' = 'Toshiba'
    'VSC' = 'ViewSonic'
    'ZCM' = 'Zenith'
    'UNK' = 'Unknown'
    '_YV' = 'Fujitsu'
    ## not in original
    'TMX' = 'Huawei'
    'HSD' = 'Hannspree'
    'BOE' = 'BOE Technology'
 }
 
Function Set-WindowToFront
{
    Param
    (
        [Parameter(Mandatory)]
        [IntPtr]$windowHandle
    )
    
    ## first restore window
    if( [bool]$setForegroundWindow = [user32]::ShowWindowAsync( $windowHandle , 9 ) ) ## 9 = SW_RESTORE
    {
        ## https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowpos
        ## now set window top most to bring it to front
        if( $setForegroundWindow = [user32]::SetWindowPos( $windowHandle, [IntPtr]-1 , 0 ,0 , 0 , 0 , 0x4043 ) ) ## -1 = HWND_TOPMOST , -2 = HWND_NOTOPMOST , 0x4043 = SWP_ASYNCWINDOWPOS | SWP_SHOWWINDOW | SWP_NOSIZE | SWP_NOMOVE
        {
            ## now set window to not top most but will stay on top for now (otherwise would always be on top like task manager can be)
            $setForegroundWindow  = [user32]::SetWindowPos( $windowHandle, [IntPtr]-2 , 0 ,0 , 0 , 0 , 0x4043 ) 
        }
    }

    $setForegroundWindow ## return
}

Function New-RemoteSession
{
    [CmdletBinding()]

    Param
    (
        [switch]$rethrow
    )
    
    [string]$commandLine = $null

    if( -Not [string]::IsNullOrEmpty( $address ) )
    {
        $commandLine = -join ($commandLine , " /v:$address" )
    }

    [int]$width = -1
    [int]$height = -1

    try
    {
        if( $fullScreen )
        {
            $commandLine = -join ($commandLine ,  ' /f' )
        }
        elseif( -Not $noResize )
        {
            [int]$reservedWidth  = $chosenDisplay.ScreenBounds.Width  - $chosenDisplay.ScreenWorkingArea.Width
            [int]$reservedHeight = $chosenDisplay.ScreenBounds.Height - $chosenDisplay.ScreenWorkingArea.Height
            if( -Not [string]::IsNullOrEmpty( $widthHeight ) )
            {
                [string[]]$dimensions = @( $widthHeight -split '[:x,\s]' )
                if( $dimensions.Count -ne 2 )
                {
                    Throw "Invalid parameter `"$widthHeight`" specified - must be width:height"
                }
                if( ( $width = $dimensions[0] -as [int] ) -le 0 )
                {
                    Throw "Invalid width in `"$widthHeight`""
                }
                if( ( $height = $dimensions[1] -as [int] ) -le 0 )
                {
                    Throw "Invalid height in `"$widthHeight`""
                }

            }
            elseif( -Not [string]::IsNullOrEmpty( $percentage ) )
            {
                [int]$xpercentage = 100
                [int]$ypercentage = 100
                [int]$percentageAsInt = -1

                if( $percentage -match '^(\d+):(\d+)' )
                {
                    $xpercentage = $Matches[1]
                    $ypercentage = $Matches[2]
                }
                elseif( [int]::TryParse( $percentage , [ref]$percentageAsInt ) )
                {
                    $xpercentage = $ypercentage = $percentageAsInt
                }
                else
                {
                    Write-Warning -Message "Percentage `"$percentage`" is not valid - ignoring"
                }
                $width  = $chosenDisplay.dmPelsWidth  * $xpercentage / 100
                $height = $chosenDisplay.dmPelsHeight * $ypercentage / 100
            }
            else ## fill the chosen screen
            {
                $width  = $chosenDisplay.dmPelsWidth  - $reservedWidth
                $height = $chosenDisplay.dmPelsHeight - $reservedHeight
            }
            <# ## doesn't work when font scaling not 100%
            elseif( $percentage -gt 0 )
            {
                $width  = $chosenDisplay.ScreenWorkingArea.Width  * $percentage / 100
                $height = $chosenDisplay.ScreenWorkingArea.Height * $percentage / 100
            }
            else ## fill the chosen screen
            {
                $width  = $chosenDisplay.ScreenWorkingArea.Width  - $widthReduction
                $height = $chosenDisplay.ScreenWorkingArea.Height - $heightReduction
            }
            #>
        }

        if( $width -gt 0 )
        {
            $commandLine = -join ($commandLine , " /w:$width" )
        }

        if( $height -gt 0 )
        {
            $commandLine = -join ($commandLine , " /h:$height" )
        }

        Write-Verbose -Message "Running $exe $commandLine"

        $process = $null

        [string]$tempRdpFile = $null
        [string]$windowTitle = ' - Remote Desktop Connection$'

        if( -Not ( $tempFile = New-TemporaryFile  ) )
        {
            Throw "Unable to create temporary file for rdp settings"
        }
        ## change file extension (probably .tmp but don't assume)
        $tempRdpFile = $tempFile.FullName -replace '\.[^\.]+$' , '.rdp'
        if( -Not ( Move-Item -Path $tempFile -Destination $tempRdpFile -PassThru ) )
        {
            Throw "Failed to move $tempFile to $tempRdpFile"
        }

        ## mstsc file will have things in it doesn't understand which it silently ignores
        
        [int]$screenmode = 1
        if( $fullScreen )
        {
            $screenmode = 2
        }
    
        [string]$rdpFileContents = $ExecutionContext.InvokeCommand.ExpandString( $rdpTemplate )

        if( $usemsrdc )
        {
            if( -Not ( Get-Command -Name ($exe = 'msrdc.exe') -CommandType Application -ErrorAction SilentlyContinue ) )
            {
                if( $apppathskey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msrdc.exe' -ErrorAction SilentlyContinue ) 
                {
                    if( $apppathskey.psobject.Properties[ '(default)' ] )
                    {
                        $exe = $apppathskey.'(default)'
                    }
                    elseif( $apppathskey.psobject.Properties[ 'path' ] )
                    {
                        $exe = Join-Path -Path $apppathskey.path -ChildPath 'msrdc.exe'
                    }
                    else
                    {
                        Throw "App Paths key found for msrdc.exe but it contains no usable paths"
                    }
                }
                elseif( $installPath = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.PSObject.Properties[ 'DisplayName' ] -and $_.DisplayName -eq 'Remote Desktop' -and $_.Publisher -eq 'Microsoft Corporation' } | Select-Object -ExpandProperty InstallLocation )
                {
                    $exe = Join-Path -Path $installPath -ChildPath 'msrdc.exe'
                }
                else
                {
                    $exe = 'C:\Program Files\Remote Desktop\msrdc.exe'
                }
            }
            
            if( [string]::IsNullOrEmpty( $remoteDesktopName ))
            {
                ## if there is no remote desktop name specified then the temp rdp file name is included in the Window title which is fugly
                $remoteDesktopName = "$address - Remote Desktop"
                $rdpFileContents += "`nremotedesktopname:s:$remoteDesktopName`n"
                $windowTitle = $remoteDesktopName
            }

            $commandLine = "`"$tempRdpFile`" /u:$username"
        }
        else ## mstsc
        {
            ## window title comes from the base name of the .rdp file so if we don't rename the temp file, that will be the name in the title bar which is fugly
            [string]$tempRdpFileWithName = Join-Path -Path (Split-Path -Path $tempRdpFile -Parent) -ChildPath "$($address -replace ':' , '.').$(Split-Path -Path $tempRdpFile -Leaf)"
            if( Move-Item -Path $tempRdpFile -Destination $tempRdpFileWithName -PassThru )
            {
                $tempRdpFile = $tempRdpFileWithName
            }
            $commandline = "`"$tempRdpFile`" $commandLine"
        }
        
        ## see if we already have a window with this title so we can offer to switch to that or create a new one
        $existingWindows = $null
        $existingWindows = [Api.Apidef]::GetWindows( -1 ) | Where-Object WinTitle -ieq $windowTitle

        if( $null -ne $existingWindows )
        {
            Write-Verbose -Message "Already have window `"$windowTitle`" in process $($existingWindows.PID)"
            
            $answer = [Windows.MessageBox]::Show(  'Activate Existing Window ?' , "Already Connected to $address" , 'YesNoCancel' ,'Question' )
            if( $answer -ieq 'yes' )
            {
                $otherprocess = Get-Process -Id $existingWindows.PID
                if( $otherprocess )
                {
                    if( -Not ( Set-WindowToFront -windowHandle $otherprocess.MainWindowHandle ))
                    {
                        [void][Windows.MessageBox]::Show( 'Failed to Activate Window' , "$($otherprocess.Name) (PID $($otherprocess.Id))" , 'Ok' ,'Error' )
                    }
                }
                else
                {
                    [void][Windows.MessageBox]::Show( 'Failed to Get Process' , "PID $($otherprocess.Id)" , 'Ok' ,'Error' )
                }
            
                return
            }
            elseif( $answer -ieq 'Cancel' )
            {
                return
            }
        }

        if( -Not [string]::IsNullOrEmpty( $rdpFileContents ) )
        {
            Write-Verbose -Message "Writing $($rdpFileContents.Length) bytes to $tempRdpFile"
            ## TODO do we need to make sure no duplicates?
            ( $rdpFileContents + "`n" + $(if( -Not $wpfchkboxDoNotApply.IsChecked ) { $WPFtxtBoxOtherOptions.Text } )) | Set-Content -Path $tempRdpFile -Force
            if( -Not $? )
            {
                Throw "Failed to write rdp file contents to $tempRdpFile"
            }
        }
        
        ## use Start-Process so we can get a pid and thus window handle to move to chosen display
        $process = Start-Process -FilePath $exe -ArgumentList $commandLine -PassThru -WindowStyle Normal

        if( -Not $process )
        {
            Throw "Failed to launch $exe $commandLine"
        }

        try
        {
            [void]$process.WaitForInputIdle() ## but this may only be authentication
        }
        catch
        {
            Write-Warning -Message "$_"
        }

#      if( $process.HasExited -or -not $process.MainWindowHandle ) ## msrdc reuses existing process and new process exits so have to look for this window title in another process
#      {
        ## There may be more than one, either open or closed, so we need to find the new one which will be the one not in the existingWindows collection
        [int]$windowPid = -1
        [datetime]$endTime = [datetime]::MaxValue
        [string]$baseExe = (Split-Path -Path $exe -Leaf) -replace '\.[^\.]+$'
        if( $windowWaitTimeSeconds -gt 0 )
        {
            $endTime = [datetime]::Now.AddSeconds( $windowWaitTimeSeconds )
        }
        ## can take a little time for the window to appear and get the title so we poll :-(
        do
        {
            $allWindowsNow = @( [Api.Apidef]::GetWindows( -1 ) | Where-Object WinTitle -ieq $windowTitle )
            ForEach( $window in $allWindowsNow )
            {
                ## could switch process because of the way msrdc seems to work
                if( ( $process = Get-Process -Id $window.PID -ErrorAction SilentlyContinue ) -and $process.Name -eq $baseExe )
                {
                    [bool]$existing = $false
                    ForEach( $existingWindow in $existingWindows ) ## if there weren't any existing windows with this title, there may be now but it may be for a different msrdc process if one was already running for another machine
                    {
                        if( $existing = $window.WinHwnd -eq $existingWindow.WinHwnd )
                        {
                            break
                        }
                    }
                    if( -Not $existing )
                    {
                        $windowPid = $window.PID
                        break
                    }
                }
            }
            if( $windowPid -lt 0 )
            {
                if( $usemsrdc )
                {
                    ## TODO can't simply check if process has exited as msrdc can re-use existing plus if prompting for credentials it will be a process CredentialUIBroker.exe which is not a child of msrdc
                }
                else
                {
                    if( $process.HasExited )
                    {
                        Write-Warning -Message "Process $($process.Id) has exited"
                        break
                    }
                }
                Write-Verbose -Message "$(Get-Date -Format G): waiting until $(Get-Date -Date $endTime -Format G) for PID $($process.Id) to find window title `"$windowTitle`" for $baseExe"
                Start-Sleep -Milliseconds $pollForWindowEveryMilliseconds
            }
        }
        while( $windowPid -le 0 -and [datetime]::Now -lt $endTime )

        if( $windowPid -gt 0 )
        {
            $process = Get-Process -Id $windowPid
        }

        if( -not $process -or -Not $process.MainWindowHandle )
        {
            Throw "No main window handle for process $($process.id)"
        }

        if( -Not $noMove )
        {
            ## if window is maximized, undo that first so positioning & resizing works ok - msrdc seems to ignore -WindowStyle Normal
            if( [user32]::IsZoomed( $process.MainWindowHandle ) -and -Not $fullScreen )
            {
                Write-Verbose -Message "Window is maximised so undoing"
                ## 1 is SW_NORMAL
                $unmaximiseResult = [user32]::ShowWindowAsync( $process.MainWindowHandle, 1 ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

                if( -Not $unmaximiseResult )
                {
                    Write-Warning -Message "Failed ShowWindowAsync to unmaximise - $lastError"
                }
            }

            ## https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowpos

            [int]$flags = 0x4041 ## SWP_NOSIZE (leave size alone ) | SWP_SHOWWINDOW | SWP_ASYNCWINDOWPOS

            [int]$x = $chosenDisplay.ScreenBounds.x
            [int]$y = $chosenDisplay.ScreenBounds.y

            if( -Not [string]::IsNullOrEmpty( $xy ) )
            {
                ## we change the coordinates arelatively, not absolutely otherwise the window may not change window
                [string[]]$dimensions = @( $xy -split ':' )
                [int]$deltax = 0
                [int]$deltay = 0
                if( $dimensions.Count -ne 2 )
                {
                    Throw "Invalid parameter `"$xy`" specified - must be x:y"
                }
                if( $null -eq ( $deltax = $dimensions[0] -as [int] ) )
                {
                    Throw "Invalid x coordinate in `"$xy`""
                }
                if( $null -eq ( $deltay = $dimensions[1] -as [int] ) )
                {
                    Throw "Invalid y coordinate in `"$xy`""
                }
                $x += $deltax
                $y += $deltay
            }

            Write-Verbose -Message "SetWindowPos x=$x y=$y"
            $result = [user32]::SetWindowPos( $process.MainWindowHandle , [IntPtr]::Zero , $x , $y , $width , $height , $flags ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

            if( -Not $result) 
            {
                Write-Warning -Message "Failed SetWindowPos - $lastError"
            }

            ## if msrdc and fill screen (so not percentage or x/y) we make it maximised so it fills that screen
            if( $fullScreen -or ( $usemsrdc -and [string]::IsNullOrEmpty( $percentage ) -and [string]::IsNullOrEmpty( $widthHeight ) ) )
            {
                ## if it has been moved then we may need to maximise it again
                [int]$cmdShow = 3 ## SHOWMAXIMIZED

                ## https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindowasync
                $styleResult = [user32]::ShowWindowAsync( $process.MainWindowHandle, [int]$cmdShow ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

                if( -Not $styleResult )
                {
                    Write-Warning -Message "Failed ShowWindowAsync - $lastError"
                }
            }
        }

        if( $tempRdpFile )
        {
            Remove-Item -Path $tempRdpFile
            $tempRdpFile = $null
        }
        ## TODO add computer to MRU ?
    }
    catch
    {
        if( $rethrow )
        {
            Throw $_
        }
        else
        {
            Write-Error -Message $_
            return $false
        }
    }
    finally
    {
        if( $tempRdpFile )
        {
            Remove-Item -Path $tempRdpFile -Force
        }
    }

    return $true
}
 

Function New-WPFWindow( $inputXAML )
{
    $form = $NULL
    [xml]$XAML = $inputXAML -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window'
  
    if( $reader = New-Object -TypeName Xml.XmlNodeReader -ArgumentList $xaml )
    {
        try
        {
            if( $Form = [Windows.Markup.XamlReader]::Load( $reader ) )
            {
                $xaml.SelectNodes( '//*[@Name]' ) | . { Process `
                {
                    Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName( $_.Name ) -Scope Script
                }}
            }
        }
        catch
        {
            Write-Error "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$($_.Exception.InnerException)"
            $form = $null
        }
    }

    $form ## return
}


Function Set-RemoteSessionProperties
{
    [CmdletBinding()]

    Param
    (
        ## relying on variables in parent scope being available here as this function was originally the main code body and later moved to a function to support the GUI
        [string]$connectTo
    )
    
    if( -Not $wpfcomboboxComputer.SelectedItem -and [string]::IsNullOrEmpty( $wpfcomboboxComputer.Text ) -and [string]::IsNullOrEmpty( $connectTo ) )
    {
        [void][Windows.MessageBox]::Show( 'No Computer Selected' , 'Select a computer or enter a name/address' , 'Ok' ,'Error' )
    }
    elseif( ( $WPFdatagridDisplays.SelectedItems.Count -ne 1 -and $activeDisplaysWithMonitors.Count -gt 1 ) -and -not $WPFchkboxPrimary.IsChecked )
    {
        [void][Windows.MessageBox]::Show( 'No Monitor Selected' , 'Select a Monitor' , 'Ok' ,'Error' )
    }
    elseif( $WPFradioPercentage.IsChecked -and [string]::IsNullOrEmpty( $WPFtxtboxScreenPercentage.Text ))
    {
        [void][Windows.MessageBox]::Show( 'No Screen Percentage Entered' , 'Enter a screen percentage' , 'Ok' ,'Error' )
    }
    elseif( $wpfradioWidthHeight.IsChecked -and [string]::IsNullOrEmpty( $wpftxtboxWidthHeight.Text ))
    {
        [void][Windows.MessageBox]::Show( 'No Width & Height Entered' , 'Enter width and height' , 'Ok' ,'Error' )
    }
    else
    {
        if( -Not ( $chosen = $WPFdatagridDisplays.SelectedItems ) )
        {
            $chosen = $WPFdatagridDisplays.Items[0] ## no monitor selected but only one monitor
        }
        if( [string]::IsNullOrEmpty( $connectTo ) )
        {
            $address = $wpfcomboboxComputer.Text
        }
        {
            $address = $connectTo
        }

        if( $wpfcomboboxComputer.SelectedIndex -lt 0 ) ## manually entered
        {
            [bool]$alreadyPresent = $false

            ForEach( $item in $wpfcomboboxComputer.Items )
            {
                if( $alreadyPresent = $item -ieq $address )
                {
                    break
                }
            }
            if( -not $alreadyPresent )
            {
                $wpfcomboboxComputer.Items.Insert( 0 , $address ) ## TODO should we resort it ? Need to check if already there
            }
        }

        $noMove = $wpfchkboxNoMove.IsChecked
        $usemsrdc = $WPFchkboxmsrdc.IsChecked

        ## clear from previous runs
        $fullScreen = $false
        $widthHeight = $null
        $percentage = $null
        $widthHeight = $null
        $xy = $null

        if( $wpfradioFullScreen.IsChecked )
        {
            $fullScreen = $true
        }
        elseif( $wpfradioPercentage.IsChecked )
        {
            $percentage = ($wpftxtboxScreenPercentage.Text -replace '%')
        }
        elseif( $wpfradioWidthHeight.IsChecked )
        {
            $widthHeight = $wpftxtboxWidthHeight.Text.Trim() -replace '[\&x\s,]' , ':' 
            $fullScreen = $fullScreen
        }
        elseif( $wpfradioFillScreen.IsChecked )
        {
            ## this is the default 
        }

        if( -Not [string]::IsNullOrEmpty( $wpftxtboxWindowPosition.Text ) )
        {
            $xy = $wpftxtboxWindowPosition.Text -replace ',' , ':'
        }

        $username = $wpftextboxUsername.Text
            
        $drivesToRedirect = $wpftxtboxDrivesToRedirect.Text

        if( $WPFchkboxPrimary.IsChecked )
        {
            $chosenDisplay = $activeDisplaysWithMonitors | Where-Object ScreenPrimary -eq $true
        }
        else
        {
            $chosenDisplay = $activeDisplaysWithMonitors | Where ScreenDeviceName -eq $chosen.ScreenDeviceName
        }
        if( -Not $chosenDisplay )
        {
            Write-Warning -Message "Failed to find device name $($chosen.ScreenDeviceName) in internal data"
        }
        else
        {
            New-RemoteSession
        }
    }
}

Function Set-WindowContent
{
    [CmdletBinding()]

    Param
    (
    )
    
    ## copy existing comment items so we can associate again where possible
    $itemsCopy = $null
    if( $WPFdatagridDisplays.Items -and $WPFdatagridDisplays.Items.Count -gt 0 )
    {
        $itemsCopy = New-Object -TypeName object[] -ArgumentList $WPFdatagridDisplays.Items.Count
        $WPFdatagridDisplays.Items.CopyTo( $itemsCopy , 0 )
    }

    $WPFdatagridDisplays.Clear()

    $datatable = New-Object -TypeName System.Data.DataTable

    ForEach( $property in ($activeDisplaysWithMonitors | Select-Object -Property $displayFields -First 1).Psobject.Properties )
    {
        if( $column = $Datatable.Columns.Add( $property.Name , [string] ) ) ##$property.TypeNameOfValue ) )
        {
            $column.ReadOnly = $true
        }
    }
    
    if( $column = $Datatable.Columns.Add( 'Comment' , [string] ) )
    {
        $column.ReadOnly = $false ## TODO persist to registry?
    }

    ForEach( $row in ( $activeDisplaysWithMonitors | Select-Object -Property $displayFields ))
    {
        ## check if previously had a comment and add if it did. ScreenDeviceName could be different as changes when docked/undocked/docked
        if( $itemsCopy -and $itemsCopy.Count -gt 0 )
        {
            ## we have to deal with a potential empty row
            [string]$comment = $itemsCopy | Where-Object { $_.PSobject -and $_.PSObject.Properties -and  $_.PSObject.Properties[ 'ScreenPrimary' ] -and $_.ScreenPrimary -eq $row.ScreenPrimary -and $_.Width -eq $row.Width -and $_.Height -eq $row.Height `
                -and $_.MonitorManufacturerName -eq $row.MonitorManufacturerName -and $_.MonitorManufacturerCode -eq $row.MonitorManufacturerCode -and $_.MonitorModel -eq $row.MonitorModel } | Select-Object -ExpandProperty Comment
            if( -Not [string]::IsNullOrEmpty( $comment ) )
            {
                Add-Member -InputObject $row -MemberType NoteProperty -Name Comment -Value $comment
            }
        }
        [void]$datatable.Rows.Add( @( $row.PSObject.Properties | Select-Object -ExpandProperty Value ) )
    }
  
    if( -Not [string]::IsNullOrEmpty( $percentage ) )
    {
        $WPFradioPercentage.IsChecked = $true
        $WPFtxtboxScreenPercentage.Text = $percentage
    }
    elseif( $fullScreen )
    {
        $WPFradioFullScreen.IsChecked = $true
    }
    elseif( -not [string]::IsNullOrEmpty( $widthHeight ) )
    {
        $WPFradioWidthHeight.IsChecked = $true
        $WPFtxtboxWidthHeight = $widthHeight -replace '[\s\&:]' , 'x'
    }
    else
    {
        $WPFradioFillScreen.IsChecked = $true
    }

    $wpftextboxUsername.Text = $username

    $WPFdatagridDisplays.ItemsSource = $datatable.DefaultView
    ##$WPFdatagridDisplays.IsReadOnly = $false
    $WPFdatagridDisplays.CanUserSortColumns = $true
    
    if( $null -ne $wpfcomboboxComputer.Items -and $wpfcomboboxComputer.Items.Count -eq 0 )
    {
        $previouslyUsed = @( Get-ItemProperty -Path 'HKCU:\SOFTWARE\Guy Leech\mstsc wrapper' -Name Computers -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Computers )
        if( $null -ne $previouslyUsed -and $previouslyUsed.Count -gt 0 )
        {
            ForEach( $value in $previouslyUsed )
            {
                [void]$wpfcomboboxComputer.Items.Add( $value )
            }
        }
        else
        {
            $mru = Get-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Terminal Server Client\Default' -Name MRU* -ErrorAction SilentlyContinue | Select-Object -Property MRU*
            if( $null -ne $mru )
            {
                ForEach( $value in ($mru.PSobject.properties | Select-Object -ExpandProperty Value | Sort-Object ))
                {
                    [void]$wpfcomboboxComputer.Items.Add( $value )
                }
            }
        }
    }

    $wpftxtboxDrivesToRedirect.Text = $drivesToRedirect
    $wpfchkboxNoMove.IsChecked = $noMove
    $wpfchkboxPrimary.IsChecked = $primary
    if( -Not [string]::IsNullOrEmpty( $xy ) )
    {
        $txtboxWindowPosition.Text = $xy -replace '[:]' , ','
    }
}

Function Get-DisplayInfo
{
    [CmdletBinding()]

    Param
    (
    )
    
    ## get display info for each display
             
    [array]$Screens = @( [system.windows.forms.screen]::AllScreens )

    Write-Verbose -Message "Got $($Screens.Count) screens"

    ## https://rakhesh.com/powershell/powershell-snippet-to-get-the-name-of-each-attached-monitor/
    [array]$monitors = @( Get-CimInstance -ClassName WmiMonitorID -Namespace root\wmi )

    $chosenDisplay = $null

    [array]$activeDisplaysDevices = @( [Resolution.Displays]::GetDisplays( 0 ) | Where-Object StateFlags -match 'AttachedToDesktop' )

    ForEach( $activeDisplaysDevice in $activeDisplaysDevices )
    {
        $result = [pscustomobject]$activeDisplaysDevice

        if( $monitorInfo = [Resolution.Displays]::GetDisplayDeviceInfo( $activeDisplaysDevice.DeviceName ) )
        {
            ForEach( $property in $monitorInfo.PSObject.Properties )
            {   
                if( $property.Name -ine 'cb' )
                {
                    Add-Member -InputObject $result -MemberType NoteProperty -Name "Monitor$($property.Name)" -Value $property.Value

                    if( $property.Name -ieq 'DeviceId' )
                    {
                        if( $property.Value -match '^MONITOR\\([^\\]+)\\' )
                        {
                            [string]$manufacturerCode = $Matches[1]
                            if( $monitor = $monitors | Where-Object InstanceName -Match "^DISPLAY\\$manufacturerCode\\" )
                            {
                                [string]$manufacturerCode = [System.Text.Encoding]::ASCII.GetString( $monitor.ManufacturerName )
                                [string]$manufacturer = $ManufacturerHash[ $manufacturerCode ] 
                                if( [string]::IsNullOrEmpty( $manufacturer ) )
                                {
                                    Write-Warning -Message "No monitor manufacturer for code $manufacturerCode"
                                    $manufacturer = $manufacturerCode
                                }
                                Add-Member -InputObject $result -NotePropertyMembers @{
                                    MonitorModel = $(if( [string]::IsNullOrEmpty( $monitor.UserFriendlyName ) ){ 'Generic/Unknown' } else { [System.Text.Encoding]::ASCII.GetString( $monitor.UserFriendlyName ) })
                                    MonitorManufacturerCode = $manufacturerCode
                                    MonitorManufacturerName = $manufacturer
                                    MonitorProductCodeId = [System.Text.Encoding]::ASCII.GetString( $monitor.ProductCodeID )
                                }
                            }
                        }
                    }
                }
            }
        }
        else
        {
            Write-Warning -Message "Failed to get monitor info for device $($activeDisplaysDevice.DeviceName)"
        }
        if( $screen = $screens | Where-Object DeviceName -ieq $activeDisplaysDevice.DeviceName )
        {
            ForEach( $property in $screen.PSObject.Properties )
            {   
                Add-Member -InputObject $result -MemberType NoteProperty -Name "Screen$($property.Name)" -Value $property.Value -Force
            }
        }
        else
        {
            Write-Warning -Message "Failed to get screen info for device $($activeDisplaysDevice.DeviceName)"
        }
        ## this gives us the actual resolution in dmPelsWidth & dmPelsHeight
        if( $activeDisplaysDevice.DeviceName -and ( $displaySettings = [Resolution.Displays]::GetCurrentDisplaySettings( $activeDisplaysDevice.DeviceName ) ) )
        {
            ForEach( $property in $displaysettings.PSObject.Properties )
            {   
                ## all properties start dm* so won't clash
                Add-Member -InputObject $result -MemberType NoteProperty -Name $property.Name -Value $property.Value -Force
            }
        }
        if( $result.PSObject.Properties[ 'cb' ] )
        {
            $result.PSObject.Properties.Remove( 'cb' )
        }
        $result
    }
}

## adapted from https://gist.github.com/mintsoft/22a5ae4cc68d3e51b2f2

$pinvokeCode = @" 
using System; 
using System.Runtime.InteropServices; 
using System.Collections.Generic;
namespace Resolution 
{ 
    [StructLayout(LayoutKind.Sequential)] 
    public struct DEVMODE1 
    { 
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] 
        public string dmDeviceName; 
        public short dmSpecVersion; 
        public short dmDriverVersion; 
        public short dmSize; 
        public short dmDriverExtra; 
        public int dmFields; 
        public short dmOrientation; 
        public short dmPaperSize; 
        public short dmPaperLength; 
        public short dmPaperWidth; 
        public short dmScale; 
        public short dmCopies; 
        public short dmDefaultSource; 
        public short dmPrintQuality; 
        public short dmColor; 
        public short dmDuplex; 
        public short dmYResolution; 
        public short dmTTOption; 
        public short dmCollate; 
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] 
        public string dmFormName; 
        public short dmLogPixels; 
        public short dmBitsPerPel; 
        public int dmPelsWidth; 
        public int dmPelsHeight; 
        public int dmDisplayFlags; 
        public int dmDisplayFrequency; 
        public int dmICMMethod; 
        public int dmICMIntent; 
        public int dmMediaType; 
        public int dmDitherType; 
        public int dmReserved1; 
        public int dmReserved2; 
        public int dmPanningWidth; 
        public int dmPanningHeight; 
    }; 
	
	[Flags()]
	public enum DisplayDeviceStateFlags : int
	{
		/// <summary>The device is part of the desktop.</summary>
		AttachedToDesktop = 0x1,
		MultiDriver = 0x2,
		/// <summary>The device is part of the desktop.</summary>
		PrimaryDevice = 0x4,
		/// <summary>Represents a pseudo device used to mirror application drawing for remoting or other purposes.</summary>
		MirroringDriver = 0x8,
		/// <summary>The device is VGA compatible.</summary>
		VGACompatible = 0x10,
		/// <summary>The device is removable; it cannot be the primary display.</summary>
		Removable = 0x20,
		/// <summary>The device has more display modes than its output devices support.</summary>
		ModesPruned = 0x8000000,
		Remote = 0x4000000,
		Disconnect = 0x2000000
	}
	[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Ansi)]
	public struct DISPLAY_DEVICE 
	{
		  [MarshalAs(UnmanagedType.U4)]
		  public int cb;
		  [MarshalAs(UnmanagedType.ByValTStr, SizeConst=32)]
		  public string DeviceName;
		  [MarshalAs(UnmanagedType.ByValTStr, SizeConst=128)]
		  public string DeviceString;
		  [MarshalAs(UnmanagedType.U4)]
		  public DisplayDeviceStateFlags StateFlags;
		  [MarshalAs(UnmanagedType.ByValTStr, SizeConst=128)]
		  public string DeviceID;
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst=128)]
		  public string DeviceKey;
	}
    public class User_32 
    { 
        [DllImport("user32.dll", SetLastError=true)] 
        public static extern int EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE1 devMode); 
        [DllImport("user32.dll", SetLastError=true)] 
        public static extern int ChangeDisplaySettings(ref DEVMODE1 devMode, int flags); 
		[DllImport("user32.dll", SetLastError=true)]
		public static extern bool EnumDisplayDevices(string lpDevice, uint iDevNum, ref DISPLAY_DEVICE lpDisplayDevice, uint dwFlags);
        public const int ENUM_CURRENT_SETTINGS = -1; 
        public const int CDS_UPDATEREGISTRY = 0x01; 
        public const int CDS_TEST = 0x02; 
        public const int DISP_CHANGE_SUCCESSFUL = 0; 
        public const int DISP_CHANGE_RESTART = 1; 
        public const int DISP_CHANGE_FAILED = -1; 
    } 
    public class Displays
    {
		public static IList<string> GetDisplayNames( )
		{
			var returnVals = new List<string>();
			for(var x=0U; x<1024; ++x)
			{
				DISPLAY_DEVICE outVar = new DISPLAY_DEVICE();
				outVar.cb = (short)Marshal.SizeOf(outVar);
				if(User_32.EnumDisplayDevices(null, x, ref outVar, 1U ))
				{
					returnVals.Add(outVar.DeviceName);
				}
			}
			return returnVals;
		}
		
        // added by Guy Leech in order to get all properties returned from EnumDisplayDevices()
		public static IList<object> GetDisplays( uint flags = 0 )
		{
			var returnVals = new List<object>();
			for(var x=0U; x<1024; ++x)
			{
				DISPLAY_DEVICE outVar = new DISPLAY_DEVICE();
				outVar.cb = (short)Marshal.SizeOf(outVar);
				if(User_32.EnumDisplayDevices(null, x, ref outVar, flags ))
				{
					returnVals.Add(outVar);
				}
			}
			return returnVals;
		}
		
        // added by Guy Leech in order to get properties for a specific device returned from EnumDisplayDevices()
		public static DISPLAY_DEVICE GetDisplayDeviceInfo( string deviceName , uint flags = 0 )
		{
			DISPLAY_DEVICE displayDevice = new DISPLAY_DEVICE() ;
            
            displayDevice.cb = (int)Marshal.SizeOf(displayDevice); 

			if( ! User_32.EnumDisplayDevices( deviceName , 0, ref displayDevice, flags ))
			{
				displayDevice.cb = 0 ;
			}

			return displayDevice ;
		}
		
		public static string GetCurrentResolution(string deviceName)
        {
            string returnValue = null;
            DEVMODE1 dm = GetDevMode1();
            if (0 != User_32.EnumDisplaySettings(deviceName, User_32.ENUM_CURRENT_SETTINGS, ref dm))
            {
                returnValue = dm.dmPelsWidth + "," + dm.dmPelsHeight;
            }
            return returnValue;
        }
		
        // added by Guy Leech in order to get all properties returned from EnumDisplaySettings()
		public static DEVMODE1 GetCurrentDisplaySettings(string deviceName)
        {
            DEVMODE1 dm = GetDevMode1();
            if (0 == User_32.EnumDisplaySettings(deviceName, User_32.ENUM_CURRENT_SETTINGS, ref dm))
            {
                dm.dmSize = 0 ; // denotes call failed
            }
            return dm;
        }

		public static IList<string> GetResolutions()
		{
			var displays = GetDisplayNames();
			var returnValue = new List<string>();
			foreach(var display in displays)
			{
				returnValue.Add(GetCurrentResolution(display));
			}
			return returnValue;
		}
		
        private static DEVMODE1 GetDevMode1() 
        { 
            DEVMODE1 dm = new DEVMODE1(); 
            dm.dmDeviceName = new String(new char[32]); 
            dm.dmFormName = new String(new char[32]); 
            dm.dmSize = (short)Marshal.SizeOf(dm); 
            return dm; 
        } 
    }
} 
"@

Function Add-VMwareVMsToListView
{
    Param
    (
        [string]$filter ,
        [bool]$regex
    )
    $vmwareError = $null
    $script:vms = @( Get-VM -ErrorVariable vmwareError -norecursion | Where-Object { $_.PowerState -ieq 'PoweredOn' -and (( $regex -and $_.Name -match $filter ) -or ( -Not $regex -and $_.Name -like $filter )) } | Sort-Object -Property Name | Select-Object -ExpandProperty Guest )
    if( $vmwareError )
    {
        [void][Windows.MessageBox]::Show( $vmwareError , 'VMware Error' , 'Ok' ,'Error' )
    }
    Write-Verbose -Message "Got $($vms.Count) powered on VMs"
    $WPFlistViewVMwareVMs.Items.Clear()

    ForEach( $vm in $vms )
    {
        $WPFlistViewVMwareVMs.Items.Add( $vm.VmName )
    }
    $WPFlabelVMwareVMs.Content = "$($WPFlistViewVMwareVMs.Items.Count) VMs"
}

Function Start-RemoteVMwareSession
{
    Param
    (
    )

    if( $WPFlistViewVMwareVMs.SelectedIndex -ge 0 )
    {
        if( $null -eq $script:vms -or $WPFlistViewVMwareVMs.SelectedIndex -gt $script:vms.Count )
        {
            Write-Error -Message "Internal error : selected grid view index $($WPFlistViewVMwareVMs.SelectedIndex) greater than $(($script:vms|Measure-Object).Count)"
        }
        else
        {
            $vm = $script:vms[ $WPFlistViewVMwareVMs.SelectedIndex ]
            $address = $null
            if( $WPFradioButtonConnectByIP.IsChecked )
            {
                ## TODO do we allow IPv6 ?
                $address = $vm.IPaddress | Where-Object { $_ -match '^\d+\.' -and $_ -ne '127.0.0.1' -and $_ -ne '::1' -and $_ -notmatch '^169\.254\.' }
                if( $null -eq $address )
                {
                    [void][Windows.MessageBox]::Show( "No IP address for $($vm.VmName)" , 'VMware Error' , 'Ok' ,'Error' )
                }
                elseif( $address -is [array] -and $address.Count -gt 1 )
                {
                    [void][Windows.MessageBox]::Show( "$($address.Count) IP addresses for $($vm.VmName)" , 'VMware Error' , 'Ok' ,'Error' )
                    ## TODO do we ask them to select one? Try in turn?
                    $address = $null
                }
            }
            else
            {
                $address = $vm.Hostname
            }
            if( $address )
            {
                if( -Not [string]::IsNullOrEmpty( $WPFtextBoxVMwareRDPPort.Text ) )
                {
                    if( $WPFtextBoxVMwareRDPPort.Text -notmatch '^\d+$' )
                    {
                        [void][Windows.MessageBox]::Show( "Port `"$($WPFtextBoxVMwareRDPPort.Text)`" is invalid" , 'VMware Error' , 'Ok' ,'Error' )
                        $address = $null
                    }
                    else
                    {
                        $address = "$($address):$($WPFtextBoxVMwareRDPPort.Text)"
                    }
                }
                Write-Verbose -Message "Connecting to VM $address"
                if( $address )
                {
                    ## put into main computer list if not already there
                    if( -Not $wpfcomboboxComputer.Items.Contains( $address ) )
                    {
                        $wpfcomboboxComputer.Items.Add( $address )
                    }
                    Set-RemoteSessionProperties -connectTo $address
                }
                else
                {
                    Write-Warning -Message "No address for VMware VM $($WPFlistViewVMwareVMs.SelectedItem)"
                }
            }
        }
    }
    else
    {
        [void][Windows.MessageBox]::Show( "No VM selected" , 'VMware Error' , 'Ok' ,'Error' )
    }
    }

#endregion pre-main

if( $usemsrdc -and [string]::IsNullOrEmpty( $address ) )
{
    Throw "Must specify computer to connect to via -address when using msrdc mode"
}

try
{
    Add-Type -TypeDefinition $pinvokeCode
}
catch
{
    ## hopefully because already loaded
}

Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Windows.Forms,System.Drawing

$script:activeDisplaysWithMonitors = @( Get-DisplayInfo )

if( $showDisplays )
{
    $activeDisplaysWithMonitors
    exit 0
}
if( $showManufacturerCodes )
{
    $ManufacturerHash.GetEnumerator() | Select-Object -Property @{n='Manufacturer';e={$_.Value}},@{n='Code';e={$_.Name}} | Sort-Object -Property Manufacturer
    exit 0
}

## don't need it yet but need it before we start the GUI

[string]$windowTypes = @'
    using System;
    using System.Runtime.InteropServices;
    
    [StructLayout(LayoutKind.Sequential)]

    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public static class user32
    {
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow); 
            
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow); 
            
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsIconic(IntPtr hWnd); 
        
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsZoomed(IntPtr hWnd); 
        
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd); 
        
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowUnicode(IntPtr hWnd); 
        
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
   
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos( IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    }
'@ 

try
{
    Add-Type -TypeDefinition $windowTypes
}
catch
{
    ## hopefully because we already have it
}

## https://www.linkedin.com/pulse/fun-powershell-finding-suspicious-cmd-processes-britton-manahan/

$TypeDef = @"

using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Api
{

 public class WinStruct
 {
   public string WinTitle {get; set; }
   public int WinHwnd { get; set; }
   public int PID { get; set; }
 }

 public class ApiDef
 {
   private delegate bool CallBackPtr(int hwnd, int lParam);
   private static CallBackPtr callBackPtr = Callback;
   private static List<WinStruct> _WinStructList = new List<WinStruct>();

   [DllImport("User32.dll")]
   [return: MarshalAs(UnmanagedType.Bool)]
   private static extern bool EnumWindows(CallBackPtr lpEnumFunc, IntPtr lParam);

   [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
   static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
   
   [DllImport("user32.dll")]
   static extern bool IsWindowVisible(IntPtr hWnd);
   
    [DllImport("user32.dll")]
    public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
 
   private static bool Callback(int hWnd, int pid)
   {
        if( IsWindowVisible( (IntPtr)hWnd ) )
        {
            int ipid = 0 ;
            GetWindowThreadProcessId( (IntPtr)hWnd , out ipid );
            if( ipid == pid || pid < 0 ) // -1 will return all windows
            {
                StringBuilder sb = new StringBuilder(256);
                int res = GetWindowText((IntPtr)hWnd, sb, 256);
                _WinStructList.Add( new WinStruct { WinHwnd = hWnd, WinTitle = sb.ToString() , PID = ipid  });
            }
        }
        return true;
   }   

   public static List<WinStruct> GetWindows( int pid )
   {
      _WinStructList = new List<WinStruct>();
      EnumWindows(callBackPtr, (IntPtr)pid );
      return _WinStructList;
   }

 }
}
"@

try
{
    Add-Type -TypeDefinition $TypeDef -ErrorAction Stop
}
catch
{
    ## hopefully because we already have it
}

if( $primary )
{
    $chosenDisplay = $activeDisplaysWithMonitors | Where-Object ScreenPrimary
}
elseif( $PSBoundParameters.ContainsKey( 'displayModel' ) )
{
    if( -Not ( $chosenDisplay = $activeDisplaysWithMonitors | Where-Object MonitorModel -IMatch $displayModel ) )
    {
        Throw "No displays for model `"$displayModel`" out of $(($activeDisplaysWithMonitors | Select-Object -ExpandProperty MonitorModel) -join ',')"
    }
    elseif( $chosenDisplay -is [array] -and $chosenDisplay.Count -gt 1 )
    {
        Throw "Multiple monitors for model `"$displayModel`""
    }
}
elseif( $PSBoundParameters[ 'displayManufacturer' ] )
{
    [string]$displayManufacturerCode = 'NONE'

    $displayManufacturerCodes = $ManufacturerHash.GetEnumerator() | Where-Object Value -match $displayManufacturer | Select-Object -ExpandProperty Name
    if( -Not $displayManufacturerCodes )
    {
        Throw "No monitor manufacturer found matching `"$displayManufacturer`""
    }
    elseif( $displayManufacturerCodes -is [array] -and $displayManufacturerCodes.Count -gt 1 )
    {
        Throw "Found ($displayManufacturerCodes.Count) manufacturer codes matching `"$displayManufacturer`" - use code instead - $($displayManufacturerCodes -join ' or ')"
    }
    else
    {
        $displayManufacturerCode = $displayManufacturerCodesG
    }
  
    if( -Not ( $chosenDisplay = $activeDisplaysWithMonitors | Where-Object MonitorManufacturerName -ieq $displayManufacturerCode ) )
    {
        Throw "No displays for manufacturer code `"$displayManufacturerCode`" ($displayManufacturer) out of $(($activeDisplaysWithMonitors | Select-Object -ExpandProperty MonitorManufacturerName) -join ',')"
    }
    elseif( $chosenDisplay -is [array] -and $chosenDisplay.Count -gt 1 )
    {
        Throw "Multiple monitors for manufacturer code `"$displayManufacturerCode`" ($displayManufacturer) - try model number?"
    }
}
elseif( $PSBoundParameters[ 'displayManufacturerCode' ] )
{
    if( -Not ( $chosenDisplay = $activeDisplaysWithMonitors | Where-Object MonitorManufacturerName -ieq $displayManufacturerCode ) )
    {
        Throw "No displays for manufacturer code `"$displayManufacturerCode`" out of $(($activeDisplaysWithMonitors | Select-Object -ExpandProperty MonitorManufacturerName) -join ',')"
    }
    elseif( $chosenDisplay -is [array] -and $chosenDisplay.Count -gt 1 )
    {
        Throw "Multiple monitors for manufacturer code `"$displayManufacturerCode`" - try model number?"
    }
}
else ## if not passed displayNumber or displaymanufacturer , display a GUI with the choices
{
    [array]$displayFields = @( 'ScreenPrimary','ScreenDeviceName',@{n='Width';e={$_.dmPelswidth}},@{n='Height';e={$_.dmPelsHeight}},'MonitorManufacturerName','MonitorManufacturerCode','MonitorModel' )
    if( $usegridviewpicker )
    {
        if( -Not ( $chosen = $activeDisplaysWithMonitors | Select-Object -Property $displayFields | Out-GridView -Title "Select monitor for $exe" -PassThru ) )
        {
            Throw "Please select a monitor"
        }
    }
    else
    {
        if( -Not ( $mainWindow = New-WPFWindow -inputXAML $mainwindowXAML ) )
        {
            Throw 'Failed to create WPF from XAML'
        }
        
        Set-WindowContent
        
        $wpfbtnRefresh.Add_Click({
            $_.Handled = $true
            Write-Verbose "Refresh clicked"
            $script:activeDisplaysWithMonitors = @( Get-DisplayInfo )
            Set-WindowContent
        })

        $wpfbtnLaunch.IsDefault = $true
        $wpfbtnLaunch.Add_Click({
            $_.Handled = $true
            Write-Verbose "Launch clicked"
            Set-RemoteSessionProperties
        })
        
        $WPFbtnLaunchMstscOptions.Add_Click({
            $_.Handled = $true
            Write-Verbose "Launch on mstsc options clicked"
            Set-RemoteSessionProperties
        })

        $WPFbtnLaunchOtherOptions.Add_Click({
            $_.Handled = $true
            Write-Verbose "Launch on other options clicked"
            Set-RemoteSessionProperties
        })
        
        $WPFbtnLaunchVMwareOptions.Add_Click({
            $_.Handled = $true
            Write-Verbose "Launch on VMware clicked"
            Start-RemoteVMwareSession
        })

        $WPFlistViewVMwareVMs.add_MouseDoubleClick({
            $_.Handled = $true
            Write-Verbose "Launch on VMware list item double clicked"
            Start-RemoteVMwareSession
        })
        
        $WPFbuttonVMwareConnect.Add_Click({
            $_.Handled = $true
            Write-Verbose "VMware Connect clicked"
            if( -Not [string]::IsNullOrEmpty( $WPFtextBoxVMwarevCenter.Text ) )
            {
                Import-Module -Name VMware.VimAutomation.Core
                $script:vmwareConnection = Connect-VIServer -Server $WPFtextBoxVMwarevCenter.Text -Force
                if( -Not $script:vmwareConnection )
                {
                    [void][Windows.MessageBox]::Show( "Failed to connect to $($WPFtextBoxVMwarevCenter.Text)" , 'VMware Error' , 'Ok' ,'Error' )
                }
            }
            Add-VMwareVMsToListView -filter $wpfTextBoxVMwareFilter.Text -regex $WPFcheckBoxVMwareRegEx.IsChecked
        })
        
        $WPFbuttonVMwareDisconnect.Add_Click({
            $_.Handled = $true
            Write-Verbose "VMware Disconnect clicked"
            if( $script:vmwareConnection )
            {
                Import-Module -Name VMware.VimAutomation.Core
                $disconnection = Disconnect-VIServer -Force -Confirm:$false
                $script:vmwareConnection = $null
                $WPFlistViewVMwareVMs.Items.Clear()
            }
            else
            {
                [void][Windows.MessageBox]::Show( "Not connected" , 'VMware Error' , 'Ok' ,'Error' )
            }
        })

        $WPFbuttonApplyFilter.Add_Click({
            $_.Handled = $true
            Write-Verbose "VMware Apply Filter clicked"
            Add-VMwareVMsToListView -filter $wpfTextBoxVMwareFilter.Text -regex $WPFcheckBoxVMwareRegEx.IsChecked
        })


        $WPFdatagridDisplays.add_MouseDoubleClick({
            $_.Handled = $true
            Write-Verbose "Grid item double clicked"
            $script:activeDisplaysWithMonitors = @( Get-DisplayInfo )
            Set-RemoteSessionProperties
        })
        
        ## so enter key can launch rather than move to next grid line
        $WPFdatagridDisplays.add_PreviewKeyDown({
            Param
            (
              [Parameter(Mandatory)][Object]$sender,
              [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$event
            )
            if( $event -and $event.Key -ieq 'return' )
            {
                $_.Handled = $true
                $script:activeDisplaysWithMonitors = @( Get-DisplayInfo )
                Set-RemoteSessionProperties
            }    
        })

        $WPFtxtboxWidthHeight.add_GotFocus({
            $_.Handled = $true
            $WPFradioWidthHeight.IsChecked = $true
        })
 
        $WPFtxtboxScreenPercentage.add_GotFocus({
            $_.Handled = $true
            $WPFradioPercentage.IsChecked = $true
        })
        
        if( $rdpoptions = Get-ItemProperty -Path $configKey -Name 'RDPOptions' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'RDPOptions'  )
        {
            $WPFtxtBoxOtherOptions.Text = $rdpoptions
        }
        
        $mainWindow.add_KeyDown({
            Param
            (
              [Parameter(Mandatory)][Object]$sender,
              [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$event
            )
            if( $event -and $event.Key -ieq 'F5' )
            {
                $_.Handled = $true
                $script:activeDisplaysWithMonitors = @( Get-DisplayInfo )
                Set-WindowContent
            }    
        })

        $mainWindow.Add_Loaded({
            $_.Handled = $true
            if( $_.Source -and $_.Source.WindowState -ieq 'Minimized' )
            {
                $_.Source.WindowState = 'Normal'
            }
        })

        $wpfchkboxDoNotApply.IsChecked = (-Not $useOtherOptions ) ## don't want it on by default as passes blank password to mstsc giving failed logon

        $guiResult = $mainWindow.ShowDialog()

        if( -Not [string]::IsNullOrEmpty( $WPFtxtBoxOtherOptions.Text ) -and -Not $wpfchkboxDoNotSave.IsChecked )
        {
            if( -Not ( Get-ChildItem -Path $configKey -ErrorAction SilentlyContinue ) -and -Not (New-Item -Path $configKey -ItemType Key -Force) )
            {
                Write-Warning -Message "Failed to create `"$configKey`""
            }
            if( -Not ( Set-ItemProperty -Path $configKey -Name 'RDPOptions' -Value $WPFtxtBoxOtherOptions.Text -PassThru -Force -Type MultiString))
            {
                Write-Warning -Message "Problem writing RDP options to `"$configKey`""
            }
        }

        ## persist computers to registry
        Set-ItemProperty -Path $configKey -Name Computers -Value ([string[]]@( $wpfcomboboxComputer.Items.GetEnumerator() | Sort-Object -Unique )) -Force

        exit $guiResult
    }

    if( $chosenDisplay -is [array] -and $chosenDisplay.Count -gt 1 )
    {
        Throw "Spanning monitors not yet supported"
    }
    $chosenDisplay = $activeDisplaysWithMonitors | Where ScreenDeviceName -eq $chosen.ScreenDeviceName
    if( -Not $chosenDisplay )
    {
        Throw "Failed to find device name $($chosen.ScreenDeviceName) in internal data"
    }
}

New-RemoteSession -rethrow

# SIG # Begin signature block
# MIIjcAYJKoZIhvcNAQcCoIIjYTCCI10CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUzQ6rnC97V23dNdhUhkKEvY+t
# jMyggh2OMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
# AQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcNMjgxMDIyMTIwMDAwWjByMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQg
# Q29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# +NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid2zLXcep2nQUut4/6kkPApfmJ
# 1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sjlOSRI5aQd4L5oYQjZhJUM1B0
# sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjfDPXiTWAYvqrEsq5wMWYzcT6s
# cKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzLfnQ5Ng2Q7+S1TqSp6moKq4Tz
# rGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR93O8vYWxYoNzQYIH5DiLanMg
# 0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckwEgYDVR0TAQH/BAgwBgEB/wIB
# ADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMweQYIKwYBBQUH
# AQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYI
# KwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFz
# c3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaG
# NGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0
# dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYIYIZIAYb9bAMwHQYDVR0OBBYE
# FFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1aJLPzItEVyCx8JSl2qB1dHC06
# GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUPUbRupY5a4l4kgU4QpO4/cY5j
# DhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1UcKNJK4kxscnKqEpKBo6cSgC
# PC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjFEmifz0DLQESlE/DmZAwlCEIy
# sjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM1ekN3fYBIM6ZMWM9CBoYs4Gb
# T8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhsRDKyZqHnGKSaZFHvMIIFTzCC
# BDegAwIBAgIQBP3jqtvdtaueQfTZ1SF1TjANBgkqhkiG9w0BAQsFADByMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29k
# ZSBTaWduaW5nIENBMB4XDTIwMDcyMDAwMDAwMFoXDTIzMDcyNTEyMDAwMFowgYsx
# CzAJBgNVBAYTAkdCMRIwEAYDVQQHEwlXYWtlZmllbGQxJjAkBgNVBAoTHVNlY3Vy
# ZSBQbGF0Zm9ybSBTb2x1dGlvbnMgTHRkMRgwFgYDVQQLEw9TY3JpcHRpbmdIZWF2
# ZW4xJjAkBgNVBAMTHVNlY3VyZSBQbGF0Zm9ybSBTb2x1dGlvbnMgTHRkMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAr20nXdaAALva07XZykpRlijxfIPk
# TUQFAxQgXTW2G5Jc1YQfIYjIePC6oaD+3Zc2WN2Jrsc7bj5Qe5Nj4QHHHf3jopLy
# g8jXl7Emt1mlyzUrtygoQ1XpBBXnv70dvZibro6dXmK8/M37w5pEAj/69+AYM7IO
# Fz2CrTIrQjvwjELSOkZ2o+z+iqfax9Z1Tv82+yg9iDHnUxZWhaiEXk9BFRv9WYsz
# qTXQTEhv8fmUI2aZX48so4mJhNGu7Vp1TGeCik1G959Qk7sFh3yvRugjY0IIXBXu
# A+LRT00yjkgMe8XoDdaBoIn5y3ZrQ7bCVDjoTrcn/SqfHvhEEMj1a1f0zQIDAQAB
# o4IBxTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0O
# BBYEFE16ovlqIk5uX2JQy6og0OCPrsnJMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUE
# DDAKBggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUw
# QzA3BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNl
# cnQuY29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcw
# AYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8v
# Y2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNp
# Z25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAQEAU9zO
# 9UpTkPL8DNrcbIaf1w736CgWB5KRQsmp1mhXbGECUCCpOCzlYFCSeiwH9MT0je3W
# aYxWqIpUMvAI8ndFPVDp5RF+IJNifs+YuLBcSv1tilNY+kfa2OS20nFrbFfl9QbR
# 4oacz8sBhhOXrYeUOU4sTHSPQjd3lpyhhZGNd3COvc2csk55JG/h2hR2fK+m4p7z
# sszK+vfqEX9Ab/7gYMgSo65hhFMSWcvtNO325mAxHJYJ1k9XEUTmq828ZmfEeyMq
# K9FlN5ykYJMWp/vK8w4c6WXbYCBXWL43jnPyKT4tpiOjWOI6g18JMdUxCG41Hawp
# hH44QHzE1NPeC+1UjTCCBY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJ
# KoZIhvcNAQEMBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IElu
# YzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQg
# QXNzdXJlZCBJRCBSb290IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1
# OVowYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBS
# b290IEc0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+Rd
# SjwwIjBpM+zCpyUuySE98orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20d
# q7J58soR0uRf1gU8Ug9SH8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7f
# gvMHhOZ0O21x4i0MG+4g1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRA
# X7F6Zu53yEioZldXn1RYjgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raR
# mECQecN4x7axxLVqGDgDEI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzU
# vK4bA3VdeGbZOjFEmjNAvwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2
# mHY9WV1CdoeJl2l6SPDgohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkr
# fsCUtNJhbesz2cXfSwQAzH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaA
# sPvoZKYz0YkH4b235kOkGLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxf
# jT/JvNNBERJb5RBQ6zHFynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEe
# xcCPorF+CiaZ9eRpL5gdLfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQF
# MAMBAf8wHQYDVR0OBBYEFOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaA
# FEXroq/0ksuCMS1Ri6enIZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcB
# AQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggr
# BgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNz
# dXJlZElEUm9vdENBLmNydDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQK
# MAgwBgYEVR0gADANBgkqhkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3
# v1cHvZqsoYcs7IVeqRq7IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy
# 3iS8UgPITtAq3votVs/59PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cn
# RNTnf+hZqPC/Lwum6fI0POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3
# WlxUjG/voVA9/HYJaISfb8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2
# zm8jLfR+cWojayL/ErhULSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCC
# Bq4wggSWoAMCAQICEAc2N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAwYjEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0
# MB4XDTIyMDMyMzAwMDAwMFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMCVVMx
# FzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVz
# dGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJKoZI
# hvcNAQEBBQADggIPADCCAgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh1tKD
# 0Z5Mom2gsMyD+Vr2EaFEFUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+FeoAn39
# Q7SE2hHxc7Gz7iuAhIoiGN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1decf
# BmWNlCnT2exp39mQh0YAe9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxndX7RU
# CyFobjchu0CsX7LeSn3O9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6Th+x
# tVhNef7Xj3OTrCw54qVI1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPjQ2OA
# e3VuJyWQmDo4EbP29p7mO1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlRErWHRA
# KKtzQ87fSqEcazjFKfPKqpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JMq++b
# Pf4OuGQq+nUoJEHtQr8FnGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh3pP+
# OcD5sjClTNfpmEpYPtMDiP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8ju2Tj
# Y+Cm4T72wnSyPx4JduyrXUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnSDmuZ
# DNIztM2xAgMBAAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQW
# BBS6FtltTYUvcyl2mi91jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/
# 57qYrhwPTzAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYI
# KwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFRydXN0ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9j
# cmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1Ud
# IAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEA
# fVmOwJO2b5ipRCIBfmbW2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp/GnB
# zx0H6T5gyNgL5Vxb122H+oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40BIiXO
# lWk/R3f7cnQU1/+rT4osequFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2dfNBw
# CnzvqLx1T7pa96kQsl3p/yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibBt94q
# 6/aesXmZgaNWhqsKRcnfxI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7T6NJ
# uXdmkfFynOlLAlKnN36TU6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZAmyEh
# QNC3EyTN3B14OuSereU0cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdBeHo4
# 6Zzh3SP9HSjTx/no8Zhf+yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnKcPA3
# v5gA3yAWTyf7YGcWoWa63VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/pNHz
# V9m8BPqC3jLfBInwAM1dwvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yYlvZV
# VCsfgPrA8g4r5db7qS9EFUrnEw4d2zc4GqEr9u3WfPwwggbAMIIEqKADAgECAhAM
# TWlyS5T6PCpKPSkHgD1aMA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcw
# FQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3Rl
# ZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjIwOTIxMDAw
# MDAwWhcNMzMxMTIxMjM1OTU5WjBGMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGln
# aUNlcnQxJDAiBgNVBAMTG0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDIyIC0gMjCCAiIw
# DQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAM/spSY6xqnya7uNwQ2a26HoFIV0
# MxomrNAcVR4eNm28klUMYfSdCXc9FZYIL2tkpP0GgxbXkZI4HDEClvtysZc6Va8z
# 7GGK6aYo25BjXL2JU+A6LYyHQq4mpOS7eHi5ehbhVsbAumRTuyoW51BIu4hpDIjG
# 8b7gL307scpTjUCDHufLckkoHkyAHoVW54Xt8mG8qjoHffarbuVm3eJc9S/tjdRN
# lYRo44DLannR0hCRRinrPibytIzNTLlmyLuqUDgN5YyUXRlav/V7QG5vFqianJVH
# hoV5PgxeZowaCiS+nKrSnLb3T254xCg/oxwPUAY3ugjZNaa1Htp4WB056PhMkRCW
# fk3h3cKtpX74LRsf7CtGGKMZ9jn39cFPcS6JAxGiS7uYv/pP5Hs27wZE5FX/Nurl
# fDHn88JSxOYWe1p+pSVz28BqmSEtY+VZ9U0vkB8nt9KrFOU4ZodRCGv7U0M50GT6
# Vs/g9ArmFG1keLuY/ZTDcyHzL8IuINeBrNPxB9ThvdldS24xlCmL5kGkZZTAWOXl
# LimQprdhZPrZIGwYUWC6poEPCSVT8b876asHDmoHOWIZydaFfxPZjXnPYsXs4Xu5
# zGcTB5rBeO3GiMiwbjJ5xwtZg43G7vUsfHuOy2SJ8bHEuOdTXl9V0n0ZKVkDTvpd
# 6kVzHIR+187i1Dp3AgMBAAGjggGLMIIBhzAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0T
# AQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwHwYDVR0jBBgwFoAUuhbZbU2FL3MpdpovdYxqII+e
# yG8wHQYDVR0OBBYEFGKK3tBh/I8xFO2XC809KpQU31KcMFoGA1UdHwRTMFEwT6BN
# oEuGSWh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFJT
# QTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcmwwgZAGCCsGAQUFBwEBBIGDMIGA
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wWAYIKwYBBQUH
# MAKGTGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRH
# NFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcnQwDQYJKoZIhvcNAQELBQAD
# ggIBAFWqKhrzRvN4Vzcw/HXjT9aFI/H8+ZU5myXm93KKmMN31GT8Ffs2wklRLHiI
# Y1UJRjkA/GnUypsp+6M/wMkAmxMdsJiJ3HjyzXyFzVOdr2LiYWajFCpFh0qYQitQ
# /Bu1nggwCfrkLdcJiXn5CeaIzn0buGqim8FTYAnoo7id160fHLjsmEHw9g6A++T/
# 350Qp+sAul9Kjxo6UrTqvwlJFTU2WZoPVNKyG39+XgmtdlSKdG3K0gVnK3br/5iy
# JpU4GYhEFOUKWaJr5yI+RCHSPxzAm+18SLLYkgyRTzxmlK9dAlPrnuKe5NMfhgFk
# nADC6Vp0dQ094XmIvxwBl8kZI4DXNlpflhaxYwzGRkA7zl011Fk+Q5oYrsPJy8P7
# mxNfarXH4PMFw1nfJ2Ir3kHJU7n/NBBn9iYymHv+XEKUgZSCnawKi8ZLFUrTmJBF
# YDOA4CPe+AOk9kVH5c64A0JH6EE2cXet/aLol3ROLtoeHYxayB6a1cLwxiKoT5u9
# 2ByaUcQvmvZfpyeXupYuhVfAYOd4Vn9q78KVmksRAsiCnMkaBXy6cbVOepls9Oie
# 1FqYyJ+/jbsYXEP10Cro4mLueATbvdH7WwqocH7wl4R44wgDXUcsY6glOJcB0j86
# 2uXl9uab3H4szP8XTE0AotjWAQ64i+7m4HJViSwnGWH2dwGMMYIFTDCCBUgCAQEw
# gYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1
# cmVkIElEIENvZGUgU2lnbmluZyBDQQIQBP3jqtvdtaueQfTZ1SF1TjAJBgUrDgMC
# GgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYK
# KwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG
# 9w0BCQQxFgQUwjzHvXhKrkZ+9FihQPsw5gKBsDgwDQYJKoZIhvcNAQEBBQAEggEA
# UNDxcpmGsRhp8VheSU9nmQ2hDPSOFfdAJ9EU2qNhUCezkL3fOml28YF22qC5sEk8
# E+00+w+GQ+DVxPXYfH4ZkaxossZ6W+8wEJNjuxc00I1yGmIiFOqjxKM5nnSRWFcd
# PIF34BQU75fJkyKXeSh9Myazs9eHZ5OyvEAcfplm9VNqqxcrsp5omJQD/e9YYQSc
# dtw7ydCjVf0uoLcEzEVsC9ZFb5VWtlHPRdJs1V9hKxJoeUYxLy2A9zqjkorCTuXV
# xEmkgBbSKF25q1PClH7mabh/G9Q308IhHE4ovNx9QG+Sk5bpwNvxnY9nU3XIOcde
# x0H2GuNFDDuE6HdX7Dr/YKGCAyAwggMcBgkqhkiG9w0BCQYxggMNMIIDCQIBATB3
# MGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UE
# AxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBp
# bmcgQ0ECEAxNaXJLlPo8Kko9KQeAPVowDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG
# 9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yMzAyMjQxMDE1MDda
# MC8GCSqGSIb3DQEJBDEiBCC09JJUrP89tNsSwcZ4WSAkTZ4Z6ZqH6n6DAtKuIz10
# dDANBgkqhkiG9w0BAQEFAASCAgCBaJ+QglIDOtcX68yTFYqTzxWyiqBRLQZpIgXT
# ERDUoHTsSib5gd2FYZezkIGhMcOQ0z0wTt93z73dC9sd5lPOdOyB67pRaaHKjnJB
# Nl0GtFxxqiyv6n3mCeARq8LFiHC+FyL1gnlze4R3DoyQlYTrnnhKEmN2CS+ViO1F
# IebES0VphGAOptR93Fq0fPj/P6+7JIcDwU9VQhPGs3W+q4kj+u9whYf5g6w7lufG
# wBjNggpyvfhtd25eXByHkSuBvX+ccM159r/6J92WfA8ZQ8DAJMIVwpRiaimHshVy
# 9QJET9d62p7T8v+hY1bMSNnAVJsMDqSBtsVySPDEPjZv2axnVwG+qY2b22p+RiNg
# wRv9UT5qf4VlqrvvgYRsUvAEaBPKM1zKMJ7IPZ36Fusr06FIAbmchsDKeTd4ZTiA
# KP1rDxkm8Jl3rFxN+vNqr6Wh8RvrNC0lNTxfGbk4GlP/+qMkbvozPLGQoFqMZJOZ
# fj2regtqC2ANEZ86KTM8PE7asVkaxdw+yLlo4JQiv7iJ8AlSWsbyApSPRxMwr1T4
# mTqd1IEDV/aqZ1IFMrw3XIt5abPWI+KcCLk/VZUhA7Z1mjiPFUMx7uFkQnBb1Is6
# mQyUGClXlYwNnMyz/oSa7lmLwP7ADvOmkuAQjabf/J7UrSQ5nEFDTL+nYvwWJH6H
# uQlx6g==
# SIG # End signature block
