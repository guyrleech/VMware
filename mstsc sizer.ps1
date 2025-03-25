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
    2023/03/10 @guyrleech  Added looking for msrdc.exe in program files x86
    2024/09/18 @guyrleech  Added Hyper-V support
    2024/09/19 @guyrleech  Added Hyper-V console button
    2024/09/23 @guyrleech  Added context menu for Hyper-V VMs
    2024/10/04 @guyrleech  Added more Hyper-V context menu items
                           Changed temp rdpfile naming & location
                           Uses primary monitor if no monitor selected
    2024/11/08 @guyrleech  Reverse name in .rdp file
                           Fix window title
    2024/11/12 @guyrleech  -remove added to remove characters from address for icon display differentiation improvement functionality
                           Added Hyper-V Clear Filter button
    2024/11/13 @guyrleech  Fixed not finding existing mstsc window
    2024/12/10 @guyrleech  Added snapshot management dialog
    2024/12/16 @guyrleech  Fixed VMware VM list not showing names
    2025/02/24 @guyrleech  Fixed snapshot issues
    2025/03/14 @guyrleech  Re-enable support for msrdc
    2025/03/24 @guyrleech  msrdc (Windows (365) App, was Remote Desktop (store) app) autodetection and greyed out if not available
    2025/03/25 @guyrleech  No Hyper-V host specified causes it to use localhost

    ## TODO persist the "comment" column in memory so that it is available when undocked and redocked
    ## TODO make hypervisor operations async with a watcher thread
    ## TODO add history tab which is disabled by default (and thus audit)
    ## TODO add VMware console to that tab, make mstsc.exe configurable so could use with other exes
    ## TODO can we embed mstsx ax control so we can resize windows natively without mstsc.exe etc?

#>

<#
Copyright © 2025 Guy Leech

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
    [string]$hypervHost ,
    [switch]$primary ,
    [string]$percentage ,
    [switch]$usemsrdc ,
    [switch]$noFriendlyName ,
    [switch]$noResize , ## use mstsc with no width/height parameters
    [string]$widthHeight , ## colon delimited
    [string]$xy , ## colon delimited
    [string]$drivesToRedirect = '*' ,
    [string]$extraMsrdcParameters = '/SkipAvdSignatureChecks' ,
    [string]$msrdcCopyPath ,
    [string]$msrdcCopyFolder ,
    [string]$msrdcCopyName = 'Copy of msrdc' ,
    [switch]$usegridviewpicker ,
    [switch]$fullScreen ,
    [string]$youAreHereSnapshot = 'YouAreHere'  ,
    [switch]$showDisplays ,
    [switch]$showManufacturerCodes ,
    [switch]$useOtherOptions ,
    [switch]$noMove ,
    [switch]$reverse ,
    [string]$remove , ## ^GL([AH]V)?[SW]\d+
    [int]$windowWaitTimeSeconds = 20 ,
    [int]$pollForWindowEveryMilliseconds = 333 ,
    [string]$tempFolder = $(Join-Path -Path $env:temp -ChildPath 'Guy Leech mstsc Sizer') ,
    [string]$exe = 'mstsc.exe' ,
    [string]$configKey = 'HKCU:\SOFTWARE\Guy Leech\mstsc wrapper'
)

#region data

[array]$script:vms = $null
$script:vmwareConnection = $null
[array]$script:theseSnapshots = @()

# keep user added comments so can set when displays change
##$script:itemscopy = New-Object -TypeName System.Collections.Generic.List[object]

## https://docs.microsoft.com/en-us/windows-server/remote/remote-desktop-services/clients/rdp-files
## full address:s:$address ## removed since also using /v: so causes doubling of address in mstsc title bar
[string]$rdpTemplate = @'
desktopwidth:i:$width
desktopheight:i:$height
full address:s:$address
window title:s:$address
use multimon:i:0
screen mode id:i:$screenmode
dynamic resolution:i:1
smart sizing:i:0
drivestoredirect:s:$drivesToRedirect
'@

<#  from ChatGPT after asking it to make it resize properly

<Window x:Class="mstsc_msrdc_wrapper.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Guy's mstsc Wrapper Script" Height="500" Width="809">
    <Grid>
        <TabControl HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
            <TabItem Header="Main">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="*" />
                        <RowDefinition Height="Auto" />
                    </Grid.RowDefinitions>

                    <StackPanel Orientation="Horizontal" Grid.Row="0" Margin="10">
                        <Label Content="Computer" VerticalAlignment="Center" />
                        <ComboBox x:Name="comboboxComputer" Width="200" Margin="10,0,0,0" IsEditable="True" />
                        <CheckBox x:Name="chkboxPrimary" Content="Use primary monitor" Margin="10,0,0,0" VerticalAlignment="Center" />
                    </StackPanel>

                    <DataGrid x:Name="datagridDisplays" Grid.Row="1" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="10" SelectionMode="Single" />

                    <StackPanel Orientation="Horizontal" Grid.Row="2" HorizontalAlignment="Center" Margin="10">
                        <Button x:Name="btnLaunch" Content="_Launch" Width="96" Margin="5" />
                        <Button x:Name="btnRefresh" Content="_Refresh" Width="96" Margin="5" />
                        <Button x:Name="btnCreateShortcut" Content="_Create Shortcut" Width="96" Margin="5" />
                    </StackPanel>
                </Grid>
            </TabItem>

            <TabItem Header="Hyper-V">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="*" />
                        <RowDefinition Height="Auto" />
                    </Grid.RowDefinitions>

                    <Label Content="VMs" Grid.Row="0" HorizontalAlignment="Center" Margin="10" />

                    <ListView x:Name="listViewHyperVVMs" Grid.Row="1" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="10" SelectionMode="Multiple">
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}" />
                                <GridViewColumn Header="Power State" DisplayMemberBinding="{Binding PowerState}" />
                            </GridView>
                        </ListView.View>
                        <ListView.ContextMenu>
                            <ContextMenu>
                                <MenuItem Header="Power" x:Name="PowerContextMenu">
                                    <MenuItem Header="Power On" x:Name="HyperVPowerOnContextMenu" />
                                    <MenuItem Header="Power Off" x:Name="HyperVPowerOffContextMenu" />
                                </MenuItem>
                                <MenuItem Header="Config" x:Name="ConfigContextMenu">
                                    <MenuItem Header="Detail" x:Name="HyperVDetailContextMenu" />
                                    <MenuItem Header="Rename" x:Name="HyperVRenameMenu" />
                                </MenuItem>
                            </ContextMenu>
                        </ListView.ContextMenu>
                    </ListView>

                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center" Margin="10">
                        <Button x:Name="btnLaunchHyperVOptions" Content="_Launch" Width="96" Margin="5" />
                        <Button x:Name="btnLaunchHyperVConsole" Content="Console" Width="96" Margin="5" />
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
@'

#>

#>

[string]$mainwindowXAML = @'
<Window x:Class="mstsc_msrdc_wrapper.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:mstsc_msrdc_wrapper"
        mc:Ignorable="d"
        Title="Guy's mstsc Wrapper Script" Height="500" Width="809">
    <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
        <TabControl HorizontalAlignment="Stretch" Height="432" VerticalAlignment="Stretch" Width="768">
            <TabItem Header="Main">
                <Grid  HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
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
                    <CheckBox x:Name="chkboxmsrdc" Grid.Column="4" Content="Use msrdc instead of mstsc" HorizontalAlignment="Left" Height="21" Margin="145,189,0,0" VerticalAlignment="Top" Width="292" IsEnabled="true"/>
                    <ComboBox x:Name="comboboxComputer" Grid.Column="2" HorizontalAlignment="Left" Height="27" Margin="14,137,0,0" VerticalAlignment="Top" Width="254" IsEditable="True" IsDropDownOpen="False" Grid.ColumnSpan="3">
                        <ComboBox.ContextMenu>
                            <ContextMenu>
                                <MenuItem Header="Delete" x:Name="deleteComputersContextMenu"/>
                            </ContextMenu>
                        </ComboBox.ContextMenu>
                    </ComboBox>
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
                    <Button x:Name="btnLaunch" Content="_Launch" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="25" Margin="2,0,0,10" VerticalAlignment="Bottom" Width="96" Grid.Column="1"/>
                    <Button x:Name="btnRefresh" Content="_Refresh" HorizontalAlignment="Left" Height="25" Margin="71,0,0,10" VerticalAlignment="Bottom" Width="96" Grid.Column="2" Grid.ColumnSpan="2"/>
                    <Button x:Name="btnCreateShortcut" Content="_Create Shortcut" HorizontalAlignment="Left" Height="25" Margin="14,0,0,10" VerticalAlignment="Bottom" Width="96" Grid.Column="4"/>
                    <Label Content="User" HorizontalAlignment="Left" Height="38" Margin="14,181,0,0" VerticalAlignment="Top" Width="71" Grid.ColumnSpan="3"/>
                    <TextBox x:Name="textboxUsername" Grid.Column="2" HorizontalAlignment="Left" Height="27" Margin="14,186,0,0" VerticalAlignment="Top" Width="254" Grid.ColumnSpan="3"/>
                </Grid>
            </TabItem>
            <TabItem Header="Mstsc Options">
                <Grid Margin="0,0,100,100   " Grid.Column="1" Height="200"  HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
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
                <Grid x:Name="OtherRDPOptions" Margin="55,0,528,0" Height="309"  HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
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
                <Grid x:Name="VMwareOptions" Margin="55,0,409,0" Height="342"  HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="129*"/>
                        <ColumnDefinition Width="140*"/>
                    </Grid.ColumnDefinitions>
                    <Button x:Name="btnLaunchVMwareOptions" Content="_Launch" HorizontalAlignment="Left" Height="25" VerticalAlignment="Bottom" Width="96" Margin="10,0,0,-24" IsDefault="True"/>
                    <ListView x:Name="listViewVMwareVMs" Grid.ColumnSpan="2" Height="293" Margin="10,39,70,0" VerticalAlignment="Top" SelectionMode="Multiple" >
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                    <Label x:Name="labelVMwareVMs" Content="VMs" HorizontalAlignment="Center" Height="29" Margin="0,5,0,0" VerticalAlignment="Top" Width="122"/>
                    <Label Content="Filter" HorizontalAlignment="Left" Height="29" Margin="94,2,0,0" VerticalAlignment="Top" Width="123" Grid.Column="1"/>
                    <CheckBox x:Name="checkBoxVMwareRegEx" Content="RegEx" Height="29" Width="93" Grid.Column="1" Margin="304,41,-255,272" IsChecked="True"/>
                    <TextBox x:Name="textBoxVMwareFilter" TextWrapping="Wrap" Grid.Column="1" Margin="94,39,-143,275" />
                    <Button x:Name="buttonVMwareApplyFilter" Content="Apply _Filter" Height="31" Width="117" Grid.Column="1" Margin="94,86,-69,225"/>
                    <Label Content="vCenter" HorizontalAlignment="Left" Height="29" VerticalAlignment="Top" Width="124" Grid.Column="1" Margin="100,226,0,0" />
                    <TextBox x:Name="textBoxVMwareRDPPort" TextWrapping="Wrap" Height="28" Width="189" Grid.Column="1" Margin="102,156,-136,158" />
                    <Button x:Name="buttonVMwareConnect" Content="_Connect" Height="31" Width="117" Grid.Column="1" Margin="102,301,-64,10"/>
                    <RadioButton x:Name="radioButtonVMwareConnectByIP" Content="Connect by _IP" Margin="97,202,-127,121" Grid.Column="1" GroupName="GroupBy"/>
                    <RadioButton x:Name="radioButtonVMwareConnectByName"   Content="Connect by _Name" Margin="216,202,-202,121" Grid.Column="1" GroupName="GroupBy" IsChecked="True"/>
                    <Label Content="RDP Port" HorizontalAlignment="Left" Height="29" VerticalAlignment="Top" Width="124" Grid.Column="1" Margin="97,122,0,0" />
                    <TextBox x:Name="textBoxVMwarevCenter" TextWrapping="Wrap" Height="28" Grid.Column="1" Margin="100,255,-202,59" />
                    <Button x:Name="buttonVMwareDisconnect" Content="_Disconnect" Height="31" Width="117" Grid.Column="1" Margin="240,301,-202,10"/>
                </Grid>
            </TabItem>
            <TabItem Header="Hyper-V">
                <Grid x:Name="HyperVOptions" Margin="55,0,409,0" Height="342"  HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="160*"/>
                        <ColumnDefinition Width="160*"/>
                    </Grid.ColumnDefinitions>
                    <Button x:Name="btnLaunchHyperVOptions" Content="_Launch" HorizontalAlignment="Left" Height="25" VerticalAlignment="Bottom" Width="96" Margin="10,0,0,-24" IsDefault="True"/>
                    <Button x:Name="btnLaunchHyperVConsole" Content="C_onsole" HorizontalAlignment="Left" Height="25" VerticalAlignment="Bottom" Width="96" Margin="10,0,0,-24"  Grid.Column="2" IsDefault="False"/>
                    <ListView x:Name="listViewHyperVVMs" Grid.ColumnSpan="2" Height="293" Margin="10,39,70,0" VerticalAlignment="Top" SelectionMode="Multiple" >
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}"/>
                                <GridViewColumn Header="Power State" DisplayMemberBinding="{Binding PowerState}"/>
                            </GridView>
                        </ListView.View>
                        <ListView.ContextMenu>
                            <ContextMenu>
                                <MenuItem Header="Power" x:Name="PowerContextMenu" >
                                    <MenuItem Header="Power On" x:Name="HyperVPowerOnContextMenu" />
                                    <MenuItem Header="Power Off" x:Name="HyperVPowerOffContextMenu" />
                                    <MenuItem Header="Shutdown" x:Name="HyperVShutdownContextMenu" />
                                    <MenuItem Header="Restart" x:Name="HyperVRestartContextMenu" />
                                    <MenuItem Header="Resume" x:Name="HyperVResumeContextMenu" />
                                    <MenuItem Header="Save" x:Name="HyperVSaveContextMenu" />
                                    <MenuItem Header="Suspend" x:Name="HyperVSuspendContextMenu" />
                                </MenuItem>
                                <MenuItem Header="Config" x:Name="ConfigContextMenu" >
                                    <MenuItem Header="Detail" x:Name="HyperVDetailContextMenu" />
                                    <MenuItem Header="Rename" x:Name="HyperVRenameMenu" />
                                    <MenuItem Header="Reconfigure" x:Name="HyperVReconfigureMenu" />
                                    <MenuItem Header="Enable Resource Metering" x:Name="HyperVEnableResourceMeteringContextMenu" />
                                    <MenuItem Header="Performance Data" x:Name="HyperVMeasureContextMenu" />
                                    <MenuItem Header="Disable Resource Metering" x:Name="HyperVDisableResourceMeteringContextMenu" />
                                </MenuItem>
                                <MenuItem Header="Delete" x:Name="DeletionContextMenu" >
                                    <MenuItem Header="Delete VM" x:Name="HyperVDeleteContextMenu" />
                                    <MenuItem Header="Delete VM + Disks" x:Name="HyperVDeleteAllContextMenu" />
                                </MenuItem>
                                <MenuItem Header="CD" x:Name="CDContextMenu" >
                                    <MenuItem Header="Mount" x:Name="HyperVMountCDContextMenu" />
                                    <MenuItem Header="Eject" x:Name="HyperVEjectCDContextMenu" />
                                </MenuItem>
                                <MenuItem Header="Snapshots" x:Name="SnapshotsContextMenu" >
                                    <MenuItem Header="Manage" x:Name="HyperVManageSnapshotContextMenu" />
                                    <MenuItem Header="Take Snapshot" x:Name="HyperVTakeSnapshotContextMenu" />
                                    <MenuItem Header="Revert to Latest Snapshot" x:Name="HyperVRevertLatestSnapshotContextMenu" />
                                    <MenuItem Header="Delete Latest Snapshot" x:Name="HyperVDeleteLatestSnapshotContextMenu" />
                                </MenuItem>
                                <MenuItem Header="New" x:Name="NewContextMenu" >
                                    <MenuItem Header="Brand New" x:Name="HyperVNewVMContextMenu" />
                                    <MenuItem Header="Templated" x:Name="HyperVNewVMFromTemplateContextMenu" />
                                </MenuItem>
                                <MenuItem Header="Name to Clipboard" x:Name="HyperVNameToClipboard" />
                                <MenuItem Header="NICS" x:Name="NICSContextMenu" >
                                    <MenuItem Header="Disconnect NIC" x:Name="HyperVDisconnectNICContextMenu" />
                                    <MenuItem Header="Connect To" x:Name="ConnectNICContextMenu" >
                                        <MenuItem Header="Internal" x:Name="HyperVConnectNICInternalContextMenu" />
                                        <MenuItem Header="External" x:Name="HyperVConnectNICExternalContextMenu" />
                                        <MenuItem Header="Private"  x:Name="HyperVConnectNICPrivateContextMenu" />
                                    </MenuItem>
                                </MenuItem>
                            </ContextMenu>
                        </ListView.ContextMenu>
                    </ListView>
                    <Label x:Name="labelHyperVVMs" Content="VMs" HorizontalAlignment="Center" Height="29" Margin="0,5,0,0" VerticalAlignment="Top" Width="122"/>
                    <Label Content="Filter" HorizontalAlignment="Left" Height="29" Margin="94,2,0,0" VerticalAlignment="Top" Width="40" Grid.Column="1"/>
                    <CheckBox x:Name="checkBoxHyperVRegEx" Content="RegE_x" Height="29" Width="93" Grid.Column="1" Margin="304,35,-255,272" IsChecked="True"/>
                    <CheckBox x:Name="checkBoxHyperVAllVMs" Content="_All VMs" Height="29" Width="93" Grid.Column="2" Margin="304,55,-255,272" IsChecked="False"/>
                    <TextBox x:Name="textBoxHyperVFilter" TextWrapping="Wrap" Grid.Column="1" Margin="94,39,-143,275" />
                    <Button x:Name="buttonHyperVApplyFilter" Content="Apply _Filter" Height="31" Width="117" Grid.Column="1" Margin="94,86,-69,225"/>
                    <Button x:Name="buttonHyperVClearFilter" Content="Clea_r Filter" Height="31" Width="117" Grid.Column="1" Margin="240,86,-202,225"/>
                    <Label Content="Host" HorizontalAlignment="Left" Height="29" VerticalAlignment="Top" Width="124" Grid.Column="1" Margin="100,226,0,0" />
                    <TextBox x:Name="textBoxHyperVRDPPort" TextWrapping="Wrap" Height="28" Width="189" Grid.Column="1" Margin="102,156,-136,158" />
                    <Button x:Name="buttonHyperVConnect" Content="_Connect" Height="31" Width="117" Grid.Column="1" Margin="102,301,-64,10"/>
                    <RadioButton x:Name="radioButtonHyperVConnectByIP" Content="Connect by _IP" Margin="97,202,-127,121" Grid.Column="1" GroupName="GroupBy"/>
                    <RadioButton x:Name="radioButtonHyperVConnectByName"   Content="Connect by _Name" Margin="216,202,-202,121" Grid.Column="1" GroupName="GroupBy" IsChecked="True"/>
                    <Label Content="RDP Port" HorizontalAlignment="Left" Height="29" VerticalAlignment="Top" Width="154" Grid.Column="1" Margin="97,122,0,0" />
                    <TextBox x:Name="textBoxHyperVHost" TextWrapping="Wrap" Height="28" Grid.Column="1" Margin="100,255,-202,59" />
                </Grid>
            </TabItem>
            <TabItem Header="Active Directory" IsEnabled="false">
                <Grid x:Name="ActiveDirectory" Margin="55,0,409,0" Height="342"  HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="129*"/>
                        <ColumnDefinition Width="140*"/>
                    </Grid.ColumnDefinitions>
                    <Button x:Name="btnLaunchADOptions" Content="_Launch" HorizontalAlignment="Left" Height="25" VerticalAlignment="Bottom" Width="96" Margin="10,0,0,-24" IsDefault="True"/>
                    <ListView x:Name="listViewAD" Grid.ColumnSpan="2" Height="293" Margin="10,39,70,0" VerticalAlignment="Top" SelectionMode="Multiple" >
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                    <Label x:Name="labelADVMs" Content="VMs" HorizontalAlignment="Center" Height="29" Margin="0,5,0,0" VerticalAlignment="Top" Width="122"/>
                    <Label Content="Filter" HorizontalAlignment="Left" Height="29" Margin="94,2,0,0" VerticalAlignment="Top" Width="123" Grid.Column="1"/>
                    <CheckBox x:Name="checkBoxADRegEx" Content="RegEx" Height="29" Width="93" Grid.Column="1" Margin="304,41,-255,272" IsChecked="True"/>
                    <TextBox x:Name="textBoxADFilter" TextWrapping="Wrap" Grid.Column="1" Margin="94,39,-143,275" />
                    <Button x:Name="buttonADApplyFilter" Content="Apply _Filter" Height="31" Width="117" Grid.Column="1" Margin="94,86,-69,225"/>
                    <Label Content="vCenter" HorizontalAlignment="Left" Height="29" VerticalAlignment="Top" Width="124" Grid.Column="1" Margin="100,226,0,0" />
                    <TextBox x:Name="textBoxADRDPPort" TextWrapping="Wrap" Height="28" Width="189" Grid.Column="1" Margin="102,156,-136,158" />
                    <Button x:Name="buttonADConnect" Content="_Connect" Height="31" Width="117" Grid.Column="1" Margin="102,301,-64,10"/>
                    <RadioButton x:Name="radioButtonADConnectByIP" Content="Connect by _IP" Margin="97,202,-127,121" Grid.Column="1" GroupName="GroupBy"/>
                    <RadioButton x:Name="radioButtonADConnectByName"   Content="Connect by _Name" Margin="216,202,-202,121" Grid.Column="1" GroupName="GroupBy" IsChecked="True"/>
                    <Label Content="RDP Port" HorizontalAlignment="Left" Height="29" VerticalAlignment="Top" Width="124" Grid.Column="1" Margin="97,122,0,0" />
                    <TextBox x:Name="textBoxADvCenter" TextWrapping="Wrap" Height="28" Grid.Column="1" Margin="100,255,-202,59" />
                    <Button x:Name="buttonADDisconnect" Content="_Disconnect" Height="31" Width="117" Grid.Column="1" Margin="240,301,-202,10"/>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
'@
#>

[string]$textInputXAML = @'
<Window x:Class="WPF_Scratchpad.Window1"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WPF_Scratchpad"
        mc:Ignorable="d"
        Title="Window1" Height="450" Width="800">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="4*"/>
            <ColumnDefinition Width="30*"/>
            <ColumnDefinition/>
            <ColumnDefinition Width="365*"/>
        </Grid.ColumnDefinitions>
        <TextBox x:Name="textboxInputText" HorizontalAlignment="Left" Height="85" Margin="45,52,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="347" Grid.ColumnSpan="3" Grid.Column="1"/>
        <Label x:Name="lblInputTextLabel" Content="Label" HorizontalAlignment="Left" Height="27" Margin="45,10,0,0" VerticalAlignment="Top" Width="253" Grid.ColumnSpan="3" Grid.Column="1"/>
        <Button x:Name="btnInputTextOK" Content="OK" HorizontalAlignment="Left" Height="39" Margin="45,157,0,0" VerticalAlignment="Top" Width="89" Grid.ColumnSpan="3" IsDefault="True" Grid.Column="1"/>
        <Button x:Name="btnInputTextCancel" Content="Cancel" HorizontalAlignment="Left" Height="39" Margin="110,157,0,0" VerticalAlignment="Top" Width="88" Grid.Column="3" IsCancel="True"/>

    </Grid>
</Window>
'@

[string]$snapshotsXAML = @'
<Window x:Class="mstsc_GUI.Snapshots"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:VMWare_GUI"
        mc:Ignorable="d"
        Title="Snapshots" Height="450" Width="800">
    <Grid>
        <TreeView x:Name="treeSnapshots" HorizontalAlignment="Left" Height="246" Margin="85,74,0,0" VerticalAlignment="Top" Width="471"/>
        <Grid Margin="593,88,85,128" >
            <Button x:Name="btnTakeSnapshot" Content="_Take Snapshot" HorizontalAlignment="Left" VerticalAlignment="Top" Width="92"/>
            <Button x:Name="btnDeleteSnapshot" Content="De_lete" HorizontalAlignment="Left" Margin="0,172,0,0" VerticalAlignment="Top" Width="92"/>
            <Button x:Name="btnRevertSnapshot" Content="_Revert" HorizontalAlignment="Left" Margin="0,83,0,0" VerticalAlignment="Top" Width="92"/>
            <Button x:Name="btnConsolidateSnapshot" Content="_Consolidate" HorizontalAlignment="Left" Margin="0,129,0,0" VerticalAlignment="Top" Width="92"/>
            <Button x:Name="btnDetailsSnapshot" Content="_Details" HorizontalAlignment="Left" Margin="0,42,0,0" VerticalAlignment="Top" Width="92"/>
        </Grid>
        <Button x:Name="btnSnapshotsOk" Content="OK" HorizontalAlignment="Left" Margin="95,365,0,0" VerticalAlignment="Top" Width="75" IsDefault="True"/>
        <Button x:Name="btnSnapshotsCancel" Content="Cancel" HorizontalAlignment="Left" Margin="198,365,0,0" VerticalAlignment="Top" Width="75" IsCancel="True"/>
        <Label x:Name="lblLastRevert" Content="Last Revert" HorizontalAlignment="Left" Height="29" Margin="85,23,0,0" VerticalAlignment="Top" Width="622"/>
        <Button x:Name="btnDeleteSnapShotTree" Content="Delete _Tree" HorizontalAlignment="Left" Margin="593,300,0,0" VerticalAlignment="Top" Width="92"/>
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
 
 Function Get-Msrdc
 {
    [string]$msrdc = $null

    if( -Not [string]::IsNullOrEmpty( $msrdcCopyPath ) )
    {
        $exe = $msrdcCopyPath
    }
    elseif( -Not ( Get-Command -Name ($exe = 'msrdc.exe') -CommandType Application -ErrorAction SilentlyContinue ) )
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
        elseif( $appx = Get-AppxPackage -Name '*.Windows365' | Sort-Object -Property Version -Descending | Select-Object -First 1 )
        {
            ## cannot execute it directly from here so we take a copy (hopefully Ivanti Application Control's Trusted Ownership won't bite us :-) )
            [string]$copyToFolder = $msrdcCopyFolder
            if( [string]::IsNullOrEmpty( $copyToFolder ) )
            {
                $copyToFolder = [System.IO.Path]::Combine( $env:LOCALAPPDATA , 'Programs' , $msrdcCopyName )
            }
            if( -Not (Test-Path -Path $copyToFolder) )
            {
                New-Item -Path $copyToFolder -ItemType Directory -Force ## if it errors, so be it
            }
            $appxMsrdcVersion = Get-ItemProperty -Path ( Join-Path $appx.InstallLocation -ChildPath 'msrdc\msrdc.exe' ) | Select-Object -ExpandProperty VersionInfo | Select-Object -ExpandProperty FileVersionRaw
            if( -Not ( $copyProperties = Get-ItemProperty -Path (Join-Path -Path $copyToFolder -ChildPath 'msrdc.exe') -ErrorAction SilentlyContinue) -or $appxMsrdcVersion -gt $copyProperties.VersionInfo.FileVersionRaw ) 
            {
                Write-Verbose "Copying from $($appx.InstallLocation ) to $copyToFolder"
                Copy-Item -Path (Join-Path $appx.InstallLocation -ChildPath 'msrdc\*') -Destination $copyToFolder -Recurse -Force
            }
            else
            {
                Write-Verbose -Message "msrdc copy folder `"$copyToFolder`" already exists"
            }
            $exe = Join-Path -Path $copyToFolder -ChildPath 'msrdc.exe' 
        }
        else
        {
            $exe = [System.IO.Path]::Combine( ([Environment]::GetFolderPath( [Environment+SpecialFolder]::ProgramFiles )) , 'Remote Desktop' , 'msrdc.exe' )
            if( -Not ( Test-Path -Path $exe -PathType Leaf ) )
            {
                $exe = [System.IO.Path]::Combine( ([Environment]::GetFolderPath( [Environment+SpecialFolder]::ProgramFilesX86 )) , 'Remote Desktop' , 'msrdc.exe' )
                ## TODO what if per user install?
            }
        }
    }

    if( -Not [string]::IsNullOrEmpty( $exe ) -and ( Test-Path -Path $exe -PathType Leaf -ErrorAction SilentlyContinue ) )
    {
        $exe ## return
    }
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

Function Process-Snapshot
{
    Param
    (
        $GUIobject ,
        $Operation ,
        $VM
    )
    
    $_.Handled = $true
    
    [bool]$closeDialogue = $true

    if( $null -eq $VM )
    {
        [void][Windows.MessageBox]::Show( 'No VM passed to Process-Snapshot' , $Operation , 'Ok' ,'Error' )
        return
    }

    if( $operation -eq 'ConsolidateSnapshot' )
    {
        ## $VM = Get-VM -Id $VMId
        if( ! ( $task = $VM.ExtensionData.ConsolidateVMDisks_Task() ) `
            -or ! ($taskStatus = Get-Task -Id $task) `
                -or $taskStatus.State -eq 'Error' )
        {
            [void][Windows.MessageBox]::Show( 'Task Failed' , $Operation , 'Ok' ,'Error' )
        }
        else
        {
            [void][Windows.MessageBox]::Show( 'Task Started' , $Operation , 'Ok' ,'Information' )
        }
        $closeDialogue = $false
    }
    elseif( $Operation -eq 'TakeSnapShot' )
    {
        if( $takeSnapshotForm = New-Form -inputXaml $takeSnapshotXAML )
        {
            $takeSnapshotForm.Title += " of $($vm.name)"

            $WPFbtnTakeSnapshotOk.Add_Click({ 
                $_.Handled = $true
                $takeSnapshotForm.DialogResult = $true 
                $takeSnapshotForm.Close()  })

            if( $VM = Get-VM -Id $VMId )
            {
                if( $VM.PowerState -eq 'PoweredOff' ) 
                {
                    $wpfchkSnapshotQuiesce.IsEnabled = $false
                    $WPFchckSnapshotMemory.IsEnabled = $false
                }
                elseif( $VM.Guest.State -eq 'NotRunning' )
                {
                    $wpfchckSnapshotShutdownStart.IsEnabled = $false ## won't be able to shut it down cleanly so don't offer it
                }

                if( $takeSnapshotForm.ShowDialog() )
                {
                    if( $VM = Get-VM -Id $VMId )
                    {
                        ## Get Data from form and take snapshot
                        [hashtable]$parameters = @{ 'VM' = $vm ; 'RunAsync' = $true }
                        if( ! [string]::IsNullOrEmpty( $WPFtxtSnapshotName.Text ) )
                        {
                            $parameters.Add( 'Name' , $WPFtxtSnapshotName.Text )
                        }
                        else
                        {
                            return
                        }
                        if( ! [string]::IsNullOrEmpty( $WPFtxtSnapshotDescription.Text ) )
                        {
                            $parameters.Add( 'Description' , $WPFtxtSnapshotDescription.Text )
                        }
                        $parameters.Add( 'Quiesce' , $wpfchkSnapshotQuiesce.IsChecked )
                        $parameters.Add( 'Memory' , $WPFchckSnapshotMemory.IsChecked )
         
                        [string]$answer = 'yes'

                        if( $wpfchckSnapshotShutdownStart.IsChecked )
                        {
                            if( $VM.PowerState -ne 'PoweredOff' -and ( $answer = [Windows.MessageBox]::Show( "VM $($VM.Name)" , "Confirm Shutdown & Startup" , 'YesNo' ,'Question' ) ) -and $answer -ieq 'yes' )
                            {
                                $shutdownError = $null
                                if( -Not ( $guest = Shutdown-VMGuest -VM $VM -ErrorVariable shutdownError -Confirm:$false ) )
                                {
                                    $answer = 'abort'
                                    [void][Windows.MessageBox]::Show( $shutdownError , "Shutdown Error for $($VM.Name)" , 'Ok' ,'Error' )
                                }
                                else
                                {
                                    [datetime]$endWaitTime = [datetime]::Now.AddSeconds( $powerActionWaitTimeSeconds )
                                    Write-Verbose -Message "$(Get-Date -Format G): waiting for VM to shutdown until $(Get-Date -Date $endWaitTime -Format G)"
                                    do
                                    {
                                        Start-Sleep -Seconds 5
                                        $VM = Get-VM -Id $VMId
                                    }
                                    while( $VM -and $VM.PowerState -ne 'PoweredOff' -and [datetime]::Now -le $endWaitTime )
                                    
                                    Write-Verbose -Message "$(Get-Date -Format G): finished waiting for VM $($VM.Name) to shutdown - power state is $($VM|Select-Object -ExpandProperty PowerState)"

                                    if( -Not $VM -or $VM.PowerState -ne 'PoweredOff' )
                                    {
                                        [void][Windows.MessageBox]::Show( "Timed out waiting $powerActionWaitTimeSeconds seconds for shutdown to complete" , "Shutdown Error for $($VM.Name)" , 'Ok' ,'Error' )
                                        $answer = 'abort'
                                    }
                                    $parameters[ 'RunAsync' ] = $false
                                }
                            }
                        }

                        if( $answer -ieq 'yes' )
                        {
                            New-Snapshot @parameters
                            if( $? -and $wpfchckSnapshotShutdownStart.IsChecked )
                            {
                                Start-VM -VM $VM -RunAsync
                            }
                        }
                    }
                    else
                    {
                        [void][Windows.MessageBox]::Show( "Failed to get VM" , "Snapshot Error for $($VM.Name)" , 'Ok' ,'Error' )
                    }
                }
            }
            else
            {
                [void][Windows.MessageBox]::Show( "Failed to get VM" , "Snapshot Error for $($VM.Name)" , 'Ok' ,'Error' )
            }
        }
    }
    elseif( $Operation -eq 'DetailsSnapshot' )
    {
        $closeDialogue = $false
        if( $VM ) ## =  Get-VM -Id $VMId )
        {
            [string]$tag = $null
            if( $GUIobject.SelectedItem -and $GUIobject.PSObject.Properties[ 'SelectedItem' ] )
            {
                $tag = $GUIobject.SelectedItem.Tag
            }
            elseif( $GUIobject.Items.Count -eq 1 )
            {
                ## if only one snapshot then report on that one
                $tag = $GUIobject.Items[0].Tag
            }
            else
            {
                [void][Windows.MessageBox]::Show( 'No snapshot selected' , $Operation , 'Ok' ,'Error' )
                return
            }
            if( $tag -eq $youAreHereSnapshot )
            {
                return
            }
            if( $snapshot = $script:theseSnapshots | Where-Object Id -eq $tag)
            {
                [uint64]$size = 0
                ForEach( $disk in $snapshot.HardDrives  )
                {
                    $file = $null
                    ## needs backslashes escaping - file could be remote
                    $file = Get-CimInstance -ClassName cim_datafile -Filter "Name = '$($disk.Path -replace '\\' , '\\')'" -CimSession $snapshot.CimSession
                    if( $null -ne $file )
                    {
                        $size += $file.FileSize
                    }
                }
                [string]$details = "Name = $($snapshot.Name)`n`rNotes = $($snapshot.Notes)`n`rCreated = $($snapshot.CreationTime.ToString('G'))`n`rSize = $([math]::Round( $size / 1GB , 2))GB`n`rType = $($snapshot.SnapshotType)`n`rPower State = $($snapshot.State)`n`rAutomatic = $(if( $snapshot.IsAutomaticCheckpoint ) { 'Yes' } else {'No' })"
                [void][Windows.MessageBox]::Show( $details , 'Snapshot Details' , 'Ok' ,'Information' )
            }
            else
            {
                Write-Warning "Unable to get snapshot $($GUIobject.SelectedItem.Tag) for `"$($VM.Name)`""
            }
        }
        else
        {
            Write-Warning "Unable to get vm for vm id $vmid"
        }
    }
    elseif( ! $GUIobject -or ( $GUIobject.SelectedItem -and $GUIobject.SelectedItem.Tag -and $GUIobject.SelectedItem.Tag -ne $youAreHereSnapshot ) )
    {
        ##$VM = Get-VM -Id $VMId
        if( $VM )
        {
            if( $Operation -eq 'LatestSnapshotRevert' )
            {
                if( ! ( $snapshot = Get-VMSnapshot -VM $vm -verbose:$false | Sort-Object -Property CreationTime -Descending|Select-Object -First 1 ))
                {
                    [Windows.MessageBox]::Show( "No snapshots found for $($vm.Name)" , 'Snapshot Revert Error' , 'OK' ,'Error' )
                    return
                }
            }
            else
            {
                ##$snapshot = Get-VMSnapshot -Name $GUIobject.SelectedItem.Tag -VM $vm -verbose:$false
                $snapshot = $script:theseSnapshots | Where-Object Id -eq $GUIobject.SelectedItem.Tag
            }
            if( $snapshot )
            {
                [string]$answer = 'no'
                [string]$questionText = $null

                if( $Operation -eq 'DeleteSnapShotTree' )
                {
                    $questionText = 'From snapshot'
                }
                else
                {
                    $questionText = 'Snapshot'
                }

                $questionText += " `"$($snapshot.Name)`" on $($vm.Name), taken $($snapshot.CreationTime.ToString('G')) ?"
                $answer = [Windows.MessageBox]::Show( $questionText , "Confirm $($operation -creplace '([a-zA-Z])([A-Z])' , '$1 $2')" , 'YesNo' ,'Question' )
        
                if( $answer -eq 'yes' )
                {
                    if( $Operation -eq 'DeleteSnapShot' )
                    {
                        Remove-VMSnapshot -VMSnapshot $snapshot -AsJob -Confirm:$false
                    }
                    elseif( $Operation -eq 'DeleteSnapShotTree' )
                    {
                        Remove-VMSnapshot -VMSnapshot $snapshot -Confirm:$false -IncludeAllChildSnapshots -AsJob
                    }
                    elseif( $Operation -eq 'RevertSnapShot' -or $Operation -eq 'LatestSnapshotRevert' )
                    {
                        $answer = $null
                        if( $snapshot.State -ieq 'Off' )
                        {
                            $answer = [Windows.MessageBox]::Show( "Power on after snapshot restored on $($vm.Name)?" , 'Confirm Power Operation' , 'YesNo' ,'Question' )
                        }
                        [hashtable]$revertParameters = @{ 'VMSnapshot' = $snapshot ;'Confirm' = $false ; AsJob = ($answer -ne 'Yes') }
                        Restore-VMSnapshot @revertParameters
                        if( $answer -eq 'Yes' )
                        {
                            Start-VM -VM $vm -Confirm:$false
                        }
                    }
                    else
                    {
                        Write-Warning "Unexpected snapshot operation $operation"
                    }
                }
                else
                {
                    $closeDialogue = $false
                }
            }
            else
            {
                Write-Warning "Unable to get snapshot $($GUIobject.SelectedItem.Tag) for `"$($VM.Name)`""
            }
        }
        else
        {
            Write-Warning "Unable to get vm for vm id $vmid"
        }
    }
    else
    {
        $closeDialogue = $false
    }

    if( $closeDialogue -and (Get-Variable -Name snapshotsForm -ErrorAction SilentlyContinue) -and $snapshotsForm )
    {
        ## Close dialog since needs refreshing
        $snapshotsForm.DialogResult = $true 
        $snapshotsForm.Close()
        $snapshotsForm = $null
    }
}

Function Find-TreeItem
{
    Param
    (
        [array]$controls ,
        [string]$tag
    )

    $result = $null

    ForEach( $control in $controls )
    {
        if( $null -eq $result )
        {
            if( $control.Tag -eq $tag )
            {
                $result = $control
            }
            elseif( $control.PSobject.Properties[ 'Items' ] )
            {
                ForEach( $item in $control.Items )
                {
                    if( $item.Tag -eq $tag )
                    {
                        $result = $item
                    }
                    elseif( $item.PSobject.Properties[ 'Items' ] -and $item.Items.Count )
                    {
                        $result = Find-TreeItem -control $item.Items -tag $tag
                    }
                }
            }
        }
    }

    $result
}

## https://blog.ctglobalservices.com/powershell/kaj/powershell-wpf-treeview-example/

Function Add-TreeItem
{
    Param
    (
          $Name,
          $Parent,
          $Tag 
    )

    $ChildItem = New-Object System.Windows.Controls.TreeViewItem
    $ChildItem.Header = $Name
    $ChildItem.Name = $Name -replace '[/\s,;:\.\-\+)(]' , '_' -replace '%252f' , '_' -replace '&' , '_' # default snapshot names have / for date which are escaped
    $ChildItem.Tag = $Tag
    $ChildItem.IsExpanded = $true
    ##[Void]$ChildItem.Items.Add("*")
    [Void]$Parent.Items.Add($ChildItem)
}

Function Show-SnapShotWindow
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        $vmName
    )

    $thisVM = Hyper-V\Get-VM -Name $vmName @hypervParameters
    $script:theseSnapshots = @( Get-VMCheckpoint -VMName $vmName @hypervParameters )
    if( $null -eq $script:theseSnapshots -or $script:theseSnapshots.Count -eq 0)
    {
        [void][Windows.MessageBox]::Show( "No snapshots found for $vmName" , 'Snapshot Management' , 'Ok' ,'Warning' )
        return
    }

    $snapshotsForm = New-WPFWindow -inputXAML $snapshotsXAML

    [bool]$result = $false
    if( $snapshotsForm )
    {
        $snapshotsForm.Title += " for $vmName"

        ForEach( $snapshot in $script:theseSnapshots )
        {
            ## if has a parent we need to find that and add to it
            if( $snapshot.ParentSnapshotId )
            {
                ## find where to add our node
                $parent = $script:theseSnapshots | Where-Object Id -eq $snapshot.ParentSnapshotId
                if( $parent )
                {
                    $parentNode = Find-TreeItem -control $WPFtreeSnapshots -tag $snapshot.ParentSnapshotId
                    if( $parentNode )
                    {
                        Add-TreeItem -Name $snapshot.Name -Parent $parentNode -Tag $snapshot.Id 
                    }
                    else
                    {
                        Write-Warning "Unable to locate tree view item for parent snapshot `"$($snapshot.parent)`""
                    }
                }
                else
                {
                    Write-Warning "Unable to locate parent snapshot `"$($snapshot.Parent)`" for snapshot `"$($snapshot.Name)`" for $vmName"
                }
            }
            else ## no parent then needs to be top level node but check not already created because we enountered a child previously
            {
                Add-TreeItem -Name $snapshot.Name -Parent $WPFtreeSnapshots -Tag $snapshot.Id
            }
        }
        
        [string]$currentSnapShotId = $null
        if( $currentSnapShot = $script:theseSnapshots | Where-Object Id -eq $thisVM.ParentCheckpointId)
        {
            $currentSnapShotId = $currentSnapShot.Id
        }
        if( $currentSnapShotId )
        {
            if( ($currentSnapShotItem = Find-TreeItem -control $WPFtreeSnapshots -tag $currentSnapShotId ))
            {
                Add-TreeItem -Name '__You are here__' -Parent $currentSnapShotItem -Tag $youAreHereSnapshot
            }
            else
            {
                Write-Warning "Unable to locate tree view item for current snapshot"
            }
        }
        else
        {
            Write-Warning "No current snapshot set for $vmName"
        }

        $WPFbtnSnapshotsOk.Add_Click({
            $_.Handled = $true 
            $snapshotsForm.DialogResult = $true 
            $snapshotsForm.Close()  })

        ## get last revert operation
## TODO event logs?
<#
        $lastRevert = Get-VIEvent -Entity $vm.Name -ErrorAction SilentlyContinue | Where-Object { $_.PSObject.Properties[ 'EventTypeId' ] -and $_.EventTypeId -eq 'com.vmware.vc.vm.VmStateRevertedToSnapshot' -and $_.FullFormattedMessage -match 'has been reverted to the state of snapshot (.*), with ID \d' } | Select-Object -First 1
        [string]$text = $null
        if( $lastRevert )
        {
            $text = "Last revert was to snapshot `"$($Matches[1])`" on $(Get-Date -Date $lastRevert.CreatedTime -Format G)"
        }
        else
        {
            $text = "No snapshot revert event found"
        }
        $wpflblLastRevert.Content = $text
        ## see if consolidation is required so that we enable/disable the consolidation button
        $wpfbtnConsolidateSnapshot.IsEnabled = $vm.Extensiondata.Runtime.ConsolidationNeeded
                     
#>   
        $WPFbtnTakeSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'TakeSnapShot' -VM $thisVM } )
        $WPFbtnDeleteSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'DeleteSnapShot' -VM $thisVM} )
        $WPFbtnDeleteSnapShotTree.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'DeleteSnapShotTree' -VM $thisVM} )
        $WPFbtnRevertSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'RevertSnapShot' -VM $thisVM} )
        $WPFbtnDetailsSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'DetailsSnapShot' -VM $thisVM} )
        $WPFbtnConsolidateSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'ConsolidateSnapShot' -VMId $thisVM } )

        $result = $snapshotsForm.ShowDialog()
     }
}

Function Sort-Columns( $control )
{
    if( $view =  [Windows.Data.CollectionViewSource]::GetDefaultView($control.ItemsSource) )
    {
        [string]$direction = 'Ascending'
        if(  $view.PSObject.Properties[ 'SortDescriptions' ] -and $view.SortDescriptions -and $view.SortDescriptions.Count -gt 0 )
        {
	        $sort = $view.SortDescriptions[0].Direction
	        $direction = if( $sort -and 'Descending' -eq $sort){ 'Ascending' } else { 'Descending' }
	        $view.SortDescriptions.Clear()
        }

        Try
        {
            [string]$column = $_.OriginalSource.Column.DisplayMemberBinding.Path.Path ## has to be name of the binding, not the header unless no binding
            if( [string]::IsNullOrEmpty( $column ) )
            {
                $column = $_.OriginalSource.Column.Header
            }
	        $view.SortDescriptions.Add( ( New-Object ComponentModel.SortDescription( $column , $direction ) ) )
        }
        Catch
        {
        }
    }
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
        
        ## recreate in case temp been tidied since script started
        if( -Not ( Test-Path -Path $tempFolder -PathType Container ) -and -Not ( New-Item -Path $tempFolder -ItemType Directory -Force ) )
        {
            Throw "Failed to create temp folder $tempFolder"
        }

        ## mstsc title bar uses filename before first dot and address so if we use address in filename it duplicates it
        ## as differentiator in host names is at the right hand side then reverse the name used in the rdp file so the last characters are first
        [string]$filename = $address.ToUpper()
        if( $reverse )
        {
            [array]$array = $address.ToCharArray()
            [array]::Reverse( $array )
            $filename = ($array -join '').ToUpper()
        }
        elseif( -Not [string]::IsNullOrEmpty( $remove ) )
        {
            $filename = $address -replace $remove ## designed to remove the leading characters of the name which will all be the same/similar eg GLHVS22 and makes icons difficult to distinguish
        }
        [string]$tempRdpFile = Join-Path -Path $tempFolder -ChildPath "$filename.rdp"
        [string]$windowTitle = "^$filename - $address - Remote Desktop Connection$"

        ## using our own folder so drop tmp from name since makes right click on mstsc taskbar icon look crap
        <#
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
        #>

        ## mstsc file will have things in it doesn't understand which it silently ignores
        
        [int]$screenmode = 1
        if( $fullScreen )
        {
            $screenmode = 2
        }
    
        [string]$rdpFileContents = $ExecutionContext.InvokeCommand.ExpandString( $rdpTemplate )

        if( $usemsrdc )
        {
           $exe = Get-Msrdc
            
            if( [string]::IsNullOrEmpty( $remoteDesktopName ))
            {
                ## if there is no remote desktop name specified then the temp rdp file name is included in the Window title which is fugly
                $remoteDesktopName = "$address - Remote Desktop"
                $rdpFileContents += "`nremotedesktopname:s:$remoteDesktopName`n"
                $windowTitle = $remoteDesktopName
            }

            $commandLine = "`"$tempRdpFile`""
            if( -Not [string]::IsNullOrEmpty( $username ) )
            {
                $commandLine = "$commandLine /u:$username"
            }
            if( -Not [string]::IsNullOrEmpty( $extraMsrdcParameters ) )
            {
                $commandLine = "$commandLine $extraMsrdcParameters"
            }
            if( -Not $nofriendlyName )
            {
                $commandLine = "$commandLine /friendlyname:`"$filename`""
                $windowTitle = $filename
            }
        }
        else ## mstsc
        {
        <#
            ## window title comes from the base name of the .rdp file so if we don't rename the temp file, that will be the name in the title bar which is fugly
            [string]$tempRdpFileWithName = Join-Path -Path (Split-Path -Path $tempRdpFile -Parent) -ChildPath "$($address -replace ':' , '.').$(Split-Path -Path $tempRdpFile -Leaf)"
            if( Move-Item -Path $tempRdpFile -Destination $tempRdpFileWithName -PassThru )
            {
                $tempRdpFile = $tempRdpFileWithName
            }
        #>
            $commandline = "`"$tempRdpFile`"" ## $commandLine" ## everything is in the rdp file
        }
        
        ## see if we already have a window with this title so we can offer to switch to that or create a new one
        $existingWindows = $null
        $existingWindows = [Api.Apidef]::GetWindows( -1 ) | Where-Object WinTitle -imatch $windowTitle
        $otherProcess = $null

        if( $null -ne $existingWindows )
        {
            Write-Verbose -Message "Already have window `"$windowTitle`" in process $($existingWindows.PID)"
            
            $otherprocess = Get-Process -Id $existingWindows.PID
            $answer = [Windows.MessageBox]::Show(  "Activate Existing Window ?`nLaunched $($otherprocess.StartTime.ToString('G'))" , "Already Connected to $address" , 'YesNoCancel' ,'Question' )
            if( $answer -ieq 'yes' )
            {
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
            else
            {
                $otherProcess = $null
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
        
        $process = $null

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
        $existingProcesses = @()
        if( $usemsrdc )
        {
            $existingProcesses = @( Get-Process -Name msrdc -ErrorAction SilentlyContinue )
        }
        ## can take a little time for the window to appear and get the title so we poll :-(
## TODO change search string when connecting to IP - window title is "192 - 192.168.1.32 - Remote Desktop Connection"
        do
        {
            $allWindowsNow = @( [Api.Apidef]::GetWindows( -1 ) | Where-Object WinTitle -match $windowTitle )
            ForEach( $window in $allWindowsNow )
            {
                ## Need to find our new window, not any existing one
                if( ( $existingProcess = Get-Process -Id $window.PID -ErrorAction SilentlyContinue ) -and $existingProcess.Name -eq $baseExe  )
                {
                    ## if msrdc then it may have used an existing process :(
                    if( $existingProcess.StartTime -ge $process.StartTime -or $existingProcesses.Count -gt 0 )
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
                    if( $null -eq $existingProcesses -or $existingProcesses.Count -eq 0 )
                    {
                        Write-Warning -Message "Process $($process.Id) has exited and no previous instances"
                        break
                    }

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
            Write-Warning "No main window handle for process $($process.id)"
            return
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
            Remove-Item -Path $tempRdpFile -Force -ErrorAction SilentlyContinue
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
    ## get primary monitor
    <#
    elseif( ( $WPFdatagridDisplays.SelectedItems.Count -ne 1 -and $activeDisplaysWithMonitors.Count -gt 1 ) -and -not $WPFchkboxPrimary.IsChecked )
    {
        [void][Windows.MessageBox]::Show( 'No Monitor Selected' , 'Select a Monitor' , 'Ok' ,'Error' )
    }
    #>
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
            $chosen = $WPFdatagridDisplays.items | Where-Object -Property ScreenPrimary -eq 'true' -ErrorAction SilentlyContinue
            ##$chosen = $WPFdatagridDisplays.Items[0] ## no monitor selected but only one monitor
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
            [string]$comment = $itemsCopy | Where-Object { $_.PSobject -and $_.PSObject.Properties -and $_.PSObject.Properties[ 'ScreenPrimary' ] -and $_.ScreenPrimary -eq $row.ScreenPrimary -and $_.Width -eq $row.Width -and $_.Height -eq $row.Height `
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
    Write-Verbose -Message "Got $($vms.Count) powered on VMware VMs"
    $WPFlistViewVMwareVMs.Items.Clear()

    ForEach( $vm in $vms )
    {
        $WPFlistViewVMwareVMs.Items.Add( [pscustomobject]@{ Name = $vm.VmName  } )
        Write-Verbose -Message "Added $($vm.VmName)"
    }
    $WPFlabelVMwareVMs.Content = "$($WPFlistViewVMwareVMs.Items.Count) VMs"
}

Function Add-HyperVVMsToListView
{
    Param
    (
        [string]$hyperVhost ,
        [string]$filter ,
        [bool]$regex ,
        [bool]$allVMs
    )
    $hyperVError = $null
    [string]$powerState = '.'
    if( -Not $allVMs )
    {
        $powerState = 'Running'
    }
    ## module qualifying in case clash with VMware PowerCLI. Deal with multiple hosts
    $script:vms = @( Hyper-V\Get-VM -ErrorVariable hyperVError -ComputerName ($hyperVhost -split ',') | Where-Object { $_.State -match $powerState -and (( $regex -and $_.Name -match $filter ) -or ( -Not $regex -and $_.Name -like $filter )) } | Sort-Object -Property Name )
    if( $hyperVError )
    {
        [void][Windows.MessageBox]::Show( $hyperVError , "Hyper-V Error from $hyperVhost" , 'Ok' ,'Error' )
    }
    Write-Verbose -Message "Got $($vms.Count) powered on Hyper-V VMs"
    $WPFlistViewHyperVVMs.Items.Clear()
    ForEach( $vm in $vms )
    {
        $WPFlistViewHyperVVMs.Items.Add( [pscustomobject]@{ Name = $vm.Name ; PowerState = $vm.State } ) ## value comes from what is in Binding property for the grid view column
    }
    $WPFlabelHyperVVMs.Content = "$($WPFlistViewHyperVVMs.Items.Count) VMs"
}

## TODO make this generic for detecting which hypervisor is connected to and use that 

Function Start-RemoteSessionFromHypervisor
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [ValidateSet('VMware','Hyper-V')]
        [string]$hypervisorType ,
        [switch]$console
    )
    [int]$loopIterations = 0
    [string]$answer = $null
    $listView = $WPFlistViewVMwareVMs
    $radioButton = $WPFradioButtonVMwareConnectByIP
    $rdpportTextbox = $WPFtextBoxVMwareRDPPort
    if( $hypervisorType -ieq 'Hyper-V' )
    {
        $listView = $WPFlistViewHyperVVMs
        $radioButton = $WPFradioButtonHyperVConnectByIP
        $rdpportTextbox = $WPFtextBoxHyperVRDPPort
    }
    if( $listview.SelectedIndex -ge 0 )
    {
        if( $null -eq $script:vms -or $listView.SelectedIndex -gt $script:vms.Count )
        {
            Write-Error -Message "Internal error : selected grid view index $($listView.SelectedIndex) greater than $(($script:vms|Measure-Object).Count)"]
            return
        }
        ForEach( $selection in $listView.selectedItems )
        {
            $loopIterations++
            if( $hypervisorType -ieq 'Hyper-V' )
            {
                $vm = $script:vms | Where-Object Name -eq $selection.Name
            }
            else
            {
                $vm = $script:vms | Where-Object VMName -eq $selection.Name
            }
            if( $null -eq $vm )
            {
                Write-Warning "Could not find VM for selected item $($selection.Name) out of $($script:vms.Count)"
                continue
            }
            if( $listview.selectedItems.Count -gt 1 -and ( [string]::IsNullOrEmpty( $answer ) -or $answer -eq 'No' ))
            {
                [string]$buttons = 'YesNoCancel'
                [string]$prompt = "Are you sure you want to connect to $($listview.selectedItems.Count - $loopIterations + 1) VMs?`nYes for all , No for $($selection.Name) only or Cancel for none"
                ## make it modal to the main window
                $answer = [Windows.MessageBox]::Show( $mainWindow , $prompt , 'Confirm Multiple Connections' , $buttons ,'Question' )
                if( $answer -ieq 'cancel' )
                {
                    return
                }
            }
            $address = $null
            
            if( $console )
            {
                if( $hypervisorType -ieq 'Hyper-V' )
                {
                    ## TODO see if there is already a running process with these arguments (ish) and offer to activate that instead
                    $consoleProcess = $null
                    [hashtable]$consoleArguments = @{
                        FilePath = 'vmconnect.exe' 
                        ArgumentList = @( $WPFtextBoxHyperVHost.Text , "`"$($vm.Name)`"" , '-G' , $vm.Id )
                        PassThru = $true
                    }
                    $consoleProcess = Start-Process @consoleArguments
                    if( $null -eq $consoleProcess )
                    {
                        [void][Windows.MessageBox]::Show( "Failed to run $($consoleArguments[ 'filepath']) $($consoleArguments[ 'argumentlist' ] -join ' ') :`r`n$(Error[0])" , 'Hypervisor Console Error' , 'Ok' ,'Error' )
                    }

                }
                else
                {
                    ## TODO launch VMware console - web or local app?
                }
            }
            elseif( $radioButton.IsChecked )
            {
                ## TODO do we allow IPv6 ?
                if( $hypervisorType -ieq 'vmware' )
                {
                    $address = $vm.IPaddress | Where-Object { $_ -match '^\d+\.' -and $_ -ne '127.0.0.1' -and $_ -ne '::1' -and $_ -notmatch '^169\.254\.' }
                }
                else
                {
                    $address = Get-VMNetworkAdapter -VM $vm | Select-Object -ExpandProperty IPAddresses | Where-Object { $_ -match '^\d+\.' -and $_ -ne '127.0.0.1' -and $_ -ne '::1' -and $_ -notmatch '^169\.254\.' }
                }
                if( $null -eq $address )
                {
                    [void][Windows.MessageBox]::Show( "No IP address for $($vm.VmName)" , 'Hypervisor Error' , 'Ok' ,'Error' )
                }
                elseif( $address -is [array] -and $address.Count -gt 1 )
                {
                    [void][Windows.MessageBox]::Show( "$($address.Count) IP addresses for $($vm.VmName)" , 'Hypervisor Error' , 'Ok' ,'Error' )
                    ## TODO do we ask them to select one? Try in turn?
                    $address = $null
                }
            }
            elseif( $hypervisorType -ieq 'Hyper-V' )
            {
                $address = $vm.Name
            }
            else
            {
                $address = $vm.Hostname
            }
            if( $address )
            {
                if( -Not [string]::IsNullOrEmpty( $rdpportTextbox.Text ) )
                {
                    if( $rdpportTextbox.Text -notmatch '^\d+$' )
                    {
                        [void][Windows.MessageBox]::Show( "Port `"$($rdpportTextbox.Text)`" is invalid" , 'Hypervisor Error' , 'Ok' ,'Error' )
                        $address = $null
                    }
                    else
                    {
                        $address = "$($address):$($rdpportTextbox.Text)"
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
                else
                {
                    Write-Warning -Message "No address for $hypervisorType VM $($listView.SelectedItem)"
                }
            }
        }
    }
    else
    {
        [void][Windows.MessageBox]::Show( "No VM selected" , "$hypervisorType Error" , 'Ok' ,'Error' )
    }
}
    
Function Process-Action
{
    Param
    (
        $GUIobject , 
        [string]$Operation 
        ##$context  ,
        ##$thisObject
    )

    $thisObject = $_

    $thisObject.Handled = $true

    Write-Verbose -Message "Process-Action $operation "

    if( $GUIobject )
    {  
        [array]$selectedVMs = @( $GUIobject.selectedItems )
        if( $null -eq $selectedVMs -or $selectedVMs.Count -eq 0 )
        {
            Write-Verbose -Message "No items selected"
            [Windows.MessageBox]::Show( $mainWindow , 'No VMs selected' , 'Error' , 'OK' ,'Exclamation' )
            return
        }

        [hashtable]$hypervParameters = @{}
        if( -Not [string]::IsNullOrEmpty( $WPFtextBoxHyperVHost.Text ) )
        {
            $hypervParameters.Add( 'ComputerName' , $WPFtextBoxHyperVHost.Text.Trim() )
        }
        [string]$answer = $null
        [int]$loopIterations = 0
        [hashtable]$clipboardParameters = @{}

        [array]$jobs = @( ForEach( $selection in $selectedVMs )
        {
            $loopIterations++
            
            if( $operation -match 'Hyper' )
            {
                $vm = $script:vms | Where-Object Name -eq $selection.Name
            }
            else
            {
                $vm = $script:vms | Where-Object VMName -eq $selection.Name
            }
            if( $null -eq $vm )
            {
                Write-Warning "Could not find VM for selected item $($selection.Name) out of $($script:vms.Count)"
                ## might not be fatal as some operations don't use vm
            }
            if( $operation -ieq 'DeleteComputer' ) ## not Hyper-V context
            {
                Write-Verbose -Message "Deleting computer $selection"
                $GUIobject.Items.RemoveAt( $selection )
                continue
            }
            if( $Operation -match 'PowerOn|Detail|Resume|Clipboard|TakeSnapshot|((Manage|Revert|Delete).*Snapshot)' ) ## don't need to prompt or will prompt with more information later
            {
                $answer = 'yes'
            }
            elseif( [string]::IsNullOrEmpty( $answer ) -or $answer -eq 'No' )
            {
                [string]$buttons = 'YesNo'
                [string]$prompt = $(if( $selectedVMs.Count -gt 1 )
                {
                    "Are you sure you want to $($operation -replace '^HyperV_' -creplace '([a-zA-Z])([A-Z])' , '$1 $2') $($selectedVMs.Count - $loopIterations + 1) VMs?`nYes for All , No for $($selection.Name) Only or Cancel for None"
                    $buttons = 'YesNoCancel'
                }
                else
                {
                    "Are you sure you want to $($operation -replace '^HyperV_' -creplace '([a-zA-Z])([A-Z])' , '$1 $2') $($selection.Name)?"
                })
                ## make it modal to the main window
                $answer = [Windows.MessageBox]::Show( $mainWindow , $prompt , 'Confirm Power Operation' , $buttons ,'Question' )
            }
            if( [string]::IsNullOrEmpty( $answer ) -or $answer -ieq 'cancel' )
            {
                $answer = $null
                break
            }
            ## else if $answer = no then we are just performing on this VM and will prompt again next time round this loop           
            if( $operation -ieq 'NameToClipboard' )
            {
                $selection.Name | Set-Clipboard @clipboardParameters
            }
            elseif( $operation -ieq 'HyperV_PowerOn' )
            {
                Hyper-V\Start-VM -VMName $selection.Name -Passthru @hypervParameters
            }
            elseif( $operation -ieq 'HyperV_Shutdown' )
            {
                Hyper-V\Stop-VM -VMName $selection.Name -Passthru @hypervParameters
            }
            elseif( $operation -ieq 'HyperV_PowerOff' )
            {
                Hyper-V\Stop-VM -VMName $selection.Name -Passthru -TurnOff @hypervParameters -Force -Confirm:$false
            }
            elseif( $operation -ieq 'HyperV_Restart' )
            {
                Hyper-V\Restart-VM -VMName $selection.Name -Passthru @hypervParameters -Force -Confirm:$false
            }
            elseif( $operation -ieq 'HyperV_Resume' )
            {
                Hyper-V\Resume-VM -VMName $selection.Name -Passthru @hypervParameters
            }
            elseif( $operation -ieq 'HyperV_Suspend' )
            {
                Hyper-V\Suspend-VM -VMName $selection.Name -Passthru @hypervParameters
            }
            elseif( $operation -ieq 'HyperV_RevertLatestSnapshot' -or $operation -ieq 'HyperV_DeleteLatestSnapshot')
            {
                $latestCheckPoint = $null
                $latestCheckPoint = Get-VMCheckpoint -VMName $selection.Name @hypervParameters | Sort-Object -Property CreationTime -Descending | Select-Object -First 1
                if( $null -ne $latestCheckPoint )
                {
                    if( [Windows.MessageBox]::Show( $mainWindow , "$($latestCheckPoint.CreationTime.ToString('g')) : $($latestCheckPoint.Name)",
                        "$(if( $operation -match 'delete' ) { 'Delete' } else { 'Restore' }) Snapshot on $($selection.Name)" , 'YesNoCancel' ,'Question' ) -ieq 'yes' )
                    {
                        ## do not need hypervparameters as cannot user -computername with a snapshot object as that contains the remote details
                        if( $operation -match 'delete' )
                        {
                            Hyper-V\Remove-VMCheckpoint  -VMSnapshot $latestCheckPoint -Confirm:$false -Passthru
                        }
                        else
                        {
                            Hyper-V\Restore-VMCheckpoint -VMSnapshot $latestCheckPoint -Confirm:$false -Passthr
                        }
                    }
                }
                else
                {
                    Write-Warning -Message "No snapshots found for $($selection.Name)"
                    ## TODO No snapshots message
                }
            }         
            elseif( $operation -ieq 'HyperV_TakeSnapshot' )
            {        
                if( $textInputWindow = New-WPFWindow -inputXAML $textInputXAML )
                {
                    $WPFbtnInputTextOk.Add_Click({ 
                        $_.Handled = $true
                        $textInputWindow.DialogResult = $true 
                        $textInputWindow.Close()  })
                    $textInputWindow.Title = "New Snapshot"
                    $WPFlblInputTextLabel.Content = "Enter Snapshot Name"
                    if( $textInputWindow.ShowDialog() )
                    {
                        [hashtable]$snapshotParameters = @{}
                        if( $WPFtextboxInputText.Text.Length )
                        {
                            $snapshotParameters.Add( 'SnapshotName' , $WPFtextboxInputText.Text.Trim() )
                        }
                        Hyper-V\Checkpoint-VM -VMName $selection.Name -Passthru @hypervParameters @snapshotParameters
                    }
                }
            }
            elseif( $operation -ieq 'HyperV_Rename' )
            {        
                if( $textInputWindow = New-WPFWindow -inputXAML $textInputXAML )
                {
                    $WPFbtnInputTextOk.Add_Click({ 
                        $_.Handled = $true
                        $textInputWindow.DialogResult = $true 
                        $textInputWindow.Close()  })
                    $textInputWindow.Title = "Rename $($selection.Name)"
                    $WPFlblInputTextLabel.Content = "Enter New VM Name"
                    if( $textInputWindow.ShowDialog() )
                    {
                        [string]$newname = $WPFtextboxInputText.Text.Trim().Trim('"')
                        if( [string]::IsNullOrEmpty( $newname ) )
                        {
                            Write-Error "New name `"$newname`" too short"
                        }
                        else
                        {
                            if( $newname -ieq $selection.Name )
                            {
                                Write-Error "New name $newname is the same"
                            }
                            elseif( $null -ne ($existingVM = Hyper-V\Get-VM -Name $newname -ErrorAction SilentlyContinue @hypervParameters ) )
                            {
                                Write-Error "VM $newname already exists"
                            }
                            else
                            {
                                Hyper-V\Rename-VM -VM $vm -Passthru -NewName $newname
                            }
                        }
                    }
                }
            }
            elseif( $operation -ieq 'HyperV_ManageSnapshot' )
            {
                Show-SnapShotWindow -vm $selection.Name
            }
            elseif( $operation -ieq 'HyperV_Save' )
            {
                Hyper-V\Save-VM -VMName $selection.Name -Passthru @hypervParameters
            }
            elseif( $operation -ieq 'HyperV_Delete' -or $operation -ieq 'HyperV_DeleteIncludingDisks' )
            {
                $disks = $null
                if( $operation -ieq 'HyperV_DeleteIncludingDisks' )
                {
                    $disks = @( Hyper-V\Get-VMHardDiskDrive -VMName $selection.Name @hypervParameters )
                    Write-Verbose -Message "Got $($disks.Count) disks for VM $($disks.VMName)"
                }
                $removal = $null
                $removal = Hyper-V\Remove-VM -VMName $selection.Name -Passthru -Force -Confirm:$false @hypervParameters
                if( $? -and $null -ne $removal )
                {
                    ForEach( $disk in $disks )
                    {
                        Write-Verbose -Message "Deleting disk $($disk.Path)"
                        ## Could be remote so we use WMI with the CIM session in the disks object
                        $file = $null
                        ## needs backslashes escaping
                        $file = Get-CimInstance -ClassName cim_datafile -Filter "Name = '$($disk.Path -replace '\\' , '\\')'" -CimSession $disk.CimSession
                        if( $null -ne $file )
                        {
                            Remove-CimInstance -InputObject $file -CimSession $disk.CimSession -Confirm:$false
                        }
                        else
                        {
                            Write-Warning -Message "Failed to get file for disk $($disk.Path)"
                        }
                    }
                }
                else
                {
                    Write-Verbose -Message "Not deleting disks for $($selection.Name) as deleting VM errored"
                }
            }
            elseif( $operation -ieq 'HyperV_Detail' )
            {
                $details = $null
                $details = Hyper-V\Get-VM -VMName $selection.Name @hypervParameters
                if( $null -ne $details )
                {
                    [array]$hardDrives = @( Get-VMHardDiskDrive -VM $details )
                    [string[]]$diskDetails = @( ForEach( $disk in $hardDrives )
                    {
                        [string]$size = ''
                        $file = $null
                        $file = Get-CimInstance -ClassName cim_datafile -Filter "Name = '$($disk.Path -replace '\\' , '\\')'" -CimSession $disk.CimSession
                        if( $null -ne $file )
                        {
                            $size = "$([math]::Round( $file.filesize / 1GB , 1 ))"
                        }
                        "$($disk.Path) ($($size) GB)"
                    })
                    [array]$snapshots = @( Get-VMSnapshot -VM $details | Sort-Object -Property CreationTime )
                    [array]$NICs = @( Get-VMNetworkAdapter -VM $details )
                    $form = New-Object System.Windows.Forms.Form
                    $form.Text = $selection.Name
                    $form.Size = New-Object System.Drawing.Size(800, 400)
                    $form.StartPosition = "CenterScreen"

                    $listView = New-Object System.Windows.Forms.ListView
                    $listView.View = 'Details'
                    $listView.FullRowSelect = $true
                    $listView.GridLines = $true
                    $listView.Dock = 'Fill'
                    $listView.Columns.Add("Setting", 180)
                    $listView.Columns.Add("Value", 600 )

                    $data = @(
                        @{ Setting = "Notes"; Value = $details.Notes }
                        @{ Setting = "State"; Value = $details.State.ToString() }
                        @{ Setting = "vCPU"; Value = $details.ProcessorCount }
                        @{ Setting = "Resource Metering Enabled"; Value = $details.ResourceMeteringEnabled.ToString() }
                        @{ Setting = "Uptime"; Value = $details.Uptime.ToString() }
                        @{ Setting = "Version"; Value = $details.Version.ToString() }
                        @{ Setting = "Memory Startup MB"; Value = $details.MemoryStartup / 1MB }
                        @{ Setting = "Memory Assigned MB"; Value = $details.MemoryAssigned / 1MB }
                        @{ Setting = "Memory Minimum MB"; Value = $details.MemoryMinimum / 1MB }
                        @{ Setting = "Memory Maximum MB"; Value = $details.MemoryMaximum / 1MB }
                        @{ Setting = "Dynamic Memory Enabled"; Value = $details.DynamicMemoryEnabled.ToString() }
                        @{ Setting = "Hard Drives"; Value = $diskDetails -join "`n" }
                        @{ Setting = "NICs"; Value = $NICs.Count }                        
                        @{ Setting = "IP Addresses"; Value = ( $NICs | Select-Object -ExpandProperty IPAddresses -ErrorAction SilentlyContinue | Where-Object { $_ -notmatch ':' } ) -join ' , ' }
                        @{ Setting = "Snapshots"; Value = $snapshots.Count }
                        @{ Setting = "Created"; Value = $details.CreationTime.ToString('G') }
                        @{ Setting = "Oldest Snapshot"; Value = $(if( $null -ne $snapshots -and $snapshots.Count -gt 0 ) { "$($snapshots[0].CreationTime.ToString('G')) ($($snapshots[0].Name) ($($snapshots[0].Notes)))" })}
                        @{ Setting = "Latest Snapshot"; Value = $(if( $null -ne $snapshots -and $snapshots.Count -gt 0 ) { "$($snapshots[-1].CreationTime.ToString('G')) ($($snapshots[-1].Name) ($($snapshots[-1].Notes)))" })}  
                    )

                    foreach ($item in $data)
                    {
                        $listItem = New-Object System.Windows.Forms.ListViewItem($item.Setting)
                        $listItem.SubItems.Add($item.Value)
                        $null = $listView.Items.Add($listItem)
                    }

                    $form.Controls.Add($listView)
                    $form.Add_Shown({ $form.Activate() })
                    [void]$form.ShowDialog()

                }
                else
                {
                    ## TODO error dialogue
                }
            }
            elseif( $operation -ieq 'HyperV_EnableResourceMetering' )
            {
                Hyper-V\Enable-VMResourceMetering -VMName $selection.Name @hypervParameters
            }
            elseif( $operation -ieq 'HyperV_DisableResourceMetering' )
            {
                Hyper-V\Disable-VMResourceMetering -VMName $selection.Name @hypervParameters
            }
            elseif( $operation -ieq 'HyperV_DisConnectNIC' )
            {
                ## TODO what if more than one NIC?
                Hyper-V\Disconnect-VMNetworkAdapter -VMName $selection.Name @hypervParameters
            }
            elseif( $operation -imatch 'HyperV_ConnectNIC*' )
            {
                [string]$switchType = $Operation -replace '^HyperV_ConnectNIC'
                ## TODO need to get virtual switch names and if more than one prompt for the required one
                [array]$switches = @( Hyper-V\Get-VMSwitch -SwitchType $switchType @hypervParameters )
                if( $switches.Count -gt 1 )
                {
                    Write-Warning "VM $($selecion.Name) has $($switches.Count) NICs which isn't yet implemented sorry"
                }
                Hyper-V\Connect-VMNetworkAdapter -VMName $selection.Name -SwitchName $switches[ 0 ].Name @hypervParameters
            }      
            else
            {
                Write-Warning -Message "Unimplemented operation $Operation"
            }
            $clipboardParameters[ 'Append' ] = $true
        })
        if( $null -ne $jobs -and $jobs.count -gt 0 )
        {
            $jobs | Write-Verbose
        }
    }
}

#endregion pre-main


if( [string]::IsNullOrEmpty( $tempFolder ) )
{
    Throw "No temp folder"
}
if( -Not ( Test-Path -Path $tempFolder -PathType Container ) -and -Not ( New-Item -Path $tempFolder -ItemType Directory -Force ) )
{
    Throw "Failed to create temp folder $tempFolder"
}
         
<#
if( $usemsrdc -and [string]::IsNullOrEmpty( $address ) )
{
    Throw "Must specify computer to connect to via -address when using msrdc mode"
}
#>

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
   
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowRect(IntPtr hWnd, string lpString);
   
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
   
   [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
   public static extern int SetWindowText(IntPtr hWnd, string lpString );
   
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
            Start-RemoteSessionFromHypervisor -hypervisorType VMware
        })

        $WPFbtnLaunchHyperVOptions.Add_Click({
            $_.Handled = $true
            Write-Verbose "Launch on Hyper-V clicked"
            Start-RemoteSessionFromHypervisor -hypervisorType Hyper-V
        })
        
        $WPFbtnLaunchHyperVConsole.Add_Click({
            $_.Handled = $true
            Write-Verbose "Launch Console on Hyper-V clicked"
            Start-RemoteSessionFromHypervisor -hypervisorType Hyper-V -console
        })

        $WPFlistViewVMwareVMs.add_MouseDoubleClick({
            $_.Handled = $true
            Write-Verbose "Launch on VMware list item double clicked"
            Start-RemoteSessionFromHypervisor -hypervisorType VMware
        })
        
        $WPFlistViewHyperVVMs.add_MouseDoubleClick({
            $_.Handled = $true
            Write-Verbose "Launch on Hyper-V list item double clicked"
            Start-RemoteSessionFromHypervisor -hypervisorType Hyper-V
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
        
        $WPFbuttonHyperVConnect.Add_Click({
            $_.Handled = $true
            Write-Verbose "Hyper-V Connect clicked"
            $hyperVhost = $WPFtextBoxHyperVHost.Text.Trim() -replace '"'
            if( [string]::IsNullOrEmpty( $hyperVhost ) )
            {
                $hyperVhost = 'localhost'
            }
            Import-Module -Name Hyper-V
            
            Add-HyperVVMsToListView -hyperVhost $hyperVhost -filter $wpfTextBoxHyperVFilter.Text -regex $WPFcheckBoxHyperVRegEx.IsChecked -all $WPFcheckBoxHyperVAllVMs.IsChecked
        })
        
        if( -Not [string]::IsNullOrEmpty( $hypervHost ) )
        {
            $WPFtextBoxHyperVHost.Text = $hypervHost
        }
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

        $WPFbuttonVMwareApplyFilter.Add_Click({
            $_.Handled = $true
            Write-Verbose "VMware Apply Filter clicked"
            Add-VMwareVMsToListView -filter $wpfTextBoxVMwareFilter.Text -regex $WPFcheckBoxVMwareRegEx.IsChecked
        })
        
        $WPFbuttonHyperVApplyFilter.Add_Click({
            $_.Handled = $true
            Write-Verbose "Hyper-V Apply Filter clicked"
            Add-HyperVVMsToListView -filter $WPFtextBoxHyperVFilter.Text -regex $WPFcheckBoxHyperVRegEx.IsChecked -hyperVhost $WPFtextBoxHyperVHost.Text -all $WPFcheckBoxHyperVAllVMs.IsChecked
        })
        
        $WPFbuttonHyperVClearFilter.Add_Click({
            $_.Handled = $true
            Write-Verbose "Hyper-V Clear Filter clicked"
            $WPFtextBoxHyperVFilter.Text = ''
            Add-HyperVVMsToListView -filter $WPFtextBoxHyperVFilter.Text -regex $WPFcheckBoxHyperVRegEx.IsChecked -hyperVhost $WPFtextBoxHyperVHost.Text -all $WPFcheckBoxHyperVAllVMs.IsChecked
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
        
        $WPFdeleteComputersContextMenu.Add_Click( { Process-Action -GUIobject $WPFcomboboxComputer -Operation 'DeleteComputer' -Context $_  -thisObject $this } )

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
        
        $WPFHyperVPowerOnContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_PowerOn' })
        $WPFHyperVPowerOffContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_PowerOff' })
        $WPFHyperVShutdownContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_ShutDown' })
        $WPFHyperVRestartContextMenu.Add_Click(  { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Restart' })
        $WPFHyperVDetailContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Detail' })
        $WPFHyperVDeleteContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Delete' })
        $WPFHyperVDeleteAllContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_DeleteIncludingDisks' })
        
        $WPFHyperVEjectCDContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_EjectCD' })
        $WPFHyperVMountCDContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_MountCD' })
        $WPFHyperVNameToClipboard.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'NameToClipboard' })
        $WPFHyperVSaveContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Save' })
        $WPFHyperVNewVMFromTemplateContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_New' })
        $WPFHyperVNewVMContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_NewFromTemplate' })
        $WPFHyperVReconfigureMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Reconfigure' })
        $WPFHyperVConnectNICInternalContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_ConnectNICInternal' })
        $WPFHyperVConnectNICExternalContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_ConnectNICExternal' })
        $WPFHyperVConnectNICPrivateContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_ConnectNICPrivate' })
        $WPFHyperVDisconnectNICContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_DisconnectNIC' })
        $WPFHyperVRenameMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Rename' })
        $WPFHyperVSuspendContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Suspend' })
        
        $WPFHyperVResumeContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Resume' })
        $WPFHyperVEnableResourceMeteringContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_EnableResourceMetering' })
        $WPFHyperVDisableResourceMeteringContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_DisableResourceMetering' })
        $WPFHyperVTakeSnapshotContextMenu.Add_Click(  { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_TakeSnapshot' })
        $WPFHyperVManageSnapshotContextMenu.Add_Click(  { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_ManageSnapshot' })
        $WPFHyperVRevertLatestSnapshotContextMenu.Add_Click(  { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_RevertLatestSnapshot' })
        $WPFHyperVDeleteLatestSnapshotContextMenu.Add_Click(  { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_DeleteLatestSnapshot' })

        $wpfchkboxDoNotApply.IsChecked = (-Not $useOtherOptions ) ## don't want it on by default as passes blank password to mstsc giving failed logon

        [string]$msrdcExe = Get-Msrdc
        $WPFchkboxmsrdc.IsEnabled = -Not [string]::IsNullOrEmpty( $msrdcExe )
        $WPFchkboxmsrdc.IsChecked = $usemsrdc -or -Not [string]::IsNullOrEmpty($msrdcExe ) 
        
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
        if( -Not (Test-Path -Path $configKey ) )
        {
            if( -Not ( New-Item -Path $configKey -ItemType Key -Force ) )
            {
                Write-Warning -Message "Problem creating $configKey"
            }
        }
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
