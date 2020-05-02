#requires -version 3

<#
    Modification History:

    @guyrleech 04/11/2019  Initial release
    @guyrleech 04/11/2019  Added client drive mapping to rdp file created for credentials and added -extraRDPSettings parameter
    @guyrleech 05/11/2019  Added code to reconnect to VMware if errors with "not connected", e.g. times out. Fixed bug where can't take snapshot or consolidate disks if no snapshots exist. Improved mstsc address detection logic. Fixes when already connected to VMware
    @guyrleech 06/11/2019  Fixed reconnection code if connection has timed out.
                           Added Details button to Snapshot window and show date of last revert.
                           Add functionality to find existing mstsc window & activate unless -noreusemstsc.
                           Don't pass username via rdp file to mstsc if find saved credential for that connection.
    @guyrleech 06/11/2019  Added message if VMware operation fails with "not connected" that should try clicking Refresh
                           Pull username from existing VMware connection
                           Event context menu added
                           Implemented policies for maximum memory and cpus when reconfiguring
                           Changed -webconsole to -consoleBrowser
                           Added setting of PowerCLI configuration before connecting
    @guyrleech 06/11/2019  Improved mstsc working address determination when DNS resolution fails
    @guyrleech 18/11/2019  Activate existing vmrc.exe console process if already running from this script for the selected VM
                           Added "Mstsc (new)" option
                           Changed -webconsole to -consoleBrowser to take path to 32 bit web browser
    @guyrleech 16/12/2019  Added menu option to revert to latest snapshot and option to power on VM after snapshot restore.
                           Gives snapshot name and date of creation when prompting for snapshot operation
    @guyrleech 23/12/2019  Added -credential parameter
                           Added folder, hardware version and datastore names to grid view
    @guyrleech 07/01/2020  Added vMware resource usage to grid view
    @guyrleech 08/01/2020  Added F5 keyboard handler
    @guyrleech 29/03/2020  Added check box and parameter for performance data display. Backup function added. Hosts display added.
    @guyrleech 20/04/2020  Added date/time connected to title and number of VMs shown
                           Added snapshot tree removal button & code
    @guyrleech 02/05/2020  Added console screenshot feature & -screenshotfolder to persist the image files
#>

<#
.SYNOPSIS

Simple VMware vSphere/ESXi managament UI - power, delete, snapshot, reconfigure and console/mstsc

.DESCRIPTION

Uses VMware PowerCLI. Permissions to perform actions will be as per the permissions defined for the account used to connect to VMware as defined in vSphere/ESXi

.PARAMETER server

One or more VMware vCenter or ESXi hosts to connect to, separated by commas. Automatically saved to the registry for subsequent invocations

.PARAMETER username

The username, domain qualified or UPN, to connect to VMware with. Use -passthru to pass through credentials used to run script. Will prompt for credentials if no saved credentials in user's local profile or via New-VICredentialStoreItem

.PARAMETER credential

Credential object for connecting to vCenter/ESXi

.PARAMETER backupSnapshotName

The name of the snapshot taken when performing a backup 

.PARAMETER vmName

The name of a VM or pattern to include in the GUI

.PARAMETER performanceData

Enable fetching and displaying of per-VM performance data. Can be changed via a toggle within the UI too

.PARAMETER rdpPort

The port used to connect to RDP via mstsc. Only use when not using the default port of 3389

.PARAMETER mstscParams

Additional parameters to pass to mstsc.exe such as window dimenstions via /w: and /h:

.PARAMETER doubleClick

The action to perform when a VM is double clicked in the GUI. The default is mstsc

.PARAMETER saveCredentials

Saves credentials to the user's local profile. The password is encrypted such that only the user running the script can decrypt it and only on the machine running the script

.PARAMETER ipv6

Include IPv6 addresses in the display

.PARAMETER passThru

Connect to vCenter using the credentials running the script (where windows domain authentication is configured and working in VMware)

.PARAMETER consoleBrowser

Do not attempt to run vmrc.exe when launching the VMware console for a VM but use this browser instead (must be 32 bit)

.PARAMETER noRdpFile

Do not create an .rdp file to specify the username for mstsc

.PARAMETER noReuseMstsc

Always create a new mstsc process otherwise it will find an existing one for that VM and restore the window

.PARAMETER showPoweredOn

Include powered on VMs in the GUI display

.PARAMETER showPoweredOff

Include powered on VMs in the GUI display

.PARAMETER showSuspended

Include suspended VMs in the GUI display

.PARAMETER showAll

Show VMs in all power states in the GUI

.PARAMETER extraRDPSettings

Extra settings to put into the .rdp file created to specify username for mstsc to connect to

.PARAMETER maxEvents

The maximum number of events to retrieve from VMware

.PARAMETER sortProperty

The VM property to sort the display on.

.PARAMETER screenshotFolder

Folder in which to store console screenshot files. Uuser's temp folder is used if not specified and files deleted.

.EXAMPLE

& '.\VMware GUI.ps1' -server grl-vcenter01 -username guy@leech.com -saveCredentials -doubleClick console

Show all powered on VMs (the default if no -show* parameters are specified) on the VMware vCenter server grl-vcenter01, connecting as guy@leech.com for which the password will be prompted and then saved to the user's local profile.
Double clicking on any VM will launch the VMware Remote Console (vmrc) if installed otherwise a browser based console. The server(s) will be stored in HKCU for the user running the script so becomes the default if none is specified

.EXAMPLE

& '.\VMware GUI.ps1' -showAll -mstscParams "/w:1916 /h:1016"

Show all VMs on the VMware server stored in the registry using the encrypted credentials stored in the user's local profile. If the mstsc option is used, additionally pass "/w:1916 /h:1016" (so the window fits nicely on a full HD monitor)

.NOTES

The latest VMware PowerCLI module required by the script can be installed if you have internet access by running the following PowerShell command as an administrator "Install-Module -Name VMware.PowerCLI"

RDP file settings for -extraRDPSettings option available at https://docs.microsoft.com/en-us/windows-server/remote/remote-desktop-services/clients/rdp-files

#>

[CmdletBinding()]

Param
(
    [string[]]$server ,
    [string]$username ,
    [System.Management.Automation.PSCredential]$credential ,
    [string]$vmName = '*' ,
    [int]$rdpPort = 3389 ,
    [string]$mstscParams ,
    [ValidateSet('mstsc','console','PowerOn','reconfigure','snapshots','Delete','backup')]
    [string]$doubleClick = 'mstsc',
    [switch]$saveCredentials ,
    [switch]$ipv6 ,
    [switch]$passThru ,
    [string]$consoleBrowser ,
    [switch]$noRdpFile ,
    [bool]$showPoweredOn = $true,
    [bool]$showPoweredOff ,
    [bool]$showSuspended ,
    [switch]$showAll ,
    [switch]$performanceData ,
    [string]$backupSnapshotName = 'For Backup Taking via Script' ,
    [int]$lastSeconds = 300 ,
    [string]$exclude = 'uptime' ,
    [switch]$noReuseMstsc ,
    [int]$maxEvents = 10000 ,
    [string]$screenShotFolder ,
    [string]$sortProperty = 'Name' ,
    [string[]]$extraRDPSettings = @( 'drivestoredirect:s:*' ) ,
    [string]$regKey = 'HKCU:\Software\Guy Leech\Simple VMware Console' ,
    [string]$vmwareModule = 'VMware.VimAutomation.Core'
)

[hashtable]$powerstate = @{
    'Suspended'  = 'Suspended'
    'PoweredOn'  = 'On'
    'PoweredOff' = 'Off'
}

[hashtable]$powerOperation = @{
    'PowerOn'  = 'Start-VM'
    'PowerOff' = 'Stop-VM'
    'Suspend' = 'Suspend-VM'
    'Reset' = 'Restart-VM'
    'Shutdown' = 'Shutdown-VMGuest'
    'Restart' = 'Restart-VMGuest'
}

$pinvokeCode = @'
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow); 
'@

[string]$youAreHereSnapshot = 'YouAreHere' 

#region XAML
[string]$mainwindowXAML = @'
<Window x:Class="VMWare_GUI.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:VMWare_GUI"
        mc:Ignorable="d"
        Title="Simple VMware Console (SVC) by @guyrleech" Height="500" Width="1000">
    <DockPanel Margin="10,10,-70,2" HorizontalAlignment="Left" >
        <Button x:Name="btnHosts" Content="_Hosts" HorizontalAlignment="Left" Height="35"  Margin="10,0,0,0" VerticalAlignment="Bottom" Width="145" DockPanel.Dock="Bottom"/>
        <Button x:Name="btnFilter" Content="_Filter" HorizontalAlignment="Left" Height="35" Margin="247,0,0,-35" VerticalAlignment="Bottom" Width="145" DockPanel.Dock="Bottom"/>
        <Button x:Name="btnRefresh" Content="_Refresh" HorizontalAlignment="Left" Height="35" Margin="500,0,0,-35" Width="145" VerticalAlignment="Bottom" DockPanel.Dock="Bottom" />
        <Button x:Name="btnDatastores" Content="_Datastores" HorizontalAlignment="Left" Height="35" Margin="690,0,0,-35" Width="145" VerticalAlignment="Bottom" DockPanel.Dock="Bottom" />
        <CheckBox x:Name="chkPerfData" Content="_Performance Data" HorizontalAlignment="Left" Height="35" Margin="850,0,0,-35" Width="150" VerticalAlignment="Bottom" DockPanel.Dock="Bottom"/>
        <DataGrid HorizontalAlignment="Stretch" VerticalAlignment="Top" x:Name="VirtualMachines" >
            <DataGrid.ContextMenu>
                <ContextMenu>
                    <MenuItem Header="Console" x:Name="ConsoleContextMenu" />
                    <MenuItem Header="Mstsc" x:Name="MstscContextMenu" />
                    <MenuItem Header="Mstsc (New)" x:Name="MstscNewContextMenu" />
                    <MenuItem Header="Snapshots" x:Name="SnapshotContextMenu" />
                    <MenuItem Header="Revert to latest snapshot" x:Name="LatestSnapshotRevertContextMenu" />
                    <MenuItem Header="Reconfigure" x:Name="ReconfigureContextMenu" />
                    <MenuItem Header="Events" x:Name="EventsContextMenu" />
                    <MenuItem Header="Backup" x:Name="BackupContextMenu" />
                    <MenuItem Header="Delete" x:Name="DeleteContextMenu" />
                    <MenuItem Header="Screenshot" x:Name="ScreenshotContextMenu" />
                    <MenuItem Header="Power" x:Name="PowerContextMenu">
                        <MenuItem Header="Power On" x:Name="PowerOnContextMenu" />
                        <MenuItem Header="Power Off" x:Name="PowerOffContextMenu" />
                        <MenuItem Header="Suspend" x:Name="SuspendContextMenu" />
                        <MenuItem Header="Reset" x:Name="ResetContextMenu" />
                        <MenuItem Header="Shutdown" x:Name="ShutdownContextMenu" />
                        <MenuItem Header="Restart" x:Name="RestartContextMenu" />
                    </MenuItem>
                </ContextMenu>
            </DataGrid.ContextMenu>
        </DataGrid>
    </DockPanel>
</Window>
'@

[string]$screenshotXAML = @'
<Window x:Class="VMWare_GUI.Screenshot"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:VMWare_GUI"
        mc:Ignorable="d"
        Title="Screenshot" Height="700" Width="950">
    <Grid>
        <Image x:Name="imgScreenshot" HorizontalAlignment="Left" Height="620" Margin="11,31,0,0" VerticalAlignment="Top" Width="922"/>

    </Grid>
</Window>
'@

[string]$hostsXAML = @'
<Window x:Class="VMWare_GUI.Hosts"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:VMWare_GUI"
        mc:Ignorable="d"
        Title="Hosts" Height="450" Width="800">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition />
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <DataGrid Name="HostList">
            <DataGrid.Columns>
                <DataGridTextColumn Binding="{Binding Name}" Header="Name"/>
                <DataGridTextColumn Binding="{Binding Parent}" Header="Parent"/>
                <DataGridTextColumn Binding="{Binding PowerState}" Header="Power State"/>
                <DataGridTextColumn Binding="{Binding ConnectionState}" Header="Connection State"/>
                <DataGridTextColumn Binding="{Binding Manufacturer}" Header="Manufacturer"/>
                <DataGridTextColumn Binding="{Binding Model}" Header="Model"/>
                <DataGridTextColumn Binding="{Binding Version}" Header="Version"/>
                <DataGridTextColumn Binding="{Binding Build}" Header="Build"/>
                <DataGridTextColumn Binding="{Binding RunningVMs}" Header="Running VMs"/>
                <DataGridTextColumn Binding="{Binding MemoryTotalGB}" Header="Memory Total (GB)"/>
                <DataGridTextColumn Binding="{Binding MemoryUsageGB}" Header="Memory Used (GB)"/>
                <DataGridTextColumn Binding="{Binding MemoryUsedPercent}" Header="Memory Used %"/>
                <DataGridTextColumn Binding="{Binding NumCPU}" Header="CPUs"/>
                <DataGridTextColumn Binding="{Binding CPUUsedPercent}" Header="CPU Used %"/>
            </DataGrid.Columns>
        </DataGrid>
    </Grid>
</Window>
'@

[string]$datastoresXAML = @'
<Window x:Class="VMWare_GUI.Datastores"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:VMWare_GUI"
        mc:Ignorable="d"
        Title="Datastores" Height="450" Width="800">
    <Grid>
            <Grid.RowDefinitions>
                <RowDefinition />
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <DataGrid Name="DatastoreList">
                <DataGrid.Columns>
                    <DataGridTextColumn Binding="{Binding Name}" Header="Name"/>
                    <DataGridTextColumn Binding="{Binding FreeSpaceGB}" Header="Free Space (GB)"/>
                    <DataGridTextColumn Binding="{Binding CapacityGB}" Header="Capacity (GB)"/>
                    <DataGridTextColumn Binding="{Binding UsedPercent}" Header="Used %"/>
                </DataGrid.Columns>
            </DataGrid>
        </Grid>
</Window>
'@

[string]$backupXAML = @'
<Window x:Class="VMWare_GUI.Backup"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:VMWare_GUI"
        mc:Ignorable="d"
        Title="Backup" Height="402.918" Width="465.895">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="333*"/>
            <ColumnDefinition Width="68*"/>
        </Grid.ColumnDefinitions>
        <ComboBox x:Name="comboDatastore" HorizontalAlignment="Left" Height="30" Margin="159,56,0,0" VerticalAlignment="Top" Width="255" Grid.ColumnSpan="2"/>
        <Label HorizontalAlignment="Left" Margin="39,59,0,0" VerticalAlignment="Top"/>
        <ComboBox x:Name="comboFolder" HorizontalAlignment="Left" Margin="159,120,0,0" VerticalAlignment="Top" Width="255" Height="30" Grid.ColumnSpan="2"/>
        <Label Content="Folder" HorizontalAlignment="Left" Height="32" Margin="23,120,0,0" VerticalAlignment="Top" Width="80"/>
        <Label Content="Datastore" HorizontalAlignment="Left" Height="30" Margin="23,56,0,0" VerticalAlignment="Top" Width="79"/>
        <CheckBox x:Name="chkAsync" Content="_Asynchronous " HorizontalAlignment="Left" Height="24" Margin="159,235,0,0" VerticalAlignment="Top" Width="161"/>
        <Button x:Name="btnBackupOK" Content="OK" HorizontalAlignment="Left" Height="35" Margin="23,298,0,0" VerticalAlignment="Top" Width="110" IsDefault="True"/>
        <Button x:Name="btnBackupCancel" Content="Cancel" HorizontalAlignment="Left" Height="35" Margin="195,298,0,0" VerticalAlignment="Top" Width="110" IsCancel="True"/>
        <TextBox x:Name="txtBackupName" HorizontalAlignment="Left" Height="30" Margin="159,178,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="255" Grid.ColumnSpan="2"/>
        <Label Content="Backup VM Name" HorizontalAlignment="Left" Height="30" Margin="23,178,0,0" VerticalAlignment="Top" Width="131"/>

    </Grid>
</Window>
'@

[string]$filtersXAML = @'
<Window x:Name="formFilters"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Filters" Height="358.391" Width="474.411" ShowInTaskbar="False">
    <Grid HorizontalAlignment="Left" Height="300" Margin="18,24,0,0" VerticalAlignment="Top" Width="433">
        <Label Content="VM Name/Pattern" HorizontalAlignment="Left" Margin="58,41,0,0" VerticalAlignment="Top" Width="129"/>
        <TextBox x:Name="txtVMName" HorizontalAlignment="Left" Height="26" Margin="187,45,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="217"/>
        <CheckBox x:Name="chkPoweredOn" Content="Powered _On" HorizontalAlignment="Left" Height="32" Margin="187,94,0,0" VerticalAlignment="Top" Width="217"/>
        <CheckBox x:Name="chkPoweredOff" Content="Powered O_ff" HorizontalAlignment="Left" Height="32" Margin="187,131,0,0" VerticalAlignment="Top" Width="217"/>
        <CheckBox x:Name="chkSuspended" Content="_Suspended" HorizontalAlignment="Left" Height="32" Margin="187,163,0,0" VerticalAlignment="Top" Width="217"/>
        <Button x:Name="btnFiltersOk" Content="OK" HorizontalAlignment="Left" Margin="45,245,0,0" VerticalAlignment="Top" Width="75" IsDefault="True"/>
        <Button x:Name="btnFiltersCancel" Content="Cancel" HorizontalAlignment="Left" Margin="148,245,0,0" VerticalAlignment="Top" Width="75" IsCancel="True"/>
    </Grid>
</Window>
'@

[string]$reconfigureXAML = @'
<Window x:Class="VMWare_GUI.Reconfigure"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:VMWare_GUI"
        mc:Ignorable="d"
        Title="Reconfigure" Height="450" Width="800">
    <Grid>
        <Label Content="vCPUs" HorizontalAlignment="Left" Margin="53,54,0,0" VerticalAlignment="Top" Width="105"/>
        <Label Content="Memory" HorizontalAlignment="Left" Margin="53,95,0,0" VerticalAlignment="Top" Width="105"/>
        <TextBox x:Name="txtvCPUs" HorizontalAlignment="Left" Height="26" Margin="158,58,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="72" />
        <TextBox x:Name="txtMemory" HorizontalAlignment="Left" Height="26" Margin="158,99,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="72"/>
        <ComboBox x:Name="comboMemory" HorizontalAlignment="Left" Margin="259,99,0,0" VerticalAlignment="Top" Width="120" IsReadOnly="True" IsEditable="True">
            <ComboBoxItem Content="MB"/>
            <ComboBoxItem Content="GB" IsSelected="True"/>
        </ComboBox>
        <Button x:Name="btnReconfigureOk" Content="OK" HorizontalAlignment="Left" Margin="52,365,0,0" VerticalAlignment="Top" Width="75" IsDefault="True"/>
        <Button x:Name="btnReconfigureCancel" Content="Cancel" HorizontalAlignment="Left" Margin="155,365,0,0" VerticalAlignment="Top" Width="75" IsCancel="True"/>
        <TextBox x:Name="txtNotes" HorizontalAlignment="Left" Height="56" Margin="158,159,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="221"/>
        <Label Content="Notes" HorizontalAlignment="Left" Height="25" Margin="53,150,0,0" VerticalAlignment="Top" Width="65"/>

    </Grid>
</Window>
'@

[string]$snapshotsXAML = @'
<Window x:Class="VMWare_GUI.Snapshots"
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

[string]$takeSnapshotXAML = @'
<Window x:Class="VMWare_GUI.Take_Snapshot"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:VMWare_GUI"
        mc:Ignorable="d"
        Title="Take Snapshot" Height="450" Width="800">
    <Grid x:Name="gridTakeSnapshotOptions">
        <Label Content="Name" HorizontalAlignment="Left" Margin="63,67,0,0" VerticalAlignment="Top" Height="34" Width="99"/>
        <Label Content="Description" HorizontalAlignment="Left" Margin="63,138,0,0" VerticalAlignment="Top" Height="34" Width="99"/>
        <CheckBox x:Name="chckSnapshotMemory" Content="Snapshot the virtual machine's memory" HorizontalAlignment="Left" Margin="63,248,0,0" VerticalAlignment="Top" Height="22" Width="368"/>
        <CheckBox x:Name="chkSnapshotQuiesce" Content="Quiesce guest file system (Needs VMware Tools installed)" HorizontalAlignment="Left" Margin="63,309,0,0" VerticalAlignment="Top" Width="403"/>
        <Button x:Name="btnTakeSnapshotOk" Content="OK" HorizontalAlignment="Left" Margin="68,366,0,0" VerticalAlignment="Top" Width="75" IsDefault="True"/>
        <Button x:Name="btnTakeSnapshotCancel" Content="Cancel" HorizontalAlignment="Left" Margin="171,366,0,0" VerticalAlignment="Top" Width="75" IsCancel="True"/>
        <TextBox x:Name="txtSnapshotName" HorizontalAlignment="Left" Height="22" Margin="196,67,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="333"/>
        <TextBox x:Name="txtSnapshotDescription" HorizontalAlignment="Left" Height="89" Margin="196,147,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="326"/>

    </Grid>
</Window>
'@

Function Refresh-Form
{
    $getError = $null
    $script:snapshots = @( VMware.VimAutomation.Core\Get-Snapshot -ErrorVariable getError -VM $(if( [string]::IsNullOrEmpty( $script:vmName ) ) { '*' } else { $script:vmName }))

    Write-Verbose "Got $($getError.Count) errors"
    if( $getError )
    {
        $getError[0].Exception.Message|Write-Verbose
    }
    if( $getError -and $getError.Count -and $getError[0].Exception.Message -match 'Not Connected' )
    {
        $connection = Connect-VIServer @connectParameters
        [void][Windows.MessageBox]::Show( "Server was not connected, please retry" , 'Connection Error' , 'Ok' ,'Exclamation' )
    }

    $script:datastores.Clear()
    Get-Datastore | ForEach-Object { $script:datastores.Add( $PSItem.Id , $PSItem.Name ) }

    Update-Form -form $mainForm -datatable $script:datatable -vmname $script:vmName
}

Function Load-GUI
{
    Param
    (
        [Parameter(Mandatory=$true)]
        $inputXaml
    )

    $form = $null
    $inputXML = $inputXaml -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window'
 
    [xml]$xaml = $inputXML

    if( $xaml )
    {
        $reader = New-Object -TypeName Xml.XmlNodeReader -ArgumentList $xaml

        try
        {
            $form = [Windows.Markup.XamlReader]::Load( $reader )
        }
        catch
        {
            Throw "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$_"
        }
 
        $xaml.SelectNodes( '//*[@Name]' ) | ForEach-Object `
        {
            Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global
        }
    }
    else
    {
        Throw "Failed to convert input XAML to WPF XML"
    }

    $form
}

Function Set-Filters
{
    Param
    (
        [string]$name
    )

    $filtersForm = Load-GUI -inputXaml $filtersXAML
    [bool]$result = $false
    if( $filtersForm )
    {
        $WPFtxtVMName.Text = $name
        $wpfchkPoweredOn.IsChecked = $showPoweredOn
        $wpfchkPoweredOff.IsChecked = $showPoweredOff
        $wpfchkSuspended.IsChecked = $showSuspended

        $WPFbtnFiltersOk.Add_Click({ 
            $filtersForm.DialogResult = $true 
            $filtersForm.Close()  })

        $WPFbtnFiltersCancel.Add_Click({
            $filtersForm.DialogResult = $false 
            $filtersForm.Close() })

        $result = $filtersForm.ShowDialog()
    }
    $result
}

Function Set-Configuration
{
    Param
    (
        $vm
    )

    $reconfigureForm = Load-GUI -inputXaml $reconfigureXAML
    [bool]$result = $false
    if( $reconfigureForm )
    {
        $reconfigureForm.Title += " $($vm.Name)"
        $WPFtxtvCPUs.Text = $vm.NumCPU
        $WPFtxtMemory.Text = $vm.MemoryGB
        $WPFcomboMemory.SelectedItem = $WPFcomboMemory.Items.GetItemAt(1)
        $wpftxtNotes.Text = $vm.Notes
        $WPFbtnReconfigureOk.Add_Click({ 
            $reconfigureForm.DialogResult = $true 
            $reconfigureForm.Close()  })

        $WPFbtnReconfigureCancel.Add_Click({
            $reconfigureForm.DialogResult = $false 
            $reconfigureForm.Close() })

        $result = $reconfigureForm.ShowDialog()
    }
    $result
}

#endregion XAML

#region Functions

Function Process-Stats
{
    Param
    (
        $stats ,
        [int]$lastSeconds ,
        [switch]$averageOnly ,
        [AllowNull()]
        [string]$exclude
    )
    
    [int]$minimum = [int]::MaxValue
    [int]$maximum = 0 
    [int]$average = 0 
    [long]$total = 0
    [long]$count = 0
    [string]$unit = $null

    If( ! $exclude -or $stats.Name -notmatch $exclude )
    {
        [datetime]$startTime = (Get-Date).AddSeconds( -$lastSeconds )
        ForEach( $stat in $stats.Group )
        {
           If( $stat.Timestamp -ge $startTime -and [string]::IsNullOrEmpty( $stat.Instance ) ) ## not getting resource specific info, e.g. each CPU
           {
                $count++
                If( $stat.Value -lt $minimum )
                {
                    $minimum = $stat.Value
                }
                If( $stat.Value -gt $maximum )
                {
                    $maximum = $stat.Value
                }
                If( ! $unit -and $stat.PSObject.Properties[ 'unit' ] )
                {
                    $unit = $stat.Unit
                }
                $total += $stat.Value
           }
        }
        [string]$label = (Get-Culture).TextInfo.ToTitleCase( ( $stats.Name -replace '\.([a-z])' , ' $1'))
        $result = [hashtable]@{ "$label$(If(! $averageOnly ){ ' Average'})" = [int]( $total / $count ) }
        If( ! $averageOnly )
        {
            $result += @{
                "$label Minimum" = $minimum
                "$label Maximum" = $maximum
            }
        }
        $result
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
        if( ! $result )
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
    $ChildItem.Name = $Name -replace '[\s,;:\.\-]' , '_' -replace '%252f' , '_' # default snapshot names have / for date which are escaped
    $ChildItem.Tag = $Tag
    $ChildItem.IsExpanded = $true
    ##[Void]$ChildItem.Items.Add("*")
    [Void]$Parent.Items.Add($ChildItem)
}

Function Process-Snapshot
{
    Param
    (
        $GUIobject ,
        $Operation ,
        $VMId
    )
    
    $_.Handled = $true
    
    [bool]$closeDialogue = $true

    if( ! $vmId )
    {
        [void][Windows.MessageBox]::Show( 'No VM id passed to Process-Snapshot' , $Operation , 'Ok' ,'Error' )
        return
    }

    if( $operation -eq 'ConsolidateSnapshot' )
    {
        $VM = Get-VM -Id $VMId
        $VM.ExtensionData.ConsolidateVMDisks_Task()
        $closeDialogue = $false
    }
    elseif( $Operation -eq 'TakeSnapShot' )
    {
        $takeSnapshotForm = Load-GUI -inputXaml $takeSnapshotXAML
        if( $takeSnapshotForm )
        {
            $takeSnapshotForm.Title += " of $($vm.name)"

            $WPFbtnTakeSnapshotOk.Add_Click({ 
                $takeSnapshotForm.DialogResult = $true 
                $takeSnapshotForm.Close()  })

            $WPFbtnTakeSnapshotCancel.Add_Click({
                $takeSnapshotForm.DialogResult = $false 
                $takeSnapshotForm.Close() })

            if( $takeSnapshotForm.ShowDialog() )
            {
                $VM = Get-VM -Id $VMId

                ## Get Data from form and take snapshot
                [hashtable]$parameters = @{ 'VM' = $vm ; 'RunAsync' = $true }
                if( ! [string]::IsNullOrEmpty( $WPFtxtSnapshotName.Text ) )
                {
                    $parameters.Add( 'Name' , $WPFtxtSnapshotName.Text )
                }
                if( ! [string]::IsNullOrEmpty( $WPFtxtSnapshotDescription.Text ) )
                {
                    $parameters.Add( 'Description' , $WPFtxtSnapshotDescription.Text )
                }
                $parameters.Add( 'Quiesce' , $wpfchkSnapshotQuiesce.IsChecked )
                $parameters.Add( 'Memory' , $WPFchckSnapshotMemory.IsChecked )
         
                New-Snapshot @parameters
            }
        }
    }
    elseif( $Operation -eq 'DetailsSnapshot' )
    {
        $closeDialogue = $false
        $VM = Get-VM -Id $VMId
        if( $VM )
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
            $snapshot = Get-Snapshot -Id $tag -VM $vm
            if( $snapshot )
            {
                [string]$details = "Name = $($snapshot.Name)`n`rDescription = $($snapshot.Description)`n`rCreated = $(Get-Date -Date $snapshot.Created -Format G)`n`rSize = $([math]::Round( $snapshot.SizeGB , 2 ))GB`n`rPower State = $($snapshot.PowerState)`n`rQuiesced = $(if( $snapshot.Quiesced ) { 'Yes' } else {'No' })"
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
        $VM = Get-VM -Id $VMId
        if( $VM )
        {
            if( $Operation -eq 'LatestSnapshotRevert' )
            {
                $snapshot = Get-Snapshot -VM $vm|Sort-Object -Property Created -Descending|Select-Object -First 1
                if( ! $snapshot )
                {
                    [Windows.MessageBox]::Show( "No snapshots found for $($vm.Name)" , 'Snapshot Revert Error' , 'OK' ,'Error' )
                    return
                }
            }
            else
            {
                $snapshot = Get-Snapshot -Id $GUIobject.SelectedItem.Tag -VM $vm
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

                $questionText += " `"$($snapshot.Name)`" on $($vm.Name), taken $(Get-Date -Date $snapshot.Created -Format G)?"
                $answer = [Windows.MessageBox]::Show( $questionText , "Confirm $($operation -creplace '([a-zA-Z])([A-Z])' , '$1 $2')" , 'YesNo' ,'Question' )
        
                if( $answer -eq 'yes' )
                {
                    if( $Operation -eq 'DeleteSnapShot' )
                    {
                        Remove-Snapshot -Snapshot $snapshot -RunAsync -Confirm:$false
                    }
                    elseif( $Operation -eq 'DeleteSnapShotTree' )
                    {
                        Remove-Snapshot -Snapshot $snapshot -Confirm:$false -RemoveChildren -RunAsync
                    }
                    elseif( $Operation -eq 'RevertSnapShot' -or $Operation -eq 'LatestSnapshotRevert' )
                    {
                        $answer = $null
                        if( $snapshot.PowerState -eq 'PoweredOff' )
                        {
                            $answer = [Windows.MessageBox]::Show( "Power on after snapshot restored on $($vm.Name)?" , 'Confirm Power Operation' , 'YesNo' ,'Question' )
                        }
                        [hashtable]$revertParameters = @{ 'VM' = $vm ; 'Snapshot' = $snapshot ;'Confirm' = $false ; RunAsync = ($answer -ne 'Yes') }
                        Set-VM @revertParameters
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

Function Show-DatastoresWindow
{
    $datastoresForm = Load-GUI -inputXaml $datastoresXAML

    if( $datastoresForm )
    {
        [array]$datastores = @( Get-Datastore | ForEach-Object `
        {
            $datastore = $_
            
            [pscustomobject][ordered]@{
                'Name' = $datastore.Name
                'FreeSpaceGB' = [math]::Round( $datastore.FreeSpaceGB , 1 )
                'CapacityGB' = [math]::Round( $datastore.CapacityGB , 1 )
                'UsedPercent' = [math]::Round( ($datastore.CapacityGB - $datastore.FreeSpaceGB ) / $datastore.CapacityGB * 100 , 1 )
            }
        })
        $datastoresForm.Title =  "$($datastores.Count) datastores"
        $WPFDatastoreList.ItemsSource = $datastores
        $WPFDatastoreList.IsReadOnly = $true
        $WPFDatastoreList.CanUserSortColumns = $true
        $datastoresForm.ShowDialog()
    }
}


Function Show-HostsWindow
{
    $hostsForm = Load-GUI -inputXaml $hostsXAML

    if( $hostsForm )
    {
        [hashtable]$VMsByHost = Get-VM | Where-Object PowerState -eq 'PoweredOn' | Group-Object -Property VMHost -AsHashTable -AsString
        Get-VMHost| ForEach-Object `
        {
            $thisHost = $_
           
            $WPFhostList.Items.Add( [pscustomobject][ordered]@{
                'Name' = $thisHost.Name
                'Parent' = $thisHost.Parent
                'PowerState' = $thisHost.PowerState
                'ConnectionState' = $thisHost.ConnectionState
                'Manufacturer' = $thisHost.Manufacturer
                'Model' = $thisHost.Model
                'Version' = $thisHost.Version
                'Build' = $thisHost.Build
                'MemoryTotalGB' = [math]::Round( $thisHost.MemoryTotalGB , 1 )
                'MemoryUsageGB' = [math]::Round( $thisHost.MemoryUsageGB , 1 )
                'NumCPU' = $thisHost.NumCPU
                'RunningVMs' = $(if( $VMsByHost[ $thisHost.name ] ) { $VMsByHost[ $thisHost.name ].Count } else { [int]0 })
                'MemoryUsedPercent' = [math]::Round( $thisHost.MemoryUsageGB / $thisHost.MemoryTotalGB * 100 , 1 )
                'CPUUsedPercent' = [math]::Round( $thisHost.CpuUsageMhz / $thisHost.CpuTotalMhz * 100 , 1 )
            } )
        }
        $hostsForm.Title =  "$($WPFhostList.Items.Count) hosts"
        $WPFhostList.IsReadOnly = $true
        $WPFhostList.CanUserSortColumns = $true
        $hostsForm.ShowDialog()
    }
}


Function Show-SnapShotWindow
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        $vm
    )

    [array]$theseSnapshots = @( VMware.VimAutomation.Core\Get-Snapshot -VM $vm -ErrorAction SilentlyContinue )
    if( ! $theseSnapshots -or ! $theseSnapshots.Count )
    {
        [void][Windows.MessageBox]::Show( "No snapshots found for $($vm.name)" , 'Snapshot Management' , 'Ok' ,'Warning' )
    }

    $snapshotsForm = Load-GUI -inputXaml $snapshotsXAML

    [bool]$result = $false
    if( $snapshotsForm )
    {
        $snapshotsForm.Title += " for $($vm.Name)"

        ForEach( $snapshot in $theseSnapshots )
        {
            ## if has a parent we need to find that and add to it
            if( $snapshot.ParentSnapshotId )
            {
                ## find where to add our node
                $parent = $theseSnapshots | Where-Object Id -eq $snapshot.ParentSnapshotId
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
                    Write-Warning "Unable to locate parent snapshot `"$($snapshot.Parent)`" for snapshot `"$($snapshot.Name)`" for $($vm.name)"
                }
            }
            else ## no parent then needs to be top level node but check not already created because we enountered a child previously
            {
                Add-TreeItem -Name $snapshot.Name -Parent $WPFtreeSnapshots -Tag $snapshot.Id
            }
        }
        
        [string]$currentSnapShotId = $null
        if( $vm.ExtensionData -and $vm.ExtensionData.SnapShot -and $vm.ExtensionData.SnapShot.CurrentSnapshot )
        {
            $currentSnapShotId = $vm.ExtensionData.SnapShot.CurrentSnapshot.ToString()
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
            Write-Warning "No current snapshot set for $($vm.Name)"
        }

        $WPFbtnSnapshotsOk.Add_Click({ 
            $snapshotsForm.DialogResult = $true 
            $snapshotsForm.Close()  })

        $WPFbtnSnapShotsCancel.Add_Click({
            $snapshotsForm.DialogResult = $false 
            $snapshotsForm.Close() })

        ## get last revert operation
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
                        
        $WPFbtnTakeSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'TakeSnapShot' -VMId $vm.Id } )
        $WPFbtnDeleteSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'DeleteSnapShot' -VMId $vm.Id} )
        $WPFbtnDeleteSnapShotTree.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'DeleteSnapShotTree' -VMId $vm.Id} )
        $WPFbtnRevertSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'RevertSnapShot' -VMId $vm.Id} )
        $WPFbtnDetailsSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'DetailsSnapShot' -VMId $vm.Id} )
        $WPFbtnConsolidateSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'ConsolidateSnapShot' -VMId $vm.Id } )

        $result = $snapshotsForm.ShowDialog()
    }
}

## See if we have HKCU or HKLM policy setting for this
Function Get-PolicySettings
{
    [CmdletBinding()]

    Param
    (
        [string]$regKey ,
        [string]$setting
    )

    [decimal]$result = -1 ## signifies not set

    [string[]]$policyKeys = @( ($regKey -replace '^HK.*:\\Software\\' , 'HKCU:\Software\Policies\') , ($regKey -replace '^HK.*:\\Software\\' , 'HKLM:\Software\Policies\') )
    ForEach( $policyKey in $policyKeys )
    {
        if( $result -lt 0 -and ( Test-Path -Path $policyKey -PathType Container -ErrorAction SilentlyContinue ) )
        {
            $regValue = Get-ItemProperty -Path $policyKey -Name $setting -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $setting -ErrorAction SilentlyContinue
            if( $regValue )
            {
                Write-Verbose "Policy setting for $setting found in $regKey, = $regValue"
                $result = $regValue
            }
        }
    }

    $result ## return
}

Function Process-Action
{
    Param
    (
        $GUIobject , 
        [string]$Operation ,
        $context
    )

    $_.Handled = $true
    ## get selected items from control
    [array]$selectedVMs = @( $GUIobject.selectedItems )
    if( ! $selectedVMs -or ! $selectedVMs.Count )
    {
        return
    }
    ForEach( $selectedVM in $selectedVMs )
    {
        ## Don't use VM cache here in case changed since cache made
        $getError = $null
        $vm = VMware.VimAutomation.Core\Get-VM -Name $selectedVM.Name -ErrorVariable getError | Where-Object { $_.VMHost.Name -eq $selectedVM.Host }
        if( ! $vm )
        {
            [string]$message = "Unable to get VM $($selectedVM.Name)"
            if( $getError -and $getError.Count -and $getError[0].Exception.Message -match 'Not Connected' )
            {
                $string += '. Try clicking "Refresh" as no longer connected'
            }
            [void][Windows.MessageBox]::Show( $message , 'Action Error' , 'Ok' ,'Error' )
            return
        }
        elseif( $vm -is [array] )
        {
            Write-Warning -Message "Found $($vm.Count) VMware objects matching $($selectedVM.Name)"
        }
        else
        {
            [string]$address = $null
            if( $operation -eq 'backup' )
            {
                ## http://www.simonlong.co.uk/blog/2010/05/05/powercli-a-simple-vm-backup-script/
                
                $backupForm = Load-GUI -inputXaml $backupXAML

                if( $backupForm )
                {
                    $backupForm.Title += " of $($vm.Name)"
                    $wpftxtBackupName.Text = "Backup $($vm.Name) $((Get-Date -Format d) -replace '/')"

                    $wpfcomboDatastore.Items.Clear()
                    Get-Datastore | Sort-Object -Property Name | ForEach-Object `
                    {
                        $wpfcomboDatastore.Items.Add( $_.Name )
                    }
                    
                    $wpfcomboFolder.Items.Clear()
                    Get-Folder | Sort-Object -Property Name | ForEach-Object `
                    {
                        $wpfcomboFolder.Items.Add( $_.Name )
                    }
                    
                    $WPFbtnBackupOK.Add_Click({
                        if( ! $wpfcomboDatastore.SelectedItem )
                        {
                            [void][Windows.MessageBox]::Show( "Must select the destination datastore for the backup" , 'Backup Error' , 'Ok' ,'Error' )
                        }
                        elseif( ! $wpfcomboFolder.SelectedItem )
                        {
                            [void][Windows.MessageBox]::Show( "Must select the destination folder for the backup" , 'Backup Error' , 'Ok' ,'Error' )
                        }
                        elseif( [string]::IsNullOrEmpty( $wpftxtBackupName.Text ) )
                        {
                            [void][Windows.MessageBox]::Show( "Must enter a name for the backup VM" , 'Backup Error' , 'Ok' ,'Error' )
                        }
                        else
                        {
                            $backupForm.DialogResult = $true 
                            $backupForm.Close()  }
                        })

                    $result = $backupForm.ShowDialog()

                    if( $result )
                    {
                        if( $cloneSnapshot = New-Snapshot -VM $vm -Name ( [System.Environment]::ExpandEnvironmentVariables( $backupSnapshotName ) ) )
                        {
                            $vmView = $vm | Get-View
                            $cloneSpec = New-Object -Typename Vmware.Vim.VirtualMachineCloneSpec
                            $cloneSpec.Snapshot = $vmView.Snapshot.CurrentSnapshot
 
                            # Make linked disk specification
                            $cloneSpec.Location = New-Object -Typename Vmware.Vim.VirtualMachineRelocateSpec
                            $cloneSpec.Location.Datastore = (Get-Datastore -Name $wpfcomboDatastore.SelectedItem | Get-View).MoRef
                            $cloneSpec.Location.Transform =  [Vmware.Vim.VirtualMachineRelocateTransformation]::sparse
 
                            $cloneName = $wpftxtBackupName.Text
 
                            $cloneFolder = (Get-Folder -Name $wpfcomboFolder.SelectedItem | Get-View).Moref

                            if( $backupTask = $vmView.CloneVM_Task( $cloneFolder , $cloneName , $cloneSpec ) )
                            {
                                if( ! $WPFchkAsync.IsChecked )
                                {
                                    $oldCursor = $backupForm.Cursor
                                    $backupForm.Cursor = [Windows.Input.Cursors]::Wait
                                    $result = Wait-Task -Task (Get-Task -Id $backupTask)
                                    $backupForm.Cursor = $oldCursor

                                    if( ! $result )
                                    {
                                        [void][Windows.MessageBox]::Show( "Failed waiting for completion of backup of $($vm.Name)" , 'Backup Error' , 'OK' ,'Error' )
                                    }
                                    else
                                    {
                                        Remove-Snapshot -Snapshot $cloneSnapshot -Confirm:$false
                                    }
                                }
                                ## else async so cannot remove snapshot as in use
                            }
                            else
                            {
                                [void][Windows.MessageBox]::Show( "Failed to create backup of $($vm.Name)" , 'Backup Error' , 'OK' ,'Error' )
                                Remove-Snapshot -Snapshot $cloneSnapshot -Confirm:$false
                            }
                        }
                        else
                        {
                            [void][Windows.MessageBox]::Show( "Failed to create snapshot of $($vm.Name) for backup" , 'Backup Error' , 'OK' ,'Error' )
                        }
                    }
                }
            }
            elseif( $Operation -match '^mstsc' )
            {
                if( $vm.PowerState -ne 'PoweredOn' )
                {
                    [void][Windows.MessageBox]::Show( "Power state of $($vm.Name) is $($vm.PowerState) so cannot mstsc to it" , 'Mstsc Error' , 'Ok' ,'Error' )
                    return
                }
                ## See if we can resolve its name otherwise we will try its IP addresses
                [string]$address = $null

                if( $vm.PSObject.Properties[ 'Guest' ] -and $vm.guest )
                {
                    $address = $vm.guest.HostName
                }

                if( ! $address -or ! ( Resolve-DnsName -Name $address -ErrorAction SilentlyContinue -QuickTimeout ))
                {
                    Write-Warning "Unable to resolve $address for $($vm.Name) so looking for an IP address to use"

                    $address = $null
                    ## Sort addresses so get 192. addresses before 172. or 10. ones which are more likely to be reachable
                    $vm.Guest | Select-Object -ExpandProperty IPAddress | Where-Object { $_ -match $script:addressPattern -and $_ -ne '127.0.0.1' -and $_ -ne '::1' -and $_ -notmatch '^169\.254\.' } | Sort-Object -Descending |Sort-Object -Descending -Property @{e={($_ -split '\.')[0] -as [int]}} | ForEach-Object `
                    {
                        if( ! $address )
                        {
                            Write-Verbose "Testing connectivity to port $rdpPort on $PSItem"
                            $connection = Test-NetConnection -ComputerName $PSItem -Port $rdpPort -ErrorAction SilentlyContinue -InformationLevel Quiet
                            if( $connection )
                            {
                                $address = $PSItem
                            }
                        }
                    }
                }
                elseif( ! $vm.PSObject.Properties[ 'Guest' ] -or ! $vm.Guest )
                {
                    [void][Windows.MessageBox]::Show( "Cannot resolve $address for $($vm.Name) and no guest info" , 'Mstsc Error' , 'Ok' ,'Error' )
                    return
                }
        
                if( $address )
                {
                    [bool]$setForegroundWindow = $false
                    if( ! $noReuseMstsc -and $Operation -ne 'MstscNew' )
                    {
                        ## see if we have a running mstsc for this host already so we can restore and bring to foreground
                        $mstscProcess = Get-CimInstance -ClassName win32_Process -Filter "Name = 'mstsc.exe' and ParentProcessId = '$pid'" | Where-Object CommandLine -match "/v:$address\b" | Sort-Object -Property CreationDate -Descending | Select-Object -First 1
                        if( $mstscProcess )
                        {
                            Write-Verbose "Found existing mstsc process pid $($mstscProcess.ProcessId) for $address"
                            $windowHandle = Get-Process -Id $mstscProcess.ProcessId | Select-Object -ExpandProperty MainWindowHandle
                            if( $windowHandle )
                            {
                                [bool]$setForegroundWindow = [win32.user32]::ShowWindowAsync( $windowHandle , 9 ) ## restore
                                if( ! $setForegroundWindow )
                                {
                                    Write-Warning "Failed to set mstsc.exe process id $($mstsc.ProcessId) window to foreground"
                                }
                            }
                        }
                    }

                    if( ! $setForegroundWindow )
                    {
                        ## See if we have a stored credential for this and if so only create .rdp file if extraRDPSettings set
                        [string]$storedCredential = cmdkey.exe /list:TERMSRV/$address | Where-Object { $_ -match 'User:' }
                        [string]$thisRdpFileName = $rdpFileName
                        if( $storedCredential )
                        {
                            Write-Verbose "Got stored credential `"$storedCredential`" for $address"
                            $thisRdpFileName = $rdpExtrasFileName
                        }
                        Write-Verbose "Launching mstsc to $address"
                        [string]$arguments = "$thisRdpFileName /v:$($address):$rdpPort"
                        if( ! [string]::IsNullOrEmpty( $mstscParams ))
                        {
                            $arguments += " $mstscParams"
                        }
                        $mstscProcess = Start-Process -FilePath 'mstsc.exe' -ArgumentList $arguments -PassThru
                        if( ! $mstscProcess )
                        {
                            Write-Error -Message "Failed to run mstsc.exe"
                        }
                    }
                }
                else
                {
                    [void][Windows.MessageBox]::Show( "No address for $($vm.Name)" , 'Connection Error' , 'Ok' ,'Warning' )
                }
            }
            elseif( $Operation -eq 'console' )
            {
                $vmrcProcess = $null
                if( $session -and ! $consoleBrowser )
                {
                    ## see if we already have a running console process for this VM
                    $vmrcProcess = Get-CimInstance -ClassName win32_Process -Filter "Name = 'vmrc.exe' and ParentProcessId = '$pid'" | Where-Object CommandLine -match "/\?moid=$($vm.ExtensionData.MoRef.value)[^\d]" | Sort-Object -Property CreationDate -Descending | Select-Object -First 1
                    if( $vmrcProcess )
                    {
                        Write-Verbose "Found existing vmrc process pid $($vmrcProcess.ProcessId) for $address"
                        $windowHandle = Get-Process -Id $vmrcProcess.ProcessId | Select-Object -ExpandProperty MainWindowHandle
                        if( $windowHandle )
                        {
                            [bool]$setForegroundWindow = [win32.user32]::ShowWindowAsync( $windowHandle , 9 ) ## restore
                            if( ! $setForegroundWindow )
                            {
                                Write-Warning "Failed to set vmrc.exe process id $($vmrc.ProcessId) window to foreground"
                                $vmrcProcess = $null
                            }
                        }
                    }
                    
                    if( ! $vmrcProcess )
                    {
                        try
                        {
                            $ticket = $Session.AcquireCloneTicket()
                        }
                        catch
                        {
                            $ticket = $null
                        }
                        if( ! $ticket )
                        {
                            if( $Error.Count -and $Error[0].Exception.Message -match 'The session is not authenticated' )
                            {
                                Write-Verbose "Gettting new session for ticket"
                                $Session = Get-View -Id Sessionmanager -ErrorAction SilentlyContinue
                                if( $session )
                                {
                                    $ticket = $Session.AcquireCloneTicket()
                                }
                                else
                                {
                                    $connection = Connect-VIServer @connectParameters
                                    if( ! $connection )
                                    {
                                        Write-Warning "Failed to connect to VI server"
                                    }
                                    $Session = Get-View -Id Sessionmanager -ErrorAction SilentlyContinue
                                    if( $session )
                                    {
                                        $ticket = $Session.AcquireCloneTicket()
                                    }
                                    else
                                    {
                                        Write-Warning "Failed to get session"
                                    }
                                }
                            }
                            else
                            {
                                Write-Warning -Message "Failed to acquire ticket: $Error"
                            }
                        }
                        If( $ticket -and $connection -and $vm )
                        {
                            $vmrcProcess = Start-Process -FilePath 'vmrc.exe' -ArgumentList "vmrc://clone:$ticket@$($connection.ToString())/?moid=$($vm.ExtensionData.MoRef.value)" -PassThru -WindowStyle Normal -Verb Open -ErrorAction SilentlyContinue
                        }
                        Else
                        {
                            Write-Warning "Unable to launch vmrc.exe - ticket $ticket, connection $connection, vm $vm"
                        }
                    }
                }
                ## fallback to PowerCLI console but it doesn't persist connection across power operations
                if( ! $vmrcProcess )
                {
                    if( ! [string]::IsNullOrEmpty( $consoleBrowser ) )
                    {
                        [string]$URL = Open-VMConsoleWindow -VM $vm -UrlOnly
                        if( ! [string]::IsNullOrEmpty( $URL ) )
                        {
                            & $consoleBrowser `"$URL`"
                        }
                        else
                        {
                            Write-Warning "Failed to get vmrc console URL for $($vm.name)"
                        }
                    }
                    else
                    {
                        Open-VMConsoleWindow -VM $vm
                    }
                }
            }
            elseif( ( $powerCmdlet = $powerOperation[ $Operation ] ) )
            {
                [string]$answer = 'yes'
                if( $Operation -ne 'PowerOn' )
                {
                    $answer = [Windows.MessageBox]::Show( "Are you sure you want to $($operation -creplace '([a-zA-Z])([A-Z])' , '$1 $2') $($vm.Name)?" , 'Confirm Power Operation' , 'YesNo' ,'Question' )
                }
                if( $answer -eq 'yes' )
                {
                    $command = Get-Command -Module VMware.VimAutomation.Core -Name $powerCmdlet
                    if( $command )
                    {
                        [hashtable]$parameters = @{ 'Confirm' = $false ; 'VM' = $vm }
                        if( $command.Parameters[ 'RunAsync' ] )
                        {
                            $parameters.Add( 'RunAsync' , $true ) ## so that command comes back immediately
                        }
                        & $command @parameters
                    }
                    else
                    {
                        Write-Error "Unable to find $powerCmdlet cmdlet"
                    }
                }
            }
            elseif( $Operation -eq 'Reconfigure' )
            {
                if( Set-Configuration -vm $vm )
                {
                    [double]$newvCPUS = $WPFtxtvCPUs.Text
                    [double]$newMemory = $WPFtxtMemory.Text
                    [string]$memoryUnits = $WPFcomboMemory.SelectedItem.Content
                    if( $memoryUnits -eq 'MB' )
                    {
                        $newMemory = $newMemory / 1024
                    }
                    if( $newvCPUS -ne $vm.NumCPU )
                    {
                        [int]$maxCPUS = Get-PolicySettings -regKey $regKey -setting maxCPUS
                        if( $maxCPUS -lt 0 -or $newvCPUS -le $maxCPUS )
                        {
                            VMware.VimAutomation.Core\Set-VM -VM $vm -NumCpu $newvCPUS -Confirm:$false
                            if( ! $? )
                            {
                                [void][Windows.MessageBox]::Show( "Failed to change vCPUs from $($vm.NumCPU) to $newvCPUS in $($vm.Name)" , 'Reconfiguration Error' , 'Ok' ,'Exclamation' )
                            }
                        }
                        else
                        {
                            [void][Windows.MessageBox]::Show( "The maximum CPUs allowed is $maxCPUS" , 'Reconfiguration Error' , 'Ok' ,'Exclamation' )
                        }
                    }
                    if( $newMemory -ne $vm.MemoryGB )
                    {
                        [decimal]$maxMemory = Get-PolicySettings -regKey $regKey -setting maxMemory
                        if( $maxMemory -lt 0 -or $newMemory -le $maxMemory )
                        {
                            VMware.VimAutomation.Core\Set-VM -VM $vm -MemoryGB $newMemory -Confirm:$false
                            if( ! $? )
                            {
                                [void][Windows.MessageBox]::Show( "Failed to change memory from $($vm.MemoryGB)GB to $($newMemory)GB in $($vm.Name)" , 'Reconfiguration Error' , 'Ok' ,'Exclamation' )
                            }
                        }
                        else
                        {
                            [void][Windows.MessageBox]::Show( "The maximum memory allowed is $($maxMemory)GB" , 'Reconfiguration Error' , 'Ok' ,'Exclamation' )
                        }
                    }
                    if( $WPFtxtNotes.Text -ne $vm.Notes )
                    {
                        VMware.VimAutomation.Core\Set-VM -VM $vm -Notes $WPFtxtNotes.Text -Confirm:$false
                    }
                }
            }
            elseif( $Operation -eq 'Screenshot' )
            {
                $view = $vm | Get-View
                $shot = $view.CreateScreenshot()
                [datetime]$snapshotTime = [datetime]::Now
                if( $shot )
                {
                    $datacenter = Get-DataCenter -VM $vm
                    [string]$sourceFile = "vmstore:/$($datacenter.name)/$($shot -replace '\[' -replace '\]\s*' , '/')"
                    if( ! [string]::IsNullOrEmpty( $screenShotFolder ) -and ! ( Test-Path -Path $screenShotFolder -ErrorAction SilentlyContinue -PathType Container ) )
                    {
                        $null = New-Item -Path $screenShotFolder -ItemType Directory
                    }
                    [string]$localfile = Join-Path -Path $( if( ! [string]::IsNullOrEmpty( $screenShotFolder ) ) { $screenShotFolder } else { $env:temp }) -ChildPath "$(Get-Date -Format 'HHmmss-ddMMyy')-$(Split-Path -Path $shot -Leaf)"
                    if( ( $copied = Copy-DatastoreItem -Item $sourceFile -Destination $localFile -PassThru ) )
                    {
                        if( $screenshotWindow = Load-GUI -inputXaml $screenshotXAML )
                        {
                            if( $filestream = New-Object System.IO.FileStream -ArgumentList $copied.FullName , Open , Read )
                            {
                                $bitmap = New-Object -Typename System.Windows.Media.Imaging.BitmapImage
                                $bitmap.BeginInit()
                                $bitmap.StreamSource = $filestream
                                $bitmap.EndInit()
                                $wpfimgScreenshot.Source = $bitmap
                                $screenshotWindow.Title = "Screenshot of $($vm.Name) at $(Get-Date -Date $snapshotTime)"
                                $screenshotWindow.ShowDialog()
                                $bitmap.StreamSource = $null
                                $filestream.Close()
                                $filestream = $null
                                $bitmap = $null
                            }
                            else
                            {
                                [void][Windows.MessageBox]::Show( "Failed to open screenshot file $($copied.FullName)" , 'Snapshot Error' , 'Ok' ,'Exclamation' )
                            }
                        }
                        if( [string]::IsNullOrEmpty( $screenShotFolder ) )
                        {
                            Remove-Item -Path $copied.FullName
                        }
                    }
                    else
                    {
                        [void][Windows.MessageBox]::Show( "Failed to copy console snapshot of $($vm.Name) from $sourceFile" , 'Snapshot Error' , 'Ok' ,'Exclamation' )
                    }
                }
                else
                {
                    [void][Windows.MessageBox]::Show( "Failed to create console snapshot of $($vm.Name)" , 'Snapshot Error' , 'Ok' ,'Exclamation' )
                }
            }
            elseif( $Operation -eq 'Snapshots' )
            {
                Show-SnapShotWindow -vm $vm
            }
            elseif( $Operation -eq 'LatestSnapshotRevert' )
            {
                Process-Snapshot -Operation 'LatestSnapshotRevert' -VMId $vm.Id
            }            
            elseif( $Operation -eq 'Delete' )
            {
                [string]$answer = [Windows.MessageBox]::Show( "Are you sure you want to delete $($vm.Name)?" , 'Confirm Delete Operation' , 'YesNo' ,'Question' )
                if( $answer -eq 'yes' )
                {
                    VMware.VimAutomation.Core\Remove-VM -DeletePermanently -VM $vm -Confirm:$false
                }
            }
            elseif( $Operation -eq 'Events' )
            {
                [array]$events = @( Get-VIEvent -Entity $vm.name -MaxSamples $maxEvents )
                if( $events -and $events.Count )
                {
                    $events | Select-Object -Property CreatedTime,UserName,FullFormattedMessage | Out-GridView -Title "$($events.Count) events for $($vm.Name)"
                }
                else
                {
                    $oldestEvent = Get-VIEvent -MaxSamples $maxEvents | Select-Object -Last 1
                    [string]$message = "No events found for $($vm.Name)"
                    if( $oldestEvent )
                    {
                        $message += " oldest of event is from $(Get-Date -Date $oldestEvent.CreatedTime -Format G). Retrieving $maxEvents at most"
                    }
                    [void][Windows.MessageBox]::Show( $message , 'Events Error' , 'Ok' ,'Warning' )
                }
            }
            else
            {
                Write-Warning "Unknown operation $Operation"
            }
        }
    }
}

Function Update-Form
{
    Param
    (
        $form ,
        $datatable ,
        $vmname
    )

    ## update form
    $oldCursor = $form.Cursor
    $form.Cursor = [Windows.Input.Cursors]::Wait
    $datatable.Rows.Clear()
    $global:vms = Get-VMs -datatable $datatable -pattern $vmName -poweredOn $showPoweredOn -poweredOff $showPoweredOff -suspended $showSuspended -datastores $script:datastores
    
    $mainForm.Title = $mainForm.Title -replace 'connected to.*' , "connected to $($server -join ' , ') at $(Get-Date -Format G), $($global:vms.Count) VMs"
    $WPFVirtualMachines.Items.Refresh()
    $form.Cursor = $oldCursor
}

Function Get-VMs
{
    [CmdletBinding()]

    Param
    (
        $datatable ,
        [string]$pattern ,
        [bool]$poweredOn = $true ,
        [bool]$poweredOff = $true,
        [bool]$suspended = $true ,
        [hashtable]$datastores
    )
    
    [hashtable]$params = @{}
    if( ! [string]::IsNullOrEmpty( $pattern ) )
    {
        $params.Add( 'Name' , $pattern )
    }
     
    [array]$vms = @( VMware.VimAutomation.Core\Get-VM @params | Sort-Object -Property $script:sortProperty | . { Process
    {
        $vm = $PSItem
        [bool]$include = $false
        if( $poweredOn -and $vm.PowerState -eq 'PoweredOn' )
        {
            $include = $true
        }
        if( $poweredOff -and $vm.PowerState -eq 'PoweredOff' )
        {
            $include = $true
        }
        if( $suspended -and $vm.PowerState -eq 'Suspended' )
        {
            $include = $true
        }
        if( $include )
        {
            [string[]]$IPaddresses = @( $vm.guest | Select-Object -ExpandProperty IPAddress | Where-Object { $_ -match $addressPattern } | Sort-Object )
            [hashtable]$additionalProperties = @{
                'Host' = $vm.VMHost.Name
                'vCPUs' = $vm.NumCpu
                'Power State' = $powerState[ $vm.PowerState.ToString() ]
                'IP Addresses' = $IPaddresses -join ','
                'Snapshots' = ( $snapshots | Where-Object VMId -eq $VM.Id | Measure-Object|Select-Object -ExpandProperty Count)
                'Started' = $vm.ExtensionData.Runtime.BootTime
                'Used Space (GB)' = [Math]::Round( $vm.UsedSpaceGB , 1 )
                'Datastore(s)' = ($vm.DatastoreIdList | ForEach-Object { $datastores[ $PSItem ] }) -join ','
                'Memory (GB)' = $vm.MemoryGB
                'Guest OS' = $(If( $vm.Guest -and $vm.Guest.OSFullName ) { $vm.Guest.OSFullName } Else { $vm.GuestId })
                'HW Version' = $vm.HardwareVersion -replace 'vmx-(\d*)$', '$1'
                'VMware Tools' = $vm.Guest.ToolsVersion
            }

            if( $lastSeconds -gt 0 -and $WPFchkPerfData.IsChecked -and $vm.PowerState -eq 'PoweredOn' )
            {
                ## Doing get-stat for all VMs in one call is no quicker and because can throw exception would potentially mean missing some machines
                Get-Stat -Entity $vm -Common -Realtime -ErrorAction SilentlyContinue | Group-Object -Property MetricId | ForEach-Object `
                {
                    if( $stats = (Process-Stats -stats $_ -lastSeconds $lastSeconds -exclude $exclude -averageOnly) )
                    {
                        $additionalProperties += $stats
                    }
                }
            }
  
            Add-Member -InputObject $vm -NotePropertyMembers $additionalProperties
            $vm
            $items = New-Object -TypeName System.Collections.ArrayList
            ForEach( $field in $displayedFields )
            {
                if( $vm.PSObject.Properties[ $field ] )
                {
                    [void]$items.Add( $vm.$field )
                }
            }
            [void]$datatable.Rows.Add( [array]$items )
        }
    }})
    Write-Verbose "Got $(if( $vms ) { $vms.Count } else { 0 }) vms"
    $vms
}
#endregion Functions

if( $showAll )
{
    $showPoweredOn = $showPoweredOff = $showSuspended = $true
}

$importError = $null
$importedModule = Import-Module -name $vmwareModule -Verbose:$false -PassThru -ErrorAction SilentlyContinue -ErrorVariable importError

if( ! $importedModule )
{
    Throw "Failed to import VMware module $vmwareModule. It can be installed via Install-Module if required. Error was `"$($importError|Select-Object -ExpandProperty Exception|Select-Object -ExpandProperty Message)`""
}

Get-Variable -Name WPF*|Select-Object -ExpandProperty Name | Write-Debug

[bool]$alreadyConnected = $false

Add-Type -MemberDefinition $pinvokeCode -Name 'User32' -Namespace 'Win32' -UsingNamespace System.Text -Debug:$false

## see if already connected and if so if it is the server we are told to connect to
if( Get-Variable -Scope Global -Name DefaultVIServers -ErrorAction SilentlyContinue )
{
    $existingServers = @( $global:DefaultVIServers )
    if( $existingServers -and $existingServers.Count )
    {
        [string]$existingServer = $existingServers|Select-Object -ExpandProperty Name
        Write-Verbose -Message "Already connected to $($existingServer -join ' , ')"
        $server = $existingServer
        $alreadyConnected = $true
        if( ! $PSBoundParameters[ 'username' ] )
        {
            $username = $existingServers|Group-Object -Property User|Sort-Object -Property Count -Descending|Select-Object -First 1 -ExpandProperty Name
        }
        Write-Verbose -Message "Already connected to $($existingServer -join ' , ') username $username"
        ## Check not connected to same server
        [hashtable]$connections = @{}
        ForEach( $serverConnection in $existingServer )
        {
            try
            {
                $dns = Resolve-DnsName -Name $serverConnection -ErrorAction SilentlyContinue -Verbose:$false
                if( $dns )
                {
                    $connections.Add( $dns.IPAddress , $serverConnection  )
                }
            }
            catch
            {
                $sameServer = $connections[ $dns.IPAddress ]
                Throw "Multiple connections to same VMware server $serverConnection and $sameServer with IP address $($dns.IPAddress)"
            }
        }
        Remove-Variable -Name connections
    }
}

## if we have no server then see if saved in registry
if( [string]::IsNullOrEmpty( $server ) )
{
    $server = (Get-ItemProperty -Path $regKey -Name 'Server' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'Server') -split ','
    Write-Verbose -Message "Retrieved server $server from registry"
}

if( ! $PSBoundParameters[ 'mstscParams' ] )
{
    $mstscParams = Get-ItemProperty -Path $regKey -Name 'mstscParams' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'mstscParams'
}

if( ! $alreadyConnected -and ( ! $server -or ! $server.Count ) )
{
    ## See if we already have a connection
    if( $global:DefaultVIServers -and $global:DefaultVIServers.Count )
    {
        $server = $global:DefaultVIServers | Select-Object -ExpandProperty Name
    }
    if( ! $server -or ! $server.Count )
    {
        Throw 'Must specify the ESXi or vCenter to connect to via -server'
    }
}

[hashtable]$connectParameters = @{ 'Server' = $server ; 'Force' = $true }

if( ! $alreadyConnected )
{
    ## Connect and retrieve VMs, applying filter if there is one
    [string]$configFolder = Join-Path -Path ( [System.Environment]::GetFolderPath( [System.Environment+SpecialFolder]::LocalApplicationData )) -ChildPath 'Guy Leech'
    if( ! ( Test-Path -Path $configFolder -PathType Container -ErrorAction SilentlyContinue ) )
    {
        $created = New-Item -Path $configFolder -ItemType Directory
        if( ! $created )
        {
            Write-Warning -Message "Failed to create config folder `"$configFolder`""
        }
    }

    if( ! $passThru -and ! $credential )
    {
        [string]$joinedServers = ($server|Sort-Object) -join ','
        [string]$credentialFilter = $(if( ! [string]::IsNullOrEmpty( $username ) ) { "$($username -replace '\\' , ';' ).$joinedServers.crud"  } else { "*.$joinedServers.crud"  })

        Write-Verbose "Looking for credentials $credentialFilter in `"$configFolder`""

        [array]$savedCredentials = @( Get-ChildItem -Path $configFolder -ErrorAction SilentlyContinue -File -Filter $credentialFilter | ForEach-Object `
        {
            $file = $PSItem
            New-Object System.Management.Automation.PSCredential( (($file.name -replace ([regex]::Escape( ".$joinedServers.crud" )) , '') -replace ';' , '\') , (Get-Content -Path $file.FullName|ConvertTo-SecureString) )
        })

        if( $savedCredentials -and $savedCredentials.Count )
        {
            if( $savedCredentials.Count -eq 1 )
            {
                $credential = $savedCredentials[0]
                Write-Verbose "Got single saved credentials for $($credential.username)"
            }
            else
            {
                Throw "Found $($savedCredentials.Count) saved credentials in `"$configFolder`" for $joinedServers, use -username to pick the required one"
            }
        }

        if( ! $credential -and ! $passThru )
        {
            ## see if we have stored credential for the servers
            [int]$storedCredentials = 0
            ForEach( $thisServer in $server )
            {
                if( Get-VICredentialStoreItem -Host $thisServer -ErrorAction SilentlyContinue )
                {
                    $storedCredentials++
                }
            }

            Write-Verbose "Got $storedCredentials stored credentials"

            if( $storedCredentials -ne $server.Count )
            {
                [hashtable]$credentialPromptParameters = @{ 'Message' = "Enter credentials to connect to $server" }
                if( $PSBoundParameters[ 'username' ] )
                {
                    $credentialPromptParameters.Add( 'Username' , $username )
                }
                $credential = Get-Credential @credentialPromptParameters
            }
        }
    }

    if( $credential )
    {
        Write-Verbose "Connecting to $($server -join ',') as $($credential.username)"
        $connectParameters.Add( 'Credential' , $credential )
    }
    
    [hashtable]$powerCLISettings = @{
        'ParticipateInCeip' = $false 
        'InvalidCertificateAction' = 'Ignore'
        'DisplayDeprecationWarnings' = $false
        'Confirm' = $false
        'Scope' = 'Session'
    }

    if( $PSBoundParameters[ 'consoleBrowser' ] )
    {
        $powerCLISettings.Add( 'VMConsoleWindowBrowser' , $consoleBrowser )
    }

    [void](Set-PowerCLIConfiguration @powerCLISettings )

    $connection = Connect-VIServer @connectParameters

    if( ! $? -or ! $connection )
    {
        ##[void][Windows.MessageBox]::Show( "Failed to connect to $server" , 'Connection Error' , 'Ok' ,'Exclamation' )
        Exit 1
    }

    if( $saveCredentials -and $credential )
    {
        $credential.Password | ConvertFrom-SecureString | Set-Content -Path (Join-Path -Path $configFolder -ChildPath ( ( "{0}.{1}.crud" -f ( $credential.username -replace '\\' , ';') , ( $server -join ',') ) ))
    }
}

[string[]]$rdpCredential = $null

## Write to an rdp file so we can pass to mstsc in case on non-domain joined device or if using different credentials
[string]$theUser = $null
if( ($theUser = ($credential|Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue)) -or ($theUser = $username ))
{
    Write-Verbose "Connecting as $theUser"
    [string]$rdpusername = $theUser
    [string]$rdpdomain = $null
    if( $theUser.IndexOf( '@' ) -lt 0 )
    {
        $rdpdomain,$rdpUsername = $theUser -split '\\'
    }
    $rdpCredential = @( "username:s:$rdpusername" , "domain:s:$rdpdomain" )
}
elseif( $passThru )
{
    $rdpCredential = @( "username:s:$env:USERNAME" , "domain:s:$env:USERDOMAIN" )
}

[string]$rdpFileName = $null
[string]$rdpExtrasFileName = $null ## use this when find we have a stored credential so don't need username in .rdp but still need to pass extra settings

if( ! $noRdpFile )
{
    if( $rdpCredential -and $rdpCredential.Count )
    {
        ## Write username and domain to rdp file to pass to mstsc
        $rdpFileName = Join-Path -Path $env:temp -ChildPath "grl.$pid.rdp"
        Write-Verbose "Writing $($rdpUsername -join ' , ') to $rdpFileName"
        $rdpCredential + $extraRDPSettings | Out-File -FilePath $rdpFileName
    }
    if( $extraRDPSettings -and $extraRDPSettings.Count )
    {
        $rdpExtrasFileName = Join-Path -Path $env:temp -ChildPath "grl.extras.$pid.rdp"
        Write-Verbose "Writing $($extraRDPSettings -join ' , ') only to $rdpExtrasFileName"
        $extraRDPSettings | Out-File -FilePath $rdpExtrasFileName
    }
}

if( ! ( Test-Path -Path $regKey -ErrorAction SilentlyContinue ) )
{
    [void](New-Item -Path $regKey -Force)
}

Set-ItemProperty -Path $regKey -Name 'Server' -Value ($server -join ',')
Set-ItemProperty -Path $regKey -Name 'MstscParams' -Value $mstscParams

[string]$addressPattern = $null

if( ! $ipV6 )
{
    $addressPattern = '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$'
}

[void][Reflection.Assembly]::LoadWithPartialName( 'Presentationframework' )

$mainForm = Load-GUI -inputXaml $mainwindowXAML

if( ! $mainForm )
{
    return
}

[array]$snapshots = @( VMware.VimAutomation.Core\Get-Snapshot -VM $vmName -ErrorVariable getError )

if( $getError -and $getError.Count -and $getError[0].Exception.Message -match 'Not Connected' )
{
    Write-Warning "Not connected to $($server -join ',') so reconnecting"
    $connectParameters|Write-Verbose
    $connection = Connect-VIServer @connectParameters
}

[hashtable]$script:datastores = @{}
Get-Datastore | ForEach-Object `
{
    $script:datastores.Add( $PSItem.Id , $PSItem.Name )
}

## so we can acquire a ticket if required for vmrc remote console (will fail on ESXi)
$Session = Get-View -Id Sessionmanager -ErrorAction SilentlyContinue

$datatable = New-Object -TypeName System.Data.DataTable

##[string[]]$displayedFields =  @( "Name" , "Power State" , "Host" , "Notes" , "Started" , "vCPUs" , "Memory (GB)" , "Snapshots" , "IP Addresses" , "VMware Tools" , "HW Version" , "Guest OS" , "Datastore(s)" , "Folder" , "Used Space (GB)" )
$displayedFields = New-Object -TypeName System.Collections.Generic.List[String] 
$displayedFields += @( "Name" , "Power State" , "Host" , "Notes" , "Started" , "vCPUs" , "Memory (GB)" , "Snapshots" , "IP Addresses" , "VMware Tools" , "HW Version" , "Guest OS" , "Datastore(s)" , "Folder" , "Used Space (GB)" )
If( $lastSeconds -gt 0 )
{
    $displayedFields += @( 'Mem Usage Average','Cpu Usagemhz Average','Net Usage Average','Cpu Usage Average','Disk Usage Average' )
}

## need to ensure numeric columns are of that type so can sort on them
ForEach( $field in $displayedFields )
{
    $type = $(Switch -Regex ( $field )
    {
        'vCPUS|Memory|Snapshots|HW|Used|Average' { 'int' }
        'Started' { 'datetime' }
        default { 'string' }
    })

    [void]$Datatable.Columns.Add( $field , ( $type -as [type] ) )
}

$WPFchkPerfData.IsChecked = $performanceData

[array]$global:vms = Get-VMs -datatable $datatable -pattern $vmName -poweredOn $showPoweredOn -poweredOff $showPoweredOff -suspended $showSuspended -datastores $script:datastores

$mainForm.Title += " connected to $($server -join ' , ') at $(Get-Date -Format G), $($global:vms.Count) VMs"

$WPFVirtualMachines.ItemsSource = $datatable.DefaultView
$WPFVirtualMachines.IsReadOnly = $true
$WPFVirtualMachines.CanUserSortColumns = $true
##$WPFVirtualMachines.GridLinesVisibility = 'None'
$WPFVirtualMachines.add_MouseDoubleClick({
    Process-Action -GUIobject $WPFVirtualMachines -Operation $doubleClick
})
$WPFConsoleContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Console' -Context $_ })
$WPFMstscContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Mstsc'} )
$WPFMstscNewContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'MstscNew'} )
$WPFReconfigureContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Reconfigure'} )
$WPFPowerOnContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'PowerOn'} )
$WPFPowerOffContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'PowerOff'} )
$WPFSuspendContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Suspend'} )
$WPFResetContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Reset'} )
$WPFShutdownContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Shutdown'} )
$WPFRestartContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Restart'} )
$WPFSnapshotContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Snapshots'} )
$WPFLatestSnapshotRevertContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'LatestSnapshotRevert'} )
$WPFDeleteContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Delete'} )
$WPFScreenshotContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Screenshot'} )
$WPFBackupContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Backup'} )
$WPFEventsContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Events'} )
$WPFbtnDatastores.Add_Click({
    Show-DatastoresWindow
})
$WPFbtnHosts.Add_Click({
    Show-HostsWindow
})
$WPFbtnFilter.Add_Click({
    if( Set-Filters -name $vmName )
    {
        $script:showPoweredOn = $wpfchkPoweredOn.IsChecked
        $script:showPoweredOff = $wpfchkPoweredOff.IsChecked
        $script:showSuspended = $wpfchkSuspended.IsChecked
        $script:vmName = $WPFtxtVMName.Text
        $getError = $null
        $script:snapshots = @( VMware.VimAutomation.Core\Get-Snapshot -ErrorVariable getError -VM $(if( [string]::IsNullOrEmpty( $script:vmName ) ) { '*' } else { $script:vmName }))
        if( $getError -and $getError.Count -and $getError[0].Exception.Message -match 'Not Connected' )
        {
            $connection = Connect-VIServer @connectParameters
            [void][Windows.MessageBox]::Show( "Server was not connected, please retry" , 'Connection Error' , 'Ok' ,'Exclamation' )
        }
        Update-Form -form $mainForm -datatable $script:datatable -vmname $script:vmName
    }
    $_.Handled = $true
})

$WPFbtnRefresh.Add_Click({
    Refresh-Form
    $_.Handled = $true
})

$mainForm.add_KeyDown({
    Param
    (
      [Parameter(Mandatory)][Object]$sender,
      [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$event
    )
    if( $event -and $event.Key -eq 'F5' )
    {
        $_.Handled = $true
        Refresh-Form
    }    
})

$result = $mainForm.ShowDialog()

if( $rdpFileName )
{
    Remove-Item -Path $rdpFileName -Force -ErrorAction SilentlyContinue
}

if( $rdpExtrasFileName )
{
    Remove-Item -Path $rdpExtrasFileName -Force -ErrorAction SilentlyContinue
}

if( ! $alreadyConnected -and $connection )
{
    $connection | Disconnect-VIServer -Force -Confirm:$false
}

