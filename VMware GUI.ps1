#requires -version 3
<#
    Modification History:

    @guyrleech 04/11/2019  Initial release
    @guyrleech 04/11/2019  Added client drive mapping to rdp file created for credentials and added -extraRDPSettings parameter
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

.PARAMETER vmName

The name of a VM or pattern to include in the GUI

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

.PARAMETER webConsole

Do not attempt to run vmrc.exe when launching the VMware console for a VM

.PARAMETER noRdpFile

DO not create an .rdp file to specify the username for mstsc

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

.PARAMETER sortProperty

The VM property to sort the display on.

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
    [string]$vmName = '*' ,
    [int]$rdpPort = 3389 ,
    [string]$mstscParams ,
    [ValidateSet('mstsc','console','PowerOn','reconfigure','snapshots','Delete')]
    [string]$doubleClick = 'mstsc',
    [switch]$saveCredentials ,
    [switch]$ipv6 ,
    [switch]$passThru ,
    [switch]$webConsole ,
    [switch]$noRdpFile ,
    [bool]$showPoweredOn = $true,
    [bool]$showPoweredOff ,
    [bool]$showSuspended ,
    [switch]$showAll ,
    [string]$sortProperty = 'Name' ,
    [string[]]$extraRDPSettings = @( 'drivestoredirect:s:*' ) ,
    [string]$regKey = 'HKCU:\Software\Guy Leech\Simple VMware Console' ,
    [string]$vmwareModule = 'VMware.PowerCLI'
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
    <DockPanel Margin="10,10,10,10">
        <Button x:Name="btnConnect" Content="Connect" HorizontalAlignment="Left" Height="35"  Margin="10,0,0,0" VerticalAlignment="Bottom" Width="145" DockPanel.Dock="Bottom"/>
        <Button x:Name="btnFilter" Content="Filter" HorizontalAlignment="Left" Height="35" Margin="247,0,0,-35" VerticalAlignment="Bottom" Width="145" DockPanel.Dock="Bottom"/>
        <Button x:Name="btnRefresh" Content="Refresh" HorizontalAlignment="Left" Height="35" Margin="500,0,0,-35" Width="145" VerticalAlignment="Bottom" DockPanel.Dock="Bottom" />
        <DataGrid HorizontalAlignment="Stretch" VerticalAlignment="Top" x:Name="VirtualMachines" >
            <DataGrid.ContextMenu>
                <ContextMenu>
                    <MenuItem Header="Console" x:Name="ConsoleContextMenu" />
                    <MenuItem Header="Mstsc" x:Name="MstscContextMenu" />
                    <MenuItem Header="Snapshots" x:Name="SnapshotContextMenu" />
                    <MenuItem Header="Reconfigure" x:Name="ReconfigureContextMenu" />
                    <MenuItem Header="Delete" x:Name="DeleteContextMenu" />
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
        <CheckBox x:Name="chkPoweredOn" Content="Powered On" HorizontalAlignment="Left" Height="32" Margin="187,94,0,0" VerticalAlignment="Top" Width="217"/>
        <CheckBox x:Name="chkPoweredOff" Content="Powered Off" HorizontalAlignment="Left" Height="32" Margin="187,131,0,0" VerticalAlignment="Top" Width="217"/>
        <CheckBox x:Name="chkSuspended" Content="Suspended" HorizontalAlignment="Left" Height="32" Margin="187,163,0,0" VerticalAlignment="Top" Width="217"/>
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
        <Button x:Name="btnTakeSnapshot" Content="Take Snapshot" HorizontalAlignment="Left" Margin="597,88,0,0" VerticalAlignment="Top" Width="92"/>
        <Button x:Name="btnDeleteSnapshot" Content="Delete" HorizontalAlignment="Left" Margin="597,280,0,0" VerticalAlignment="Top" Width="92"/>
        <Button x:Name="btnSnapshotsOk" Content="OK" HorizontalAlignment="Left" Margin="95,365,0,0" VerticalAlignment="Top" Width="75" IsDefault="True"/>
        <Button x:Name="btnSnapshotsCancel" Content="Cancel" HorizontalAlignment="Left" Margin="198,365,0,0" VerticalAlignment="Top" Width="75" IsCancel="True"/>
        <Button x:Name="btnRevertSnapshot" Content="Revert" HorizontalAlignment="Left" Margin="597,148,0,0" VerticalAlignment="Top" Width="92"/>
        <Button x:Name="btnConsolidateSnapshot" Content="Consolidate" HorizontalAlignment="Left" Margin="597,208,0,0" VerticalAlignment="Top" Width="92"/>
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
    $ChildItem.Name = $Name -replace '[\s,;:]' , '_' -replace '%252f' , '_' # default snapshot names have / for date which are escaped
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

    if( $operation -eq 'ConsolidateSnapshot' )
    {
        $VM = Get-VM -Id $VMId
        $VM.ExtensionData.ConsolidateVMDisks_Task()
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
    elseif( $GUIobject.SelectedItem -and $GUIobject.SelectedItem.Tag -and $GUIobject.SelectedItem.Tag -ne 'YouAreHere' )
    {
        [string]$answer = 'yes'
        $answer = [Windows.MessageBox]::Show( "Are you sure you want to $($operation -creplace '([a-zA-Z])([A-Z])' , '$1 $2') on $($vm.Name)?" , 'Confirm Snapshot Operation' , 'YesNo' ,'Question' )
        
        if( $answer -eq 'yes' )
        {
            $VM = Get-VM -Id $VMId
            if( $VM )
            {
                $snapshot = Get-Snapshot -Id $GUIobject.SelectedItem.Tag -VM $vm
                if( $snapshot )
                {
                    if( $Operation -eq 'DeleteSnapShot' )
                    {
                        Remove-Snapshot -Snapshot $snapshot -RunAsync -Confirm:$false
                    }
                    elseif( $Operation -eq 'RevertSnapShot' )
                    {
                        Set-VM -VM $vm -Snapshot $snapshot -RunAsync -Confirm:$false
                    }
                    else
                    {
                        Write-Warning "Unexpcted snapshot operation $operation"
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
    }

    ## Close dialog since needs refreshing
    $snapshotsForm.DialogResult = $true 
    $snapshotsForm.Close()
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
            return
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
                $parent = $theseSnapshots | Where-Object { $_.Id -eq $snapshot.ParentSnapshotId }
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
                Add-TreeItem -Name 'You are here' -Parent $currentSnapShotItem -Tag 'YouAreHere'
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

        ## see if consolidation is required so that we enable/disable the consolidation button
        $wpfbtnConsolidateSnapshot.IsEnabled = $vm.Extensiondata.Runtime.ConsolidationNeeded
                        
        $WPFbtnTakeSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'TakeSnapShot' -VMId $vm.Id } )
        $WPFbtnDeleteSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'DeleteSnapShot' -VMId $vm.Id} )
        $WPFbtnRevertSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'RevertSnapShot' -VMId $vm.Id} )
        $WPFbtnConsolidateSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'ConsolidateSnapShot' -VMId $vm.Id } )

        $result = $snapshotsForm.ShowDialog()
    }
}

Function Process-Action
{
    Param
    (
        $GUIobject , 
        [string]$Operation
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
        $vm = VMware.VimAutomation.Core\Get-VM -Name $selectedVM.Name | Where-Object { $_.VMHost.Name -eq $selectedVM.Host }
        if( ! $vm )
        {
            Write-Warning -Message "Unable to get VM $($selectedVM.Name)"
            return
        }
        elseif( $vm -is [array] )
        {
            Write-Warning -Message "Found $($vm.Count) VMware objects matching $($selectedVM.Name)"
        }
        else
        {
            [string]$address = $null
            if( $Operation -eq 'mstsc' )
            {
                ## TODO implement width, height, admin, etc in settings
                ## See if we can resolve its name otherwise we will try its IP addresses
                if( ! ( $address = $vm.guest.HostName ) )
                {
                    $vm.'IP Addresses' -split ',' | ForEach-Object `
                    {
                        $connection = Test-NetConnection -ComputerName $PSItem -Port $rdpPort -ErrorAction SilentlyContinue
                        if( $connection -and ! $address )
                        {
                            $address = $PSItem
                        }
                    }
                }
        
                if( $address )
                {
                    [string]$arguments = "$rdpFileName /v:$($address):$rdpPort"
                    if( ! [string]::IsNullOrEmpty( $mstscParams ))
                    {
                        $arguments += " $mstscParams"
                    }
                    $mstscProcess = Start-Process -FilePath 'mstsc.exe' -ArgumentList $arguments -PassThru
                }
                else
                {
                    [void][Windows.MessageBox]::Show( "No address for $($vm.Name)" , 'Connection Error' , 'Ok' ,'Warning' )
                }
            }
            elseif( $Operation -eq 'console' )
            {
                $vmrcProcess = $null
                if( $session -and ! $webConsole )
                {
                    $ticket = $Session.AcquireCloneTicket()
                    $vmrcProcess = Start-Process -FilePath 'vmrc.exe' -ArgumentList "vmrc://clone:$ticket@$($connection.ToString())/?moid=$($vm.ExtensionData.MoRef.value)" -PassThru -WindowStyle Normal -Verb Open -ErrorAction SilentlyContinue
                }
                ## fallback to PowerCLI console but it doesn't persist connection across power operations
                if( ! $vmrcProcess )
                {
                    Open-VMConsoleWindow -VM $vm
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
                        VMware.VimAutomation.Core\Set-VM -VM $vm -NumCpu $newvCPUS -Confirm:$false
                        if( ! $? )
                        {
                            [void][Windows.MessageBox]::Show( "Failed to change vCPUs from $($vm.NumCPU) to $newvCPUS in $($vm.Name)" , 'Reconfiguration Error' , 'Ok' ,'Exclamation' )
                        }
                    }
                    if( $newMemory -ne $vm.MemoryGB )
                    {
                        VMware.VimAutomation.Core\Set-VM -VM $vm -MemoryGB $newMemory -Confirm:$false
                        if( ! $? )
                        {
                            [void][Windows.MessageBox]::Show( "Failed to change memory from $($vm.MemoryGB)GB to $($newMemory)GB in $($vm.Name)" , 'Reconfiguration Error' , 'Ok' ,'Exclamation' )
                        }
                    }
                    if( $WPFtxtNotes.Text -ne $vm.Notes )
                    {
                        VMware.VimAutomation.Core\Set-VM -VM $vm -Notes $WPFtxtNotes.Text -Confirm:$false
                    }
                }
            }
            elseif( $Operation -eq 'Snapshots' )
            {
                Show-SnapShotWindow -vm $vm
            }
            elseif( $Operation -eq 'Delete' )
            {
                [string]$answer = [Windows.MessageBox]::Show( "Are you sure you want to delete $($vm.Name)?" , 'Confirm Delete Operation' , 'YesNo' ,'Question' )
                if( $answer -eq 'yes' )
                {
                    VMware.VimAutomation.Core\Remove-VM -DeletePermanently -VM $vm -Confirm:$false
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
    $global:vms = Get-VMs -datatable $datatable -pattern $vmName -poweredOn $showPoweredOn -poweredOff $showPoweredOff -suspended $showSuspended
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
        [bool]$suspended = $true
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
                'Snapshots' = ( $snapshots | Where-Object { $_.VMId -eq $VM.Id } | Measure-Object|Select-Object -ExpandProperty Count)
                'Started' = $vm.ExtensionData.Runtime.BootTime
                'Used Space (GB)' = [Math]::Round( $vm.UsedSpaceGB , 1 )
                'Memory (GB)' = $vm.MemoryGB
                'VMware Tools' = $vm.Guest.ToolsVersion
            }
            Add-Member -InputObject $vm -NotePropertyMembers $additionalProperties
            $vm
            $items = New-Object -TypeName System.Collections.ArrayList
            ForEach( $field in $displayedFields )
            {
                [void]$items.Add( $vm.$field )
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

## see if already connected and if so if it is the server we are told to connect to
if( Get-Variable -Scope Global -Name DefaultVIServers -ErrorAction SilentlyContinue )
{
    $existingServer = @( $global:DefaultVIServers | Select-Object -ExpandProperty Name )
    if( $existingServer -and $existingServer.Count )
    {
        Write-Verbose -Message "Already connected to $($existingServer -join ' , ')"
        $server = $existingServer

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
    }
}

## if we have no server then see if saved in registry
if( ! $PSBoundParameters[ 'server' ] )
{
    $server = (Get-ItemProperty -Path $regKey -Name 'Server' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'Server') -split ','
}

if( ! $PSBoundParameters[ 'mstscParams' ] )
{
    $mstscParams = Get-ItemProperty -Path $regKey -Name 'mstscParams' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'mstscParams'
}

if( ! $alreadyConnected -and ! $server -or ! $server.Count )
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

$credential = $null

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

    if( ! $passThru )
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

    [hashtable]$connectParameters = @{ 'Server' = $server ; 'Force' = $true }
    if( $credential )
    {
        Write-Verbose "Connecting to $($server -join ',') as $($credential.username)"
        $connectParameters.Add( 'Credential' , $credential )
    }
    
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
if( $credential )
{
    Write-Verbose "Connecting as $($credential.username)"
    [string]$rdpusername = $credential.UserName
    [string]$rdpdomain = $null
    if( $credential.UserName.IndexOf( '@' ) -lt 0 )
    {
        $rdpdomain,$rdpUsername = $credential.UserName -split '\\'
    }
    $rdpCredential = @( "username:s:$rdpusername" , "domain:s:$rdpdomain" )
}
elseif( $passThru )
{
    $rdpCredential = @( "username:s:$env:USERNAME" , "domain:s:$env:USERDOMAIN" )
}

[string]$rdpFileName = $null
if( ! $noRdpFile )
{
    if( $rdpCredential -and $rdpCredential.Count )
    {
        ## Write username and domain to rdp file to pass to mstsc
        $rdpFileName = Join-Path -Path $env:temp -ChildPath "grl.$pid.rdp"
        Write-Verbose "Writing $($rdpUsername -join ' , ') to $rdpFileName"
        $rdpCredential + $extraRDPSettings | Out-File -FilePath $rdpFileName
    }
}

if( ! ( Test-Path -Path $regKey -ErrorAction SilentlyContinue ) )
{
    [void](New-Item -Path $regKey -Force)
}

Set-ItemProperty -Path $regKey -Name 'Server' -Value ($server -join ',')
Set-ItemProperty -Path $regKey -Name 'MstscParams' -Value $mstscParams

## so we can acquire a ticket if required for vmrc remote console (will fail on ESXi)
$Session = Get-View -Id Sessionmanager -ErrorAction SilentlyContinue

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

[array]$snapshots = @( VMware.VimAutomation.Core\Get-Snapshot -VM $vmName )

$datatable = New-Object -TypeName System.Data.DataTable

[string[]]$displayedFields = @( "Name" , "Power State" , "Host" , "Notes" , "Started" , "vCPUs" , "Memory (GB)" , "Snapshots" , "IP Addresses" , "VMware Tools" , "Used Space (GB)" )
[void]$Datatable.Columns.AddRange( $displayedFields )

[array]$global:vms = Get-VMs -datatable $datatable -pattern $vmName -poweredOn $showPoweredOn -poweredOff $showPoweredOff -suspended $showSuspended

$mainForm.Title += " connected to $($server -join ' , ')"

$WPFVirtualMachines.ItemsSource = $datatable.DefaultView
$WPFVirtualMachines.IsReadOnly = $true
$WPFVirtualMachines.CanUserSortColumns = $true
##$WPFVirtualMachines.GridLinesVisibility = 'None'
$WPFVirtualMachines.add_MouseDoubleClick({
    Process-Action -GUIobject $WPFVirtualMachines -Operation $doubleClick
})
$WPFConsoleContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Console'} )
$WPFMstscContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Mstsc'} )
$WPFReconfigureContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Reconfigure'} )
$WPFPowerOnContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'PowerOn'} )
$WPFPowerOffContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'PowerOff'} )
$WPFSuspendContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Suspend'} )
$WPFResetContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Reset'} )
$WPFShutdownContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Shutdown'} )
$WPFRestartContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Restart'} )
$WPFSnapshotContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Snapshots'} )
$WPFDeleteContextMenu.Add_Click( { Process-Action -GUIobject $WPFVirtualMachines -Operation 'Delete'} )

$WPFbtnFilter.Add_Click({
    if( Set-Filters -name $vmName )
    {
        $script:showPoweredOn = $wpfchkPoweredOn.IsChecked
        $script:showPoweredOff = $wpfchkPoweredOff.IsChecked
        $script:showSuspended = $wpfchkSuspended.IsChecked
        $script:vmName = $WPFtxtVMName.Text
        $script:snapshots = @( VMware.VimAutomation.Core\Get-Snapshot -VM $(if( [string]::IsNullOrEmpty( $script:vmName ) ) { '*' } else { $script:vmName }))
        Update-Form -form $mainForm -datatable $script:datatable -vmname $script:vmName
    }
    $_.Handled = $true
})

$WPFbtnRefresh.Add_Click({
    $script:snapshots = @( VMware.VimAutomation.Core\Get-Snapshot -VM $(if( [string]::IsNullOrEmpty( $script:vmName ) ) { '*' } else { $script:vmName }))
    Update-Form -form $mainForm -datatable $script:datatable -vmname $script:vmName
    $_.Handled = $true
})

$result = $mainForm.ShowDialog()

if( $rdpFileName )
{
    Remove-Item -Path $rdpFileName -Force -ErrorAction SilentlyContinue
}

$connection | Disconnect-VIServer -Force -Confirm:$false
