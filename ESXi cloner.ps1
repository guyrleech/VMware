<#
    Clone one or more VMware ESXi VMs from a 'template' VM

    @guyrleech 2018

    Modification History:
#>

<#
.SYNOPSIS

Clone one ore more VMware ESXi VMs from a 'template' VM

.DESCRIPTION

Since ESXi without vCenter does not have templates and therefore a cloning mechanism built in, this script mimics this by allowing existing VMs to
be treated as templates and this script will create one or more new VMs with the same specification as the chosen template and copy its hard disks.

.PARAMETER esxihost

The name or IP address of the ESXi host to use. Credentials will be prompted for if saved ones are not available

.PARAMETER templateName

The exact name or regular expression matching the template VM to use. Must only match one VM

.PARAMETER dataStore

The datastore to create the copied hard disks in

.PARAMETER vmName

The name of the VM to create. If creating more than one, it must contain %d which will be replaced by the number of the clone

.PARAMETER snapshot

The name of the snapshot to use when creating a linked clone. Specifying this automatically enables the linked clone disk feature

.PARAMETER noGUI

Do not show the user interface and go straight to cloning the VM using the supplied command line parameters

.PARAMETER count

The number of clones to create

.PARAMETER notes

Notes to assign to the created VM(s)

.PARAMETER powerOn

Power on the VM(s) once creation is complete

.PARAMETER noConnect

Do not automatically connect to ESXi

.PARAMETER disconnect

Disconnect from ESXi before exit

.PARAMETER numCpus

Override the number of CPUs defined in the template with the number specified here

.PARAMETER numCores

Override the number of cores per CPU defined in the template with the number specified here

.PARAMETER MB

Override the allocated memory defined in the template with the number specified here in MB

.PARAMETER maxVmdkDescriptorSize

If the vmdk file exceeds this size then the script will not attempt to edt it because it is probably a binary file and not a text descriptor

.PARAMETER configRegKey

The registry key used to store the ESXi host and template name to save having to enter them every time

.EXAMPLE

& '.\ESXi cloner.ps1'

Run the user interface which will require various fields to be completed and when "OK" is clicked, create VMs as per these fields

.EXAMPLE

& '.\ESXi cloner.ps1' -noGUI -templateName 'Server 2016 Sysprep' -dataStore 'Datastore1' -vmName 'GRL-2016-%d' -count 4 -notes "Guy's for product testing" -snapshot 'Linked Clone' -poweron

Do not show a user interface and go straight to creating four VMs from the given template name in the datastore1 datastore, creating linked clone (delta) disks from the 
snapshot if the "Server 2016 Sysprep" VM called "Linked Clone". Power each one on once its creation is complete.

.NOTES

Credentials for ESXi can be stored by running "Connect-VIServer -Server <esxi> -User <username> -Password <password> -SaveCredentials

#>

[CmdletBinding()]

Param
(
    [string]$esxihost ,
    [string]$templateName ,
    [string]$dataStore ,
    [string]$vmName ,
    [string]$snapshot ,
    [switch]$noGui ,
    [int]$count = 1 ,
    [string]$notes ,
    [switch]$powerOn ,
    [switch]$noConnect ,
    [switch]$disconnect ,
    ## Parameters to override what comes from template
    [int]$numCpus ,
    [int]$numCores ,
    [int]$MB ,
    ## it is advised not to use the following parameters
    [int]$maxVmdkDescriptorSize = 10KB ,
    [string]$configRegKey = 'HKCU:\Software\Guy Leech\ESXi Cloner' 
)

#region Functions
Function Get-Templates
{
    Param
    (
        $GUIobject ,
        $pattern
    )
    $_.Handled = $true
    [hashtable]$params = @{}
    if( $pattern )
    {
        $params.Add( 'Name' , "*$pattern*" )
    }
    [string[]]$templates = @( Get-VM @params | Where-Object { $_.PowerState -eq 'PoweredOff' } | Sort Name | Select -ExpandProperty Name)
    if( ! $templates -or ! $templates.Count )
    {      
        $null = Display-MessageBox -window $GUIobject -text "Failed to get any powered off templates matching `"$pattern`"" -caption 'Unable to Clone' -buttons OK -icon Error
        return
    }
    ## Stuff in the template list            
    $WPFcomboTemplate.Items.Clear()
    $WPFcomboTemplate.IsEnabled = $true
    $templates | ForEach-Object { $WPFcomboTemplate.items.add( $_ ) }
}

Function Connect-Hypervisor( $GUIobject , $servers , [bool]$pregui, [ref]$vServer )
{
    $vServer.Value = Connect-VIServer -Server $servers -ErrorAction Continue

    if( ! $vServer.Value )
    {
        $null = Display-MessageBox -window $GUIobject -text "Failed to connect to $($servers -join ' , ')" -caption 'Unable to Connect' -buttons OK -icon Error
    }
    elseif( ! $pregui )
    {   
        $_.Handled = $true
        $WPFbtnFetch.IsEnabled = $true
        $WPFcomboDatastore.Items.Clear()
        $WPFcomboDatastore.IsEnabled = $true          
        $WPFcomboTemplate.Items.Clear()
        Get-Datastore | Select -ExpandProperty Name | ForEach-Object { $WPFcomboDatastore.items.add( $_ ) }
        if( $WPFcomboDatastore.items.Count -eq 1 )
        {
            $WPFcomboDatastore.SelectedIndex = 0
        }
    }
}

Function PopulateFrom-Template( $guiobject , $templateName )
{
    if( ! [string]::IsNullOrEmpty( $templateName ) )
    {
        $vm = Get-VM -Name $templateName

        if( $vm )
        {
            $WPFtxtCoresPerCpu.Text = $vm.CoresPerSocket
            $WPFtxtCPUs.Text = $vm.NumCpu
            $WPFtxtMemory.Text = $vm.MemoryMB
            ## select MB in the drop down unit comboMemoryUnits
            $wpfcomboMemoryUnits.SelectedItem = $WPFcomboMemoryUnits.Items.GetItemAt(0)
            $wpfcomboMemoryUnits.BringIntoView()
            if( $WPFchkLinkedClone.IsChecked )
            {
                $WPFcomboSnapshot.Items.Clear()
                $WPFcomboSnapshot.IsEnabled = $true
                Get-Snapshot -VM $VM | Select -ExpandProperty Name | ForEach-Object { $WPFcomboSnapshot.items.add( $_ ) }
            }
        }
        else
        {
            $null = Display-MessageBox -window $guiobject -text "Failed to retrieve template `"$templateName`"" -caption 'Unable to Clone' -buttons OK -icon Error
        }
    }
}

Function Validate-Fields( $guiobject )
{
    $_.Handled = $true
    
    if( [string]::IsNullOrEmpty( $WPFtxtVMName.Text ) )
    {
        $null = Display-MessageBox -window $guiobject -text 'No VM Name Specified' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( ! $WPFcomboTemplate.SelectedItem )
    {
        $null = Display-MessageBox -window $guiobject -text 'No Template VM Selected' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( ! $WPFcomboDatastore.SelectedItem )
    {
        $null = Display-MessageBox -window $guiobject -text 'No Datastore Selected' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( $WPFchkLinkedClone.IsChecked -and ! $WPFcomboSnapshot.SelectedItem )
    {
        $null = Display-MessageBox -window $guiobject -text 'No Snapshot Selected for Linked Clone' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    $result = $null
    if( ! [int]::TryParse( $WPFtxtNumberClones.Text , [ref]$result ) -or ! $result )
    {
        $null = Display-MessageBox -window $guiobject -text 'Specified number of clones is invalid' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( $result -gt 1 -and $WPFtxtVMName.Text -notmatch '%d' )
    {
        $null = Display-MessageBox -window $guiobject -text 'Must specify %d replacement pattern in the VM name when creating more than one clone' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( $result -eq 1 -and $WPFtxtVMName.Text -match '%d' )
    {
        $null = Display-MessageBox -window $guiobject -text 'Illegal character ''%'' in VM name - did you mean to make more than one clone ?' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( ! [int]::TryParse( $WPFtxtMemory.Text , [ref]$result ) -or ! $result )
    {
        $null = Display-MessageBox -window $guiobject -text 'Specified memory value is invalid' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( ! [int]::TryParse( $WPFtxtCPUs.Text , [ref]$result ) -or ! $result )
    {
        $null = Display-MessageBox -window $guiobject -text 'Specified number of CPUs is invalid' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( ! [int]::TryParse( $WPFtxtCoresPerCpu.Text , [ref]$result ) -or ! $result )
    {
        $null = Display-MessageBox -window $guiobject -text 'Specified number of cores is invalid' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    return $true
}

Function Display-MessageBox( $window , $text , $caption , [System.Windows.MessageBoxButton]$buttons , [System.Windows.MessageBoxImage]$icon )
{
    if( $window -and $window.Handle )
    {
        [int]$modified = switch( $buttons )
            {
                'OK' { [System.Windows.MessageBoxButton]::OK }
                'OKCancel' { [System.Windows.MessageBoxButton]::OKCancel }
                'YesNo' { [System.Windows.MessageBoxButton]::YesNo }
                'YesNoCancel' { [System.Windows.MessageBoxButton]::YesNo }
            }
        [int]$choice = [PInvoke.Win32.Windows]::MessageBox( $Window.handle , $text , $caption , ( ( $icon -as [int] ) -bor $modified ) )  ## makes it app modal so UI blocks
        switch( $choice )
        {
            ([MessageBoxReturns]::IDYES -as [int]) { 'Yes' }
            ([MessageBoxReturns]::IDNO -as [int]) { 'No' }
            ([MessageBoxReturns]::IDOK -as [int]) { 'Ok' } 
            ([MessageBoxReturns]::IDABORT -as [int]) { 'Abort' } 
            ([MessageBoxReturns]::IDCANCEL -as [int]) { 'Cancel' } 
            ([MessageBoxReturns]::IDCONTINUE -as [int]) { 'Continue' } 
            ([MessageBoxReturns]::IDIGNORE -as [int]) { 'Ignore' } 
            ([MessageBoxReturns]::IDRETRY -as [int]) { 'Retry' } 
            ([MessageBoxReturns]::IDTRYAGAIN -as [int]) { 'TryAgain' } 
        }       
    }
    else
    {
        [Windows.MessageBox]::Show( $text , $caption , $buttons , $icon )
    }
}

## Based on code from http://www.lucd.info/2010/03/31/uml-diagram-your-vm-vdisks-and-snapshots/
Function Get-SnapshotDisk{
    param($vmName , $snapshotname)

    Function Get-SnapHash{
        param($parent)
        Process {
            $snapHash[$_.Snapshot.Value] = @($_.Name,$parent)
            if($_.ChildSnapshotList){
                $newparent = $_
                $_.ChildSnapShotList | Get-SnapHash $newparent
            }
        }
    }

    $snapHash = @{}
    Get-View -ViewType VirtualMachine -Filter @{'Name' = $vmName} | ForEach-Object {
        $vm = $_
        if($vm.Snapshot){
            $vm.Snapshot.RootSnapshotList | Get-SnapHash $vm
        }
        else{
            return
        }
        $firstHD = $true
        $_.Config.Hardware.Device | Where-Object {$_.DeviceInfo.Label -match '^Hard disk' -and $_.Backing.DiskMode -notmatch 'independent' } | ForEach-Object {
            $hd = $_
            $hdNr = $hd.DeviceInfo.Label.Split('\s')[-1]
            $exDisk = $vm.LayoutEx.Disk | Where-Object {$_.Key -eq $hd.Key}
            $diskFiles = @()
            $exDisk.Chain | ForEach-Object {$_.FileKey} | ForEach-Object { $diskFiles += $_ }

            $snapHash.GetEnumerator() | ForEach-Object {
                $key = $_.Key
                $value = $_.Value
                if( [string]::IsNullOrEmpty( $snapshotname ) -or $snapshotname -eq $value[0] ) {
                    $vm.LayoutEx.Snapshot | Where-Object {$_.Key.Value -eq $key} | ForEach-Object {
                        $vmsnId = $_.DataKey
                        $_.Disk | Where-Object{$_.Key -eq $hd.Key} | ForEach-Object {
                            if($diskFiles -notcontains $_.Chain[-1].FileKey[0] -and $diskFiles -notcontains $_.Chain[-1].FileKey[1]){
                                $chain = $_.Chain[-1]
                            }
                            else{
                                $preSnapFiles = $_.Chain | ForEach-Object {$_.FileKey} | ForEach-Object {$_}
                                $vm.layoutEx.Disk | Where-Object {$_.Key -eq $hd.Key} | ForEach-Object {
                                    foreach($chain in $_.Chain){
                                        if($preSnapFiles -notcontains $chain.FileKey[0] -and $preSnapFiles -notcontains $chain.FileKey[1]){
                                            break
                                        }
                                    }
                                }
                            }
                            $chain.FileKey | ForEach-Object {
                                if( ! $vm.LayoutEx.File[$_].Size ){
                                    $vm.LayoutEx.File[$_].Name
                                }
                            }
                        }
                    }
                }
            }
            $firstHD = $false
        }
    }
}

Function Add-NewDisk
{
    [CmdletBinding()]

    Param
    (
        $VM ,
        $sourceDisk ,
        $newDisk
    )
    
    $sourceController = $sourceDisk | Get-ScsiController
    if( ! $sourceController )
    {
        Write-Warning "Source disk $($sourceDisk.Name) was not attached to a SCSI controller so creating new one"
        [hashtable]$diskParams = @{
            'VM' = $VM
            'DiskPath' = $newDisk
            'ErrorAction' =  'Continue'
        } 
        New-HardDisk @diskParams | New-ScsiController -Type Default -BusSharingMode NoSharing
    }
    else
    {
        $dsName = $newDisk.Split(']')[0].TrimStart('[')
        $ds = Get-Datastore -Name $dsName
        $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
        $spec.deviceChange = @()
        $spec.deviceChange += New-Object VMware.Vim.VirtualDeviceConfigSpec
        $spec.deviceChange[0].device = New-Object VMware.Vim.VirtualDisk
        $spec.deviceChange[0].device.backing = New-Object VMware.Vim.VirtualDiskFlatVer2BackingInfo
        $spec.deviceChange[0].device.backing.datastore = $ds.ExtensionData.MoRef
        $spec.deviceChange[0].device.backing.fileName = $newDisk
        $spec.deviceChange[0].device.backing.diskMode = switch( $sourceDisk.Persistence )
        {
            'IndependentNonPersistent' { 'independent_nonpersistent' }
            'IndependentPersistent'    { 'independent_persistent' }
            default { $_ }
        }
        $spec.deviceChange[0].device.unitnumber = -1
        $spec.deviceChange[0].device.controllerKey = $scsiControllers[ $sourceController.ExtensionData.Key ]
        $spec.deviceChange[0].operation = 'add'
        $VM.ExtensionData.ReconfigVM($spec)

        Get-HardDisk -VM $VM | Where-Object { $_.Filename -eq $newDisk }
    }
}

Function New-ClonedDisk
{
    [CmdletBinding()]
    Param
    (
        $cloneVM ,
        $sourceDisk ,
        $destinationDatastore ,
        [string]$parent ,
        [int]$diskCount
    )
    [string]$baseDiskName = [io.path]::GetFileNameWithoutExtension( ( Split-Path $sourceDisk.FileName -Leaf ) )
    Write-Verbose "Cloning `"$baseDiskName`" , format $($sourceDisk.StorageFormat) to [$($destinationDatastore.Name)] $($cloneVM.Name)"
    [hashtable]$copyParams = @{}
    if( $sourceDisk.PSObject.properties[ 'DestinationStorageFormat' ] )
    {
        $copyParams.Add( 'DestinationStorageFormat' , $sourceDisk.StorageFormat )
    }
    $clonedDisk = $sourceDisk | Copy-HardDisk -DestinationPath "[$($destinationDatastore.Name)] $($cloneVM.Name)" @copyParams
    if( $clonedDisk )
    {
        [string]$diskToAdd = $null
        [string]$diskProviderPath = $null
        [string]$qualifier = $null
        ## Now rename the disk files for this disk 
        Get-ChildItem -Path (Join-Path $destinationFolder ($baseDiskName + '*.vmdk')) | ForEach-Object `
        {
            [string]$restOfName = $_.Name -replace "^$baseDiskName(.*)\.vmdk$" , '$1'
            [string]$newDiskName = $cloneVM.Name
            if( $sourceDisks.Count -gt 1 )
            {
                $qualifier = ".disk$diskCount"
                $newDiskName += $qualifier
            }

            $newDiskName += "$restOfName.vmdk"
            Rename-Item -Path $_.FullName -NewName $newDiskName -ErrorAction Stop ## PassThru doesn't work
            if( ! $diskToAdd -and ! $restOfName )
            {
                $diskToAdd = "{0}/{1}" -f $_.FolderPath , $newDiskName ## can't use Join-Path as that uses a backslash
                $diskProviderPath = Join-Path -Path (Split-Path $_.FullName -Parent) -ChildPath $newDiskName
            }
        }
        if( $diskToAdd -and $diskProviderPath )
        {
            ## Now we have to edit the file to point its extents at the renamed file. Can't edit in situ so have to copy to local file system, change and copy back
            [int]$fileLength = Get-ChildItem -Path $diskProviderPath | Select -ExpandProperty Length
            if( $fileLength -lt $maxVmdkDescriptorSize )
            {
                [string]$tempDisk = Join-Path $env:temp ( $baseDiskName + '.' + $pid + '.vmdk' )
                if( Test-Path -Path $tempDisk -ErrorAction SilentlyContinue )
                {
                    Remove-Item -Path $tempDisk -Force -ErrorAction Stop
                }
                Copy-DatastoreItem -Item $diskProviderPath -Destination $tempDisk
                $existingContent = Get-Content -Path $tempDisk
                $newContent = $existingContent | ForEach-Object `
                {
                    if( $_ -match "^(.*) `"$baseDiskName(.*)\.vmdk\`"$" )
                    {
                        "$($matches[1]) `"$($cloneVM.Name)$qualifier$($matches[2]).vmdk`""
                    }
                    else
                    {
                        $_
                    }
                }
                if( $existingContent -eq $newContent )
                {
                    Write-Warning "Unexpectedly, no changes were made to renamed disk $diskToAdd"
                }
                ## Need to output without CR/LF that Windows normally does with text files. It's ASCII not UTF8 despite what the "encoding" says in the file
                $streamWriter = New-Object System.IO.StreamWriter( $tempDisk  , $false , [System.Text.Encoding]::ASCII )
                $newcontent | ForEach-Object { $streamwriter.Write( ( $_ + "`n" ) ) }
                $streamWriter.Close()

                Copy-DatastoreItem -Destination $diskProviderPath -Item $tempDisk -Force -ErrorAction Continue
                Remove-Item -Path $tempDisk -Force
                $result = Add-NewDisk -VM $cloneVM -sourceDisk $sourceDisk -newDisk $diskToAdd
                if( ! $result )
                {
                    Throw "Failed to add cloned disk `"$diskToAdd`" to cloned VM"
                }
            }
            else
            {
                Write-Error "`"$diskProviderPath`" is large ($($fileLength/1KB)KB) which means it probably isn't a text descriptor file that can be changed"
            }
        }
        else
        {
            Write-Error "No disk path to add to new VM for disk $baseDiskName"
        }
    }
    else
    {
        Throw "Failed to clone disk $($sourceDisk.FileName)"
    }
}
#endregion Functions

Remove-Module -Name Hyper-V -ErrorAction SilentlyContinue ## lest it clashes
Import-Module -Name VMware.PowerCLI -ErrorAction Stop

if( Test-Path -Path $configRegKey -PathType Container -ErrorAction SilentlyContinue )
{
    if( [string]::IsNullOrEmpty( $esxihost ) )
    {
        $esxihost = Get-ItemProperty -Path $configRegKey -Name 'ESXi Host' -ErrorAction SilentlyContinue | select -ExpandProperty 'ESXi Host'
    }
    if( [string]::IsNullOrEmpty( $templateName ) )
    {
        $templateName = Get-ItemProperty -Path $configRegKey -Name 'Template Pattern' -ErrorAction SilentlyContinue | select -ExpandProperty 'Template Pattern' 
    }
}

$vServer = $null

if( ! $noConnect -and ! [string]::IsNullOrEmpty( $esxihost ) )
{
    Connect-Hypervisor -GUIobject $null -servers $esxihost -pregui $true -vServer ([ref]$vServer)
}

#region XAML&Modules

[string]$mainwindowXAML = @'
<Window x:Class="MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApplication1"
        mc:Ignorable="d"
        Title="ESXi VM Cloner" Height="487.015" Width="430.5" FocusManager.FocusedElement="{Binding ElementName=txtTargetComputer}">
    <Grid Margin="10,10,46,13" >
        <Grid VerticalAlignment="Top" Height="345" Margin="0,20,0,0">
            <Grid.RowDefinitions>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="200*"></ColumnDefinition>
                <ColumnDefinition Width="250*"></ColumnDefinition>
                <ColumnDefinition Width="150*"></ColumnDefinition>
            </Grid.ColumnDefinitions>

            <TextBlock Grid.Row="0" Grid.Column="0" Text="ESXi Host"></TextBlock>
            <TextBlock Grid.Row="1" Grid.Column="0" Text="Template Pattern"></TextBlock>
            <TextBlock Grid.Row="2" Grid.Column="0" Text="Template VM"></TextBlock>
            <TextBlock Grid.Row="3" Grid.Column="0" Text="Snapshot"></TextBlock>
            <TextBlock Grid.Row="4" Grid.Column="0" Text="New VM Name"></TextBlock>
            <TextBlock Grid.Row="5" Grid.Column="0" Text="CPUs"></TextBlock>
            <TextBlock Grid.Row="6" Grid.Column="0" Text="Cores per CPU"></TextBlock>
            <TextBlock Grid.Row="7" Grid.Column="0" Text="Memory"></TextBlock>
            <TextBlock Grid.Row="8" Grid.Column="0" Text="Datastore"></TextBlock>
            <TextBlock Grid.Row="9" Grid.Column="0" Text="Notes"></TextBlock>
            <TextBlock Grid.Row="10" Grid.Column="0" Text="Number of Clones"></TextBlock>
            <TextBlock Grid.Row="11" Grid.Column="0"></TextBlock>
            <TextBlock Grid.Row="12" Grid.Column="0" Text="Options:"></TextBlock>
            <TextBlock Grid.Row="13" Grid.Column="0" Text=""></TextBlock>

            <TextBox x:Name="txtESXiHost" Grid.Row="0" Grid.Column="1"></TextBox>
            <Button x:Name="btnConnect" Grid.Row="0"  Grid.Column="2" Content="_Connect"></Button>
            <Button x:Name="btnFetch" Grid.Row="1"  Grid.Column="3" Content="_Fetch"></Button>
            <TextBox x:Name="txtTemplate" Grid.Row="1" Grid.Column="1"/>
            <ComboBox x:Name="comboTemplate" Grid.Row="2" Grid.Column="1"></ComboBox>
            <ComboBox x:Name="comboSnapshot" Grid.Row="3" Grid.Column="1"></ComboBox>
            <TextBox x:Name="txtVMName" Grid.Row="4" Grid.Column="1" Text=""></TextBox>
            <TextBox x:Name="txtCPUs" Grid.Row="5" Grid.Column="1" Text="2"></TextBox>
            <TextBox x:Name="txtCoresPerCpu" Grid.Row="6" Grid.Column="1" Text="1"></TextBox>
            <TextBox x:Name="txtMemory" Grid.Row="7" Grid.Column="1" Text="2"></TextBox>
            <ComboBox x:Name="comboDatastore" Grid.Row="8" Grid.Column="1" Text="5"></ComboBox>
            <TextBox x:Name="txtNotes" Grid.Row="9" Grid.Column="1"></TextBox>
            <TextBox x:Name="txtNumberClones" Grid.Row="10" Grid.Column="1" Text="1"></TextBox>
            <ComboBox x:Name="comboMemoryUnits" Grid.Row="7" Grid.Column="2" Text="5">
                <ComboBoxItem Content="MB" FontSize="9" IsSelected="True"/>
                <ComboBoxItem Content="GB" FontSize="9"/>
            </ComboBox>

            <CheckBox x:Name="chkLinkedClone" Content="_Linked Clone Disks" Grid.Row="12" Grid.Column="1"/>
            <CheckBox x:Name="chkStart" Content="_Start after Creation" Grid.Row="13" Grid.Column="1"/>
            <CheckBox x:Name="chkSaveSettings" Content="Sa_ve Settings" Grid.Row="14" Grid.Column="1" IsChecked="True"/>
            <CheckBox x:Name="chkDisconnect" Content="_Disconnect on Exit" Grid.Row="15" Grid.Column="1"/>
        </Grid>
        <Button x:Name="btnOk" Content="OK" HorizontalAlignment="Left" Height="31" Margin="8,381,0,0" VerticalAlignment="Top" Width="90" IsDefault="True"/>
        <Button x:Name="btnCancel" Content="Cancel" HorizontalAlignment="Left" Height="31" Margin="122,381,0,0" VerticalAlignment="Top"  Width="90" IsDefault="False" IsCancel="True"/>
    </Grid>
</Window>
'@

Function Load-GUI( $inputXml )
{
    $form = $NULL
    $inputXML = $inputXML -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window'
 
    [xml]$XAML = $inputXML
 
    $reader = New-Object Xml.XmlNodeReader $xaml

    try
    {
        $Form = [Windows.Markup.XamlReader]::Load( $reader )
    }
    catch
    {
        Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$_"
        return $null
    }
 
    $xaml.SelectNodes('//*[@Name]') | ForEach-Object `
    {
        Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global
    }

    return $form
}

#endregion XAML&Modules

if( ! $noGui )
{
    [void][Reflection.Assembly]::LoadWithPartialName('Presentationframework')

    $mainForm = Load-GUI $mainwindowXAML

    if( ! $mainForm )
    {
        return
    }

    if( $DebugPreference -eq 'Inquire' )
    {
        Get-Variable -Name WPF*
    }
    ## set up call backs

    $WPFbtnConnect.Add_Click({
        $_.Handled = $true
        if( [string]::IsNullOrEmpty( $wpftxtESXiHost.Text ) )
        {
            Display-MessageBox -window $mainForm -text 'No ESXi Host Specified' -caption 'Unable to Connect' -buttons OK -icon Error
        }
        else
        {
            Connect-Hypervisor -GUIobject $mainForm -servers $wpftxtESXiHost.Text -pregui $false -vServer ([ref]$vServer)
        }
    })
    $WPFbtnFetch.Add_Click({
        $_.Handled = $true
        if( [string]::IsNullOrEmpty( $WPFtxtTemplate.Text ) )
        {
            Display-MessageBox -window $mainForm -text 'No Template Pattern Specified' -caption 'Unable to Fetch Templates' -buttons OK -icon Error
        }
        else
        {
            Get-Templates -GUIobject $mainForm -pattern $WPFtxtTemplate.Text
        }
    })
    $WPFcomboTemplate.add_SelectionChanged({
        $_.Handled = $true
        PopulateFrom-Template -guiobject $mainForm -template $WPFcomboTemplate.SelectedItem
    })
    $WPFbtnOk.add_Click({
        $_.Handled = $true
        if( Validate-Fields -guiobject $mainForm )
        {
            $mainForm.DialogResult = $true
            $mainForm.Close()
        }
    })
    $WPFchkLinkedClone.add_Checked({
        $_.Handled = $true
        $WPFcomboSnapshot.IsEnabled = $true
        if( $vServer )
        {
            $WPFcomboSnapshot.Items.Clear()
            $WPFcomboSnapshot.IsEnabled = $true
            if( $WPFcomboTemplate.SelectedItem )
            {
                Get-Snapshot -VM (Get-VM -Name $WPFcomboTemplate.SelectedItem) | Select -ExpandProperty Name | ForEach-Object { $WPFcomboSnapshot.items.add( $_ ) }
            }
        }
    })
    $WPFchkLinkedClone.add_Unchecked({
        $_.Handled = $true
        $WPFcomboSnapshot.IsEnabled = $false
    })
    $WPFtxtESXiHost.Text = $esxihost
    $WPFtxtVMName.Text = $vmName
    $WPFtxtTemplate.Text = $templateName
    $WPFbtnFetch.IsEnabled = $vServer -ne $null
    $WPFcomboTemplate.IsEnabled = $false
    $WPFcomboDatastore.IsEnabled = $vServer -ne $null
    if( $vServer )
    {
        Get-Datastore | Select -ExpandProperty Name | ForEach-Object { $null = $WPFcomboDatastore.items.add( $_ ) }
        if( $WPFcomboDatastore.items.Count -eq 1 )
        {
            $WPFcomboDatastore.SelectedIndex = 0
        }
        $WPFbtnConnect.Content = 'Connected'
    }
    $WPFcomboSnapshot.IsEnabled = $false

    $result = $mainForm.ShowDialog()
    
    $disconnect = $wpfchkDisconnect.IsChecked
    if( ! $result )
    {
        if( $disconnect )
        {
            $vServer | Disconnect-VIServer -Confirm:$false
        }

        return
    }
    $chosenTemplate = Get-VM -Name $WPFcomboTemplate.SelectedItem
    $chosenDatastore = Get-Datastore -Name $WPFcomboDatastore.SelectedItem
    $chosenName = $WPFtxtVMName.Text
    $chosenMemory = [int]$WPFtxtMemory.Text
    ## Chosen memory may now need adjusting for units
    if( $WPFcomboMemoryUnits.SelectedItem.Content -eq 'GB' ) 
    {
        $chosenMemory *= 1024
    }
    ## else MB already and that's what we pass to New-VM
    $chosenCPUs = $WPFtxtCPUs.Text
    $chosenCores = $WPFtxtCoresPerCpu.Text
    $count = $WPFtxtNumberClones.Text
    $notes = $WPFtxtNotes.Text
    if( $WPFchkLinkedClone.IsChecked )
    {
        $snapshot = $WPFcomboSnapshot.SelectedItem
    }
    $powerOn = $WPFchkStart.IsChecked
    if( $WPFchkSaveSettings )
    {   
        if( ! ( Test-Path $configRegKey -ErrorAction SilentlyContinue ) )
        {
            $null = New-Item -Path $configRegKey -Force
        }
        Set-ItemProperty -Path $configRegKey -Name 'ESXi Host' -Value $WPFtxtESXiHost.Text
        Set-ItemProperty -Path $configRegKey -Name 'Template Pattern' -Value $WPFtxtTemplate.Text
    }
}
else
{
    ## Get list of template VMs
    [array]$templates = @( Get-VM -Name "*$templateName*" | Where-Object { $_.PowerState -eq 'PoweredOff' } )

    if( ! $templates -or ! $templates.Count )
    {
        Throw "Found no powered off VMs named with $templateName"
    }
    if( $templates.Count -gt 1 )
    {
        Throw "Found $($templates.Count) templates, please be more specific ($(($templates|Select -ExpandProperty Name) -join ','))"
    }
    [array]$dataStores = @( Get-Datastore )

    if( ! $dataStores -or ! $dataStores.Count )
    {
        Throw "Found no datastores!"
    }
    $chosenTemplate = $templates[0]
    $chosenDatastore = if( $PSBoundParameters[ 'DataStore' ] ) { $dataStores | Where-Object { $_.Name -match $dataStore } } else { $dataStores[0] }
    $chosenName = $vmName
    $chosenMemory = if( $PSBoundParameters[ 'MB' ] ) { $MB } else { $chosenTemplate.MemoryMB }
    $chosenCPUs = if( $PSBoundParameters['NumCPUs'] ) { $numCpus } else { $chosenTemplate.NumCPU }
    $chosenCores = if( $PSBoundParameters['NumCores'] ) { $numCores} else { $chosenTemplate.CoresPerSocket }
}

if( $count -gt 1 -and $chosenName -notmatch '%d' )
{
    Throw "When creating multiple clones, the name must contain %d which will be replaced by the number of the clone"
}

[array]$sourceDisks = @( $chosenTemplate | Get-HardDisk )

1..$count | ForEach-Object `
{
    [int]$vmNumber = $_
    [string]$thisChosenName = if( $count -gt 1 ) { $chosenName -replace '%d',$vmNumber } else { $chosenName }

    Write-Verbose "Creating VM # $vmNumber : $thisChosenName"

    $cloneVM = New-VM -Name $thisChosenName -MemoryMB $chosenMemory -NumCpu $chosenCPUs -CoresPerSocket $chosenCores -Datastore $chosenDatastore -DiskMB 1 -DiskStorageFormat Thin -GuestId $chosenTemplate.ExtensionData.Config.GuestId -Notes $notes

    if( ! $cloneVM )
    {
        Throw "Failed to clone new VM `"$thisChosenName`""
    }

    $thisChosenName = $cloneVM.Name

    ## Remove the auto added hard disk and NICs as we'll add from template
    $cloneVM | Get-HardDisk | Remove-HardDisk -DeletePermanently -Confirm:$false
    $cloneVM | Get-NetworkAdapter | Remove-NetworkAdapter -Confirm:$false

    ## if video memory has been changed, e.g. for full HD screen resolution, then change the clone to match
    $sourceVideoSettings = $chosenTemplate.ExtensionData.Config.Hardware.Device | Where-Object { $_.GetType().Name -eq 'VirtualMachineVideoCard' }
    if( $sourceVideoSettings )
    {
        $destinationVideoSettings = $cloneVM.ExtensionData.Config.Hardware.Device | Where-Object { $_.GetType().Name -eq 'VirtualMachineVideoCard' }
        if( $destinationVideoSettings )
        {
            if( $sourceVideoSettings.VideoRamSizeInKB -ne $destinationVideoSettings.VideoRamSizeInKB `
                -or $sourceVideoSettings.NumDisplays -ne $destinationVideoSettings.NumDisplays `
                -or $sourceVideoSettings.GraphicsMemorySizeInKB -ne $destinationVideoSettings.GraphicsMemorySizeInKB `
                -or $sourceVideoSettings.Use3dRenderer -ne $destinationVideoSettings.Use3dRenderer `
                -or $sourceVideoSettings.UseAutoDetect -ne $destinationVideoSettings.UseAutoDetect )
            {
                ## https://code.vmware.com/forums/2530/vsphere-powercli#580657
                $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
                $deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec
                $deviceChange.Operation = 'edit'
                $deviceChange.Device += $sourceVideoSettings
                $spec.DeviceChange += $deviceChange
                $cloneVM.ExtensionData.ReconfigVM($spec)
            }
        }
        else
        {
            Write-Warning 'Failed to get video settings for cloned VM'
        }
    }
    else
    {
        Write-Warning 'Failed to get video settings for source VM'
    }

    Get-CDDrive -VM $chosenTemplate | ForEach-Object `
    {
        [hashtable]$cdparams = @{
            'StartConnected' = $_.ConnectionState.StartConnected
            'VM' = $cloneVM
        }
        if( $_.HostDevice )
        {
            $cdparams.Add(  'HostDevice' , $_.HostDevice )
        }
        if( $_.RemoteDevice )
        {
            $cdparams += @{
                'IsoPath' = $_.IsoPath
                'RemoteDevice' = $_.RemoteDevice
            }
        }
        $newcd = New-CDDrive @cdparams
    }

    Get-FloppyDrive -VM $chosenTemplate | ForEach-Object `
    {
        [hashtable]$floppyparams = @{
            'VM' = $cloneVM
            'StartConnected' = $_.ConnectionState.StartConnected
        }
        if( $_.HostDevice )
        {
            $floppyparams.Add(  'HostDevice' , $_.HostDevice )
        }
        if( $_.RemoteDevice )
        {
            $floppyparams += @{
                'FloppyImagePath' = $_.FloppyImagePath
                'RemoteDevice' = $_.RemoteDevice
            }
        }
        $newfloppy = New-FloppyDrive @floppyparams
    }

    ## if EFI then set it
    ## https://github.com/vmware/PowerCLI-Example-Scripts/blob/master/Scripts/SecureBoot.ps1
    $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
    $spec.Firmware = $chosenTemplate.ExtensionData.Config.Firmware
    $bootOptions = New-Object VMware.Vim.VirtualMachineBootOptions
    $bootOptions.EfiSecureBootEnabled = $chosenTemplate.ExtensionData.Config.BootOptions.EfiSecureBootEnabled
    $bootOptions.BootDelay = $chosenTemplate.ExtensionData.Config.BootOptions.BootDelay
    $bootOptions.BootOrder = $chosenTemplate.ExtensionData.Config.BootOptions.BootOrder
    $bootOptions.BootRetryDelay = $chosenTemplate.ExtensionData.Config.BootOptions.BootRetryDelay
    $bootOptions.BootRetryEnabled = $chosenTemplate.ExtensionData.Config.BootOptions.BootRetryEnabled
    $bootOptions.EnterBIOSSetup = $false
    $bootOptions.NetworkBootProtocol = $chosenTemplate.ExtensionData.Config.BootOptions.NetworkBootProtocol
    $spec.BootOptions = $bootOptions
    $cloneVM.ExtensionData.ReconfigVM( $spec )

    [array]$snapshots = @( Get-Snapshot -VM $chosenTemplate | Where-Object { $_.Name -match $snapshot } )

    $view = $chosenTemplate | Get-View
    $newView = $cloneVM | Get-View
    [int]$busNumber = 0
    [int]$key = 1
    [hashtable]$scsiControllers = @{}
    $Controllers = $view.Config.Hardware.Device | Where-Object {$_ -is [VMware.Vim.ParaVirtualSCSIController] -or $_ -is [VMware.Vim.VirtualLsiLogicSASController] -or $_ -is [VMware.Vim.VirtualLsiLogicController]} | ForEach-Object `
    {
        $existingController = $_

        $storagespec = New-Object VMware.Vim.VirtualMachineConfigSpec
        $NewSCSIDevice = New-Object VMware.Vim.VirtualDeviceConfigSpec
        $NewSCSIDevice.operation = 'add'
        $NewSCSIDevice.device = New-Object -TypeName "VMware.Vim.$($existingcontroller.GetType().Name)"
        $NewSCSIDevice.device.key = $key++
        $NewSCSIDevice.device.busNumber = $BusNumber++
        $NewSCSIDevice.device.sharedBus = $existingController.SharedBus
        $storageSpec.deviceChange += $NewSCSIDevice
        [array]$deviceBefore = @( $newView.Config.Hardware.Device | Select -ExpandProperty Key )
        $newView.ReconfigVM($storageSpec)
        $newView.UpdateViewData()
        [array]$devicesAfter = @( $newView.Config.Hardware.Device )
        [int]$controllersBefore = $scsiControllers.Count
        $devicesAfter | Where-Object { $_.Key -notin $deviceBefore } | ForEach-Object `
        {
            ## So we can map disks to the correct controller
            $scsiControllers.Add( $existingController.Key , $_.Key )
        }
        if( $scsiControllers.Count -eq $controllersBefore )
        {
            Write-Warning "Failed to find newly added $($existingcontroller.GetType().Name) SCSI controller so disks may not be attached or attached incorrectly"
        }
    }

    Get-NetworkAdapter -VM $chosenTemplate | ForEach-Object `
    {   
        $newNIC = New-NetworkAdapter -NetworkName $_.NetworkName -Type $_.Type -WakeOnLan:$_.WakeOnLanEnabled -VM $cloneVM -StartConnected:$_.ConnectionState.StartConnected
        if( ! $newNIC )
        {
            Write-Warning "Failed to create NIC for `"$($_.NetworkName)`""
        }
    }

    ## Update VM now we have changed its hardware
    $cloneVM = Get-VM -Id $cloneVM.Id

    if( ! $sourceDisks -or ! $sourceDisks.Count )
    {
        Write-Warning "Template $($chosenTemplate.Name) has no disks"
    }
    else
    {
        ## Check chosen datastore has enough free space for the copies
        $totalDiskSize = $sourceDisks | Measure-Object -Property CapacityGB -Sum | Select -ExpandProperty Sum
        if( $totalDiskSize -gt $chosenDatastore.FreeSpaceGB )
        {
            Write-Warning "Disks require $($totalDiskSize)GB but datastore $($chosenDatastore.Name) only has $($chosenDatastore.FreeSpaceGB)GB free space"
        }
        $destinationDatastore = $cloneVM | Get-Datastore
        $destinationDatacentre = $cloneVM | Get-Datacenter
        $destinationFolder = "vmstore:\$destinationDatacentre\$destinationDatastore\$thisChosenName"
        if( ! ( Test-Path -Path $destinationFolder ) )
        {
            Throw "Path $destinationFolder does not exist"
        }
        [int]$diskCount = 1
   
        if( ! [string]::IsNullOrEmpty( $snapshot ) )
        {
            ## Linked clone from a named snapshot of this VM
            if( ! $snapshots -or ! $snapshots.Count )
            {
                Throw "Failed to find snapshot matching `"$snapshot`" for template `"$($chosenTemplate.Name)`""
            }
            elseif( $snapshots.Count -gt 1 )
            {
                Throw "Snapshot `"$snapshot`" is ambiguous as matches $($snapshots.Count) snapshots - $(($snapshots | Select -ExpandProperty Name) -join ',')"
            }
            $theSnapshot = $snapshots[0]
            if( $theSnapshot.PowerState -ne 'PoweredOff' )
            {
                Throw "Snapshot `"$($snapshot.Name)`" is in power state $($theSnapshot.PowerState) but only powered off ones can be cloned from"
            }
            ## Get differencing disk  for this snapshot as we will copy that, rename it and change the parent disk since its location will be different
            [array]$differencingDisks = @( Get-SnapshotDisk -vmName $chosenTemplate.Name -snapshotname $snapshot )
            if( ! $differencingDisks -or ! $differencingDisks.Count )
            {
                Throw "Failed to get the differencing disk for snapshot `"$snapshot`""
            }
            ForEach( $differencingDisk in $differencingDisks )
            {
                [string]$thisDataStore = $null
                if( $differencingDisk -match '^\[(.*)\] ' )
                {
                    $thisDataStore = $Matches[1]
                }
                $sourceDisk = Get-HardDisk -DatastorePath $differencingDisk -Datastore $thisDataStore
                if( ! $sourceDisk )
                {
                    Throw "Failed to get the hard disk from snapshot disk `"$differencingDisk`""
                }
                ## Since we will be copying this disk, we need to change the parentFileNameHint field to point to the original base disk for which this is the differencing disk
                ## Can't use Copy-HardDisk on the delta as that will copy the parent too
                if( $differencingDisk -match '^\[(.*)\]\s(.*)/(.*)\.vmdk$' )
                {
                    [string]$sourceDatacentre = $chosenTemplate | Get-Datacenter
                    [string[]]$diskBits = $Matches[3] -split '-'
                    [string]$sourceBaseDiskName = $diskBits[0]
                    [string]$sourceDatastore = $Matches[1]
                    [string]$sourceParentFolder = $Matches[2]
                    [string]$sourceFolder = "vmstore:\$sourceDatacentre\$sourceDatastore\$sourceParentFolder"
                    [string]$source = "$sourceFolder\$($matches[3]).vmdk"
                    [string]$newDiskName = $cloneVM.Name
                    if( $differencingDisks.Count -gt 1 )
                    {
                        $newDiskName += ".disk$diskCount"      
                    }
                    $newDiskName += '.vmdk'
                    [string]$tempDisk = Join-Path $env:temp $newDiskName
                    if( Test-Path -Path $tempDisk -ErrorAction SilentlyContinue )
                    {
                        Remove-Item -Path $tempDisk -Force -ErrorAction Stop
                    }
                    Copy-DatastoreItem -Item $source -Destination $tempDisk
                    $destinationBinaryDisk = $null
                    $sourceBinaryDisk = $null
                    $existingContent = Get-Content -Path $tempDisk
                    $null = Get-Random -SetSeed ([datetime]::Now.Ticks % [int]::MaxValue)
                    $newCID = ([int]"0x$(((1..4 | %{ "{0:x}" -f (Get-Random -Max 256) } ) -join ''))").ToString('x8') ## needs to be 8 hex digits
                    $newContent = $existingContent | ForEach-Object `
                    {
                        if( $_-match '^CID=' )
                        {
                            "CID=$newCID" ## needs to be different from the parentCID
                        }
                        elseif( $_ -match  "parentFileNameHint=""(.*)""" )
                        {
                            if( $sourceDatastore -eq $destinationDatastore )
                            {
                                "parentFileNameHint=""../$sourceParentFolder/$($Matches[1])"""
                            }
                            else
                            {
                                "parentFileNameHint=""../../$($sourceDatastore)/$sourceParentFolder/$($Matches[1])"""
                            }
                        }
                        elseif( $_ -match "(.*) ""$sourceBaseDiskName-(\d{6})-(.*)\.vmdk""$" )
                        {
                            $destinationBinaryDisk = "{0}-{1}.vmdk" -f ($newDiskName -replace '\.vmdk$' , '') , $Matches[3]
                            "{0} ""{1}""" -f $Matches[1] , $destinationBinaryDisk
                            $sourceBinaryDisk = ($_ -split '"')[1]
                        }
                        else
                        {
                            $_
                        }
                    }
                    if( $existingContent -eq $newContent )
                    {
                        Write-Warning "Unexpectedly, no changes were made to renamed disk $diskToAdd"
                    }
                    if( [string]::IsNullOrEmpty( $destinationBinaryDisk ) )
                    {
                        Write-Warning "Failed to get binary difference disk name from descriptor file $source"
                    }
                    [string]::Join( "`n" , $newContent ) | Set-Content -Path $tempDisk -Force
                    ## Copy back to cloned VM's folder
                    Copy-DatastoreItem -Destination $destinationFolder -Item $tempDisk -Force -ErrorAction Continue
                    Remove-Item -Path $tempDisk -Force

                    ## Copy the binary differencing disk but can't do with Copy-HardDisk as will copy the parent too
                    Copy-DatastoreItem -Item ($sourceFolder + '\' + $sourceBinaryDisk ) -Destination $destinationFolder
                    Move-Item -Path (Join-Path $destinationFolder $sourceBinaryDisk) -Destination (Join-Path $destinationFolder $destinationBinaryDisk)
                    ## Need to find the base disk so we can find what SCSI controller it is attached to
                    [string]$baseSourceDiskName = (($sourceDisk.FileName -replace '-\d{6}\.vmdk$' , '') -split '/')[-1]
                    $theSourceDisk = $sourceDisks | Where-Object { $_.FileName -match "$($baseSourceDiskName)\.vmdk" -or $_.FileName -match "$($baseSourceDiskName)-\d{6}\.vmdk" }
                    [string]$destinationDisk = ( "[{0}] {1}/{2}" -f $destinationDatastore , $thisChosenName , $newDiskName )
                    ## Add disk to VM
                    $result = Add-NewDisk -VM $cloneVM -sourceDisk $theSourceDisk -newDisk $destinationDisk
                    if( ! $result )
                    {
                        Throw "Failed to add cloned disk `"$destinationDisk`" to cloned VM"
                    }
                }
                else
                {
                    Throw "Unexpected format of source snapshot disk `"$differencingDisk`""
                }
                $diskCount++
            }
            ## Now copy non-persistent disks
            ForEach( $sourceDisk in $sourceDisks )
            {
                if( $sourceDisk.Persistence -ne 'Persistent' )
                {
                    New-ClonedDisk -cloneVM $cloneVM -sourceDisk $sourceDisk -destinationDatastore $destinationDatastore -diskCount $diskCount
                    $diskCount++
                }
            }
        }
        else
        {
            if( $snapshots -and $snapshots.Count )
            {
                Write-Warning "Template has $($snapshots.Count) snapshots, cloning from the current state of the VM"
            }
            ForEach( $sourceDisk in $sourceDisks )
            {
                New-ClonedDisk -cloneVM $cloneVM -sourceDisk $sourceDisk -destinationDatastore $destinationDatastore -diskCount $diskCount
                $diskCount++
            }
        }
    }
    if( $powerOn )
    {
        $poweredOn = Start-VM -VM $cloneVM
        if( ! $poweredOn -or $poweredOn.PowerState -ne 'PoweredOn' )
        {
            Write-Warning "Failed to power on `"$($cloneVM.Name)`""
        }
    }
}

if( $disconnect )
{
    $vServer|Disconnect-VIServer -Confirm:$false
}
