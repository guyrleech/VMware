#requires -version 3
<#
    Show VMware ESXi/vSphere or Hyper-V VM details in a grid view, standard output or text file that Sysinternals BGinfo can use as a custom field

    @guyrleech 2018
#>

<#
.SYNOPSIS

Show VMware ESXi/vSphere or Hyper-V VM details in a grid view, standard output or text file that Sysinternals BGinfo can use as a custom field

.DESCRIPTION

BGinfo allows custom fields but only via WMI queries, VBS scripts or text file contents so this script can create a text file with VM details in
which can then be used in a custom field in a BGInfo .bgi file

.PARAMETER viServers

A comma separated list of ESXi servers to connect to. If not specified then Hyper-V is assumed as the hypervisor

.PARAMETER hypervServer

The name of a Hyper-V server to connect to. Cannot be used with -viServers. Will use the local computer if not specified

.PARAMETER vmName

The name or pattern of a virtual machine to include. If not specified then all VMs found will be included.

.PARAMETER wait

Wait for keyboard input when the script has finished. Useful if the script is being run from a shortcut so the output doesn't disappear until the enter key is pressed.

.PARAMETER gridView

Display the results in an on screen filterable and sortable grid view. Any rows selected when OK is clicked are put in the clipboard. If not specified output goes to standard output

.PARAMETER ipv4only

Do not include IPv6 addresses

.PARAMETER noHeaders

Do not output any headers

.PARAMETER bginfoFile

A non-existent text file, unless -overwrite is specified, which will have the VM details written to for use as a custom field in SysInternals BGinfo

.PARAMETER overWrite

Will overwrite an existing file when the -bginfoFile option is used

.PARAMETER bgInfoFields

The VM fields which will be placed into the file specified -bginfofile

.PARAMETER fields

The VM fields to display or output

.PARAMETER tabStop

The tab stop size used by BGinfo. Do not change unless output is misaligned

.EXAMPLE

'.\Get VMware or Hyper-V powered on vm details.ps1' -bginfoFile C:\temp\vmaddresses.txt -overwrite -viservers 192.168.0.69  -ipv4only -vmName 'GRL*'

Retrieve details for powered on VMware VMs called GRL* from the ESXi/vCenter server at 192.168.0.69 and write the VM details to the file C:\temp\vmaddresses.txt which can then be used in a custom rule in BGinfo.

.EXAMPLE

'.\Get VMware or Hyper-V powered on vm details.ps1' -gridview

Retrieve details for all powered on VMs from the local Hyper-V server and output the results to an on screen gridview                          

.EXAMPLE

'.\Get VMware or Hyper-V powered on vm details.ps1' -hypervServer GRL-HYPERV07

Retrieve details for all powered on VMs from the Hyper-V server GRL-HYPERV07 and output the results to standard output which can then be piped into other scripts/cmdlets or a file                         

.NOTES

To apply a BGinfo .bgi file automatically, run the fulling where the .bgi file already has the custom field in where the output file from -bginfofile is specified:

Bginfo.exe' 'c:\temp\your-custom-background.bgi' /timer:0 /nolicprompt

Run at logon and/or as a periodically scheduled scheduled task

Requires VMware PowerCLI if querying VMware machines or the Microsoft Hyper-V module if querying Hyper-V VMs

BGInfo is available at https://docs.microsoft.com/en-us/sysinternals/downloads/bginfo

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true, ParameterSetName = "VMware")]
    [string[]]$viservers ,
    [Parameter(Mandatory=$false, ParameterSetName = "HyperV")]
    [string]$hypervServer ,
    [string]$vmName ,
    [switch]$wait ,
    [switch]$gridView ,
    [switch]$ipv4only ,
    [switch]$noHeaders ,
    [string]$bginfoFile ,
    [switch]$overwrite ,
    [string[]]$bginfoFields = @( 'Name' , 'IPAddresses' ) ,
    [string[]]$fields = @( 'Name' , 'NumCPU' , 'CoresPerSocket' , 'MemoryMB' , 'IPAddresses' , 'UsedSpaceGB' , 'ProvisionedSpaceGB' ) ,
    [int]$tabStop = 6 ## for bginfo
)

if( $PSBoundParameters[ 'bginfoFile' ] -and $gridView )
{
    Throw '-bginfoFile and -gridView cannot be used together'
}

Function Get-VHDSize
{
    [CmdletBinding()]

    Param
    (
        $vhd
    )
    
    [long]$usedSpaceSoFar = 0
    if( $vhd )
    {
        $disk = Get-VHD -Path $vhd.Path
        $usedSpaceSoFar = $disk.FileSize
        if( $vhd.ParentPath )
        {
            $usedSpaceSoFar += Get-VHDSize -vhd (Get-VHD -Path $vhd.ParentPath)
        }
    }
    $UsedSpaceSoFar
}

[hashtable]$getvmparams = @{}

if( $PSBoundParameters[ 'vmName' ] )
{
    $getvmparams.Add( 'Name' , $vmName )
}

[string]$addressPattern = $null

if( $ipv4only )
{
    $addressPattern = '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$'
}

$connectedServer = $null

if( $PSBoundParameters[ 'viservers' ] )
{
    Import-Module VMware.PowerCLI -ErrorAction Stop

    $connectedServer = Connect-VIServer -Server $viservers -ErrorAction Stop
    $vms = @( Get-VM @getvmparams | Where-Object PowerState -eq 'PoweredOn'|Select *,@{n='IPAddresses';e={($_.guest.IPaddress | Where-Object { $_ -match $addressPattern } | Sort ) -join ' , '}} | Select -Property $fields )
}
else
{
    Remove-Module VMware.* -ErrorAction SilentlyContinue
    Import-Module Hyper-V -ErrorAction Stop
    if( $PSBoundParameters[ 'hypervServer' ] )
    {
        $getvmparams.Add( 'ComputerName' , $hypervServer )
    }
    ## if they have multiple NICs then it produces multiple entries for NIC so process per VM to get single entry
    $vms = @( Get-VM @getvmparams | Where-Object State -eq 'Running' | ForEach-Object `
    {
        $VM = $_
        $result = [pscustomobject]@{ 
            'Name' = $VM.Name
            'NumCPU' = $VM.ProcessorCount
            'CoresPerSocket' = 1 -as [int]
            'IPAddresses' = ''
            'MemoryMB' = [int]($VM.MemoryAssigned / 1MB )
            'UsedSpaceGB' = 0 -as [long]
            'ProvisionedSpaceGB' = 0 -as [long]
        }
        $VM | Get-VMNetworkAdapter | Select -ExpandProperty IPAddresses -ErrorAction SilentlyContinue | Sort | ForEach-Object `
        {
            if( $_ -match $addressPattern )
            {
                $result.IPAddresses += ( "{0}{1}" -f $(if( $result.IPAddresses.Length ) { ' , ' } ) , $_ )
            }
        }
        Get-VHD -VMid $VM.VMid | ForEach-Object `
        {
            $result.ProvisionedSpaceGB  += $_.Size
            ## will recurse to parent disks if snapshots
            $result.UsedSpaceGB = Get-VHDSize -vhd $_
        }
        $result.UsedSpaceGB = [int]($result.UsedSpaceGB / 1GB)
        $result.ProvisionedSpaceGB = [int]($result.ProvisionedSpaceGB / 1GB)
        $result
    })
}

if( ! $vms -or ! $vms.Count )
{
    Write-Warning 'Found no VMs'
}
elseif( $gridView )
{
    $selected = $vms | Out-GridView -PassThru

    if( $selected )
    {
        $selected | Set-Clipboard
    }
}
else
{
    if( $PSBoundParameters[ 'bginfoFile' ] )
    {
        ## output data in format that bginfo can read as a text file as a custom field in a bgi file
        [hashtable]$outfileParams = @{ 'FilePath' = $bginfoFile ; 'NoClobber' = (!$overwrite) }
        ## figure out longest name so we know how many tabs to use
        [string]$tabs = "`t`t`t`t`t`t`t`t`t" ## overkill!
        [string]$firstPad = $null
        [int]$longestName = -1
        $vms | ForEach-Object { $longestName = [math]::Max( $longestName , $_.Name.Length ) }
        $vms | Select -Property $bginfoFields | ForEach-Object `
        {
            "{0}{1}:`t{2}{3}" -f $firstPad , $_.Name , $tabs.Substring( 0 , ( $longestName - $_.Name.Length ) / $tabStop ) , ( $_.PSObject.Properties | Where Name -eq $bginfoFields[-1] | Select -ExpandProperty Value )
            $firstPad = "`t"
        } | Out-File @outfileParams
    }
    else
    {
        [hashtable]$ftParams =  @{ 'AutoSize' = $true ; 'HideTableHeaders' = $noHeaders }
        $vms | Format-Table @ftParams
    }
}

if( $wait )
{
    $null = Read-Host "Hit <enter> to continue"
}
