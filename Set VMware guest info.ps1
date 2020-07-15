
<#
.SYNOPSIS

Set VM guest information so it can be retrieved in VMs 

.DESCRIPTION

The VMware properties to set are those returned from the Get-VM cmdlet

.PARAMETER server

The vCenter server(s) to connect to

.PARAMETER properties

Comma separated list of properties to set where the property name to set is before the delimiter, default is =, and the VMware property to use is after the delimiter

.PARAMETER VMs

Specific VM name or pattern to operate on. If not specified will operate on all powered on VMs.

.PARAMETER credential

Credentials to use to connect to the vCenter(s)

.PARAMETER port

The port to connect to on the vCenter(s)

.PARAMETER protocol

The protocol to use with the specified vCenter(s)

.PARAMETER allLinked

Connect to all vCenters linked to the specified vCenter(s)

.PARAMETER guestinfo

The prefix to use for the property name in the VM

.PARAMETER delimiter

The delimiter to specify the property name from the VMware property name

.PARAMETER literal

The string prefix which denotes that a VMware property name is a string literal rather than a VMware property name

.PARAMETER powerState

A regular expression that matches the power state of VMs to operate on

.PARAMETER remove

Remove the specified properties rather than set them

.EXAMPLE

& '.\Set VMware guest info.ps1' -server grl-vcenter04 -VMs GRL-W10* -properties host.name=VMhost,host.version=VMhost.Version,host.build=VMhost.Build,host.parent=VMhost.parent,owner='*Guy Leech'

Set the specified properties for all VMs matching the pattern GRL-W10* managed by the vCenter server grl-vcenter04
   
.NOTES

Retrieve the properties set via this script in a Windows VM via vmtoolsd.exe https://www.virtuallyghetto.com/2011/01/how-to-extract-host-information-from.html

@guyrleech
#>

[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]

Param
(
    [Parameter(Mandatory,HelpMessage='VMware vCenter to use')]
    [string[]]$server ,
    [Parameter(Mandatory,HelpMessage='Properties to set name=property')]
    [string[]]$properties ,
    [string]$VMs ,
    [System.Management.Automation.PSCredential]$credential ,
    [int]$port ,
    [ValidateSet( 'http' , 'https' )]
    [string]$protocol ,
    [switch]$allLinked ,
    [string]$guestinfo = 'guestinfo' ,
    [string]$delimiter = '=' ,
    [string]$literal = '*' ,
    [string]$powerState = 'PoweredOn' ,
    [switch]$remove
)

Function Get-VMProperty
{
    Param
    (
        [Parameter(Mandatory)]
        $VM ,
        [Parameter(Mandatory)]
        [string]$Property
    )

    ## we have dots in the property name which are taken literally so we have to get to the property 1 property at a time
    [string[]]$individualProperties = $Property -split '\.'

    if( ! $individualProperties -or $individualProperties.Count -le 1 )
    {
        $VM.$Property.ToString() ## top level property
    }
    else
    {
        $parent = $VM
        ForEach( $individualProperty in $individualProperties )
        {
            if( $parent -and $parent.PSObject.Properties[ $individualProperty ] )
            {
                $parent = $parent.$individualProperty
            }
            else
            {
                $parent = $null
            }
        }
        if( $parent )
        {
            $parent.ToString()
        }
        else
        {
            Write-Warning -Message "Failed to get property $property for VM $($VM.Name)"
        }
    }
}

$connection = $null

Import-Module -Name VMware.VimAutomation.Core -Verbose:$false

[hashtable]$connectionParameters = @{ 'Server' = $server ; 'AllLinked' = $allLinked }

if( $PSBoundParameters[ 'credential' ] )
{
    $connectionParameters.Add( 'Credential' , $credential )
}

if( $PSBoundParameters[ 'port' ] )
{
    $connectionParameters.Add( 'port' , $port )
}

if( $PSBoundParameters[ 'protocol' ] )
{
    $connectionParameters.Add( 'protocol' , $protocol )
}

if( ! ( $connection = Connect-VIServer @connectionParameters ) )
{
    Throw "Failed to connect to $server"
}

try
{
    ## either get all powered on VMs or jsut a specific set
    [hashtable]$getvmParameters = @{ }
    if( $PSBoundParameters[ 'VMs' ] )
    {
        $getvmParameters.Add( 'Name' , $VMs )
    }

    [array]$virtualMachines = @( Get-VM @getvmParameters | Where-Object { $_.PowerState -match $powerState } )

    if( ! $virtualMachines -or ! $virtualMachines.Count )
    {
        Throw "No virtual machines retrieved"
    }

    Write-Verbose -Message "Retrieved $($virtualMachines.Count) powered on virtual machines"

    [int]$counter = 0

    ForEach( $virtualMachine in $virtualMachines )
    {
        $counter++
        Write-Verbose -Message "$counter / $($virtualMachines.Count) $($virtualMachine.Name)"
        if( $PSCmdlet.ShouldProcess( "$($properties.Count) properties in VM $($virtualMachine.Name)" , 'Set' ) )
        {
            [int]$propertyCounter = 0
            ForEach( $property in $properties )
            {
                $propertyCounter++
                [string]$advancedSettingName,[string]$vmPropertyName = $property -split $delimiter , 2
                ## don't check strings as empty/null vmProperty can be set
                if( $remove )
                {
                    if( $advancedSetting = Get-AdvancedSetting -Entity $virtualMachine -Name $advancedSettingName )
                    {
                        Remove-AdvancedSetting -AdvancedSetting $advancedSetting
                    }
                }
                else
                {
                    $vmPropertyValue = $null
                    if( ! [string]::IsNullOrEmpty( $vmPropertyName ) )
                    {
                        if( ! [string]::IsNullOrEmpty( $literal ) -and $vmPropertyName.StartsWith( $literal ) )
                        {
                            $vmPropertyValue = $vmPropertyName.SubString( $literal.Length )
                        }
                        else
                        {
                            $vmPropertyValue = Get-VMProperty -VM $virtualMachine -Property $vmPropertyName
                        }
                    }
                    if( $null -ne $vmPropertyValue )
                    {
                        [string]$fullPropertyName = "$guestinfo.$advancedSettingName"
                        Write-Verbose -Message "Setting property $propertyCounter / $($properties.Count) $fullPropertyName to `"$vmPropertyValue`" in VM $($virtualMachine.Name)"
                        if( ! ( $setting = New-AdvancedSetting -Name $fullPropertyName -Value $vmPropertyValue -Entity $virtualMachine -Confirm:$false -Force ) )
                        {
                            Write-Warning -Message "Failed to set property $fullPropertyName to `"$vmPropertyValue`" in VM $($virtualMachine.Name)"
                        }
                    }
                }
            }
        }
    }
}
catch
{
    Throw $_
}
finally
{
    $connection | Disconnect-VIServer -Confirm:$false
    $connection = $null
}
