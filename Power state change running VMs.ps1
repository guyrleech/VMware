<#
    Pause or shutdown running VMs - designed to be run by UPS shutdown software

    @guyrleech 2019

    Modification History:

    09/09/19   GRL   Added manual input of password when encrypting it for later use
#>

<#
.SYNOPSIS

Pause or shutdown running VMs and the ESXi host - designed to be run by UPS shutdown software

.PARAMETER viServer

The ESXi server to connect to

.PARAMETER username

The username to use when connecting to the ESXi server

.PARAMETER password

The clear text password for the username specified. If none is specified and there are no saved credentials or pass thru then the contents of the environment varinble _Mval12 are used as the password if set

.PARAMETER securePassword

A password previously encrypted using the -encryptPassword argument

.PARAMETER encryptPassword

Encrypt the password specified via -password, in the environment variable _Mval12 or prompt if neither is present

.PARAMETER vm

The name or pattern to match for the VMs to operate on

.PARAMETER selfName

If the name of the VM running the script is not the same as the NetBIOS name, specify it with this argument as long as -notSelf is not specified or -vm does not match it since it must be shutdown/paused last

.PARAMETER logFile

The name and path of a log file to append to

.PARAMETER hostShutdown

Shut down the host specified via -viServer when power operations are complete

.PARAMETER notSelf

Do not shutdown/pause the VM running this script. If the name of the VM is not its NetBIOS name use -selfName to specify the VM name

.PARAMETER shutdown

Shutdown the VMs instead of pausing them. Note that if there are outstanding Windows Updates to install, shutdown can take a long time

.EXAMPLE

& '.\Power state change running VMs.ps1" -Verbose -viServer esxi01 -logFile c:\scripts\vm.shutdown.log -notSelf -hostShutdown

Pause all VMs running on esxi01, except for the VM running this script and then shutdown the host esxi01. A log file will be written to c:\scripts\vm.shutdown.log

.EXAMPLE

& '.\Power state change running VMs.ps1" -encryptPassword

Prompt for a password, encrypt it and output it such that it can be passed in another script invocation via the -securePassword parameter

.NOTES

Save credentials first with New-VICredentialStoreItem, for the account that will run the script, if pass thru won't work and not passing username and password

If -notSelf is used with -hostShutdown, an automatic power action should be configured for that VM.

#>

[CmdletBinding()]

Param
(
    [string]$viServer ,
    [string]$username ,
    [string]$password ,
    [string]$securePassword ,
    [switch]$encryptPassword ,
    [string]$vm = '*' ,
    [string]$selfName ,
    [string]$logFile ,
    [switch]$hostShutdown ,
    [switch]$notSelf ,
    [switch]$shutdown
)

if( $encryptPassword )
{
    if( ! $PSBoundParameters[ 'password' ] -and ! ( $password = $env:_Mval12 ) )
    {
        $enteredPassword = Read-Host -Prompt "Enter password to encrypt" -AsSecureString
        if( $enteredPassword -and $enteredPassword.Length )
        {
            $enteredPassword | ConvertFrom-SecureString
        }
        else
        {
            Throw 'No password entered'
        }
    }
    else
    {
        ConvertTo-SecureString -AsPlainText -String $password -Force | ConvertFrom-SecureString
    }
    Exit 0
}

if( $PSBoundParameters[ 'LogFile' ] )
{
    Start-Transcript -Path $logFile -Append
}

Try
{
    $oldVerbosePreference = $VerbosePreference
    $VerbosePreference = 'SilentlyContinue'
    Import-Module -Name VMware.PowerCLI -Verbose:$false
    $VerbosePreference = $oldVerbosePreference

    [hashtable]$connectParams = @{}
    if( $PSBoundParameters[ 'viServer' ] )
    {
        $connectParams.Add( 'Server' , $viServer )
    }
    if( $PSBoundParameters[ 'username' ] )
    {
        $connectParams.Add( 'user' , $username )
    
        if( $PSBoundParameters[ 'password' ] -or ( $password = $env:_Mval12 ) )
        {
            $connectParams.Add( 'password' , $password )
        }
        elseif( $PSBoundParameters[ 'securePassword' ] )
        {
            $credential = New-Object -Typename System.Management.Automation.PSCredential -Argumentlist $username , (ConvertTo-SecureString -String $securePassword)
            $connectParams.Add( 'Credential' , $credential )
        }
    }

    $connection = Connect-VIServer @connectParams -Force -ErrorAction Continue
    $securePassword = $password = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
    if( ! [string]::IsNullOrEmpty( $env:_Mval12 ) )
    {
        $env:_Mval12 = $password
    }
    
    if( ! $connection )
    {
        Throw "Unable to connect to $viServer"
    }

    [int]$counter = 0
    $self = $null
    Get-VM -Name $vm | Where-Object { $_.PowerState -eq 'PoweredOn' } | ForEach-Object `
    {
        $thisVM = $_
        $counter++
        Write-Verbose -Message "$(Get-Date -Format G) : $counter : $($thisVM.Name)"
        if( ( $PSBoundParameters[ 'selfName' ] -and $selfName -eq $thisVM.name ) -or $thisVM.name -eq $env:COMPUTERNAME )
        {
            $self = $thisVM
            Write-Verbose -Message "Got self $($thisVM.Name) so deferring until the end"
        }
        else
        {
            if( $shutdown )
            {
                Shutdown-VMGuest -VM $thisVM -Confirm:$false
            }
            else
            {
                Suspend-VM -VM $thisVM -RunAsync -Confirm:$false
            }
        }
    }

    if( ! $notSelf )
    {
        if( $self )
        {
            Write-Verbose -Message "$(Get-Date -Format G): Performing power action on self"

            if( $shutdown )
            {
                Shutdown-VMGuest -VM $self -Confirm:$false
            }
            else
            {
                Suspend-VM -VM $self -RunAsync -Confirm:$false
            }
        }
        else
        {
            Write-Warning -Message "Did not find self $env:COMPUTERNAME"
        }
    }
    
    if( $hostShutdown )
    {
        Write-Verbose -Message "$(Get-Date -Format G): Shutting down host $viServer"
        Stop-VMHost -VMHost $viServer -Force -RunAsync -Confirm:$false
    }
}
Catch
{
    Throw $_
}
Finally
{
    if( $connection )
    {
        Disconnect-VIServer -Server $connection -Force -Confirm:$false
    }

    if( $PSBoundParameters[ 'LogFile' ] )
    {
        Stop-Transcript
    }
}