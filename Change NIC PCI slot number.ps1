
<#
.SYNOPSIS

Change the PCI slot number for the network adapters in specified virtual machines

.DESCRIPTION

Citrix Provisioning Services (PVS) target devices can fail to boot with a BSoD which may be due to the virtual NIC being in a different PCI slot number from the machine it was created on

.PARAMETER vcenter

The VMware Virtual Center to connect to

.PARAMETER slotNumber

The PCI slot number to set the vNICS to

.PARAMETER nicType

The type of the NIC to pick the slot number from in the source VM (use when it has multiple NICs)

.PARAMETER network

The name of the network that the NIC to pick the slot number from in the source VM is connected to (use when it has multiple NICs)

.PARAMETER vms

The name/pattern of the VMs to change

.PARAMETER port

The vCenter port to connect on if not the default

.PARAMETER protocol

The vCenter protocol to use if not the default

.PARAMETER allLinked

Connect to all linked vCenters

.PARAMETER force

Suppresses all user interface prompts during vCenter connection

.PARAMETER fromVM

The name/pattern of the VM to get the NIC PCI slot number from to apply to the VMs to change

.EXAMPLE

& '.\Change NIC PCI slot number.ps1' -fromVM GLXAPVSMASTS19 -vms GLXA19PVS40* -vcenter grl-vcenter04.guyrleech.local

Change the PCI port number in VMs matching GLXA19PVS40* to the PCI port number in the VM GLXAPVSMASTS19 using the VMware vCenter server grl-vcenter04.guyrleech.local

.EXAMPLE

& '.\Change NIC PCI slot number.ps1' -slotNumber 256 -vms GLXA19PVS40* -vcenter grl-vcenter04.guyrleech.local

Change the PCI port number in VMs matching GLXA19PVS40* to 256 using the VMware vCenter server grl-vcenter04.guyrleech.local

.EXAMPLE

& '.\Change NIC PCI slot number.ps1' -slotNumber 256 -vms GLXA19PVS40* -vcenter grl-vcenter04.guyrleech.local -network "Internal Network"

Change the PCI port number on the NIC connected to the "Internal Network" network in VMs matching GLXA19PVS40* to 256 using the VMware vCenter server grl-vcenter04.guyrleech.local
Use this when the VMs have multiple NICs

.NOTES

VMs must be powered off.

Requires VMware PowerCLI.

Modification History:

    04/07/2021 @guyrleech   Initial public release
    05/07/2021 @guyrleech   Deal with VMs to change having multiple NICs or no slot number (never booted).
                            Added -poweron option
#>

<#
Copyright © 2021 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]

Param
(
    [string]$vcenter ,
    [int]$slotNumber ,
    [string]$fromVM ,
    [ValidateSet('e1000','e1000e','vmxnet2','vmxnet3','flexible','enhancedvmxnet','SriovEthernetCard','vmxnet3','Vmxnet3Vrdma')]
    [string]$nicType ,
    [string]$network ,
    [string]$vms ,
    [switch]$powerOn ,
    [int]$port ,
    [ValidateSet('http','https')]
    [string]$protocol ,
    [switch]$allLinked ,
    [switch]$force
)

if( $PSBoundParameters.ContainsKey( 'slotNumber' ) -and $PSBoundParameters[ 'fromVM' ] )
{
    Throw "Only one of -slotNumber and -fromVM is allowed"
}

Import-Module -Name VMware.VimAutomation.Core -Verbose:$false

[hashtable]$vcenterParameters = @{ 'Server' = $vcenter ; 'AllLinked' = $allLinked ; 'Force' = $force }

if( $PSBoundParameters[ 'port' ] )
{
    $vcenterParameters.Add( 'port' , $port )
}

if( $PSBoundParameters[ 'protocol' ] )
{
    $vcenterParameters.Add( 'protocol' , $protocol )
}

if( ! ( $viconnection = Connect-VIServer @vcenterParameters ) )
{
    Throw "Unable to connect to $vcenter"
}

if( ! $PSBoundParameters[ 'slotNumber' ] )
{
    if( ! ( $sourceVM = Get-VM -Name $fromVM ) )
    {
        Throw "Unable to find source VM $fromVM"
    }
    if( $sourceVM -is [array] )
    {
        Throw "Found $($sourceVM.Count) VMs matching $fromVM"
    }
    if( ! ( $nic = Get-NetworkAdapter -VM $sourceVM ) )
    {
        Throw "VM $($sourceVM.Name) has no NICs"
    }
    if( $nic -is [array] -and $nic.Count -gt 1 )
    {
        [array]$slots = @( $nic | Select-Object -ExpandProperty ExtensionData | Select-Object -ExpandProperty SlotInfo |  Group-Object -Property PciSlotNumber )
        if( $slots.Count -gt 1 )
        {
            ## More than 1 NIC and different slots so need to filter on nic type and/or network name
            [array]$filtered = $nic
            if( $PSBoundParameters[ 'nicType' ] )
            {
                $filtered = @( $filtered.Where( { $_.Type -eq $nicType } ))
            }
            if( $PSBoundParameters[ 'network' ] )
            {
                $filtered = @( $filtered.Where( { $_.NetworkName -eq $network } ))
            }
            $slots = @( $filtered | Select-Object -ExpandProperty ExtensionData | Select-Object -ExpandProperty SlotInfo |  Group-Object -Property PciSlotNumber )
            if( $slots.Count -ne 1 )
            {
                Throw "Unable to refine NIC list to find a single slot number"
            }
        }

        if( ! $PSBoundParameters[ 'nicType' ] )
        {
            ## find the NIC for which we have the slot number so we can get its type as destination VMs probably have multiple NICs too
            if( $thisNic = $nic.Where( { $_.ExtensionData.SlotInfo.PciSlotNumber -eq $slots[0].Group.PciSlotNumber -and ( [string]::IsNullOrEmpty( $nicType ) -or $_.Type -eq $nicType ) -and ( [string]::IsNullOrEmpty( $network ) -or $_.NetworkName -eq $network ) } )  )
            {
                $nicType = $thisNic.Type
                Write-Verbose -Message "$($nic.Count) NICs found so setting type to $nicType"
            }
            else
            {
                Write-Warning -Message "Failed to find filtered NIC in $($sourceVM.Name) for slot $($slots[0].Name)"
            }
            if( $thisNIC -and ! $PSBoundParameters[ 'network' ] )
            {
                $network = $thisNIC.NetworkName
                Write-Verbose -Message "$($nic.Count) NICs found so setting network name to `"$network`""
            }
        }

        $slotNumber = $slots[0].Name
    }
    else
    {
        $slotNumber = $nic.ExtensionData.SlotInfo.PciSlotNumber
        if( ! $PSBoundParameters[ 'nicType' ] )
        {
            $nicType = $nic.type
        }
        if( ! $PSBoundParameters[ 'network' ] )
        {
            $network = $nic.NetworkName
        }
    }

    Write-Verbose -Message "Slot number is $slotNumber in $($sourceVM.Name)"
}

[array]$vmsToChange = @( Get-VM -Name $vms | Sort-Object -Property Name )

if( ! $vmsToChange -or ! $vmsToChange.Count )
{
    Throw "No VMs found matching $vms"
}

ForEach( $vmToChange in $vmsToChange )
{
    if( $nic = Get-NetworkAdapter -VM $vmToChange | Where-Object { ( [string]::IsNullOrEmpty( $network ) -or $_.NetworkName -eq $network ) -and ( [string]::IsNullOrEmpty( $nicType ) -or $_.Type -eq $nicType ) } )
    {
        if( $nic -is [array] )
        {
            Write-Warning -Message "Unable to change nic for $($vmToChange.Name) as there are $($nic.Count) - use -nictype and/or -network to be more specific"
        }
        else
        {
            [int]$existingSlotNumber = -1
            
            try
            {
                $existingSlotNumber  = $nic.ExtensionData.SlotInfo.PciSlotNumber
            }
            catch
            {
                Write-Warning -Message "Unable to get existing slot number for nic in $($vmToChange.Name)"
            }

            if( $existingSlotNumber -ne $slotNumber )
            {
                if( $vmToChange.PowerState -ne 'PoweredOff' )
                {
                    Write-Warning -Message "Cannot change $($vmToChange.Name) from $existingSlotNumber to $slotNumber because VM is not powered off, it is $($vmToChange.PowerState)"
                }
                elseif( $PSCmdlet.ShouldProcess( $vmToChange.Name , "Change $($nic.Name) ($($nic.Type)) on `"$($nic.NetworkName)`" PCI slot number from $existingSlotNumber to $slotNumber" ))
                {
                    $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
                    $device = New-Object VMware.Vim.VirtualDeviceConfigSpec

                    $device.Operation = [VMware.Vim.VirtualDeviceConfigSpecOperation]::edit
                    $device.Device = $nic.ExtensionData
                    ## if never booted then may not have slot number
                    if( $null -eq $device.Device.SlotInfo )
                    {
                        $device.Device.SlotInfo = New-Object -TypeName VMware.Vim.VirtualDevicePciBusSlotInfo
                        Write-Warning -Message "$($vmToChange.Name) did not have slot info so added"
                    }

                    $device.Device.SlotInfo.PciSlotNumber = $slotNumber

                    $spec.deviceChange = @( $device )

                    $vmToChange.ExtensionData.ReconfigVM($spec)
                    if( ! $? )
                    {
                        Write-Error "Problem changing slot from $existingSlotNumber to $slotNumber in $($vmToChange.Name)"
                    }
                    elseif( $powerOn )
                    {
                        Start-VM -VM $vmToChange
                    }
                }
            }
            else
            {
                Write-Warning "NIC in $($vmToChange.Name) is already in slot $slotNumber"
            }
        }
    }
    else
    {
        [string]$message = "No NIC found in $($vmToChange.Name)"
        if( $network )
        {
            $message += " on network $network"
        }
        Write-Warning -Message $message
    }
}

# SIG # Begin signature block
# MIINRQYJKoZIhvcNAQcCoIINNjCCDTICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU8f9s/eZEep70mcyeHSF2RQEw
# GhqgggqHMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
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
# hH44QHzE1NPeC+1UjTGCAigwggIkAgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EC
# EAT946rb3bWrnkH02dUhdU4wCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFPT9aochjwS743ZsUWX8
# kH1lWhwAMA0GCSqGSIb3DQEBAQUABIIBADOUL0Iy6hYrhqR+3SOHfOJAUu3s/iF1
# hkedlzB0Y7a7Dc77j90BAyV+X1Uc7muyq/ABpFdrrmrg1tNs1iqJ8gehoh+ai2Ig
# 4Po9xdZtUjpyCG7olf3pozpyPJOdRxHyqTJqmaLc8z9MfQ2EImMgVXboHgblP2nX
# jHM8qVBFQsjgwburRxLOIlonCqUtWloqH3TqZnjvsUcwE4LZ5YgjECU/ckSacEpx
# d4RNuxwNfuar4gP+eMSXEoGspuSNIDlkJoco4TRqFRuGF4PjINcQoi0qsDmAegOE
# QIL3/DXmFbpMdX+H1eYZMJB28DayVDBCuyB6HO0s4yFZZeucSx1Qu9k=
# SIG # End signature block
