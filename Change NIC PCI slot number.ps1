
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
                            Added -poweron
    09/12/2021 @guyrleech   Added -credential
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
    [ValidateSet('e1000','e1000e','vmxnet2','vmxnet3','flexible','enhancedvmxnet','SriovEthernetCard','Vmxnet3Vrdma')]
    [string]$nicType ,
    [string]$network ,
    [string]$vms ,
    [switch]$powerOn ,
    [pscredential]$credential ,
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

if( $PSBoundParameters[ 'credential' ] )
{
    $vcenterParameters.Add( 'credential' , $credential )
}

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
# MIIZsAYJKoZIhvcNAQcCoIIZoTCCGZ0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUJ8T9vrdWO07o63eW3itopnTz
# 4yWgghS+MIIE/jCCA+agAwIBAgIQDUJK4L46iP9gQCHOFADw3TANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgVGltZXN0YW1waW5nIENBMB4XDTIxMDEwMTAwMDAwMFoXDTMxMDEw
# NjAwMDAwMFowSDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAMLmYYRnxYr1DQikRcpja1HXOhFCvQp1dU2UtAxQ
# tSYQ/h3Ib5FrDJbnGlxI70Tlv5thzRWRYlq4/2cLnGP9NmqB+in43Stwhd4CGPN4
# bbx9+cdtCT2+anaH6Yq9+IRdHnbJ5MZ2djpT0dHTWjaPxqPhLxs6t2HWc+xObTOK
# fF1FLUuxUOZBOjdWhtyTI433UCXoZObd048vV7WHIOsOjizVI9r0TXhG4wODMSlK
# XAwxikqMiMX3MFr5FK8VX2xDSQn9JiNT9o1j6BqrW7EdMMKbaYK02/xWVLwfoYer
# vnpbCiAvSwnJlaeNsvrWY4tOpXIc7p96AXP4Gdb+DUmEvQECAwEAAaOCAbgwggG0
# MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMEEGA1UdIAQ6MDgwNgYJYIZIAYb9bAcBMCkwJwYIKwYBBQUHAgEWG2h0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAfBgNVHSMEGDAWgBT0tuEgHf4prtLk
# YaWyoiWyyBc1bjAdBgNVHQ4EFgQUNkSGjqS6sGa+vCgtHUQ23eNqerwwcQYDVR0f
# BGowaDAyoDCgLoYsaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJl
# ZC10cy5jcmwwMqAwoC6GLGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtdHMuY3JsMIGFBggrBgEFBQcBAQR5MHcwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBPBggrBgEFBQcwAoZDaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRFRpbWVzdGFtcGluZ0NB
# LmNydDANBgkqhkiG9w0BAQsFAAOCAQEASBzctemaI7znGucgDo5nRv1CclF0CiNH
# o6uS0iXEcFm+FKDlJ4GlTRQVGQd58NEEw4bZO73+RAJmTe1ppA/2uHDPYuj1UUp4
# eTZ6J7fz51Kfk6ftQ55757TdQSKJ+4eiRgNO/PT+t2R3Y18jUmmDgvoaU+2QzI2h
# F3MN9PNlOXBL85zWenvaDLw9MtAby/Vh/HUIAHa8gQ74wOFcz8QRcucbZEnYIpp1
# FUL1LTI4gdr0YKK6tFL7XOBhJCVPst/JKahzQ1HavWPWH1ub9y4bTxMd90oNcX6X
# t/Q/hOvB46NJofrOp79Wz7pZdmGJX36ntI5nePk2mOHLKNpbh6aKLzCCBTAwggQY
# oAMCAQICEAQJGBtf1btmdVNDtW+VUAgwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4X
# DTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAPjTsxx/DhGvZ3cH0wsx
# SRnP0PtFmbE620T1f+Wondsy13Hqdp0FLreP+pJDwKX5idQ3Gde2qvCchqXYJawO
# eSg6funRZ9PG+yknx9N7I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrskacLCUvIUZ4qJ
# RdQtoaPpiCwgla4cSocI3wz14k1gGL6qxLKucDFmM3E+rHCiq85/6XzLkqHlOzEc
# z+ryCuRXu0q16XTmK/5sy350OTYNkO/ktU6kqepqCquE86xnTrXE94zRICUj6whk
# PlKWwfIPEvTFjg/BougsUfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8np+mM6n9Gd8l
# k9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQD
# AgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0wazAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsME8GA1UdIARI
# MEYwOAYKYIZIAYb9bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdp
# Y2VydC5jb20vQ1BTMAoGCGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7KgqjpepxA8Bg
# +S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG
# 9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh134LYP3DPQ/E
# r4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63XX0R58zYUBor3
# nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbHJyqhKSgaOnEoAjwukaPAJRHinBRHoXpo
# aK+bp1wgXNlxsQyPu6j4xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC/i9yfhzXSUWW
# 6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG/AeB+ova+YJJ
# 92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCCBTEwggQZoAMCAQICEAqhJdbW
# Mht+QeQF2jaXwhUwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIG
# A1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTE2MDEwNzEyMDAw
# MFoXDTMxMDEwNzEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGln
# aUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBDQTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAL3QMu5LzY9/3am6gpnFOVQoV7YjSsQOB0Uz
# URB90Pl9TWh+57ag9I2ziOSXv2MhkJi/E7xX08PhfgjWahQAOPcuHjvuzKb2Mln+
# X2U/4Jvr40ZHBhpVfgsnfsCi9aDg3iI/Dv9+lfvzo7oiPhisEeTwmQNtO4V8CdPu
# XciaC1TjqAlxa+DPIhAPdc9xck4Krd9AOly3UeGheRTGTSQjMF287DxgaqwvB8z9
# 8OpH2YhQXv1mblZhJymJhFHmgudGUP2UKiyn5HU+upgPhH+fMRTWrdXyZMt7HgXQ
# hBlyF/EXBu89zdZN7wZC/aJTKk+FHcQdPK/P2qwQ9d2srOlW/5MCAwEAAaOCAc4w
# ggHKMB0GA1UdDgQWBBT0tuEgHf4prtLkYaWyoiWyyBc1bjAfBgNVHSMEGDAWgBRF
# 66Kv9JLLgjEtUYunpyGd823IDzASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB
# /wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB5BggrBgEFBQcBAQRtMGswJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBQBgNV
# HSAESTBHMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cu
# ZGlnaWNlcnQuY29tL0NQUzALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggEB
# AHGVEulRh1Zpze/d2nyqY3qzeM8GN0CE70uEv8rPAwL9xafDDiBCLK938ysfDCFa
# KrcFNB1qrpn4J6JmvwmqYN92pDqTD/iy0dh8GWLoXoIlHsS6HHssIeLWWywUNUME
# aLLbdQLgcseY1jxk5R9IEBhfiThhTWJGJIdjjJFSLK8pieV4H9YLFKWA1xJHcLN1
# 1ZOFk362kmf7U2GJqPVrlsD0WGkNfMgBsbkodbeZY4UijGHKeZR+WfyMD+NvtQEm
# tmyl7odRIeRYYJu6DC0rbaLEfrvEJStHAgh8Sa4TtuF8QkIoxhhWz0E0tmZdtnR7
# 9VYzIi8iNrJLokqV2PWmjlIwggVPMIIEN6ADAgECAhAE/eOq2921q55B9NnVIXVO
# MA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
# ZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwHhcNMjAwNzIwMDAw
# MDAwWhcNMjMwNzI1MTIwMDAwWjCBizELMAkGA1UEBhMCR0IxEjAQBgNVBAcTCVdh
# a2VmaWVsZDEmMCQGA1UEChMdU2VjdXJlIFBsYXRmb3JtIFNvbHV0aW9ucyBMdGQx
# GDAWBgNVBAsTD1NjcmlwdGluZ0hlYXZlbjEmMCQGA1UEAxMdU2VjdXJlIFBsYXRm
# b3JtIFNvbHV0aW9ucyBMdGQwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQCvbSdd1oAAu9rTtdnKSlGWKPF8g+RNRAUDFCBdNbYbklzVhB8hiMh48LqhoP7d
# lzZY3YmuxztuPlB7k2PhAccd/eOikvKDyNeXsSa3WaXLNSu3KChDVekEFee/vR29
# mJuujp1eYrz8zfvDmkQCP/r34Bgzsg4XPYKtMitCO/CMQtI6Rnaj7P6Kp9rH1nVO
# /zb7KD2IMedTFlaFqIReT0EVG/1ZizOpNdBMSG/x+ZQjZplfjyyjiYmE0a7tWnVM
# Z4KKTUb3n1CTuwWHfK9G6CNjQghcFe4D4tFPTTKOSAx7xegN1oGgifnLdmtDtsJU
# OOhOtyf9Kp8e+EQQyPVrV/TNAgMBAAGjggHFMIIBwTAfBgNVHSMEGDAWgBRaxLl7
# KgqjpepxA8Bg+S32ZXUOWDAdBgNVHQ4EFgQUTXqi+WoiTm5fYlDLqiDQ4I+uyckw
# DgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4w
# NaAzoDGGL2h0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3Mt
# ZzEuY3JsMDWgM6Axhi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1
# cmVkLWNzLWcxLmNybDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgGCCsGAQUF
# BwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYI
# KwYBBQUHAQEEeDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wTgYIKwYBBQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFNIQTJBc3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAA
# MA0GCSqGSIb3DQEBCwUAA4IBAQBT3M71SlOQ8vwM2txshp/XDvfoKBYHkpFCyanW
# aFdsYQJQIKk4LOVgUJJ6LAf0xPSN7dZpjFaoilQy8Ajyd0U9UOnlEX4gk2J+z5i4
# sFxK/W2KU1j6R9rY5LbScWtsV+X1BtHihpzPywGGE5eth5Q5TixMdI9CN3eWnKGF
# kY13cI69zZyyTnkkb+HaFHZ8r6binvOyzMr69+oRf0Bv/uBgyBKjrmGEUxJZy+00
# 7fbmYDEclgnWT1cRROarzbxmZ8R7Iyor0WU3nKRgkxan+8rzDhzpZdtgIFdYvjeO
# c/IpPi2mI6NY4jqDXwkx1TEIbjUdrCmEfjhAfMTU094L7VSNMYIEXDCCBFgCAQEw
# gYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1
# cmVkIElEIENvZGUgU2lnbmluZyBDQQIQBP3jqtvdtaueQfTZ1SF1TjAJBgUrDgMC
# GgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYK
# KwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG
# 9w0BCQQxFgQUDbYn/VvMWhiPGNIcFtochC5jrdwwDQYJKoZIhvcNAQEBBQAEggEA
# kHzsaDOzWlqst2Dml6T3RQdgl5sbcgx9POwMvt3R1IgUEND0aJpNQqK6Alb35aMr
# zZWPJYo0qdvt4pDGE36J8sw2DaPJ+w9mlFZYYh/Vjksdv1tl/0dHs3FFsrJ1NBcP
# BCyWGTl5qD0pB2L8DSUp8CzHsjZScEYOqD1CskRJu/PvS95bF4YQaswNLPogtZMC
# OYMQFEMHY1OrqKytSHxDkKVG9JxNMUIsxY9RbkFkRk8zAH3y0pXGXHZwbg7y11wh
# CLv8XXfNE7sgIQqmn9+ZgVzrHYppMhLN/Q6julCo4vFiJ/eE0MswCndXWBYZeJGh
# UrMnT7RfVQmakNrQoCmI+KGCAjAwggIsBgkqhkiG9w0BCQYxggIdMIICGQIBATCB
# hjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3Vy
# ZWQgSUQgVGltZXN0YW1waW5nIENBAhANQkrgvjqI/2BAIc4UAPDdMA0GCWCGSAFl
# AwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMjExMjA5MTYzNDQ2WjAvBgkqhkiG9w0BCQQxIgQg/j3TrX9wk0nUu+cs10Tt
# rFjBj8m+KHxl44asEmhSjW4wDQYJKoZIhvcNAQEBBQAEggEAk9TvRz9rQXMGO17h
# D0JeirmYW5Fs5U5SwZ9OcoJzXx/82ZrDuEQAzKVwb/GvKlzBRofhbj2OSMZhNY8V
# Ouqs/D4rt6ucxmlsWZZr5o/0tR+LG9r4Jk8CTE+uABjuZFpTuyO2hj6oounJiOfh
# x2091iyTZ4f0KGMNva99bc23nOyOW3gFnKRzfx2As4MlShtZvqkGFogRQNOec525
# TD8h918RCc6xzHqR51OSoyNcOHplaFNMfddWB1zc4xTjxa/K8CjylOkdMC80AAsI
# jXMFW1RF5LTHAqBp80taW/28OhS6f8YHTSBMTThJSeumeI1w+xuLSZ7O8+0SdTBN
# U0VL0A==
# SIG # End signature block
