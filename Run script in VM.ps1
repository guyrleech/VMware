#requires -version 3

<#
.SYNOPSIS
    Copy a script to a remote machine via Invoke-VMScript, as doesn't require explicitly specified credentials, and run it.

.DESCRIPTION
    Works around an issue that limits PowerShell script length to 8K because VMware run powershell.exe via cmd.exe

.PARAMETER scriptFile
    The full path to the PowerShell script file which will be run in the VM(s) specified.
    Can be local or remote to the machine running the script

.PARAMETER VMs
    The powered on VMware VM(s) to operate on. VMware Tools must be running.

.PARAMETER copyTo
    A share on the VM(s) to copy the script to directly rather than base64 encoding the script and copying via Invoke-VMscript (which can be slow).
    If the folder does not exist, it, and parents, will be created

.PARAMETER scriptParameters
    Parameters to pass when running the script specified by -scriptFile

.PARAMETER vCenter
    The VMware vCenter(s) to connect to

.PARAMETER chunkSizeBytes
    The maximum amount of data to send in each part of the base64 encoded script file.
    Values larger than the default may fail

.PARAMETER quitOnError
    Quit the script immediately on any errors encoutered. If not specified the script will move on to the next VM, if there is one

.PARAMETER port
    The vCenter port to connect to. Only specify if non-standard

.PARAMETER protocol
    The protocol to use to connect to vCenter

.PARAMETER forceVcenter
    Suppress any prompts when connecting to vCenter

.PARAMETER allLinked
    Connect to all vCenters linked with the ones specified by -vCennter

.PARAMETER OperationTimeoutSec
    Timeout in seconds of the WMI/CIM operation used to get the remote shares when -copyTo specified

.PARAMETER CIMprotocol
    The CIM protocol to use in the WMI/CIM operation used to get the remote shares when -copyTo specified

.EXAMPLE
    . '.\Run script in VM.ps1" -scriptFile "C:\Scripts\Remvoe Ghost NICS.ps1" -VMs GLXAPVSMASTS19 -scriptParameters '-nicregex "Intel.*Gigabit" -confirm:$false' -vCenter grl-vcenter04.guyrleech.local
    Copy the script specified to the VM specified using Invoke-VMscript. When copied, run it with the specified parameters

.EXAMPLE
    . '.\Run script in VM.ps1" -scriptFile "C:\Scripts\Remvoe Ghost NICS.ps1" -VMs GLXAPVSMASTS19 -scriptParameters '-nicregex "Intel.*Gigabit" -confirm:$false' -vCenter grl-vcenter04.guyrleech.local -copyTo 'D$\temp\guy' -Confirm:$false
    Copy the script specified to the VM via its D$ share. When copied, run it with the specified parameters

.NOTES
    Invoke-VMscript can be slow so use the -CopyTo parameter if the account running the script has write access to C$ share or another local share in each VM

    Modification History:

    2021/08/23  @guyrleech  Initial release
    2021/08/23  @guyrleech  Changed Get-CimInstance to use Dcom to avoid using PS remoting by default since if PS remoting is enabled, it would be better to use that rather than this script
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
    [Parameter(Mandatory=$true,HelpMessage='Script to copy to VM and run')]
    [string]$scriptFile ,
    [Parameter(Mandatory=$true,HelpMessage='The VM(s) to operate on')]
    [string[]]$VMs ,
    [string]$copyTo ,
    [string]$scriptParameters ,
    [string[]]$vCenter ,
    [int]$chunkSizeBytes = 8000 ,
    [switch]$quitOnError ,
    [int]$port ,
    [ValidateSet('http','https')]
    [string]$protocol ,
    [switch]$forceVcenter ,
    [switch]$allLinked ,
    [ValidateSet('dcom','wsman','default')]
    [string]$CIMprotocol = 'dcom' ,
    [int]$OperationTimeoutSec = 15
)

Import-Module -Name VMware.VimAutomation.Core -Verbose:$false

if( $PSBoundParameters[ 'vCenter' ] )
{
    [hashtable]$vcenterParameters = @{ 'Server' = $vcenter ; 'AllLinked' = $allLinked ; 'Force' = $forceVcenter }

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
}

## read whole file and base 64 encode
[byte[]]$data = $null
[string]$base64encoded = $null
[string]$share = $null 
if( -Not $PSBoundParameters[ 'copyTo' ] )
{
    $data = [System.IO.File]::ReadAllBytes( $scriptFile )
    if( ! $data -or ! $data.Count )
    {
        Throw "Failed to get any data from file $scriptFile"
    }

    $base64encoded = [System.Convert]::ToBase64String( $data )
    Write-Verbose -Message "Base64 data length is $($base64encoded.Length)"
}
elseif( -Not ( Test-Path -Path $scriptFile -PathType Leaf ) )
{
    Throw "Unable to access script file $scriptFile"
}
else
{
    [int]$skip = 0
    For( [int]$index = 0 ; $index -lt $copyTo.Length ; $index++ )
    {
        if( $copyTo[$index] -eq '\' )
        {
            $skip++
        }
        else
        {
            break
        }
    }
    $share = $copyTo -split '\\' | Select-Object -First 1 -Skip $skip
    if( [String]::IsNullOrEmpty( $share ) )
    {
        Throw "Failed to get share from $copyTo"
    }
}

## Create a temporary file and copy script to it
[scriptblock]$scriptBlock = [scriptblock]::Create(
{
    if( $tempFile = New-TemporaryFile )
    {
        [string]$destination = "$tempfile.txt"
        if( Move-Item -Path $tempFile -Destination $destination -PassThru )
        {
            $destination
        }
    }
})

if( ! ( $CIMsessionOption = New-CimSessionOption -Protocol $CIMprotocol ) )
{
    Throw "Failed to create CIM session option with protocol $CIMprotocol"
}

[string]$errorMessage = $null

## may have been flattened if called bu scheduled task
if( $VMs.Count -eq 1 -and $VMs[0].IndexOf( ',' ) -ge 0 )
{
    $VMs = $VMs -split ','
}

ForEach( $virtualMachine in $VMs )
{
    if( -Not ( $vmObjects = @( Get-VM -Name $virtualMachine ) ) )
    {
        $errorMessage = "Failed to get vm $virtualMachine"
        if( $quitOnError )
        {
            Throw $errorMessage
        }
        else
        {
            Write-Error -Message $errorMessage 
            continue
        }
    }

    ## iterate in case an array, eg had * in name
    ForEach( $VM in $vmObjects )
    {
        [string]$decodeScript = $null
        if( $VM.PowerState -ne 'PoweredOn' )
        {
            Write-Warning -Message "Unable to operate on $($VM.Name) as power state is $($VM.powerstate)"
        }
        elseif( $PSCmdlet.ShouldProcess( $VM.Name , "Run script" ) )
        {
            if( -Not $VM.extensiondata.guest -or $VM.extensiondata.guest.ToolsRunningStatus -ne 'guestToolsRunning' )
            {
                Write-Warning -Message "VMware Tools appear not to be running in $($VM.Name)"
            }
            ## if we have a location to copy to, we'll try that as may be running with credentials that allow that
            if( $PSBoundParameters[ 'copyTo' ] )
            {
                [string]$destination = Join-Path -Path "\\$($VM.Name)" -ChildPath $copyTo
                Write-Verbose -Message "Destination for script copy is $destination"
                $testPathError = $null
                ## finding that operations on non-existent shares can take upwards of 45 seconds each so check share exists if we can
                $CIMError = $null
                $errorMessage = $null

                Write-Verbose -Message "Trying to get file shares from $($VM.Name) to check for $share"
                if( ( $cimSession = New-CimSession -ComputerName $VM.Name -SessionOption $CIMsessionOption -ErrorAction SilentlyContinue ) `
                    -and ( [array]$shares = @( Get-CimInstance -ClassName win32_share -ComputerName $VM.Name -ErrorAction SilentlyContinue -OperationTimeoutSec $OperationTimeoutSec -Filter 'Path IS NOT NULL' -QueryDialect WQL -ErrorVariable CIMerror | Where-Object { $_.Path -match '^[A-Z]:\\' } ) ) `
                        -and $shares.Count -gt 0 )
                {
                    Write-Verbose -Message "Checking if share $share exists on $($VM.Name) - got $($shares.Count) file shares"
                    if( -Not ( $ourShare = $shares.Where( { $_.Name -eq $share } ) ) )
                    {
                        $errorMessage = "Unable to find share $share on $($VM.Name) - shares are $(($shares | Select-Object -ExpandProperty Name) -join ',')"
                    }
                }
                elseif( $null -eq $CIMError -or $CIMError.Count -eq 0 ) ## no error so no suitable shares found
                {
                    $errorMessage = "No file shares found on $($VM.Name) so cannot use $share"
                }

                if( $cimSession )
                {
                    Remove-CimSession -CimSession $cimSession
                    $cimSession = $null
                }

                if( $errorMessage )
                {
                    if( $quitOnError )
                    {
                        Throw $errorMessage
                    }
                    else
                    {
                        Write-Error -Message $errorMessage 
                        continue
                    }
                }
                if( -Not ( Test-Path -Path $destination -ErrorAction SilentlyContinue -ErrorVariable testPathError ) )
                {
                    Write-Verbose -Message "Creating folder $destination in $($VM.Name)"
                    if( -Not ( $newFolder = New-Item -Path $destination -Force -ItemType Directory ) )
                    {
                        Write-Warning -Message "Problem creating folder $newFolder"
                    }
                }
                Write-Verbose -Message "Copying script to $destination"
                if( -Not ( Copy-Item -Path $scriptFile -Destination $destination -Container -Force -PassThru ) )
                {
                    $errorMessage = "Problem copying script to $destination"
                    if( $quitOnError )
                    {
                        Throw $errorMessage
                    }
                    else
                    {
                        Write-Error -Message $errorMessage 
                        continue
                    }
                }
                else
                {
                    ## remove any leading \ characters since it must be a local share
                    [string]$remoteScriptFile = Join-Path -Path ($copyTo -replace '^\\+' -replace '\$' , ':') -ChildPath (Split-Path -Path $scriptFile -Leaf)
                    Write-Verbose -Message "Remote script path in $($VM.Name) is $remoteScriptFile"
                    $decodeScript = @"
                       & `"$remoteScriptFile`" $scriptParameters
                       Remove-Item -Path `"$remoteScriptFile`"
"@
                }
            }
            else
            {
                if( -Not ( $result = Invoke-VMScript -VM $VM -ScriptType Powershell -ScriptText $scriptBlock ) -or [string]::IsNullOrEmpty( $result.ScriptOutput ) )
                {
                    $errorMessage = "Failed to create temporary file on $($VM.Name)"
                    if( $quitOnError )
                    {
                        Throw $errorMessage
                    }
                    else
                    {
                        Write-Error -Message $errorMessage 
                        continue
                    }
                }

                [string]$remoteTempFile = $result.ScriptOutput.Trim()

                Write-Verbose -Message "Remote temp file is $remoteTempFile"

                [int]$chunk = 0
                [int]$chunkStart = 0
                [int]$remaining = -1
                [int]$thisChunkSize = $chunkSizeBytes

                ## read chunks of base64 encoded string , send to VM, decode and append to results file
                do
                {
                    $chunk++
                    ## if last chunk ensure not too big
                    $remaining = $base64encoded.Length - $chunkStart
                    if( $remaining -lt $chunkSizeBytes )
                    {
                        $thisChunkSize = $remaining
                    }
                    if( $thisChunkSize -gt 0 )
                    {
                        Write-Verbose -Message "Chunk $chunk : offset $chunkStart size $thisChunkSize length $($base64encoded.Length)"
                        [string]$thisChunk = $base64encoded.Substring( $chunkStart , $thisChunkSize )
                        Write-Verbose -Message "`tchunk string length $($thisChunk.Length)"
        
                        if( ! ( $result = Invoke-VMScript -VM $VM -ScriptType Bat -ScriptText "echo | set /p=`"$thisChunk`" >> $remoteTempFile" ) ) ## don't check exit code as can be non-zero when worked
                        {
                            $errorMessage = "Failed to copy $($thisChunk.Length) bytes at offset $chunkStart to $remoteTempFile on $($VM.Name)"
                            if( $quitOnError )
                            {
                                Throw $errorMessage
                            }
                            else
                            {
                                Write-Error -Message $errorMessage 
                                continue
                            }
                        }

                        $chunkStart += $chunkSizeBytes
                    }
                } while( $remaining -gt 0 )

                [string]$remoteScriptFile = $remoteTempFile -replace "\.\w*$" , '.ps1' ## change .txt extension to .ps1

                Write-Verbose -Message "Will decode $remoteTempFile to $remoteScriptFile"

                ## from https://github.com/guyrleech/Microsoft/blob/master/Bincoder%20GUI.ps1
                $decodeScript = @"
                    [byte[]]`$transmogrified = [System.Convert]::FromBase64String( (Get-Content -Path `"$remoteTempFile`" ))
                    if( `$transmogrified.Count )
                    {
                        if( `$fileStream = New-Object System.IO.FileStream( `"$remoteScriptFile`" , [System.IO.FileMode]::Create , [System.IO.FileAccess]::Write ) )
                        {
                            `$fileStream.Write( `$transmogrified , 0 , `$transmogrified.Count )
                            `$fileStream.Close()

                            Remove-Item -Path `"$remoteTempFile`"
                            & `"$remoteScriptFile`" $scriptParameters
                            Remove-Item -Path `"$remoteScriptFile`"
                        }
                        else
                        {
                            Throw "Failed to create $remoteScriptFile"
                        }
                    }
                    else
                    {
                        Throw "No data retrieved from $remoteTempFile"
                    }
"@
            }

            ## decode and run the file 
            if( ! ( $result = Invoke-VMScript -VM $VM -ScriptType Powershell -ScriptText $decodeScript ) -or $result.ExitCode -ne 0)
            {
                $errorMessage = "Failed to convert base64 data or execute script in $($VM.Name)"
                if( $result )
                {
                    $errorMessage += " exit code $($result.ExitCode) : $($result.ScriptOutput)"
                }
                if( $quitOnError )
                {
                    Throw $errorMessage
                }
                else
                {
                    Write-Error -Message $errorMessage 
                    continue
                }
            }
            else
            {
                Write-Verbose -Message "Good result from running script in $($VM.Name)"

                $result.ScriptOutput
            }
        }
    }
}

# SIG # Begin signature block
# MIINRQYJKoZIhvcNAQcCoIINNjCCDTICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUw/jmdMBa9BOMXHgMhCe6W4jS
# OQOgggqHMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
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
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFDaX3bWbbyqVWaVKQDA0
# kG0dosbUMA0GCSqGSIb3DQEBAQUABIIBADbjG1eJ1q38b7qYOq3lVwizWN4/9Wla
# Dyhgzf5uq09rvzAqORGR9pj6jF+7Klzjy1Tbq+XyB75fUEjB8+hEYtHe9adwdEzG
# KYoeyxnd0NJxmSD20qX8PSZNTAHd1XK1AfaOPLYkkhfw5iZ6W/ZZivQ4D8k0xLsf
# CTvxJKFKQsHmbADgKRsHubXs+lwkbqpb6ErVb1GLHM6GThKewg6Rf7T3NldsixsE
# VXDkBoaCNFU3AAlywTQdP8YPHbATcKzGDvoEBGQogsgC2CC2JSd+/U14nKarwRV/
# 6cz+vzGSyjotjSjVAxuWsMKsFUf+9rwlHYkmuuPEt6UXzgCvstBxNJ8=
# SIG # End signature block
