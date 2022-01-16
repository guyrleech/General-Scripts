#requires -version 3

<#
.SYNOPSIS
    Run smbstatus to get file locks from Synology NAS and parse the result into objects

.DESCRIPTION
    Help to find what user and machine files are locked from, e.g. FSlogix profile disks

.PARAMETER connection
    The username and host to connect to
    
.PARAMETER port
    The SSH port to connect to
    
.PARAMETER sshOptions
    Options to pass to ssh.exe

.PARAMETER command
    The command to run
    
.PARAMETER dividerPattern
    Regex for delimiter between different sections of output

.EXAMPLE
    & '.\Parse smbstatus.ps1' -connection root@grl-nas02

    Connect to the Synology device grl-nas02 as user root, run the smbstatus command and parse the output into objects

.NOTES
    For passwordless ssh connections, setup keys - https://kb.synology.com/en-uk/DSM/tutorial/How_to_log_in_to_DSM_with_key_pairs_as_admin_or_root_permission_via_SSH_on_computers - but protect/secure those keys

    smbstatus must be run as root

    Modification History:

    @guyrleech 2022/01/16 Initial version
#>

<#
Copyright © 2022 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage='SSH host')]
    [string]$connection ,
    [int]$port ,
    [string]$sshOptions = '-o ConnectTimeout=15 -o batchmode=yes' ,
    [string]$command = 'smbstatus' ,
    [string]$dividerPattern = '^Locked files:'
)

[string[]]$fieldHeaders = @(
    'Pid' ,
    'Id' ,
    'user' ,
    'Address' ,
    'Share' ,
    'Version'
)

[string]$portparameter = $null

if( $port -gt 0 )
{
    $portparameter = "-p $port"
}

## first half is shares (without headers), second half is locked files
<#
30706   1642246968      GUYRLEECH\admingle      192.168.0.133   homes   SMB3_11
30706   1642247068      GUYRLEECH\admingle      192.168.0.133   video   SMB3_11
29841   1642257276      GUYRLEECH\admingle      192.168.0.197   homes   SMB3_11
11271   1642246118      GUYRLEECH\admingle      10.1.1.148      Software        SMB3_11
11271   1642246117      GUYRLEECH\admingle      10.1.1.148      homes   SMB3_11
19321   1642230239      GUYRLEECH\billybob      10.1.1.151      Software        SMB3_11
30706   1642237811      GUYRLEECH\admingle      192.168.0.133   Software        SMB3_11
29841   1642257815      GUYRLEECH\admingle      192.168.0.197   Software        SMB3_11

Locked files:
Pid          Uid        DenyMode   Access      R/W        Oplock           SharePath   Name   Time
--------------------------------------------------------------------------------------------------
30706        0          DENY_NONE  0x100080    RDONLY     NONE             /volume1/homes   .   Sat Jan 15 11:45:03 2022
19321        671089749  DENY_NONE  0x100081    RDONLY     NONE             /volume1/Software   .   Sat Jan 15 07:03:59 2022
11271        0          DENY_NONE  0x100081    RDONLY     NONE             /volume1/Software   .   Sat Jan 15 11:28:37 2022
30706        0          DENY_NONE  0x100080    RDONLY     NONE             /volume1/Software   .   Sat Jan 15 11:56:00 2022
19321        671089749  DENY_NONE  0x120089    RDONLY     LEASE(R)         /volume1/Software   FSLogix/S-1-5-21-1721611859-3364803896-2099701507-1109_BillyBob/Profile_BillyBob.VHDX   Sat Jan 15 07:03:59 2022
19321        671089749  DENY_WRITE 0x12019f    RDWR       LEASE(R)         /volume1/Software   FSLogix/S-1-5-21-1721611859-3364803896-2099701507-1109_BillyBob/Profile_BillyBob.VHDX   Sat Jan 15 07:03:59 2022
#>

[bool]$gotLockedFilesLine = $false
$headers = New-Object -TypeName System.Collections.Generic.List[string]
$pids    = New-Object -TypeName System.Collections.Generic.List[object]

## found that using & to invoke the command put quotes around some arguments which breaks ssh.exe
Invoke-Expression -Command "ssh.exe $sshoptions $connection $portparameter $command" | ForEach-Object `
{
    [string]$line = $_.Trim()
    if( $line.Length )
    {
        if( $gotLockedFilesLine )
        {
            if( $headers.Count )
            {
                ## fields are delimited by two or more spaces but need to be able to deal with file names with 2 or more consecutive spaces
                ## 19321        671089749  DENY_NONE  0x120089    RDONLY     LEASE(R)         /volume1/Software   FSLogix/S-1-5-21-1721611859-3364803896-2099701507-1109_BillyBob/Profile_BillyBob.VHDX   Sat Jan 15 07:03:59 2022
                [string[]]$lineparts = $line -split '\s+'
                Write-Verbose -Message $line
                if( $lineparts -and $lineparts.Count -ge $headers.Count )
                {
                    [hashtable]$item = @{}
                    $time = $null
                    [string]$regex = $null 
                    For( [int]$index = 0 ; $index -lt $headers.Count ; $index++ )
                    {
                        $property = $lineparts[ $index ]
                        if( $headers[ $index ] -eq 'pid' )
                        {
                            if( $pidentry = $pids.Where( { $_.pid -eq $property } , 1 ) ) ## could be multiple entries
                            {
                                $item.Add( 'User' , $pidentry.User )
                                $item.Add( 'Address' , $pidentry.Address )
                            }
                            else
                            {
                                Write-Warning -Message "No entry found for pid $property on line : $line"
                            }
                        }
                        elseif( $headers[ $index ] -eq 'name' )
                        {
                            ## file name could have spaces so we grab the date off the end of the line which allows us to then work out where the file name is
                            if( $line -match '  (\w{3} \w{3} \d* \d{2}:\d{2}:\d{2} \d{4})$' )
                            {
                                $time = New-Object -TypeName datetime
                                if( [datetime]::TryParseExact( $matches[1] , 'ddd MMM d HH:mm:ss yyyy' , [System.Globalization.CultureInfo]::InvariantCulture , [System.Globalization.DateTimeStyles]::AllowWhiteSpaces , [ref]$time ) )
                                {
                                    $item.Add( 'Time' , $time )
                                    ## we have been building $regex to match the line before the file name so we can match the line up to this and calculate where the file name starts
                                    [int]$datematchLength = $matches[0].Length
                                    if( $line -match $regex )
                                    {
                                        [int]$filenameLength = $line.Length - $matches[0].Length - $datematchLength - 1
                                        if( $filenameLength -gt 0 )
                                        {
                                            $property = $line.Substring( $matches[0].Length , $filenameLength )
                                        }
                                        else
                                        {
                                            $property = $null
                                        }
                                    }
                                    else
                                    {
                                        Write-Warning -Message "Failed to match regex `"$regex`" in line $line"
                                    }
                                    ## else no space so file name is ok as is
                                }
                                else
                                {
                                    Write-Warning -Message "Failed to parse date/time $property"
                                }
                            }
                            else
                            {
                                Write-Warning -Message "Failed to parse date from end of line : $line"
                            }
                        }
                        $item.Add( $headers[ $index ] ,  $property )
                        if( $time )
                        {
                            break ## name is followed only by date which we have parsed
                        }
                        $regex = -join ( $regex , [regex]::Escape( $property ) , '\s+' )
                    }
                    [pscustomobject]$item
                }
                elseif( $line -notmatch '^----' ) 
                {
                    Write-Warning -Message "Failed to parse : $line"
                }
            }
            else ## need to read the headers
            {
                $headers = @( $line -split '\s+' )
            }   
        }
        elseif( $line -match $dividerPattern )
        {
            $gotLockedFilesLine = $true
        }
        else ## first part of the output is tab delimited, without headers
        {
            if( ( $fields = $line -split '\s+' ) -and $fields.Count )
            {
                if( $fields.Count -eq $fieldHeaders.Count )
                {
                    [hashtable]$item = @{}
                    For( [int]$index = 0 ; $index -lt $fieldHeaders.Count ; $index++ )
                    {
                        $item.Add( $fieldHeaders[ $index ] , $fields[ $index ] )
                    }
                    $pids.Add( [pscustomobject]$item )
                }
                else
                {
                    Write-Warning -Message "Got $($fields.Count) fields, not $($fieldHeaders.Count) : $line"
                }
            }
            else
            {
                Write-Warning -Message "Unable to process line : $line"
            }
        }
    }
}

# SIG # Begin signature block
# MIIZsAYJKoZIhvcNAQcCoIIZoTCCGZ0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU3GF3O64ES7CwcZ/c3z9krpQd
# eJCgghS+MIIE/jCCA+agAwIBAgIQDUJK4L46iP9gQCHOFADw3TANBgkqhkiG9w0B
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
# 9w0BCQQxFgQUB80jZNueLVewhVcBKP3XqJxUdN0wDQYJKoZIhvcNAQEBBQAEggEA
# GHg8G/OJ3T6BtpA3AHNl5EyN2MsY7VQ08giBqk9D6AI0DZxMJ4Zv+Rcv9Noz0Qy1
# Wq7+EJ6K+MmJ3lyczCZZO7qAtcicT6ddhKejNbO3C3evbh2pH+XqwIhqJiP7ALwn
# m8NOIS3diLIXssezvSHNAProtQurRSGE1O3Rpy1+oEGz3OCL3tr9TIff2hhm34UX
# 3YmVIC7blrS9BV8AWowR3C5OD0OIHTqdkHU6LKO0m1HkfQKnidqjjYs3RaDIajR/
# S58PZOsr/0KvCSHthScUxg025lnSHnIpXFks6j/Rve5+gaiRLeJgnMlBNJXu4QCl
# ACx3ZdyvPsv3qKZXI21k06GCAjAwggIsBgkqhkiG9w0BCQYxggIdMIICGQIBATCB
# hjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3Vy
# ZWQgSUQgVGltZXN0YW1waW5nIENBAhANQkrgvjqI/2BAIc4UAPDdMA0GCWCGSAFl
# AwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMjIwMTE2MTQzMDM3WjAvBgkqhkiG9w0BCQQxIgQgVvC4Ev0uDJWBdpvGD8+5
# RFjk53BRVqRTA4WFkpJDQiwwDQYJKoZIhvcNAQEBBQAEggEAenXRsSI/jrRXP0dr
# WJ0ZyG4QOGi/pfD6KcrvHndmTs4mm5KxDAahF5DvzttMS+tMuRGNaoP7Eo81tV59
# wIyDY6YyhoHml/8HtAXchgbqIEXQFGs0jEbaddM5y9zC22duEfbwI48HJbBKZFlb
# iU4nxLrW3WZ4gRFKd9LSfnEjFdlR2YG2OXH0dLrszhL395XOzNYj91q4G4cDinRP
# ceITc04RYoUgMU06SZKaCX1bjdLYwwcGMM04mC8b8Xq/QryzxSFYDAwP91iC9Yok
# 81pcbRtxkyQF/F2KKVAVwrCTnR1DGkiCkf7Ec0usjlHiT/u0hOYL9xvYsHtDzN37
# Fx1zZg==
# SIG # End signature block
