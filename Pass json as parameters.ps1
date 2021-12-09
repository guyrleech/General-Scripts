<#
.SYNOPSIS
    Read a JSON config file and call the specified script with parameters from it

.PARAMETER jsonfile
    Path to the JSON file to read parameters from

.PARAMETER script
    Name/path of the script to call with the parameters read from the JSON file and others specified by -otherParameters
    If not an absolute path, it will look in the current folder and the same folder as this script

.PARAMETER otherParameters
    Other parameters to pass to the script. These will override the same parameter if found in the json file

.EXAMPLE
    & '.\Pass json as parameters.ps1' -jsonfile C:\temp\parameters.json -script "New MCS catalog machines delivery group.ps1" -Verbose -otherParameters @{ 'Verbose' = $true }

    Read parameters from the JSON file "C:\temp\parameters.json" and call the script "New MCS catalog machines delivery group.ps1" with these parameters and also -verbose

.NOTES

    Modification History:

    2021/12/09  @guyrleech  Initial release (incorporating ideas from @KevinMarquette)
#>

[CmdletBinding()]

Param
(
    [string]$jsonfile , ## TODO read data from pipeline
    [string]$script ,
    [hashtable]$otherParameters
)

[string]$scriptWithPath = $script

if( -Not ( Test-Path -Path $scriptWithPath ) )
{
    if( -Not ( Test-Path -Path ($scriptWithPath = Join-Path -Path ((Get-Location -PSProvider FileSystem).Path) -ChildPath $script) ) )
    { 
        [string]$thisScriptPath = & { $MyInvocation.PSScriptRoot }
    
        if( -Not ( Test-Path -Path ($scriptWithPath = Join-Path -Path $thisScriptPath -ChildPath $script) ) )
        {
            Throw "Unable to find script `"$script`""
        }
    }
}

$scriptParameters = Get-Command -Name $scriptWithPath | Select-Object -ExpandProperty Parameters

if( $null -eq $scriptParameters -or $scriptParameters.Count -eq 0 )
{
    Throw "Script `"$scriptParameters`" does not take any parameters"
}

Write-Verbose -Message "Script `"$scriptWithPath`" takes $($scriptParameters.Count) parameters"

[hashtable]$parametersFromJson = @{}

if( -Not [string]::IsNullOrEmpty( $jsonfile ) )
{
    if( -Not ( Test-Path -Path $jsonfile -ErrorAction SilentlyContinue ) )
    {
        Throw "JSON file `"$jsonfile`" not found"
    }
    if( $jsonobject = ( Get-Content -Path $jsonfile | ConvertFrom-Json ) )
    {
        ForEach( $property in $jsonobject.PSObject.Properties )
        {
            if( $otherParameters.Contains( $property.Name ))
            {
                Write-Warning -Message "-$($property.Name) present in other parameters passed so ignoring json value `"$($property.Value)`""
            }
            else
            {
                $value = $property.value
                
                ## need to deal with [switch] types as will be custom objects
                if( $property.Value.GetType().Name -eq 'PSCustomObject' )
                {
                    if( $property.Value.PSObject.Properties[ 'IsPresent' ] -and $property.Value.IsPresent -is [bool] )
                    {
                        $value = $property.Value.IsPresent
                    }
                    elseif( $property.Name -match 'credentials$' -and $property.Value.PSobject.Properties[ 'username' ] -and $property.Value.PSobject.Properties[ 'password' ] )
                    {
                        $value = $null
                        $value = New-Object System.Management.Automation.PSCredential( $property.Value.username , ( ConvertTo-SecureString -String $property.Value.password ))
                        if( -not $value )
                        {
                            Throw "Failed to make PScredential for $($property.Value.username)"
                        }
                    }
                    else
                    {
                        Throw "Don't know how to process json file item `"$($property.Name)`""
                    }
                }

                try
                {
                    Write-Verbose -Message "Adding `"$($property.Name)`" = `"$Value`""
                    $parametersFromJson.Add( $property.Name , $value )
                }
                catch
                {
                    if( ( $existing = $parametersFromJson[ $property.Name ] ) -ne $value )
                    {
                        Throw "Repetition of `"$($property.Name)`" with different values `"$value`" and `"$existing`""
                    }
                }
            }
        }
    }
    if( -Not $jsonobject )
    {
        Throw "Failed to parse json from `"$jsonfile`""
    }
}

Write-Verbose -Message "Got $($parametersFromJson.Count) parameters from json to pass"

$otherParameters += $parametersFromJson

Write-Verbose -Message "Calling script `"$scriptWithPath`" with $($otherParameters.Count) parameters"

& $scriptWithPath @otherParameters

# SIG # Begin signature block
# MIIZsAYJKoZIhvcNAQcCoIIZoTCCGZ0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUB9y/n4hoOWDDAkjD1lwA/Grk
# gHCgghS+MIIE/jCCA+agAwIBAgIQDUJK4L46iP9gQCHOFADw3TANBgkqhkiG9w0B
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
# 9w0BCQQxFgQUEIlnRx6YvXMU2yy4KdS/ocsDul0wDQYJKoZIhvcNAQEBBQAEggEA
# CCzmDNIpBFAOx/G0atgt8bA48CSsD2aGMyjjqIwqzKQ41i7MjJnx8A1W6GOgiPfD
# JTzL+y1QDVR1JNyArl0KxoX11N1ZFCy0hYu62ibWTMoS6n5X8YGcOcNs+/ltlWJK
# 5AVgyszlrsgd69T8dVmSFyHcUDCNEMp6bjOHwFaDjjER0D1GPIM7rr2OM5oUkMu0
# cbxp4y+yF4eThalUMGULhHmlQq/XN/p2Klc12aS708u4s8XntiiDyLh8W8n4aB9J
# Cp2sbZnLL56uLz8lUMHJOQZrOfFt5w5lpOaFSzoHG54tR6iXRZbIK3KYboIXA3Wf
# 12Tvq8FZgN127WSn3HRzdKGCAjAwggIsBgkqhkiG9w0BCQYxggIdMIICGQIBATCB
# hjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3Vy
# ZWQgSUQgVGltZXN0YW1waW5nIENBAhANQkrgvjqI/2BAIc4UAPDdMA0GCWCGSAFl
# AwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMjExMjA5MTc1MTQ5WjAvBgkqhkiG9w0BCQQxIgQgyxFMOx7rLqvd1EW19RnN
# Tr2VE4na9khwLV/4h8dYRUMwDQYJKoZIhvcNAQEBBQAEggEAiq49wbeOLTixJqg3
# gzQecbtQZELRFNjLJMd+UmaVpdi7LdISSLKKss+fLR2VJlyA4dfW/DHXY+eVYJ1d
# r7oB5e4nKMmXq1IeREoD9fJwTQ1isOJsNOHyQYGJdYRGtPu+Xd8hUiE1XIUJNFBf
# 8BnWPNaKOVqjUS01G4vNRfkobLRMmcTHY9qhAq5cFyKO9+4wfcN9kRdRqpMZeI0b
# cfRqxhtVo5oKl40kTGBcK00yKGCNfD7LpoV0b43CW1CTXfrOZWSyYRYrtj7fgw4D
# 2vj6B/M74PjQVrcljQ2ytxAnV3V8mh76BLSSWfRx5O/n+f+1irlBBT4KcQwGuZdq
# rPVgCg==
# SIG # End signature block
