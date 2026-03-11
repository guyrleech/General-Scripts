<#
.SYNOPSIS
    Decode JWT token either passed via jwt or reads clipboard if not

.PARAMETER jwt
    Java Web Token to decode. Will try and read from the clipboard if not passed
    
.PARAMETER expand
    Which top level property to expand. Default is "claims" as that is usually what is required
    
.PARAMETER convertDates
    Convert dates to ones you humans can understand
    
.PARAMETER dateFields
    The fields in the claims that contain dates. Should not need to change. Used by -convertDates

.NOTES
    Modification History:

    2025/03/11  @guyrleech  Script born
    2026/03/11  @guyrleech  Added -expand and -convertDates
#>

[CmdletBinding()]

Param
(
    [string]$jwt ,
    [string]$expand = 'claims' ,
    [switch]$convertDates ,
    [string[]]$dateFields = @( 'exp' , 'iat' , 'nbf' )
)

if( [string]::IsNullOrEmpty( $jwt ) )
{
    $jwt = Get-Clipboard -Format Text
}

if( [string]::IsNullOrEmpty( $jwt ) )
{
    Throw "Must pass JWT token either via parameter or clipboard"
}

#region Functions

## https://gallery.technet.microsoft.com/JWT-Token-Decode-637cf001
## For decoding JWT (Java Web Tokens)

function Convert-FromBase64StringWithNoPadding([string]$data)
{
    $data = $data.Replace( '-' , '+').Replace( '_' , '/' )
    switch ($data.Length % 4)
    {
        0 { break }
        2 { $data += '==' }
        3 { $data += '=' }
        default { throw New-Object ArgumentException('data') }
    }
    return [System.Convert]::FromBase64String($data)
}

function Decode-JWT( [string]$rawToken )
{
    [string[]]$parts = $rawToken.Split('.')
    if( $parts -and $parts.Count -ge 2 )
    {
        $headers = [System.Text.Encoding]::UTF8.GetString( (Convert-FromBase64StringWithNoPadding -data $parts[0]) )
        $claims  = [System.Text.Encoding]::UTF8.GetString( (Convert-FromBase64StringWithNoPadding -data $parts[1]) )
        $signature = (Convert-FromBase64StringWithNoPadding $parts[2])
        
        Write-Verbose -Message ("JWT`r`n.headers: {0}`r`n.claims: {1}`r`n.signature: {2}`r`n" -f $headers,$claims,[System.BitConverter]::ToString($signature))

        [PSCustomObject]@{
            headers = $headers | ConvertFrom-Json
            claims = $claims | ConvertFrom-Json
            signature = $signature
        }
    }
    else
    {
        Write-Warning -Message "Bad JWT token $rawToken"
    }
}

function Get-JwtTokenData
{
    [CmdletBinding()]  
    Param
    (
        # Param1 help description
        [Parameter(Mandatory)]
        [string] $Token,
        [switch] $Recurse
    )
    
    if ($Recurse)
    {
        Decode-JWT -rawToken ([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($Token)))
    }
    else
    {
        Decode-JWT -rawToken $Token
    }
}

#endregion Functions

$decoded = Decode-JWT -rawToken $jwt
[datetime]$epoch = [datetime]'01/01/1970'  ## not US format :-)

if( [string]::IsNullOrEmpty( $expand ) )
{
    $decoded
}
else
{
    $object = $decoded | Select-Object -ExpandProperty $expand

    if( $convertDates )
    {
        $converted = [ordered]@{}
        ForEach( $property in $object.PSObject.Properties )
        {
            if( $property.Name -in $dateFields )
            {
                $date = $null
                try
                {
                    $date = $epoch.AddSeconds( $property.Value )
                }
                catch
                {
                    Write-Warning "Bad Unix date/time $($property.Valye) in $($property.Name)"
                }
                $converted.Add( $property.Name , $date )
            }
            else
            {
                $converted.Add( $property.Name , $property.Value )
            }
        }
        [pscustomobject]$converted
    }
    else
    {
        $object
    }
}

# SIG # Begin signature block
# MIIkkwYJKoZIhvcNAQcCoIIkhDCCJIACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUNLeq0Ej8gaFURjx+keEla7N1
# QLaggh9gMIIFfTCCA2WgAwIBAgIQAdazdTZfIM2RHdcv5fmTZDANBgkqhkiG9w0B
# AQsFADBaMQswCQYDVQQGEwJMVjEZMBcGA1UEChMQRW5WZXJzIEdyb3VwIFNJQTEw
# MC4GA1UEAxMnR29HZXRTU0wgRzQgQ1MgUlNBNDA5NiBTSEEyNTYgMjAyMiBDQS0x
# MB4XDTI1MDcyMTAwMDAwMFoXDTI2MDcyMDIzNTk1OVowcTELMAkGA1UEBhMCR0Ix
# EjAQBgNVBAcTCVdha2VmaWVsZDEmMCQGA1UEChMdU2VjdXJlIFBsYXRmb3JtIFNv
# bHV0aW9ucyBMdGQxJjAkBgNVBAMTHVNlY3VyZSBQbGF0Zm9ybSBTb2x1dGlvbnMg
# THRkMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAERFbrIQcZmiw2ScrP4eHxhzHvoBGn
# AnE3GpY3vjU5CpVG6JLtPgXTQz8aLW7IdGhx7x4cJ3a6y+3/6Q+OX+VVFSiuRd60
# GO22Y2eoMcBmvwc7hWbEYTtdjEzAu82sMmkAo4IB1DCCAdAwHwYDVR0jBBgwFoAU
# yfwQ71DIy2t/vQhE7zpik+1bXpowHQYDVR0OBBYEFPu1ucNQfJlsl2iXm5HJzCZH
# EgfnMD4GA1UdIAQ3MDUwMwYGZ4EMAQQBMCkwJwYIKwYBBQUHAgEWG2h0dHA6Ly93
# d3cuZGlnaWNlcnQuY29tL0NQUzAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYI
# KwYBBQUHAwMwgZcGA1UdHwSBjzCBjDBEoEKgQIY+aHR0cDovL2NybDMuZGlnaWNl
# cnQuY29tL0dvR2V0U1NMRzRDU1JTQTQwOTZTSEEyNTYyMDIyQ0EtMS5jcmwwRKBC
# oECGPmh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9Hb0dldFNTTEc0Q1NSU0E0MDk2
# U0hBMjU2MjAyMkNBLTEuY3JsMIGDBggrBgEFBQcBAQR3MHUwJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBNBggrBgEFBQcwAoZBaHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0dvR2V0U1NMRzRDU1JTQTQwOTZTSEEyNTYyMDIy
# Q0EtMS5jcnQwCQYDVR0TBAIwADANBgkqhkiG9w0BAQsFAAOCAgEACx2bHTrjHk/3
# tX7HUH4SM/9sEtfVFpRtZcV4nmDpjwI7tAwOSGYXk4DLVjJFJveoFjbYsZ8vquZ1
# /HJM7rg+O3rCNmzOBvUXFVSjdL3S2R7+kF2ROR7dqk1/BNW6n3o7Q3BmNGqjo1WH
# jov6PfAbEffCLZI1jT98RNqChMesWMmQS+nf8xwdskne4XZOFX5h/a00X7QLAJ+S
# /bOptiC0SvEEa5FCWPUcV7ML0MtoDc3HIPnmMMuYLy586eJHbE5XlfEsmWUNk3Kf
# hxzxsXpAdTSDOeb5Qm/aHGMOY+56Gnt/zxfrv2bfxPnKKZtXPjA47tm89RHpal8b
# lbCAkVfYpKSe0BFPi8FIk+zXvpoAZkNyCMm/HUMEdMbtRP7CqFmYz0YWuiS3uuUW
# qAZ1zl+n1kIJT8eOu6o01EKS8ShijHUI0vixibiNvwTFgRyX3Yc/9xkfV1Wgzli4
# ZPgoZI6FwYBdrhRF0or+CzYIoENUfUYqI7pBM5kkXuSytFD3SXIeSPx14NZSRTzk
# cdOSJWtLkjLrIrIKzzb5eXxLn/gxmJdssB7GUKZHik+cB0OUCRKHEysBj34hnvXa
# zuQ6DKLOQFy+cZ6z4f2kAeFyq7bWUxctPmF61FkmGvb9q6e3AMLg7JnfYC6EM31u
# 42oGx38b5i0NAiUzvWOAbCWTC+G44pgwggWNMIIEdaADAgECAhAOmxiO+dAt5+/b
# UOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxE
# aWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMT
# G0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBaFw0z
# MTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
# AL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/z
# G6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZ
# anMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7s
# Wxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL
# 2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfb
# BHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3
# JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3c
# AORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqx
# YxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0
# viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aL
# T8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIBNjAP
# BgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwPTzAf
# BgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMCAYYw
# eQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2Vy
# dC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0aHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# bDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0NcVec4
# X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnovLbc4
# 7/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65ZyoUi0
# mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFWjuyk
# 1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPFmCLB
# sln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9ztwGp
# n1eqXijiuZQwggahMIIEiaADAgECAhAHhD2tAcEVwnTuQacoIkZ5MA0GCSqGSIb3
# DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAX
# BgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0
# ZWQgUm9vdCBHNDAeFw0yMjA2MjMwMDAwMDBaFw0zMjA2MjIyMzU5NTlaMFoxCzAJ
# BgNVBAYTAkxWMRkwFwYDVQQKExBFblZlcnMgR3JvdXAgU0lBMTAwLgYDVQQDEydH
# b0dldFNTTCBHNCBDUyBSU0E0MDk2IFNIQTI1NiAyMDIyIENBLTEwggIiMA0GCSqG
# SIb3DQEBAQUAA4ICDwAwggIKAoICAQCtHvQHskNmiqJndyWVCqX4FtYp5FfJLO9S
# h0BuwXuvBeNYt21xf8h/pLJ/7YzeKcNq9z4zEhecqtD0xhbvSB8ksBAfWBMZO0NL
# fOT0j7WyNuD7rv+ZFza+mxIQ79s1dCiwUMwGonaoDK7mqZfDpKEExR6UyKBh3aat
# T73U2Imx/x+fYTmQFq+N8FrLs6Fh6YEGWJTgsxyw1fAChCfgtEcZkdtcgK7quqsk
# HtW6PJ9l5VNJ7T3WXpznsOOxrz3qx0CzWjwK8+3Kv2X6piWvd8YRfAOycSrT4/PM
# 0cHLFc5xs/4m/ek4FCnYSem43doFftBxZBQkHKoPW3Bt6VIrhVIwvO7hrUjhchJJ
# ZYdSld3bANDviJ5/ToP7ENv97U9MtKFvmC5dzd1p4HxFR0p5wWmYQbW+y3RFm0np
# 6H9m57MUMNp0ysmdJjb0f7+dVLX3OEBUb6H+r1LRLZT/xEOTuwOxGg2S4w25KGL9
# SCBUW4nkBljPHeJToU+THt0P8ZQf4B9IFlGxtLK0g3uOAnwSFgKtmNjhkTl8caLA
# QwbgEINCqrhc0b6k2Z8+QwgVAL0nIuzM9ckKP8xtIcWg85L3/l0cTkHQde+jKGDG
# 2CdxBHtflLIUtwqD7JA2uCxWlIzRNgwT0kH2en0+QV8KziSGaqO2r06kwboq2/xy
# 4e98CEfSYwIDAQABo4IBWTCCAVUwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4E
# FgQUyfwQ71DIy2t/vQhE7zpik+1bXpowHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5n
# P+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcG
# CCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQu
# Y29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGln
# aUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8v
# Y3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAcBgNV
# HSAEFTATMAcGBWeBDAEDMAgGBmeBDAEEATANBgkqhkiG9w0BAQsFAAOCAgEAC9sK
# 17IdmKTCUatEs7+yewhJnJ4tyrLwNEnfl6HrG8Pm7HZ0b+5Jc+GGqJT8kRc7mihu
# VrdsYNHdicueDL9imhtCusI/rUmjwhtflp+XgLkmgLGrmsEho1b+lGiRp7LC/10d
# i8SAOilDkHj5Zx142xRvBrrWj9eOdSGHwYubAsEd6CDojwcaVz9pfXMzYO3kc0O6
# PXg1TkcgkYlCUAuDHuk/sZx68W0FVj1P2iMh+VUq9lL1puroAydoeWVUh/+cMXeq
# fgpBqlAW+r8ma5F6yKL0stVQH8vYb1ES0mJSIPyIfkIjC1V0pbZS3p0QWsKaafEo
# r8fLfLNfSxntVI/ugut0+6ekluPWRpEXH+JAiNdRjbLbZchCREe3/Xl0YlwkA+eQ
# VJfM0A7XiuFtY/mOpK2AN+E25t5mQYFhpdxZX5LTDKWgDnb+A6QnEt4iNyukcLaJ
# uS8IPgPz0E2ALZLt3Rqs+lXifK/GwnNIWQNbf7FmLDB9ph8i8dvsR1hsjc2KPEW4
# bAsbvLcz8hN1zE1/QbOV92vDGoFjwZOi2koQ+UyEh0e8jDFHAKJeTI+p8EPE/mqv
# ojLFAnt31yXIA2tjt0ERtsjkhBNmZY6SEOfnIoOwvyqavLPya1Ut3/2cOFLuNQ8Q
# l6HaZsNQErnnzn+ZEAaUTkPZaeVyoHIkODECLzkwgga0MIIEnKADAgECAhANx6xX
# Bf8hmS5AQyIMOkmGMA0GCSqGSIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0yNTA1MDcwMDAwMDBa
# Fw0zODAxMTQyMzU5NTlaMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBp
# bmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4IC
# DwAwggIKAoICAQC0eDHTCphBcr48RsAcrHXbo0ZodLRRF51NrY0NlLWZloMsVO1D
# ahGPNRcybEKq+RuwOnPhof6pvF4uGjwjqNjfEvUi6wuim5bap+0lgloM2zX4kftn
# 5B1IpYzTqpyFQ/4Bt0mAxAHeHYNnQxqXmRinvuNgxVBdJkf77S2uPoCj7GH8BLux
# BG5AvftBdsOECS1UkxBvMgEdgkFiDNYiOTx4OtiFcMSkqTtF2hfQz3zQSku2Ws3I
# fDReb6e3mmdglTcaarps0wjUjsZvkgFkriK9tUKJm/s80FiocSk1VYLZlDwFt+cV
# FBURJg6zMUjZa/zbCclF83bRVFLeGkuAhHiGPMvSGmhgaTzVyhYn4p0+8y9oHRaQ
# T/aofEnS5xLrfxnGpTXiUOeSLsJygoLPp66bkDX1ZlAeSpQl92QOMeRxykvq6gby
# lsXQskBBBnGy3tW/AMOMCZIVNSaz7BX8VtYGqLt9MmeOreGPRdtBx3yGOP+rx3rK
# WDEJlIqLXvJWnY0v5ydPpOjL6s36czwzsucuoKs7Yk/ehb//Wx+5kMqIMRvUBDx6
# z1ev+7psNOdgJMoiwOrUG2ZdSoQbU2rMkpLiQ6bGRinZbI4OLu9BMIFm1UUl9Vne
# Ps6BaaeEWvjJSjNm2qA+sdFUeEY0qVjPKOWug/G6X5uAiynM7Bu2ayBjUwIDAQAB
# o4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU729TSunkBnx6
# yuKQVvYv1Ensy04wHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYD
# VR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMIMHcGCCsGAQUFBwEBBGsw
# aTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUF
# BzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVk
# Um9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2Vy
# dC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBABfO+xaAHP4HPRF2
# cTC9vgvItTSmf83Qh8WIGjB/T8ObXAZz8OjuhUxjaaFdleMM0lBryPTQM2qEJPe3
# 6zwbSI/mS83afsl3YTj+IQhQE7jU/kXjjytJgnn0hvrV6hqWGd3rLAUt6vJy9lMD
# PjTLxLgXf9r5nWMQwr8Myb9rEVKChHyfpzee5kH0F8HABBgr0UdqirZ7bowe9Vj2
# AIMD8liyrukZ2iA/wdG2th9y1IsA0QF8dTXqvcnTmpfeQh35k5zOCPmSNq1UH410
# ANVko43+Cdmu4y81hjajV/gxdEkMx1NKU4uHQcKfZxAvBAKqMVuqte69M9J6A47O
# vgRaPs+2ykgcGV00TYr2Lr3ty9qIijanrUR3anzEwlvzZiiyfTPjLbnFRsjsYg39
# OlV8cipDoq7+qNNjqFzeGxcytL5TTLL4ZaoBdqbhOhZ3ZRDUphPvSRmMThi0vw9v
# ODRzW6AxnJll38F0cuJG7uEBYTptMSbhdhGQDpOXgpIUsWTjd6xpR6oaQf/DJbg3
# s6KCLPAlZ66RzIg9sC+NJpud/v4+7RWsWCiKi9EOLLHfMR2ZyJ/+xhCx9yHbxtl5
# TPau1j/1MIDpMPx0LckTetiSuEtQvLsNz3Qbp7wGWqbIiOWCnb5WqxL3/BAPvIXK
# UjPSxyZsq8WhbaM2tszWkPZPubdcMIIG7TCCBNWgAwIBAgIQCoDvGEuN8QWC0cR2
# p5V0aDANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGln
# aUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0
# YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExMB4XDTI1MDYwNDAwMDAwMFoX
# DTM2MDkwMzIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBTSEEyNTYgUlNBNDA5NiBUaW1lc3Rh
# bXAgUmVzcG9uZGVyIDIwMjUgMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBANBGrC0Sxp7Q6q5gVrMrV7pvUf+GcAoB38o3zBlCMGMyqJnfFNZx+wvA69HF
# TBdwbHwBSOeLpvPnZ8ZN+vo8dE2/pPvOx/Vj8TchTySA2R4QKpVD7dvNZh6wW2R6
# kSu9RJt/4QhguSssp3qome7MrxVyfQO9sMx6ZAWjFDYOzDi8SOhPUWlLnh00Cll8
# pjrUcCV3K3E0zz09ldQ//nBZZREr4h/GI6Dxb2UoyrN0ijtUDVHRXdmncOOMA3Co
# B/iUSROUINDT98oksouTMYFOnHoRh6+86Ltc5zjPKHW5KqCvpSduSwhwUmotuQhc
# g9tw2YD3w6ySSSu+3qU8DD+nigNJFmt6LAHvH3KSuNLoZLc1Hf2JNMVL4Q1Opbyb
# pMe46YceNA0LfNsnqcnpJeItK/DhKbPxTTuGoX7wJNdoRORVbPR1VVnDuSeHVZlc
# 4seAO+6d2sC26/PQPdP51ho1zBp+xUIZkpSFA8vWdoUoHLWnqWU3dCCyFG1roSrg
# HjSHlq8xymLnjCbSLZ49kPmk8iyyizNDIXj//cOgrY7rlRyTlaCCfw7aSUROwnu7
# zER6EaJ+AliL7ojTdS5PWPsWeupWs7NpChUk555K096V1hE0yZIXe+giAwW00aHz
# rDchIc2bQhpp0IoKRR7YufAkprxMiXAJQ1XCmnCfgPf8+3mnAgMBAAGjggGVMIIB
# kTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBTkO/zyMe39/dfzkXFjGVBDz2GM6DAf
# BgNVHSMEGDAWgBTvb1NK6eQGfHrK4pBW9i/USezLTjAOBgNVHQ8BAf8EBAMCB4Aw
# FgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwgZUGCCsGAQUFBwEBBIGIMIGFMCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wXQYIKwYBBQUHMAKGUWh0
# dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNydDBfBgNVHR8EWDBWMFSgUqBQ
# hk5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRUaW1l
# U3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5jcmwwIAYDVR0gBBkwFzAIBgZn
# gQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQBlKq3xHCcEua5g
# QezRCESeY0ByIfjk9iJP2zWLpQq1b4URGnwWBdEZD9gBq9fNaNmFj6Eh8/YmRDfx
# T7C0k8FUFqNh+tshgb4O6Lgjg8K8elC4+oWCqnU/ML9lFfim8/9yJmZSe2F8AQ/U
# dKFOtj7YMTmqPO9mzskgiC3QYIUP2S3HQvHG1FDu+WUqW4daIqToXFE/JQ/EABgf
# ZXLWU0ziTN6R3ygQBHMUBaB5bdrPbF6MRYs03h4obEMnxYOX8VBRKe1uNnzQVTeL
# ni2nHkX/QqvXnNb+YkDFkxUGtMTaiLR9wjxUxu2hECZpqyU1d0IbX6Wq8/gVutDo
# jBIFeRlqAcuEVT0cKsb+zJNEsuEB7O7/cuvTQasnM9AWcIQfVjnzrvwiCZ85EE8L
# UkqRhoS3Y50OHgaY7T/lwd6UArb+BOVAkg2oOvol/DJgddJ35XTxfUlQ+8Hggt8l
# 2Yv7roancJIFcbojBcxlRcGG0LIhp6GvReQGgMgYxQbV1S3CrWqZzBt1R9xJgKf4
# 7CdxVRd/ndUlQ05oxYy2zRWVFjF7mcr4C34Mj3ocCVccAvlKV9jEnstrniLvUxxV
# ZE/rptb7IRE2lskKPIJgbaP5t2nGj/ULLi49xTcBZU8atufk+EMF/cWuiC7POGT7
# 5qaL6vdCvHlshtjdNXOCIUjsarfNZzGCBJ0wggSZAgEBMG4wWjELMAkGA1UEBhMC
# TFYxGTAXBgNVBAoTEEVuVmVycyBHcm91cCBTSUExMDAuBgNVBAMTJ0dvR2V0U1NM
# IEc0IENTIFJTQTQwOTYgU0hBMjU2IDIwMjIgQ0EtMQIQAdazdTZfIM2RHdcv5fmT
# ZDAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG
# 9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIB
# FTAjBgkqhkiG9w0BCQQxFgQUqatRvSGTq4e58kuICPtMKtJYvqUwCwYHKoZIzj0C
# AQUABGgwZgIxAMrIGqj2GYSO2/mLF3dnerZQwxl7ptvETMLHGK2dlWvuhFp4Mx4Y
# PUWZBYf/Cq1MDgIxAKTUNC4VxPGGqML0Tzysg98JswmX0fvOepIkmrS/k+TRLI4+
# 4BJXGPS/ktmBMsi21KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjAzMTExMjEz
# NDRaMC8GCSqGSIb3DQEJBDEiBCAu3dv+TLe5rugVoPHBJQgi5ccmMMgwgkV4NO68
# cSvr1TANBgkqhkiG9w0BAQEFAASCAgCUwHDLzdrIe97NN/bNhVOoGSBucuAqikB6
# v3Cz95BFY8tIdEN0b4vb9Cht8ZiAyPp+qAVGgs35P0oeu4bMGqAfedMvD5dbamMk
# TAsC3vdUoH271bkmC3AXzJM4sLXBcro5hun20BQTpt9q8+Nw71GCqcawYsvufLZB
# MHnlBmwun+daGmyHsgTinueoJq0YGOuO9TwTutD6gU/Q/7vRGE+hNNrbBKSlryEJ
# MwnTY5LKC2AMvujw3a6a2MofYYILl1i4mMgo/nwgtkS9DvW+YiLTeDm9ftfUPl6h
# o/n/dkz2PksdeTQV/aGYeWFUlS2/XVBtjRK0st2XFE3FKgr+ZEF4GJFPOEU6bTXR
# 1y/WsU4ZYDobv9CqHuSyOIl+N7GnOD9q6GM0ZlvlUZDvgAOGtPdDjWvieoyq5bAs
# TSr9IUwjgJ/TJ8hwmvo8Yzms7Scm94JvQ8K9yvdRxf6MSRPOBGkHhSiFNBS48iFz
# rtlQtaMBbWXxqRgHXyw2KI/Eqhs5vHr9DmbjMwn2oVzNov18WcqZJxhZM/iA0tEF
# xK0rIrJE7REXOlOl9Zx9vrtjmlzZNKdBZbAZuVM2pe5VbANucmSWjGVxCiQi5agd
# rRrbfP1y9ZxTVY30nMwVc50WhQCK/D/T9ChrJR4fnOPbkvg7yKUQ3dVw1x7T6Ka9
# SyYlh2whMg==
# SIG # End signature block
