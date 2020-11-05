#Requires -version 2.0

<#
    Get our external IP address from a web service and update our Dynamic dns provider if it has changed since last registered

    Guy Leech, 2016

    Modification History

    01/09/19  GRL Added ability to email instead of update URL
    05/11/20  GRL Made passing credentials to update optional via -webauth and can take a credential object
#>

<#
.SYNOPSIS

Update dynamic DNS provider if external IP address has changed to update the address or email the details

.DESCRIPTION

Stores the IP address used for the last update in the registry as well as the time it was updated so will only update the provider or send email if the address has changed or has not been updated in a given number of days (25 by default)

.PARAMETER url

The URL of the dynamic DNS provider which can be one of a number of forms with placeholders which will be replaced with the values passed for parameters such as the host name you are updating.

The various forms are (examples from FreeDNS from afraid.org, see https://freedns.afraid.org/dynamic/v2/):

    http://sync.afraid.org/u/dfsf73409sadfjks - A randomized update token which is unique to your host name and does not need authentication. Get your URL from your provider

    https://freedns.afraid.orgz/nic/update?hostname=+hostname&myip=+ip - The +hostname will be replaced with the host name provided via the -hostname argument, +ip will be the external IP address discovered. Needs authentication so -username and -password must be specified

    https://+username:+password@freedns.afraid.org/nic/update?hostname=+hostname&myip=+ip - As above but also passes the -username and -password arguments in the URL

.PARAMETER regKey

The registry key used to store the IP address and date of last change

.PARAMETER regValue

The name of the registry value used to store the IP address

.PARAMETER dateStamp

The name of the registry value used to store the date/time of the last time the IP address changed

.PARAMETER infoProvider

The URL of a service that will return your external IP address

.PARAMETER force

Force the update URL to be called even if the IP address hasn't changed

.PARAMETER notupdated

Force an update if the IP address has not changed in this number of days. Helps stop your account being deleted for being inactive

.PARAMETER username

The user name for your dynamic DNS provider account

.PARAMETER password

The user name password for your dynamic DNS provider account

.PARAMETER hostname

Your dynamic DNS host name

.PARAMETER history

Write IP addresses and the first date/time that they were detected to the "History" subkey of regKey

.PARAMETER encryptPassword

Encrypt the password passed by the -mailpassword option so it can be passed to -mailHashedPassword. The encrypted password is specific to the user and machine where they are encrypted.
Pipe through clip.exe or Set-ClipBoard (scb) to place in the Windows clipboard

.PARAMETER mailServer

The SMTP mail server to use

.PARAMETER proxyMailServer

If email relaying is only allowed from specific computers, try and remote the Send-Email cmdlet via the server specific via this argument

.PARAMETER noSSL

Do not use SSL to communicate with the mail server

.PARAMETER subject

The subject of the email sent with the expiring account list

.PARAMETER from

The email address to send the email from. Some mail servers must have a valid email address specified

.PARAMETER recipients

A comma separated list of the email addresses to send emails to

.PARAMETER mailUsername

The username to authenticate with at the mail server

.PARAMETER mailPassword

The clear text password for the -mailUsername argument. If the %_MVal12% environment variable is set then its contents are used for the password

.PARAMETER mailHashedPassword

The hashed mail password returned from a previous call to the script with -encryptPassword. Will only work on the same machine and for the same user that ran the script to encrypt the password

.PARAMETER mailPort

The port to use to communicate with the mail server

.PARAMETER logfile

A logfile to write information to. It will be appended to.

.EXAMPLE

& 'C:\Scripts\Update dynamic dns.ps1' -url http://sync.afraid.org/u/dfsf73409sadfjks

Updates your DNS provider, if the IP address has changed since the last check or it was updated over 25 days ago, using your unique randomized update token 

.EXAMPLE

& 'C:\Scripts\Update dynamic dns.ps1' -url https://freedns.afraid.orgz/nic/update?hostname=+hostname&myip=+ip -username bob -password LetMeIn -hostname myhost.soon.it

Updates your DNS provider for the myhost.soon.it host, if the IP address has changed since the last check or it was updated over 25 days ago, using the credentials supplied

.EXAMPLE

& 'C:\Scripts\Update dynamic dns.ps1' -recipients guyl@hell.com -mailServer smtp.hell.com -notupdated 7 -subject "External IP address of fred.dyndns.com"

Emails the current external IP address if it has changed since the script was last run or if it has not been updated in the last 7 days

.EXAMPLE

& 'C:\Scripts\Update dynamic dns.ps1' -encryptPassword -mailPassword MySecretPassword

Encrypt the given password so that it can be passed to another invocation of the script via the -mailPassword option if the SMTP email server requries authentication

.NOTES

See https://freedns.afraid.org/dynamic/v2/ for details of the URLs
#>

[CmdletBinding()]

Param
(
    [string]$url , ## unique URL from http://freedns.afraid.org/dynamic/v2/ or one with parameters that need filling in
    [string]$infoProvider = 'icanhazip.com' ,
    [switch]$force , 
    [int]$notupdated = 25 , ## days
    [string]$logfile ,
    [string]$username ,
    [string]$password ,
    [string]$hostname , 
    [switch]$history ,
    [switch]$webauth ,
    [System.Management.Automation.PSCredential]$webCredential ,
    [string]$hashedPassword ,
    [switch]$encryptPassword ,
    [string]$mailServer ,
    [string]$proxyMailServer = 'localhost' ,
    [switch]$noSSL ,
    [string]$subject = "External IP address" ,
    [string]$from = "$($env:computername)@$($env:userdnsdomain)" ,
    [string[]]$recipients ,
    [string]$mailUsername ,
    [string]$mailPassword ,
    [string]$mailHashedPassword ,
    [int]$mailport ,
    [string]$regKey = 'HKCU:\SOFTWARE\Guy Leech\DynDNS' ,
    [string]$regValue = 'External IP' ,
    [string]$dateStamp = 'Last Updated'
)

if( $encryptPassword )
{
    if( ! $PSBoundParameters[ 'mailpassword' ] -and ! ( $mailPassword = $env:_Mval12 ) )
    {
        Throw 'Must specify the mail username''s password when encrypting via -mailpassword or _Mval12 environment variable'
    }
    
    ConvertTo-SecureString -AsPlainText -String $mailPassword -Force | ConvertFrom-SecureString
    Exit 0
}

if( ! [string]::IsNullOrEmpty( $logfile ) )
{
    Start-Transcript -Path $logfile -Append
}

if( ( [string]::IsNullOrEmpty( $username ) -and ! [string]::IsNullOrEmpty( $password ) )`
    -or ( [string]::IsNullOrEmpty( $password ) -and ! [string]::IsNullOrEmpty( $username ) ) )
{
    Throw "Must specify both -username and -password, not just one"
}

[bool]$updateRegistry = $true

$externalIP = (Invoke-WebRequest $infoprovider).Content.Trim()

try
{
    $existingIP = (Get-ItemProperty -Path $regKey -EA SilentlyContinue|Select-Object -ExpandProperty $regValue -EA SilentlyContinue).Trim()
}
catch
{
    $existingIP = $null
}

## Get date of last change to see if we need to force one anyway lest we are deemed inactive
try
{
    [datetime]$dateChanged = (Get-ItemProperty -Path $regKey -EA SilentlyContinue|select -ExpandProperty $dateStamp -EA SilentlyContinue).Trim()

    $difference = New-TimeSpan -End (Get-Date) -Start $dateChanged
    Write-Verbose "Last updated `"$dateChanged`" ($($difference.Days) days ago)"
    if( $difference.TotalDays -ge $notupdated )
    {
        Write-Verbose "Forcing update due to age"
        $force = $true
    }
}
catch
{
    ## can't get date so can't do anything about it
}

if( $externalIP -eq $existingIP -and ! $force )
{
    Write-Output "IP address not changed at $externalIP so not updating"
}
elseif( $PSBoundParameters[ 'url' ] )
{
    [hashtable]$params = @{}
    
    if( $webauth )
    {
        if( ! [string]::IsNullOrEmpty( $username ) -and ! [string]::IsNullOrEmpty( $password ) )
        {
            $securePassword = $password | ConvertTo-SecureString -asPlainText -Force

            $webCredential = New-Object System.Management.Automation.PSCredential( $username , $securePassword )
        }
        else
        {
            Throw "Webauth requested but no password specified"
        }
    }
    if( $webCredential )
    {
        $params.Add( 'Credential' , $webCredential )
    }

    ## Now fill in placeholders in URL. Encode in case special characters in password
    Add-Type -AssemblyName System.Web

    $url = $url -replace '\+username' , $username -replace '\+password' , [System.Web.HttpUtility]::UrlEncode( $password ) -replace '\+ip', $externalIP -replace '\+hostname' , $hostname

    if( ( $response = Invoke-WebRequest -Uri $url @params ) -and $response.StatusCode -eq 200 )
    {
        Write-Verbose -Message "Response was $($response.Content)"
    }
    else
    {
        Write-Error "Received bad response code $($response.StatusCode) for URL `"$url`"`n$($error[0])"
        $updateRegistry = $false
    }
}
elseif( $PSBoundParameters[ 'recipients' ] -and $PSBoundParameters[ 'mailServer' ] )
{
    [hashtable]$mailParams = $null

    if( $recipients -and $recipients.Count -eq 1 -and $recipients[0].IndexOf(',') -ge 0 ) ## fix scheduled task not passing array correctly
    {
        $recipients = $recipients -split ','
    }

    $mailParams = @{
            'To' =  $recipients
            'SmtpServer' = $mailServer
            'From' =  $from
            'UseSsl' = ( ! $noSSL ) }
    if( $PSBoundParameters[ 'mailport' ] )
    {
        $mailParams.Add( 'Port' , $mailport )
    }
    if( $PSBoundParameters[ 'mailUsername' ] )
    {
        $thePassword = $null
        if( ! $PSBoundParameters[ 'mailPassword' ] )
        {
            if( $PSBoundParameters[ 'mailHashedPassword' ] )
            {
                Write-Verbose "Using hashed password of length $($mailHashedPassword.Length)"
                $thePassword = $mailHashedPassword | ConvertTo-SecureString
            }
            elseif( Get-ChildItem -Path env:_MVal12 -ErrorAction SilentlyContinue )
            {
                $thePassword = ConvertTo-SecureString -AsPlainText -String $env:_MVal12 -Force
            }
        }
        else
        {
            $thePassword = ConvertTo-SecureString -AsPlainText -String $mailPassword -Force
        }
        
        if( $thePassword )
        {
            $mailParams.Add( 'Credential' , ( New-Object System.Management.Automation.PSCredential( $mailUsername , $thePassword )))
        }
        else    
        {
            Write-Error "Must specify mail account password via -mailPassword, -mailHashedPassword or _MVal12 environment variable"
        }
        
        [string]$bodyText = "External IP address is now $externalIP, was $existingIP"

        $mailParams.Add( 'Body' , $bodyText )
        $mailParams.Add( 'BodyAsHtml' , $false )
        $mailParams.Add( 'Subject' ,  $subject )
         
        if( $PSBoundParameters[ 'proxyMailServer' ] )
        {
            Invoke-Command -ComputerName $proxyMailServer -ScriptBlock { [hashtable]$mailParams = $using:mailParams ; Send-MailMessage @mailParams }
        }
        else
        {
            Send-MailMessage @mailParams 
        }
    }
}
else
{
    Throw "No action to take"
}

if( $updateRegistry )
{
    if( ! ( Test-Path -Path $regKey ) )
    {
        if( ! (New-Item -Path $regKey -Force) )
        {
            Write-Warning -Message "Failed to create $regKey"
        }
    }
    Remove-ItemProperty -Path $regKey -Name $regValue -Force -EA SilentlyContinue
    if( $updatedValue = New-ItemProperty -Path $regKey -Name $regValue -Value $externalIP -PropertyType String -Force )
    {
        ## Store datetime of when updated so we can force it if not updated in a given time window
        Remove-ItemProperty -Path $regKey -Name $dateStamp -Force -EA SilentlyContinue
        $null = New-ItemProperty -Path $regKey -Name $dateStamp -Value (Get-Date) -PropertyType String -Force
    }

    if( $history )
    {
        [string]$historyKey = Join-Path -Path $regKey -ChildPath 'History'
        if( ! ( Test-Path -Path $historyKey ) )
        {
            if( ! (New-Item -Path $historyKey -Force ) )
            {
                Write-Warning -Message "Failed to create history key $regKey"
            }
        }
        if( ! ( New-ItemProperty -Path $historyKey -Name (Get-Date) -Value $externalIP -PropertyType String -Force ) )
        {
            Write-Warning -Message "Failed to save $externalIP to $historyKey"
        }
    }
}
if( ! [string]::IsNullOrEmpty( $logfile ) )
{
    Stop-Transcript
}

# SIG # Begin signature block
# MIINRQYJKoZIhvcNAQcCoIINNjCCDTICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUA8+RsgYClIixttSv/2AsNoJ1
# IjOgggqHMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
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
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFJJARL3svCKq1wv5liGe
# Zr+FkIujMA0GCSqGSIb3DQEBAQUABIIBAIK+FMVgIE7Blxa5VeJZ37eQIf4IHyOY
# fkE8vwMKoWUUmc8d7jytxWbwhSoUONqCiKsZQgHgI9kZLRo1nhCayH8/+TPJjJdy
# LtPpSK3VepFtPzqBGwY4VvW9MV2u2wsTJWbDm4dZRqVNi75YotDfew2dq2F6wevf
# Xx9Spdg0c12X/HgofVH3j5YsI3z8ph3ObuWZPhGu164CWDH9rslNOa5iRYH1Ukob
# xCet/ijQloXEMUyeukqWQuNOgochcwz5Op3tx/sbLf0qBLvQvt72PsZ57wAFAo8y
# Igze7xmu62yhpZLFnsokICgqJlvrgwEwUAylgjFNuCA360S/I/1jm90=
# SIG # End signature block
