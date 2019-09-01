#Requires -version 2.0

<#
    Get our external IP address from a web service and update our Dynamic dns provider if it has changed since last registered

    Guy Leech, 2016

    Modification History

    01/09/19  GRL Added ability to email instead of update URL
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
    $existingIP = (Get-ItemProperty -Path $regKey -EA SilentlyContinue|select -ExpandProperty $regValue -EA SilentlyContinue).Trim()
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
    if( ! [string]::IsNullOrEmpty( $password ) )
    {
        $securePassword = $password | ConvertTo-SecureString -asPlainText -Force

        $creds = New-Object System.Management.Automation.PSCredential( $username , $securePassword )
        $params.Add( 'Credential' , $creds )

        ## Now fill in placeholders in URL. Encode in case special characters in password
        [void](Add-Type -AssemblyName System.Web   -Debug:$false)
        $url =  $url -replace '\+username' , $username -replace '\+password' , [System.Web.HttpUtility]::UrlEncode( $password ) -replace '\+ip', $externalIP -replace '\+hostname' , $hostname
    }

    $response = Invoke-WebRequest -Uri $url @params

    if( $response -and $response.StatusCode -eq 200 )
    {
        Write-Verbose $response.Content
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
    if( ! ( Test-Path $regKey ) )
    {
        $null = New-Item $regKey -Force
    }
    Remove-ItemProperty -Path $regKey -Name $regValue -Force -EA SilentlyContinue
    $updatedValue = New-ItemProperty -Path $regKey -Name $regValue -Value $externalIP -PropertyType String -Force

    if( $updatedValue )
    {
        ## Store datetime of when updated so we can force it if not updated in a given time window
        Remove-ItemProperty -Path $regKey -Name $dateStamp -Force -EA SilentlyContinue
        $null = New-ItemProperty -Path $regKey -Name $dateStamp -Value (Get-Date) -PropertyType String -Force
    }

    if( $history )
    {
        [string]$historyKey = $regKey + '\History'
        if( ! ( Test-Path $historyKey ) )
        {
            $null = New-Item $historyKey -Force
        }
        $null = New-ItemProperty -Path $historyKey -Name (Get-Date) -Value $externalIP -PropertyType String -Force
    }
}
if( ! [string]::IsNullOrEmpty( $logfile ) )
{
    Stop-Transcript
}
