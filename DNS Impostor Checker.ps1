<#
.SYNOPSIS

Check DNS servers, using forward and reverse lookups, reports the same machine name as passed via IPv4 addresses.

.DESCRIPTION

Help find stale DNS entries for machines which no longer exist but still have DNS entries with the same IP address as running machines so appear to respond to ping, etc

.PARAMETER computers

List of computers to check

.PARAMETER server

DNS server to query otherwise use whatever has been configured

.PARAMETER quiet

Do not report any errors or warnings

.EXAMPLE

& '.\DNS Impostor Checker.ps1' -computers 'grl-dc03','grl-sql017'

Check that the configured DNS server reports that the IP addresses for machines grl-dc03 and grl-sql017 belong to the same machines
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory,HelpMessage='Computers to check')]
    [string[]]$computers ,
    [string]$server ,
    [switch]$quiet
)

[int]$counter = 0
[int]$bad = 0

[hashtable]$DNSoptions = @{}

if( $PSBoundParameters[ 'server' ] )
{
    $DNSoptions.Add( 'Server' , $server )
}

if( $quiet )
{
    $DNSoptions.Add( 'ErrorAction' , 'SilentlyContinue' )
}

ForEach( $computer in $computers )
{
    $counter++
    Write-Verbose -Message "$counter / $($computers.Count) : $computer"

    ## cope with multiple NICs/addresses
    if( ( [array]$addresses = @( Resolve-DnsName -Name $computer -Verbose:$false @DNSoptions ) ) -and $addresses.Count )
    {
        [int]$addressCounter = 0
        ForEach( $address in $addresses )
        {
            $addressCounter++
            Write-Verbose -Message "`t$addressCounter / $($addresses.Count) : $($address.IPAddress)"
            if( ! ( $name = Resolve-DnsName -Name $address.IPAddress -Verbose:$false @DNSoptions ) -or $name.NameHost -ne $address.Name )
            {
                $bad++
                if( ! $quiet )
                {
                    if( $name )
                    {
                        Write-Warning "$computer ($($address.IPAddress)) masquerading as $($name.Namehost)"
                    }
                    else
                    {
                        Write-Warning "$computer ($($address.IPAddress)) not found"
                    }
                }
            }
        }
    }
    else ## failed to resolve
    {
        $bad++
    }
}

Exit $bad
