<#
.SYNOPSIS
    Test if an IPv4 address is within the given CIDR range

.PARAMETER cidr
    The CIDR to test

.PARAMETER address
    The IP address to test against the CIDR specified

.EXAMPLE
    Test-IPRangeFromCIDR -cidr "192.168.2.1/28" -address 192.168.2.10

    Test if the specified IP address is contained within the given CIDR range

.NOTES

    Modification History:

    2021/11/03  @guyrleech  Initial Release
#>

Function Test-IPRangeFromCIDR
{
    [cmdletbinding()]

    Param
    (
        [Parameter(Mandatory=$true,HelpMessage='IP address range as CIDR')]
        [string]$cidr ,
        [Parameter(Mandatory=$true,HelpMessage='IP address to check in range')]
        [ipaddress]$address
    )

    [ipaddress]$startAddress = [ipaddress]::Any
    [ipaddress]$endAddress   = [ipaddress]::Any

    if( Get-IPRangeFromCIDR -cidr $cidr -startAddress ([ref]$startAddress) -endAddress ([ref]$endAddress) )
    {
        [byte[]]$bytes = $address.GetAddressBytes()
        [uint64]$addressToCompare =  ( ( [uint64]$bytes[0] -shl 24) -bor ( [uint64]$bytes[1] -shl 16) -bor ( [uint64]$bytes[2] -shl 8) -bor  [uint64]$bytes[3])
        $bytes = $startAddress.GetAddressBytes()
        [uint64]$startAddressToCompare =  ( ( [uint64]$bytes[0] -shl 24) -bor ( [uint64]$bytes[1] -shl 16) -bor ( [uint64]$bytes[2] -shl 8) -bor  [uint64]$bytes[3])
        $bytes = $endAddress.GetAddressBytes()
        [uint64]$endAddressToCompare =  ( ( [uint64]$bytes[0] -shl 24) -bor ( [uint64]$bytes[1] -shl 16) -bor ( [uint64]$bytes[2] -shl 8) -bor  [uint64]$bytes[3])

        $addressToCompare -ge $startAddressToCompare -and $addressToCompare -le $endAddressToCompare ## return
    }
}

<#
.SYNOPSIS
    Take a CIDR (Classless Inter-Domain Routing) notation IP v4 range and returns the first and last IPv4 addresses in the range

.PARAMETER cidr
    The CIDR to convert

.PARAMETER startAddress
    Will be set to the start address of the range if the CIDR is valid
    
.PARAMETER endAddress
    Will be set to the end address of the range if the CIDR is valid

.EXAMPLE
    Get-IPRangeFromCIDR -cidr "192.168.2.1/28" -Verbose -startAddress ([ref]$start) -endAddress ([ref]$end)

    Get the starting and ending IPv4 addresses of the specified CIDR range

.NOTES
    Results compared with https://mxtoolbox.com/SubnetCalculator.aspx

    Modification History:

    2021/11/03  @guyrleech  Initial Release
#>

Function Get-IPRangeFromCIDR
{
    [cmdletbinding()]

    Param
    (
        [Parameter(Mandatory=$true,HelpMessage='IP address range as CIDR')]
        [string]$cidr ,
        [Parameter(Mandatory=$true,HelpMessage='IP address range start result')]
        [ref]$startAddress ,
        [Parameter(Mandatory=$true,HelpMessage='IP address range end result')]
        [ref]$endAddress
    )

    [string]$ipaddressPart , [int]$bitsPart = $cidr -split '/'

    if( $bitsPart -eq $null -or $bitsPart -le 0 -or $bitsPart -gt 32 )
    {
        Write-Error -Message "/$bitsPart is invalid"
        return $false
    }

    if( -Not ( $ipaddress = $ipaddressPart -as [ipaddress] ))
    {
        Write-Error -Message "IP address $ipaddressPart is invalid"
        return $false
    }

    [uint64]$mask = ([int64][System.Math]::Pow( 2 , (32 - $bitsPart) ) - 1)
    [byte[]]$bytes = $ipaddress.GetAddressBytes()
    [uint64]$octets =  ( ( [uint64]$bytes[0] -shl 24) -bor ( [uint64]$bytes[1] -shl 16) -bor ( [uint64]$bytes[2] -shl 8) -bor  [uint64]$bytes[3])
    [uint64]$start = $octets -band ($mask -bxor 0xffffffff)
    [uint64]$end = $octets -bor $mask

    Write-Verbose -Message ('Start {0:x} end {1:x}' -f $start , $end)
    
    $startAddress.Value = [ipaddress]$start
    $endAddress.Value   = [ipaddress]$end

    return $true
}
