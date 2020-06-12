<#
.SYNOPSIS

Read specific amount of string data from a (log) file at a given offset like you might see being written in a SysInternals procmon trace

.PARAMETER path

The (log) file to read and display text data from

.PARAMETER offset

The offset in bytes from the start of the file to read data from

.PARAMETER count

The number of bytes to read from the given offset

.PARAMETER procmonDetail

The text from the "Detail" column in procmon for the given write to the (log) file. E.g. "Offset: 10,014,174, Length: 18", where you would right click the Detail column on the required line in procmon and select "Copy"

.PARAMETER minimumCount

If greater than the length from the -procmonDetail string it will display this much of the log file starting from the offset rather than what was found in the "length" string. Use to display more than one line

.EXAMPLE

& '.\Get chunk at offset.ps1' -path 'C:\temp\EMLogs\EmUser_session_1_Pid-9200_8.49.log" -minimumCount 1KB -procmonDetail 'Offset: 12,798,606, Length: 55'

Read 1KB of data starting at offset 12798606 in the specified file and show 1KB of this where the string containing "Offset" and "Length" have been copied from the detail column of a specific line of a procmon trace

.EXAMPLE

& '.\Get chunk at offset.ps1' -path 'C:\temp\EMLogs\EmUser_session_1_Pid-9200_8.49.log" -count 1KB -offset 12798606

Read 1KB of data starting at offset 12798606 in the specified file and show 1KB of this

.NOTES

Modification History:

12/06/20  @guyrleech  Bug fix procmon detail regex & add test for offset too large

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage='Test log file to read')]
    [string]$path ,
    [Parameter(Mandatory=$true,ParameterSetName='OffsetAndSize')]
    [long]$offset ,
    [Parameter(Mandatory=$false,ParameterSetName='OffsetAndSize')]
    [long]$count = 1KB,
    [Parameter(Mandatory=$true,ParameterSetName='FromProcmon')]
    [string]$procmonDetail ,
    [Parameter(Mandatory=$false,ParameterSetName='FromProcmon')]
    [long]$minimumCount
)

if( $fileStream = New-Object -TypeName System.IO.FileStream( $path , [System.IO.FileMode]::Open , [System.IO.FileAccess]::Read ) )
{
    if( $PSBoundParameters[ 'procmonDetail' ] )
    {
        ## Offset: 10,014,174, Length: 18
        if( $procmonDetail -match 'Offset:\s*([\d,]*).*Length:\s*([\d,]*)' )
        {
            ## [int] conversion copes with commas so no need to remove
            $offset = $Matches[1] -as [int]
            $count = $Matches[2] -as [int]
            if( $PSBoundParameters[ 'minimumCount' ] -and $count -lt $minimumCount )
            {
                $count = $minimumCount
            }
        }
        else
        {
            Throw "Unexpected procmon detail format `"$procmonDetail`""
        }
    }
    
    Write-Verbose -Message "Seeking to $offset and reading $count bytes"

    if( $offset -gt $fileStream.Length )
    {
        Throw "Offset $offset is too large, file is only $($fileStream.Length) bytes"
    }

    $null = $fileStream.Seek( $offset , [System.IO.SeekOrigin]::Begin )

    $chunk = New-Object byte[] $count

    [int]$read = $fileStream.Read( $chunk , 0 , $count )

    if( $read -ne $count )
    {
        Write-Warning "Read $read not $count bytes"
    }

    $fileStream.Close()

    $fileStream = $null

    [System.Text.Encoding]::ASCII.GetString( $chunk )
}
