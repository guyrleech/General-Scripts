<#
.SYNOPSIS
    Creates a file of a given size filled with random binary content.

.DESCRIPTION
    Writes a file whose size is a whole number of megabytes, filled with random
    bytes. Designed to run under PowerShell Constrained Language Mode (CLM): it
    uses only CLM-permitted constructs (no [System.Security.Cryptography] or other
    blocked type instantiation), generating randomness from [guid]::NewGuid().

    Two fill strategies:
      * Default (fast): a single random pool (8 MB by default, see -poolSize) is
        generated once, and each chunk written (1 MB by default, see -chunkSize)
        is a random slice of that pool. Much faster, but chunk-sized blocks repeat
        across the file, so the content is not uniformly unique.
      * -FullyRandom (slower): every chunk is freshly randomised, so the whole
        file is unique random data.

    Works on both Windows PowerShell 5.1 and PowerShell 7+. On 7+ it writes via a
    Latin-1 string (chars 0-255 map 1:1 to bytes) with -Encoding Latin1; on 5.1 it
    writes raw bytes with -Encoding Byte.

    Note: [guid]::NewGuid() is not a cryptographically secure RNG. Use this for
    test data / filling space, not for anything requiring unpredictable content.

.PARAMETER Path
    Full path of the file to create. Mandatory. If a file already exists at this
    path the script aborts unless -Force is specified. The parameter is passed with
    -LiteralPath, so wildcard characters ([, ], *, ?) in the path are treated
    literally.

.PARAMETER SizeMB
    Number of chunks to write. The final file size is SizeMB x chunkSize, so with
    the default 1 MB chunkSize this is the size in MB. Defaults to 500.

.PARAMETER FullyRandom
    Generate fresh random data for every chunk instead of reusing slices of a
    single pool. Produces fully unique content at the cost of speed.

.PARAMETER chunkSize
    Size in bytes of each block written to the file (one write per loop). Defaults
    to 1 MB. The final file size is SizeMB x chunkSize.

.PARAMETER poolSize
    Size in bytes of the random pool that chunks are sliced from in the default
    (pooled) strategy. Defaults to 8 MB. Must be at least chunkSize. Ignored in
    terms of content when -FullyRandom is used.

.PARAMETER Force
    Overwrite an existing file at -Path. Without it, the script aborts if the file
    already exists rather than overwriting it.

.EXAMPLE
    & ".\Create file with random content (CLM safe).ps1" -Path C:\Temp\filler.bin

    Creates a 500 MB file of random content using the fast pooled strategy.

.EXAMPLE
    & ".\Create file with random content (CLM safe).ps1" -Path D:\test\big.dat -SizeMB 2048

    Creates a 2 GB (2048 MB) file.

.EXAMPLE
    & ".\Create file with random content (CLM safe).ps1" -Path C:\Temp\unique.bin -SizeMB 100 -FullyRandom

    Creates a 100 MB file where every 1 MB block is freshly randomised (unique
    throughout), rather than reused slices of a pool.

.EXAMPLE
    & ".\Create file with random content (CLM safe).ps1" -Path C:\Temp\big.bin -SizeMB 200 -chunkSize 4MB -poolSize 32MB

    Writes 200 chunks of 4 MB each (an 800 MB file), slicing each chunk from a
    32 MB random pool.

.OUTPUTS
    None. Writes the file to -Path and prints a one-line summary (bytes written,
    full path, elapsed seconds, and the current language mode) to the host.

.NOTES
    CLM-safe: verified to run under Constrained Language Mode. The summary line
    reports the language mode it ran in.
#>

param
(
    [Parameter(Mandatory = $true)]
    [string] $Path,

    [int] $SizeMB = 500,

    [switch] $FullyRandom ,
    
    [int] $chunkSize = 1MB ,

    [int]$poolSize = 8MB ,

    [switch] $Force
)

if( ( Test-Path -LiteralPath $Path -PathType Leaf ) -and -not $Force )
{
    Throw "File `"$Path`" already exists - specify -Force to overwrite it"
}

$guidSize = [guid]::NewGuid().ToByteArray().Count
$useLatin1String = $PSVersionTable.PSVersion.Major -ge 7

function New-RandomBytes
{
    param
    (
        [int] $ByteCount
    )

    $randomBytes = [byte[]]::new($ByteCount)
    for ($offset = 0; $offset -lt $ByteCount; $offset += $guidSize)
    {
        [array]::Copy([guid]::NewGuid().ToByteArray(), 0, $randomBytes, $offset, $guidSize)
    }
    return , $randomBytes
}

function New-RandomPool
{
    param
    (
        [int] $ByteCount
    )

    $randomBytes = New-RandomBytes -ByteCount $ByteCount
    if ($useLatin1String)
    {
        # Latin-1 maps chars 0-255 to bytes 0-255 one-to-one.
        return [string]::new([char[]] $randomBytes)
    }
    return , $randomBytes
}

if ($useLatin1String)
{
    $writeParameters = @{ Encoding = 'Latin1'; NoNewline = $true }
    $emptyContent = ''
}
else
{
    $writeParameters = @{ Encoding = 'Byte' }
    $emptyContent = [byte[]]::new(0)
}

$startTime = Get-Date
Set-Content -LiteralPath $Path -Value $emptyContent @writeParameters

$pool = New-RandomPool -ByteCount $poolSize
if (-not $useLatin1String)
{
    $chunk = [byte[]]::new($chunkSize)
}

for ($chunkNumber = 1; $chunkNumber -le $SizeMB; $chunkNumber++)
{
    if ($FullyRandom)
    {
        $chunk = New-RandomPool -ByteCount $chunkSize
    }
    else
    {
        $sourceOffset = Get-Random -Minimum 0 -Maximum ($poolSize - $chunkSize + 1)
        if ($useLatin1String)
        {
            $chunk = $pool.Substring($sourceOffset, $chunkSize)
        }
        else
        {
            [array]::Copy($pool, $sourceOffset, $chunk, 0, $chunkSize)
        }
    }

    Add-Content -LiteralPath $Path -Value $chunk @writeParameters

    if ($chunkNumber % 10 -eq 0)
    {
        Write-Progress -Activity "Writing $Path" -Status "$chunkNumber / $SizeMB MB" -PercentComplete ($chunkNumber * 100 / $SizeMB)
    }
}

Write-Progress -Activity "Writing $Path" -Completed
$elapsed = (Get-Date) - $startTime
$file = Get-Item -LiteralPath $Path
Write-Host -Object ("Wrote {0:N0} bytes to {1} in {2:N1} s ({3})" -f $file.Length, $file.FullName, $elapsed.TotalSeconds, $ExecutionContext.SessionState.LanguageMode)
