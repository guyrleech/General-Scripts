#requires -version 3
<#
    Examine files to see what type they really are

    @guyrleech 2019
#>

<#
.SYNOPSIS

Look at content of specifid files or files in specified folders and show their actual type based on this content.

.DESCRIPTION

Some file formats, such as Microsoft Office .docx and .xlsx are actually zip files so this script will identify a file's type based on its contents not file extension.
Pipe to Out-Gridview for an on screen sortable/filterable window or to Export-CSV to record to file.

.PARAMETER files

A comma separated list of files to process

.PARAMETER folders

A comma separated list of folders to process

.PARAMETER recurse

Recurse each specified folder otherwise just process items in each folder

.PARAMETER list

List the file types that this script will detect

.PARAMETER types

Only include file types which match this regular expression

.PARAMETER others

Include files which have not been identified by their content

.PARAMETER adfs

Also include files which contain files in NTFS alternate data streams

.PARAMETER onlyAdfs

Only include files which contain files in NTFS alternate data streams

.PARAMETER summary

Display a summary with the count of each file type, sorted by the most prevalent

.EXAMPLE

& '.\Find file type.ps1' -list

Show all the file types that the script can identify

.EXAMPLE

& '.\Find file type.ps1' -file c:\documents\somedocument.docx

Analyse the file c:\documents\somedocument.docx and if its contents match one of the known types then output this type it otherwise no output is produced

.EXAMPLE

& '.\Find file type.ps1' -folders c:\stuff -adfs -others -recurse

Analyse all files in c:\stuff including subfolders and any with alternate data streams. Files which don't match a known type via their content will also be output

.EXAMPLE

& '.\Find file type.ps1' -folders c:\stuff -onlyAdfs -recurse

Only analyse files in c:\stuff, and subfolders, with alternate data streams.

.EXAMPLE

& '.\Find file type.ps1' -folders c:\stuff -recurse -summary -types 'zip|rar'

Analyse all files in c:\stuff including subfolders and produce a summary of the file types which match 'zip' such as "pkzip" or "7zip" or 'rar'

#>

[CmdletBinding()]

Param
(
    [Parameter(ParameterSetName='files',Mandatory=$true)]
    [string[]]$files ,
    [Parameter(ParameterSetName='folders',Mandatory=$true)]
    [string[]]$folders ,
    [Parameter(ParameterSetName='folders',Mandatory=$false)]
    [switch]$recurse ,
    [switch]$list ,
    [string]$types ,
    [switch]$others ,
    [switch]$adfs ,
    [switch]$onlyAdfs ,
    [switch]$summary
)

## a great resource for information on headers of files can be found at http://file-extension.net/seeker/

[array]$magicNumbers = @(
    [pscustomobject]@{ 'Type' = 'pkzip' ; 'Offset' = 0 ; 'Bytes' = @( 0x50 , 0x4B , 0x03 , 0x04 ) }
    [pscustomobject]@{ 'Type' = 'exe'   ; 'Offset' = 0 ; 'Bytes' = @( 0x4D , 0x5A ) }
    [pscustomobject]@{ 'Type' = 'jpg'   ; 'Offset' = 6 ; 'Bytes' = @( 0x4A , 0x46 , 0x49 , 0x46 ) }
    [pscustomobject]@{ 'Type' = 'jpg'   ; 'Offset' = 6 ; 'Bytes' = @( 0x45 , 0x78 , 0x69 , 0x66 ) }
    [pscustomobject]@{ 'Type' = 'jpg'   ; 'Offset' = 6 ; 'Bytes' = @( 0xFF , 0xD8 , 0xFF ) }
    [pscustomobject]@{ 'Type' = 'pdf'   ; 'Offset' = 0 ; 'Bytes' = @( 0x25 , 0x50 , 0x44 , 0x46 , 0x2D) }
    [pscustomobject]@{ 'Type' = 'png'   ; 'Offset' = 0 ; 'Bytes' = @( 0x89 , 0x50 , 0x4e , 0x47 ) }
    [pscustomobject]@{ 'Type' = 'mp4'   ; 'Offset' = 4 ; 'Bytes' = @( 0x66 , 0x74 , 0x79 , 0x70 , 0x6D ) }
    [pscustomobject]@{ 'Type' = 'mp4'   ; 'Offset' = 4 ; 'Bytes' = @( 0x66 , 0x74 , 0x79 , 0x70 , 0x69 ) }
    [pscustomobject]@{ 'Type' = 'bmp'   ; 'Offset' = 0 ; 'Bytes' = @( 0x42 , 0x4d ) }
    [pscustomobject]@{ 'Type' = 'gif'   ; 'Offset' = 0 ; 'Bytes' = @( 0x47 , 0x49 , 0x46 ) }
    [pscustomobject]@{ 'Type' = 'wav'   ; 'Offset' = 8 ; 'Bytes' = @( 0x57 , 0x41 , 0x56 , 0x45 ) }
    [pscustomobject]@{ 'Type' = 'avi'   ; 'Offset' = 8 ; 'Bytes' = @( 0x41 , 0x56 , 0x49 ) }
    [pscustomobject]@{ 'Type' = 'gif'   ; 'Offset' = 0 ; 'Bytes' = @( 0x52 , 0x61 , 0x72 , 0x21 , 0x1A , 0x07 , 0x00 ) }
    [pscustomobject]@{ 'Type' = 'gzip'  ; 'Offset' = 0 ; 'Bytes' = @( 0x1f , 0x8B ) }
    [pscustomobject]@{ 'Type' = '7zip'  ; 'Offset' = 0 ; 'Bytes' = @( 0x37 , 0x7A ) }
    [pscustomobject]@{ 'Type' = 'rar'   ; 'Offset' = 0 ; 'Bytes' = @( 0x52, 0x61, 0x72, 0x21, 0x1A ) }
    [pscustomobject]@{ 'Type' = 'cab'   ; 'Offset' = 0 ; 'Bytes' = @( 0x49, 0x53, 0x63, 0x28 ) }
    [pscustomobject]@{ 'Type' = 'tif'   ; 'Offset' = 0 ; 'Bytes' = @( 0x49, 0x49, 0x2A, 0x00 ) }
    [pscustomobject]@{ 'Type' = 'tif'   ; 'Offset' = 0 ; 'Bytes' = @( 0x4D, 0x4D, 0x00 ) }
    [pscustomobject]@{ 'Type' = 'cab'   ; 'Offset' = 0 ; 'Bytes' = @( 0x4D, 0x53, 0x43, 0x46, 0x00, 0x00, 0x00, 0x00 ) }
    [pscustomobject]@{ 'Type' = 'wim'   ; 'Offset' = 0 ; 'Bytes' = @( 0x4D , 0x53 , 0x57 , 0x49 , 0x4D ) }
    [pscustomobject]@{ 'Type' = 'mkv'   ; 'Offset' = 0 ; 'Bytes' = @( 0x1A , 0x45 , 0xDF , 0xA3 ) }
    [pscustomobject]@{ 'Type' = 'wmv/wma'   ; 'Offset' = 0 ; 'Bytes' = @( 0x30, 0x26, 0xB2, 0x75, 0x8E, 0x66, 0xCF, 0x11, 0xA6, 0xD9, 0x00, 0xAA, 0x00 ) }
    [pscustomobject]@{ 'Type' = 'msi'   ; 'Offset' = 0 ; 'Bytes' = @( 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 )}
    [pscustomobject]@{ 'Type' = 'vhdx'   ; 'Offset' = 0 ; 'Bytes' = @( 0x76, 0x68, 0x64, 0x78, 0x66, 0x69, 0x6C, 0x65 )}
    [pscustomobject]@{ 'Type' = 'vhd'   ; 'Offset' = 0 ; 'Bytes' = @( 0x63, 0x6F, 0x6E, 0x65, 0x63, 0x74, 0x69, 0x78 )}
    ## uncomment and use with verbose to show hex bytes found for each file to help analyse file
    ##[pscustomobject]@{ 'Type' = 'dummy' ; 'Offset' = 32 ; 'Bytes' = @( 0xDE , 0xAD , 0xBF , 0xBE ) }
)

Function Get-FileType
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$file ,
        [AllowEmptyString()]
        [AllowNull()]
        [string]$types ,
        [long]$length ,
        [int]$maxHeaderLength ### including offset
    )
    
    ## Slower than using streams but streams don't work with alternate data streams
    $fileError = $null
    [byte[]]$bytes = Get-Content -LiteralPath $file -TotalCount $maxHeaderLength -Encoding Byte -ErrorAction SilentlyContinue -ErrorVariable $fileError

    if( $bytes )
    {
        if( $VerbosePreference -eq 'Continue' )
        {
            Write-Verbose "$file : $($bytes | ForEach-Object { "0x$($_.ToString('X2'))," } )"
        }
        [bool]$matched = $false
        ForEach( $magicNumber in $magicNumbers )
        {
            [bool]$match = $true
            For( [int]$index = $magicNumber.Offset ; $index -lt $magicNumber.Offset + $magicNumber.Bytes.Count -and $index -lt $bytes.Count ; $index++ )
            {
                if( $bytes[ $index ] -ne $magicNumber.Bytes[ $index - $magicNumber.Offset ] )
                {
                    $match = $false
                    break
                }
            }
            if( $match -and $index -lt $bytes.Count )
            {
                if( ! $types -or $magicNumber.Type -match $types )
                {
                    [pscustomobject]@{ 'File' = $file ; 'Type' = $magicNumber.Type ; 'Extension' = [System.IO.Path]::GetExtension( $file ) ; 'Size (KB)' = [int]( $length / 1KB )}
                }
                ## else not interested in this type
                $matched = $true
                break
            }
        }
        if( ! $matched -and $others )
        {
            [pscustomobject]@{ 'File' = $file ; 'Type' = $null ; 'Extension' = [System.IO.Path]::GetExtension( $file ) ; 'Size (KB)' =  [int]( $length / 1KB )}
        }
    }
    elseif( $fileError )
    {
        Write-Warning "Failed to read from `"$file`" : $filError"
    }
}

if( $onlyAdfs )
{
    $adfs = $true
}

[int]$maxHeaderLength = 0

if( $summary -and [string]::IsNullOrEmpty( $types ) )
{
    Throw 'Must specify the type to summarise via the -types argument'
}

ForEach( $magicNumber in $magicNumbers )
{
    if( $magicNumber.Bytes.Count + $magicNumber.Offset -gt $maxHeaderLength )
    {
        $maxHeaderLength = $magicNumber.Bytes.Count + $magicNumber.Offset
    }
}

if( $list )
{
    $magicNumbers|Select -ExpandProperty Type|sort -Unique
}
else
{
    $results = @( if( $PSCmdlet.ParameterSetName -eq 'folders' )
    {
        [hashtable]$params = @{ 'Path' = $folders ; 'File' = $true }
        $params.Add( 'Recurse' , $recurse )
        Get-ChildItem @params | ForEach-Object `
        {
            if( ! $onlyAdfs )
            {
                Get-FileType -file $_.FullName -maxHeaderLength $maxHeaderLength -type $types -length $_.Length
            }
            if( $adfs )
            {
                try
                {
                    $file = $_
                    ## $DATA is the file content itself
                    Get-Item -LiteralPath $file.FullName -Stream '*' | Where-Object { $_.Stream  -ne ':$DATA' } | ForEach-Object `
                    {
                        $alternate = $_
                        if( ! [string]::IsNullOrEmpty( $alternate.Stream ) ) 
                        {
                            [string]$streamPath = ( $file.FullName + ':' + $alternate.Stream )
                            $item = Get-Item -LiteralPath $streamPath
                            if( $item )
                            {
                                Get-FileType -file $streamPath -maxHeaderLength $maxHeaderLength -type $types -Length $item.Length
                            }
                        }
                    }
                }
                catch
                {
                    Write-Error "$($file.FullName) : $_"
                }
            }
        }
    }
    else
    {
        ForEach( $file in $files )
        {
            $item = Get-Item -LiteralPath $file
            if( $item )
            {
                Get-FileType -file $file -maxHeaderLength $maxHeaderLength -type $types -length $item.Length
            }
        }
    })

    if( $summary )
    {
        $groupByExtension = $results | Group-Object -Property 'Extension'
        $groupByExtension | Select-Object -Property Name,Count | Sort-Object -Property Count -Descending
    }
    else
    {
        $results
    }
}
