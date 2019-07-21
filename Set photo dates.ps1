#requires -version 3
<#
    Get date created from image file metadata and set as the file's creation time

    Some code from https://nicholasarmstrong.com/2010/02/exif-quick-reference/

    @guyrleech 2019
#>


<#
.SYNOPSIS

Get the date/time created from image file metadata and set as the file's creation date/time

.DESCRIPTION

Note that generally only pictures taken with a digital camera will have the picture's creation time set in the file's metadata

.PARAMETER folders

A comma separated list of folder names to operate on

.PARAMETER files

A comma separated list of file names to operate on

.PARAMETER recurse

If specified will resource each of the given folders otherwise only the top level files are processed

.EXAMPLE

& '.\Set photo dates.ps1' -folders $env:userprofile\Pictures,f:\photos -recurse

Examine all files in the two folders given, and subfolders, and change the creation date to that found within the file's metadata, if present.

#>

[CmdletBinding()]

Param
(
    [Parameter(ParameterSetName='Folders')]
    [string[]]$folders ,
    [Parameter(ParameterSetName='Files')]
    [string[]]$files ,
    [Parameter(ParameterSetName='Folders')]
    [switch]$recurse
)

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Text")

Function Set-PhotoDate
{
    [CmdletBinding()]

    Param
    (
        [string]$filename
    )

    Write-Verbose "Examining `"$filename`" ..."

    try
    {
        $photo = [System.Drawing.Image]::FromFile( $filename )
    }
    catch
    {
        ## probably not a graphics file
        Write-Warning "Failed to get open file `"$filename`""
        $photo = $null
    }

    if( $photo )
    {
        $result = $null
        try
        {
            $dateProperty = $photo.GetPropertyItem( 0x9003 )
        }
        catch
        {
            Write-Warning "Failed to get date property from file `"$filename`""
            $dateProperty = $null
        }
        if( $dateProperty )
        {
            [string]$dateTaken = (New-Object System.Text.UTF8Encoding).GetString( $dateProperty.Value )
            if( ! [string]::IsNullOrEmpty( $dateTaken ) )
            {
                ## Seems it can get a trailing NULL character so strip this off
                while( $dateTaken.Length -and [int]$dateTaken[-1] -eq 0 )
                {
                    $dateTaken = $dateTaken.Substring( 0 , $dateTaken.Length - 1 )
                }

                $result = New-Object DateTime

                if( ! [datetime]::TryParseExact(
                    $dateTaken , 
                    'yyyy:MM:dd HH:mm:ss' , 
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [System.Globalization.DateTimeStyles]::None,
                    [ref]$result) )
                {
                    Write-Warning "Failed to parse date `"$dateTaken`" from `"$filename`""
                    $result = $null
                }
            }
         }
        $photo.Dispose() ## need to ensure not in use otherwise can't update
        $photo = $null
    
        if( $result )
        {
            $properties = Get-ItemProperty -Path $filename
            if( $properties )
            {
                if( $properties.CreationTime -ne $result )
                {
                     $properties.CreationTime = $result
                }
            }
        }
    }
}

if( $PSBoundParameters[ 'folders' ] )
{
    ForEach( $folder in $folders )
    {
        Get-ChildItem -Path $folder -File -Recurse:$recurse | . `
        {
            Process `
            {
                Set-PhotoDate -filename $_.FullName
            }
        }
    }
}
elseif( $PSBoundParameters[ 'files' ] )
{
    $files | . `
    {
        Process `
        {
            Set-PhotoDate -filename $_
        }
    }
}
