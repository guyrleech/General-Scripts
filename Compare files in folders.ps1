#requires -version 4
<#
    Compare files by version, date, checksum and size between two folders. Also show files in one or other folder but not in both.

    If using a gridview for output, text files selected in that will then be differenced into separate grid views

    @guyrleech 2019

    Modification History:

    15/02/19  GRL  Added size difference column
#>

<#
.SYNOPSIS

Compare files in folders in two different folders. Compares first by file attributes and then if they are the same by checksum. Also shows folder in one folder or other but not both.
Text files selected in the grid view will then be differenced into separate grid views to show the actual differences between the two files.

.PARAMETER folder1

A folder containing one or more files.

.PARAMETER folder2

A folder containing one or more files.

.PARAMETER recurse

Recurse the folder structures. The default is to only do the top level.

.PARAMETER include

A comma separated list of regular expressions where only files that match one of these will be processed.

.PARAMETER exclude

A comma separated list of regular expressions where a file that matches any of these will not be processed.

.PARAMETER differentOnly

Only show files which are different or do not exist on ond of the folders

.PARAMETER date

Compare date stamps on files

.PARAMETER nogridview

Do not display the results in a grid view, write them to the standard output

.PARAMETER skipBlank

When showing the differences between two files, don't show blank lines

.PARAMETER compareBinary

Compare binary file types. The default does not. This may not work as expected.

.PARAMETER binaryTypes

A comma separated list of extensions that signify the file is has binary content and should not be compared unless -compareBinary is specified.

.EXAMPLE

& '.\Compare files in folders.ps1' -folder1 "C:\Temp\Office ADMX 4810.1000" -folder2 "C:\Temp\Office ADMX 4768.1000" -recurse -skipBlank -differentOnly

Compare all files in the specified folders and subfolders, except for datestamps, display in a gridview and show the text differences in any files selected when the "OK" button is clicked except for blank lines.

.EXAMPLE

& '.\Compare files in folders.ps1' -folder1 "C:\Temp\Office ADMX 4810.1000" -folder2 "C:\Temp\Office ADMX 4768.1000" -recurse -include '.adml$'

Compare all .adml files in the specified folders and subfolders, except for datestamps, display in a gridview and show the text differences in any files selected when the "OK" button is clicked.

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true)]
    [string]$folder1 ,
    [Parameter(Mandatory=$true)]
    [string]$folder2 ,
    [string[]]$include ,
    [string[]]$exclude ,
    [switch]$recurse ,
    [switch]$differentOnly ,
    [switch]$date ,
    [switch]$noGridView ,
    [switch]$skipBlank ,
    [switch]$compareBinary ,
    [string[]]$binaryTypes = @( 'ocx' , 'exe' , 'iso' , 'vmdk' , 'vhd' , 'vhdx' , 'dll' , 'zip' , 'tar' , 'tz' , 'gz' , 'tgz' , '7z' , 'docx' , 'doc' , 'xls' , 'xlsx' , 'docm' , 'ppt' , 'pptx' , 'mdb' , 'ldb' , 'pdf' , 'bmp' , 'png' , 'jpg' , 'gif' , 'tif' , 'tiff' , 'jpeg' , 'dmp' )
)

[hashtable]$params = @{ 'File' = $true ; 'Force' = $true }

if( $recurse )
{
    $params.Add( 'Recurse' ,$true )
}

[int]$startPathLength = $folder1.Length 
[int]$totalFiles = 0
[int]$excluded = 0

[System.Collections.ArrayList]$results = @( Get-ChildItem $folder1 @params | ForEach-Object `
{
    $file1 = $_
    [bool]$processIt = $false
    if( $include -and $include.Count )
    {
        ForEach( $item in $include )
        {
            if( $file1 -match $item )
            {
                $processIt = $true
                break
            }
        }
    }
    else
    {
        $processIt = $true
    }
    if( $exclude -and $exclude.Count )
    {
        [bool]$matched = $false
        ForEach( $item in $exclude )
        {
            if( $file1 -match $item )
            {
                $matched = $true
                break
            }
        }
        $processIt = ! $matched
    }
    if( $processIt )
    {
        $totalFiles++
        [bool]$sameVersion = $false
        [bool]$sameSize = $false
        [bool]$sameModificationTime = ! $date
        [bool]$sameChecksum = $false     
        ## Get Path relative to our starting point so we can find other file
        [string]$file2path = Join-Path $folder2 ($file1.FullName).Substring( $startPathLength )
        if( Test-Path $file2path -ErrorAction SilentlyContinue )
        {
            $file2 = Get-ChildItem $file2path -ErrorAction SilentlyContinue  
        }
        else
        {
            $file2 = $null
            $file2path = $null
        }
        Write-Verbose "`"$file1`" : `"$file2`""
        if( $file2 )
        {
            $sameVersion = (($file1.VersionInfo).FileVersion -eq ($file2.VersionInfo).FileVersion) 
            $sameSize = ($file1.Length -eq $file2.Length)
            $sameModificationTime = if( ! $date ) { $true } else { ($file1.LastWriteTime -eq $file2.LastWriteTime) }
            ## Only calculate checksum if other file properties are all equal thus far otherwise must be different checksums
            if( $sameSize -and $sameVersion )
            {
                $sameChecksum = ((Get-FileHash $file1.FullName).Hash -eq (Get-FileHash $file2.FullName).Hash)
            }
        }
        else
        {
            $file2 = $null
        }
        if( ! $differentOnly -or ! $sameVersion -or ! $sameModificationTime -or ! $sameSize -or ! $sameChecksum )
        {
            $result = [pscustomobject][ordered]@{ 
                'File1' = $file1.FullName 
                'File2' = $file2path
                'Type' = [System.IO.Path]::GetExtension( $file1.FullName ) -replace '^\.' , ''
                'File1 Version' = ( $file1 | Select -ExpandProperty VersionInfo | Select -ExpandProperty FileVersion ) ##$(if( $file1.PSObject.properties[ 'VersionInfo' ] ) { $file1.VersionInfo | Select -ExpandProperty FileVersion })
                'File1 modification time' = $file1.LastWriteTime
                'File1 size' = [int]$file1.Length
                'Size Difference' = [int]($file1.Length - ($file2 | Select -ExpandProperty Length))
                'File2 Version' = ( $file2 | Select -ExpandProperty VersionInfo | Select -ExpandProperty FileVersion ) #$(if( $file2 -and $file2.PSObject.properties[ 'VersionInfo' ] ) { $file2.VersionInfo | Select -ExpandProperty FileVersion })
                'File2 size' = [int]( $file2 | Select -ExpandProperty Length)
                'File2 modification time' = ($file2 | Select -ExpandProperty LastWriteTime)
                'Same Version' = $sameVersion
                'Same size' = $sameSize
                'Same checksum' = $sameChecksum
            }
            if( $date )
            {
                Add-Member -InputObject $result -MemberType NoteProperty -Name 'Same modification time' -value $sameModificationTime
            }
            $result
        }
    }
})

## also need to look for files in $folder2 not in $folder1 but don't need to compare properties if do exist since already done
$startPathLength = $folder2.Length 
Get-ChildItem $folder2 @params | ForEach-Object `
{
    $file2 = $_
    [bool]$processIt = $false
    if( $include -and $include.Count )
    {
        ForEach( $item in $include )
        {
            if( $file2 -match $item )
            {
                $processIt = $true
                break
            }
        }
    }
    else
    {
        $processIt = $true
    }
    if( $exclude -and $exclude.Count )
    {
        [bool]$matched = $false
        ForEach( $item in $exclude )
        {
            if( $file2 -match $item )
            {
                $matched = $true
                break
            }
        }
        $processIt = ! $matched
    }
    if( $processIt )
    {
        ## Get Path relative to our starting point so we can find other file
        [string]$file1path = Join-Path $folder1 ($file2.FullName).Substring( $startPathLength )
        if( ! ( Test-Path $file1path -ErrorAction SilentlyContinue ) )
        {
            $totalFiles++
            [void]$results.Add(
                ([pscustomobject][ordered]@{ 
                    'File1' = $null 
                    'File2' = $file2.FullName
                    'Type' = [System.IO.Path]::GetExtension( $file2.FullName ) -replace '^\.' , ''
                    'File1 Version' = $null
                    'File1 modification time' = $null
                    'File1 size' = $null
                    'File2 Version' = (($file2.VersionInfo).FileVersion) 
                    'File2 size' = $file2.Length
                    'File2 modification time' = $file2.LastWriteTime
                    'Same Version' = $false
                    'Same size' = $false
                    'Same checksum' = $false
                }))
         }
    }
}

[string]$title = ( "Found {0} different files out of {1} between `"{2}`" and `"{3}`"" -f $results.Count , $totalFiles , $folder1 , $folder2 )

Write-Verbose $title

if( ! $noGridView )
{
    [array]$selected = @( $results | Out-GridView -Title $title -PassThru )
    if( $selected -and $selected.Count )
    {
        ForEach( $differentFile in $selected )
        {
            if( $differentFile.file1 -and $differentFile.file2 )
            {
                if( ! $compareBinary -and $differentFile.Type -in $binaryTypes )
                {
                    Write-Warning "Skipping `"$($differentFile.file1)`" as its extension is for a binary file, not text"
                }
                else
                {
                    ## Include equal lines so can track lines numbers
                    [int]$linenumber = 1
                    [int]$file1diffs = 0
                    [int]$file2diffs = 0
                    [array]$fileDifferences = @( Compare-Object -ReferenceObject (Get-Content -Path $differentFile.file1) -DifferenceObject (Get-Content -Path $differentFile.file2) -IncludeEqual |  Sort-Object { $_.InputObject.ReadCount } | ForEach-Object `
                    {
                        if( $_.SideIndicator -match '^=[=>]$') 
                        { 
                            $lineNumber = $_.InputObject.ReadCount 
                        }
                        if( $_.SideIndicator -ne '==' )
                        {
                            Write-Verbose ('{0}: {1} {2}' -f $lineNumber , $_.SideIndicator , $_.InputObject )
                            [string]$file1line = ""
                            [string]$file2line = ""
                            if( $_.SideIndicator -eq '<=' )
                            {
                                $file1line = $_.InputObject
                                $file1diffs++
                            }
                            elseif( $_.SideIndicator -eq '=>' )
                            {
                                $file2line = $_.InputObject
                                $file2diffs++
                            }

                            if( ! $skipBlank -or ! [string]::IsNullOrEmpty( $file1line ) -or ! [string]::IsNullOrEmpty( $file2line ) )
                            {
                                [pscustomobject]@{
                                    'Around Line Number' = $lineNumber
                                    'File1' = $file1line
                                    'File2' = $file2line
                                }
                            }
                        }
                    })
                    if( $fileDifferences -and $fileDifferences.Count )
                    {
                        [string]$title = if( $file1diffs -and $file2diffs )
                        {
                                "$($fileDifferences.Count) differences"
                        }
                        elseif( $file1diffs )
                        {
                            "$($fileDifferences.Count) lines only in first file"
                        }
                        else
                        {
                            "$($fileDifferences.Count) lines only in second file"
                        }
                        $fileDifferences | Select -Property 'Around Line Number',@{n="$($differentFile.file1)";e={$_.File1}},@{n="$($differentFile.file2)";e={$_.File2}} | Out-GridView -Title $title
                    }
                    else
                    {
                        Write-Warning "No differences detected between `"$($differentFile.file1)`" and `"$($differentFile.file2)`""
                    }
                }
            }
            else
            {
                Write-Warning "Can only compare when there are two files: `"$($differentFile.file1)`" and `"$($differentFile.file2)`""
            }
        }
    }
}
else
{
    $results
}
