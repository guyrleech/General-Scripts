#Requires -version 3.0

<#
    Get Twitter statistics such as tweets and followers and output to a csv without needing Twitter API

    Guy Leech 2018
#>

<#
.SYNOPSIS

Parse html returned from Twitter web site for specific handles, find information like tweets, likes and followers and output to a csv file

.PARAMETER handles

A comma separated list of Twitter handles to fetch statistics for

.PARAMETER csv

CSV file to output statistics to. If the file exists then it will be appended to

.PARAMETER URL

The base Twitter URL to use

.PARAMETER exclude

Information to exclude

.PARAMETER className

A regular expression which matches the class of the HTML tag which contains the required data. There should be no need to change this.

.PARAMETER tagName

The HTML tag which contains the required data. There should be no need to change this.

.PARAMETER regex

The regular expression used to match and parse the elements in the HTML that contain the required data. There should be no need to change this.

.EXAMPLE

& '.\Twitter Statistics.ps1' -csv $env:HOME\twitter.stats.csv -handles guyrleech,citrix

Gather statistics for @guyrleech and @Citrix and write to the file twitter.stats.csv in the user's home directory

.NOTES

Implement as a scheduled task to keep updating the same csv file with statistics for twitter accounts of your choices so you can see data like followers increases (or decreases).
Does not require the Twitter API.
Run with -verbose to show statistics as they are gathered

#>

[CmdletBinding()]

Param
(
    [Parameter(mandatory=$true,helpmessage='Comma separated list of twitter handles to interrogate')]
    [string[]]$handles ,
    [Parameter(mandatory=$true,helpmessage='CSV to output/append to')]
    [string]$csv ,
    [string]$URL = 'https://twitter.com/' ,
    [string]$exclude = '^List' ,
    ## There shouldn't be a need to change these unless twitter change the html
    [string]$className = '^ProfileNav-item ProfileNav-item--[ftl]' ,
    [string]$tagName = 'li' ,
    [string]$regex = 'title="([\d,]+)\s([a-z]+)"' 
)

## work around arrays not being passed correctly from a scheduled task when using Powershell.exe -File
if( $handles.Count -eq 1 -and $handles[0].IndexOf(',') -ge 0 )
{
    $handles = $handles -split ','
}

$results = @( ForEach( $handle in $handles )
{
    Write-Verbose "Fetching $handle data ..."
    try
    {
        $response = Invoke-WebRequest -Uri ( $url + $handle )
        $result = $null
        $response.ParsedHtml.getElementsByTagName($tagName) | Where-Object { $_.className -match $className } | select -ExpandProperty innerHTML | ForEach-Object `
        {
            [string]$text = ($_.Trim()) -replace "[`n`r]" , ''
            ## Lines will be like this so isolate title contents
            ## <A tabIndex=0 title="22,733 Tweets" class="ProfileNav-stat ProfileNav-stat--link
            if( $text -match $regex )
            {
                if( ! $result )
                {
                    $result = New-Object -TypeName PSCustomObject -Property (@{ 'Date' = Get-Date ; 'Handle' = $handle ;  })
                }
                ## Have to save matches since we are using a regex again
                [string]$fieldName = $matches[2]
                [string]$value = $Matches[1] -replace '\D' , '' ## remove commas and other non-numeric characters
                if( [string]::IsNullOrEmpty( $exclude ) -or $fieldName -notmatch $exclude )
                {
                    Write-Verbose "$handle : $fieldName = $value"
                    Add-Member -InputObject $result -MemberType NoteProperty -Name $fieldName -Value $value
                }
                else
                {
                    Write-Warning "Excluding $fieldName (Value=$value) for $handle"
                }
            }
        }
        if( $result )
        {
            $result
        }
        else
        {
            Write-Warning "Failed to get tags $tagName, class $className for $handle"
        }
    }
    catch
    {
        Write-Error "Failed to get page for $handle from $URL :`n$_"
    }
} )

if( $results.Count -and ! [string]::IsNullOrEmpty( $csv ) )
{
    $results | Export-Csv -Append -NoTypeInformation -Path $csv
}
