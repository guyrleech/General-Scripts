<#
    Perform a query on a SQLite database

    Requires SQLite libraries - https://sqlitebrowser.org/dl/

    @guyrleech 2019
#>

<#
.SYNOPSIS

Perform a select query on a specified SQLite database contained within a file or show the tables within the database.

.DESCRIPTION

Requires SQLite libraries - https://sqlitebrowser.org/dl/

.PARAMETER database

The path to an existing SQLite database file

.PARAMETER table

The name of the table to query. If not specified then all table names will be displayed.

.PARAMETER select

The comma separated list of columns to retrieve. Default is all.

.PARAMETER where

An optional where clause to limite the result set

.PARAMETER outputFile

The path to a non-existent csv file to retrieve the results. Will write to an on-screen sortable/filterable grid view if not specified.

.PARAMETER sqlite

The path to a 'System.Data.SQLite.dll' file in order to retrieve the required cmdlets

.EXAMPLE

& '.\Query SQLite database.ps1' -database "c:\temp\AnalysisDatabase.sqlite" -table PDM_Application -where "Computer like 'Laptop10.%'" -verbose

Retrieve all rows from the table 'PDM_Application' in the specified database file where the 'Computer' field is like 'Laptop10.' and display in a grid view

& '.\Query SQLite database.ps1' -database "c:\temp\AnalysisDatabase.sqlite" -table PDM_Application -outputFile apps.csv

Retrieve all rows from the table 'PDM_Application' in the specified database file and write to the file 'apps.csv'

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path -Path $_ -PathType 'Leaf' })] 
    [string]$database ,
    [string]$table ,
    [string]$select = '*' ,
    [string]$where ,
    [string]$outputFile ,
    [string]$sqlite = "$env:ProgramFiles\System.Data.SQLite\2010\bin\System.Data.SQLite.dll"
)

if( ! ( Get-ItemProperty -Path $database -ErrorAction Stop | Select -ExpandProperty Length ) )
{
    Throw "Database file `"$database`" is zero length"
}

Add-Type -Path $sqlite -ErrorAction Stop

$dbConnection = New-Object -TypeName System.Data.SQLite.SQLiteConnection
$dbConnection.ConnectionString = "Data Source=$database"
$dbConnection.Open()

## Get table names
$sql = $dbConnection.CreateCommand()
$adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql

if( ! $PSBoundParameters[ 'table' ] )
{
    $sql.CommandText = 'SELECT * FROM sqlite_master WHERE type=''table'''
    $data = New-Object System.Data.DataSet
    [void]$adapter.Fill($data)

    if( $data )
    {
        if( $data.Tables -and $data.Tables.Count )
        {
            $data.Tables | Format-Table -AutoSize
        }
        else
        {
            Write-Warning "No tables found in database"
        }
    }
    else
    {
        Write-Error 'Failed to retrieve table list'
    }
}
else
{
    $sql.CommandText = "SELECT $select FROM $table"
    
    if( $PSBoundParameters[ 'where' ] )
    {
        $sql.CommandText += " where $where"
    }
    $data = New-Object System.Data.DataSet
    [void]$adapter.Fill($data)
    try
    {
        if( $data.Tables -and $data.Tables.Rows -and $data.Tables.Rows.Count )
        {
            Write-Verbose ([string]$title = "Got $($data.Tables.Rows.Count) rows from query $($sql.CommandText)" )
            if( $PSBoundParameters[ 'outputFile' ] )
            {
                $data.Tables.Rows | Export-Csv -Path $outputFile -NoClobber -NoTypeInformation
            }
            else
            {
                $selected = $data.Tables.Rows | Out-GridView -Title $title -PassThru
                if( $selected -and $selected.Count )
                {
                    $selected | clip.exe
                }
            }
        }
        else
        {
            Write-Warning "No rows returned from query $($sql.CommandText)"
        }
    }
    catch
    {
        Write-Warning "Failed to get any results from query $($sql.CommandText)"
    }
}

if( $dbConnection )
{
    $dbConnection.Close()
    $dbConnection.Dispose()
}
