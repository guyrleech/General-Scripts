<#
.SYNOPSIS

Find a given attribute by its name in a JSON structure

.DESCRIPTION

By default will search for a value name with the exact name specified but this can be switched to a regular expression match.
By default will return the first match it finds but this can be overridden

.PARAMETER inputObject

The JSON object to search for the given attribute name

.PARAMETER name

The attribute name to search for

.PARAMETER regex

The -name parameter becomes a regular expression to match against each attribute name

.PARAMETER multiple

Return every matching result rather than the first encountered

.EXAMPLE

Get-JSONAttribute -inputObject $someJSON -name username

Search the JSON object contained in $someJSON and return the first attribute where the name equals 'username'

.EXAMPLE

Get-JSONAttribute -inputObject $someJSON -name connection -regex -multiple

Search the JSON object contained in $someJSON and return all attributes whose name matches 'connection'

#>

Function Get-JSONAttribute
{
    [CmdletBinding()]

    Param
    (
        [Parameter(ValueFromPipeline,Mandatory=$true,HelpMessage='JSON object to search')]
        $inputObject ,
        [Parameter(Mandatory=$true,HelpMessage='JSON property name to search for')]
        [string]$name ,
        [switch]$multiple ,
        [switch]$regex
    )

    $foundIt = $null

    If( $inputObject -and ! [string]::IsNullOrEmpty( $name ) )
    {
        ForEach( $property in $inputObject.PSObject.Properties )
        {
            If( $property.MemberType.ToString() -eq 'NoteProperty' )
            {
                If( ( ! $regex -and $property.Name -eq $name ) -or ( $regex -and $property.Name -match $name ) )
                {
                    Return $property
                }
                Elseif( $property.Value -is [PSCustomObject] )
                {
                    If( ( $multiple -or ! $foundIt ) -and ( $result = Get-JSONAttribute -name $name -inputObject $property.value -multiple:$multiple -regex:$regex ))
                    {
                        $foundIt = $result
                        $result
                    }
                }
            }
        }
    }
}