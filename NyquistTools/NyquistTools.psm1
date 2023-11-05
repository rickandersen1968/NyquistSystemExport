
using namespace System.Web


<#
.SYNOPSIS
Parse a Bogen System Controller (C4000 or E7000) system report (System Parameters -> Export Report)
and returns it as a collection of custom objects, one per report "page".
#>
function Import-NyquistReport
{
    [CmdletBinding()]
    param 
    (
        [Parameter(ValueFromPipeline,Mandatory=$true)]
        [String] $Path,

        [string] $Filter = '*'
    )

    Process
    {
        # Import the XML report file
        [xml] $rptXml = Get-Content -Path $Path 

        # Convert to a hashtable of reports, each 
        $rpts = foreach ( $ws in $rptXml.Workbook.Worksheet )
        {
            $cols = $ws.Table.Row | Select-Object -First 1 -Property @{ Name='Columns'; Expression={ $_.Cell.Data.'#text' } } | Select-Object -ExpandProperty Columns
            $entry = [PSCustomObject] @{
                # https://learn.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-pscustomobject?view=powershell-7.3#using-defaultpropertyset-the-long-way
                PSTypeName = 'BogenReport.' + $ws.Name
                Name = $ws.Name
                Columns = $cols
                Data = ($ws.Table.Row | Select-Object -Skip 1 | ForEach-Object { $_.Cell.Data.'#text' -join "`t" }) | ConvertFrom-Csv -Delimiter "`t" -Header $cols 
            }

            # Append the worksheet name to each report entry (potentially useful when filtering, sorting, and organizing report data). 
            $entry.Data | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name Worksheet -Value $ws.Name } 

            $entry 
        } 

        # Optionally filter the reports to be returned based on the specified filter.
        if ( $Filter ) 
        { 
            $rpts = $rpts | Where-Object Name -like $Filter 
        } 

        # Add a NoteProperty member to each report row. This allows flexibility in formatting and filtering.
        # $rpts | ForEach-Object { 
        #     $ws = $_.Name
        #     $_.Data | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name Worksheet -Value $ws } 
        # }

        $rpts
    }
}


<#
.SYNOPSIS
Convert the Nyquist report to one of several formats:

    * Array of objects
    * HTML text
    * Markdown text
    * Console-formatted text
#>
function ConvertFrom-NyquistReport
{
    [CmdletBinding()]
    param 
    (
        [Parameter(ValueFromPipeline,Mandatory=$true)]
        $InputObject,

        [Parameter(ParameterSetName='Markdown')]
        [switch] $Markdown,

        [Parameter(ParameterSetName='Html')]
        [switch] $Html,

        [Parameter(ParameterSetName='Show')]
        [switch] $Show
    )

    Process
    {
        if ( $Markdown ) 
        { 
            $specialCharsRegex = '`|\*|_|{|}|\[|]|<|>|\(|\)|\#|\+|-|\.|!|\|'

            $md = foreach ( $rpt in $InputObject )
            { 
                ""
                "## $($rpt.Name)"
                ""

                $data = $rpt.Data

                # Encode the table (markdown)
                "| {0} |" -f ($rpt.Columns -join ' | ')
                "| {0} |" -f (($rpt.Columns | ForEach-Object { '---' }) -join ' | ')
                foreach ( $row in $data )
                { 
                    $cols = ($rpt.Columns | ForEach-Object { $row.$_ -replace $specialCharsRegex,'\$&' -replace '\|','&#124;' }) -join ' | '
                    "| {0} |" -f $cols
                }

                # # Encode the table (HTML)
                # "<table>`n" + "  <tr>`n" + ($rpt.Columns | ForEach-Object { "  <th>$_</th>`n" } ) + " </tr>`n"
                # foreach ( $row in $data )
                # { 
                #     $cols = $rpt.Columns | ForEach-Object { "    <td>{0}</td>" -f ([System.Web.HttpUtility]::HtmlEncode( $row.$_ )) } 
                #     "  <tr>`n{0}  </tr>`n" -f ($cols -join "`n")
                # }
                # "</table>`n"
            } 

            $md -join [Environment]::NewLine 
        }
        elseif ( $Html )
        {
            $htmlBody = foreach ( $rpt in $InputObject )
            { 
                ""
                "<h2>$($rpt.Name)</h2>"
                ""

                $data = $rpt.Data

                # Encode the table (HTML)
                "<table>`n  <tr>`n" + ($rpt.Columns | ForEach-Object { "    <th>{0}</th>`n" -f [System.Web.HttpUtility]::HtmlEncode($_) } ) + " </tr>`n"
                foreach ( $row in $data )
                { 
                    $cols = $rpt.Columns | ForEach-Object { "    <td>{0}</td>" -f ([System.Web.HttpUtility]::HtmlEncode( $row.$_ )) } 
                    "  <tr>`n{0}`n  </tr>`n" -f ($cols -join "`n")
                }
                "</table>`n"
            } 

            $htmlBody -join [Environment]::NewLine 
        }
        elseif ( $Show )
        {
            # If Show was specified, format the output, else return the reports as is. 
            $InputObject | ForEach-Object { "## $($_.Name)" | Show-Markdown; $_.Data | Format-Table -AutoSize } 
#            $InputObject | ForEach-Object { "$([char]27)[4m" + $_.Name + "$([char]27)[0m" | Write-Host -ForegroundColor Red; $_.Data | Format-Table -AutoSize } 
        }
    }
}