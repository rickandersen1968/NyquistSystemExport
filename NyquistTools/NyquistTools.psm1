
using namespace System.Web


<#
.SYNOPSIS
Parse a Bogen System Controller (C4000 or E7000) system report (System Parameters -> Export Report)
and returns it as a collection of custom objects, one per report "page".

.EXAMPLE
Import-NyquistReport -Path 'system_export.xml' -Filter A* -Show

.EXAMPLE
Import-NyquistReport -Path 'system_export.xml' -Filter System* | foreach { $_.Data | Out-GridView -Title $_.Section }

Out-GridView is only available on Windows

.EXAMPLE
Import-NyquistReport -Path 'system_export.xml' -Filter A* | foreach { $_.Data | Out-GridView -Title $_.Section }

Out-GridView is only available on Windows

.EXAMPLE
Import-NyquistReport -Path 'system_export.xml' -Filter A* | ConvertFrom-NyquistReport -Markdown | Join-String | Show-Markdown -UseBrowser

Show-Markdown is only available on PowerShell 6.0+
#>
function Import-NyquistReport
{
    [CmdletBinding()]
    param 
    (
        # Path to the Nyquist report XML file.
        [Parameter(ValueFromPipeline,Mandatory=$true)]
        [String[]] $Path,

        # Pattern-matched name of the report(s) to be returned. 
        [string] $Filter = '*',

        # Formats the report(s) for display at the console.
        [switch] $Show
    )

    Process
    {
        foreach ( $rptPath in $Path )
        {
            # Import the XML report file
            [xml] $rptXml = Get-Content -Path $rptPath 

            # Convert to a hashtable of reports
            $report = foreach ( $ws in $rptXml.Workbook.Worksheet | Where-Object Name -like $Filter )
            {
                $cols = $ws.Table.Row | Select-Object -First 1 -Property @{ Name='Columns'; Expression={ $_.Cell.Data.'#text' } } | Select-Object -ExpandProperty Columns
                $worksheet = [PSCustomObject] @{
                    # https://learn.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-pscustomobject#using-defaultpropertyset-the-long-way
                    PSTypeName = 'Bogen.Nyquist.Report' 
                    Section = $ws.Name
                    Columns = $cols
                    Data = ($ws.Table.Row | Select-Object -Skip 1 | ForEach-Object { $_.Cell.Data.'#text' -join "`t" }) | ConvertFrom-Csv -Delimiter "`t" -Header $cols 
                }

                # Append the worksheet name to each report entry (potentially useful when filtering, sorting, and organizing report data). 
                #$entry.Data | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name Worksheet -Value $ws.Name } 
                #$entry.Data | Add-Member -MemberType NoteProperty -Name Worksheet -Value $ws.Name 
                $worksheet.Data | Add-Member -NotePropertyMembers @{ Worksheet = $ws.Name } -TypeName 'Bogen.Nyquist.ReportData'

                $worksheet 
            } 

            # Either return the entire report or just the data
            if ( $Show ) 
            { 
                # $report | ForEach-Object `
                # { 
                #     "## $($_.Section)" | Show-Markdown
                #     $_ | Select-Object -ExpandProperty Data | Select-Object -ExcludeProperty worksheet | Format-Table 
                # }
                $report | ConvertFrom-NyquistReport -Show
            }
            else 
            { 
                $report 
            }
        }
    }
}


<#
.SYNOPSIS
Convert the specified Nyquist report to one of several formats:

    * HTML text
    * Markdown text
    * Console-formatted text
#>
function ConvertFrom-NyquistReport
{
    [CmdletBinding()]
    param 
    (
        # The report(s) to be converted.
        [Parameter(ValueFromPipeline,Mandatory=$true)]
        [object[]] $InputObject,

        # Convert the report(s) to Markdown format. 
        # This can be displayed in a browser by piping to `| Join-String | Show-Markdown -UseBrowser`
        # or by saving it to a file (e.g., `| Out-File test.md`), which can be viewed in any Markdown-capable viewer. 
        [Parameter(ParameterSetName='Markdown')]
        [switch] $Markdown,

        # Convert the report(s) to HTML, which can be saved to a file and viewed in a browser.
        [Parameter(ParameterSetName='Html')]
        [switch] $Html,

        # Formats the report(s) for display at the console.
        [Parameter(ParameterSetName='Show')]
        [switch] $Show
    )

    Process
    {
        foreach ( $report in $InputObject )
        {
            if ( $Markdown ) 
            { 
                $specialCharsRegex = '`|\*|_|{|}|\[|]|<|>|\(|\)|\#|\+|-|\.|!|\|'

                $md = foreach ( $section in $report )
                { 
                    "## $($section.Section)`n"

                    # Encode the table (markdown)
                    "| {0} |" -f ($section.Columns -join ' | ')
                    "| {0} |" -f (($section.Columns | ForEach-Object { '---' }) -join ' | ')
                    foreach ( $row in $section.Data )
                    { 
                        $cols = ($section.Columns | ForEach-Object { $row.$_ -replace '\|','&#124;' -replace $specialCharsRegex,'\$&' }) -join ' | '
                        "| {0} |" -f $cols
                    }

                    "`n"

                    # # Encode the table (HTML)
                    # "<table>`n" + "  <tr>`n" + ($rpt.Columns | ForEach-Object { "  <th>$_</th>`n" } ) + " </tr>`n"
                    # foreach ( $row in $rpt.Data )
                    # { 
                    #     $cols = $rpt.Columns | ForEach-Object { "    <td>{0}</td>" -f ([System.Web.HttpUtility]::HtmlEncode( $row.$_ )) } 
                    #     "  <tr>`n{0}  </tr>`n" -f ($cols -join "`n")
                    # }
                    # "</table>`n"
                } 

                ($md -join [Environment]::NewLine) + [Environment]::NewLine
            }
            elseif ( $Html )
            {
                $htmlBody = foreach ( $section in $report )
                { 
                    ""
                    "<h2>$($section.Section)</h2>"
                    ""

                    $data = $section.Data

                    # Encode the table (HTML)
                    "<table>`n  <tr>`n" + ($section.Columns | ForEach-Object { "    <th>{0}</th>`n" -f [System.Web.HttpUtility]::HtmlEncode($_) } ) + " </tr>`n"
                    foreach ( $row in $data )
                    { 
                        $cols = $section.Columns | ForEach-Object { "    <td>{0}</td>" -f ([System.Web.HttpUtility]::HtmlEncode( $row.$_ )) } 
                        "  <tr>`n{0}`n  </tr>`n" -f ($cols -join "`n")
                    }
                    "</table>`n"
                } 

                ($htmlBody -join [Environment]::NewLine) + [Environment]::NewLine
            }
            elseif ( $Show )
            {
                # If Show was specified, format the output.
                $report | ForEach-Object { 
                    if ( $PSVersionTable.PSVersion -ge [Version]::new('6.0') )
                    {
                        "## $($_.Section)" | Show-Markdown
                    }
                    else 
                    {
                        # '=' * ($_.Section.Length + 6)
                        "## $($_.Section) ##"
                    }

                    $_.Data | Format-Table 
                } 
            }
        }
    }
}