# NyquistSystemExport

This project is for experimenting with the Nyquist System Export report and possibly other Nyquist-related functionality.

## Adding XML Transform to report
One idea is to define an XSL transform that will convert the XML file to an HTML page. This has potential issues, but has potential.

To view the Nyquist System Parameters report using this XSL conversion, add the following line to the XML file on the second line, immediately after `<?xml version="1.0">`:
  ```
  <?xml-stylesheet type="text/xsl" href="system_export.xsl"?>
  ```

## PowerShell module
Another idea is to define a PowerShell module that can parse a System Export report and convert it to any of multiple formats to allow the user to view and and manipulate the report data.

The NyquistTools module is a prototype of that module. The following shows a few examples of how the report can be viewed. Note that the report file, `system_export.xml`, must first be downloaded and stored locally.

```powershell
Import-NyquistReport -Path '.\system_export (fixed).xml' -Filter S* | ConvertFrom-NyquistReport -Show
```

```powershell
TODO
```
