<div align="center">

## FormatFileSize


</div>

### Description

FormatFileSize: Formats a file's size in bytes into X GB or X MB or X KB or X bytes depending on size (a la Win9x Properties tab)

* UPDATED Sept. 12, 2000 * to allow for overriding the default Format Mask.
 
### More Info
 
dblFileSize: Double; File size in bytes

Optionally, pass a standard format string (e.g.: "###.##") in strFormatMask if you need to override the default formatting

Example:

"FormatFileSize(100)" will return "100 bytes"

"FormatFileSize(5500)" will return "5.4 KB"

"FormatFileSize(15000000)" will return "14.31 MB"

String


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeff Cockayne](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeff-cockayne.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeff-cockayne-formatfilesize__1-7331/archive/master.zip)





### Source Code

```
Public Function FormatFileSize(ByVal dblFileSize As Double, _
                Optional ByVal strFormatMask As String) _
                As String
' FormatFileSize:  Formats dblFileSize in bytes into
'          X GB or X MB or X KB or X bytes depending
'          on size (a la Win9x Properties tab)
Select Case dblFileSize
  Case 0 To 1023       ' Bytes
    FormatFileSize = Format(dblFileSize) & " bytes"
  Case 1024 To 1048575    ' KB
    If strFormatMask = Empty Then strFormatMask = "###0"
    FormatFileSize = Format(dblFileSize / 1024#, strFormatMask) & " KB"
  Case 1024# ^ 2 To 1073741823 ' MB
    If strFormatMask = Empty Then strFormatMask = "###0.0"
    FormatFileSize = Format(dblFileSize / (1024# ^ 2), strFormatMask) & " MB"
  Case Is > 1073741823#    ' GB
    If strFormatMask = Empty Then strFormatMask = "###0.0"
    FormatFileSize = Format(dblFileSize / (1024# ^ 3), strFormatMask) & " GB"
End Select
End Function
```

