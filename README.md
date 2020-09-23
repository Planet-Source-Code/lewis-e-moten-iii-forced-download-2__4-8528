<div align="center">

## Forced Download 2


</div>

### Description

Allows you to force a file to be downloaded rather then displayed within the users browser. This can be used with Word documents, Excel Spreadsheets, Adobe PDF's, and other files. Script has been optimized to support large downloads and be dynamically called to download any file (except ASP, ASPX, ASA, ASAX, MDB files). Also clears up some problems with currupt files users had been having by clearing all previouse content and headers.
 
### More Info
 
Pass the location to the file via. the query string.

Download.asp?FileName=/Files/Policy.doc


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Beginner
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-forced-download-2__4-8528/archive/master.zip)





### Source Code

```
<%
Dim Stream
Dim Contents
Dim FileName
Dim FileExt
Const adTypeBinary = 1
FileName = Request.QueryString("FileName")
If FileName = "" Then
	Response.Write "Filename not specified."
	Response.End
End If
' Make sure they are not requesting your code
FileExt = Mid(FileName, InStrRev(FileName, ".") + 1)
Select Case UCase(FileExt)
	Case "ASP", "ASA", "ASPX", "ASAX", "MDB"
		Response.Write "Protected file not allowed."
		Response.End
End Select
' Download the file
Response.Clear
Response.ContentType = "application/octet-stream"
Response.AddHeader "content-disposition", "attachment; filename=" & FileName
Set Stream = server.CreateObject("ADODB.Stream")
Stream.Type = adTypeBinary
Stream.Open
Stream.LoadFromFile Server.MapPath(FileName)
While Not Stream.EOS
	Response.BinaryWrite Stream.Read(1024 * 64)
Wend
Stream.Close
Set Stream = Nothing
Response.Flush
Response.End
%>
```

