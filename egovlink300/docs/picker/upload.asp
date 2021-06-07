<%
Option Explicit
Response.Buffer = True

Dim oUpload
Dim sMessage
Dim sPath, sFilename, sFullPath

On Error Resume Next

Set oUpload = Server.CreateObject("Dundas.Upload.2")
If Err.Number <> 0 Then
	Response.Redirect "addarticle.asp?task=ADD&method=upload&Message=" & Err.Description
End If

oUpload.MaxFileSize = 1048576 * 8 '8 MBs
oUpload.SaveToMemory

If Err.Number <> 0 then
	sMessage = "Sorry, but the following error occurred: " & Err.Description
  Response.Write strMessage
Else
  sPath = oUpload.Form("txtTopic")
  sFilename = oUpload.Files(0).OriginalPath
  sFilename = Mid(sFilename, InStrRev(sFilename, "\")+1)
  sFullPath = Server.MapPath(sPath) & "\" & sFilename
  Response.Write sFullPath
  
  If oUpload.Form("chkOverwrite") = "on" Then
    If oUpload.FileExists( sFullPath ) Then oUpload.FileDelete( sFullPath )
    oUpload.Files(0).SaveAs( sFullPath )
    Response.Write "Uploaded : " & sFullPath
  Else
    If oUpload.FileExists( sFullPath ) Then
      sFilename = ""
    Else
      oUpload.Files(0).SaveAs( sFullPath )
      Response.Write "Uploaded : " & sFullPath
    End If
  End If

  sMessage = ""
End If

Set oUpload = Nothing

Response.Redirect "addarticle.asp?strTitle=" & sFileName & "&task=ADD&method=upload&Message=" & sMessage
%>