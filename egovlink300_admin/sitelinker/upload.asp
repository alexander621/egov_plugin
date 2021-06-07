
<%
Response.Buffer = True

On Error Resume Next

Dim objUpload
Dim strMessage
Dim blnSuccessful
blnSuccessful = "No"  

Set objUpload = Server.CreateObject("Dundas.Upload.2")
If Err.Number <> 0 Then
	Response.Redirect "addarticle.asp?task=ADD&method=upload&Message=" & Err.Description
End If

objUpload.MaxFileSize = 4096000 '4 MB
objUpload.SaveToMemory

Dim temp

If Err.Number <> 0 then
	strMessage = "Sorry, but the following error occurred: " & Err.Description
  Response.Write strMessage
Else


  Response.Write strFullPath

  Dim strPath, strFilename, strFullPath
  strPath = objUpload.Form("currentfolderpath")
  'strPath = Application("eCapture_AppPath")
  strFilename = objUpload.Files(0).OriginalPath
  strFileName = Mid(strFilename, InStrRev(strFilename, "\")+1)
  strFullPath = Server.MapPath(strPath) & "\" & strFilename
  session("MyDoc") = strPath & "/" & strFileName ' Used in Recording Database entry
 
  
  If objUpload.Form("chkOverwrite") = "on" Then
    If objUpload.FileExists( strFullPath ) Then objUpload.FileDelete( strFullPath )
    objUpload.Files(0).SaveAs( strFullPath )
    Response.Write "Uploaded : " & strFullPath
    blnSuccessful = "Yes"   
  Else
    If objUpload.FileExists( strFullPath ) Then
      strFileName = ""
    Else
      objUpload.Files(0).SaveAs( strFullPath )
      Response.Write "Uploaded : " & strFullPath
      blnSuccessful = "Yes"  
   
    End If
  End If

  'strMessage = ""
End If

Set objUpload = Nothing


' DONE UPLOADING...PROCESS DB AND RETURN
IF blnSuccessful = "Yes" Then
	' IF UPLOAD SUCCESSFUL REDIRECT TO A HIDDEN ASP PAGE TO UPDATE DATABASE
	Response.Redirect "adduploadtodb.asp?strTitle=" & strFileName & "&path=" & strPath & "&task=ADD&method=upload&Message=" & strMessage & "Upload=" & blnSuccessful
Else
	' UPLOAD WAS NOT SUCCESFUL RETURN TO ADD PAGE WITH MESSAGE
	Response.Redirect "default.asp?strTitle=" & strFileName & "&task=ADD&method=upload&Message=" & strMessage  
End If

%>
