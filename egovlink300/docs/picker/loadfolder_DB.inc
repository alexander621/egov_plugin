<%
Function LoadFolder(path)
  Dim sSql, oRst
  Dim objFSO, objDir, objFound, url, strNav, temp, pos, fullPath, blnFound, newPath

  'only tranlate to a local path if a virtual path is specified
  If InStr(1, path, ":") < 1 Then
    fullPath = Server.MapPath(path)
  Else
    fullPath = path
  End If

  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

  blnFound = objFSO.FolderExists(fullPath)
  If blnFound Then
	  Set objDir = objFSO.GetFolder(fullPath)
  End If

  strNav = ""
  If IsObject(objDir) And Err.Number = 0 Then
    sSql = "EXEC ListFolders " & Session("OrgID") & ", " & Session("UserID") & ", '" & path & "'"
    
    Set oRst = Server.CreateObject("ADODB.Recordset")
    oRst.Open sSql, Application("DSN"), 3, 1

    If Not oRst.EOF Then
      Do While Not oRst.EOF
        newpath = Server.URLEncode(oRst("FolderPath"))
        name = oRst("FolderName")

        If Left(name,1) <> "_" Then   'if dir starts with underscore that means it is hidden
          If Left(name,2) <> "z." Then 'if dir starts with z. that means it at the bottom of the list (sorting option)
            icon = "../../images/picker/folderclosed.gif"
          Else
            name = Right(name,Len(name)-2)
            icon = "images/help.gif"
          End If
         ' strNav = strNav & "<LI style=""position:relative;"" nodeType=""c"" id=""" & newpath & """ class=kid><IMG style=""filter:alpha(opacity=100);"" src=""" & icon &""" width=""18"" height=""18"" align=""absmiddle""> <A title=""" & name & """ href=""loadtree.asp?path="& newPath & """ target=""hiddenframe"">" & name & "</A><UL class=hdn></UL></LI>" & vbCrLf
         strNav = strNav & "<LI style=""position:relative;"" nodeType=""c"" id=""" & newpath & """ class=kid><IMG style=""filter:alpha(opacity=100);"" src=""" & icon &""" width=""18"" height=""18"" align=""absmiddle""> <A title=""" & name & """ href=""loadtree.asp?path="& newPath & """ target=""explorer"">" & name & "</A><UL class=hdn></UL></LI>" & vbCrLf
	End If
        
        oRst.MoveNext
      Loop
      oRst.Close
    End If
    Set oRst = Nothing

    For Each objFound In objDir.Files
      name = objFound.Name

      url = path & "/" & name
      url = Replace(url, " ", "%20")

      imgSrc = "images/document.gif"

      pos = InStr(1, name, ".")
      If pos > 0 Then
        Select Case Mid(name,pos+1,3)
          Case "doc"
            imgSrc = "images/msword.gif"
          Case "xls"
            imgSrc = "images/msexcel.gif"
          Case "ppt"
            imgSrc = "images/msppt.gif"
          Case "htm"
            imgSrc = "images/msie.gif"
          Case "pdf"
            imgSrc = "images/pdf.gif"
        End Select
        temp = Left(name,pos-1)
      Else
        temp = name
      End If
      
      ' CODE FOR FILES
      'strNav = strNav & "<LI style=""position:relative; filter: Alpha(opacity=50);"" nodeType=""a"" id=""" & url & """><IMG class=clsNoHand style=""filter:alpha(opacity=100);"" src=""" & imgSrc & """ width=""18"" height=""18"" align=""absmiddle""> <A title=""" & temp & """ target=""fraTopic"" href=""" & url & """>" & temp & "</A></LI>"
       strNav = strNav & "<LI style=""position:relative; filter: Alpha(opacity=50);"" nodeType=""a"" id=""" & url & """><IMG class=clsNoHand style=""filter:alpha(opacity=100);"" src=""" & imgSrc & """ width=""16"" height=""16"" align=""absmiddle""> <A title=""" & temp & """ href=""#"" onclick=""parent.document.all.FilePath.value='" & name & "'"">" & temp & "</A></LI>"
    Next

    Set objDir = Nothing
    Set objFSO = Nothing
     
    ' For New Folder 
    
    If strNav = "" Then strNav = "<li><font color=""#003366"">&nbsp;<i>(Empty)</i></font>"
    LoadFolder = strNav
  Else
    Set objDir = Nothing
    Set objFSO = Nothing
    Response.Write "Error Occured."
  End If
End Function
%>