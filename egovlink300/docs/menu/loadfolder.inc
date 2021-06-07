<%
Function LoadFolder(path)
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
    For Each objFound in objDir.SubFolders
      If Left(objFound.Name,1) <> "_" Then
        If Left(objFound.Name,2) <> "z." Then
          temp = objFound.Name
          icon = "images/book_closed.gif"
        Else
          temp = Right(objFound.Name,Len(objFound.Name)-2)
          icon = "images/help.gif"
        End If
        newPath = path & "/" & Server.URLEncode(objFound.Name)
        strNav = strNav & "<LI style=""position:relative;"" nodeType=""c"" id=""" & newPath & """ class=kid><IMG style=""filter:alpha(opacity=100);"" src=""" & icon &""" width=""16"" height=""16"" align=""absmiddle""> <A title=""" & temp & """ href=""loadtree.asp?path="& newPath & """ target=""hiddenframe"">" & temp & "</A><UL class=hdn></UL></LI>" & vbCrLf
      End If
    Next

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

      strNav = strNav & "<LI style=""position:relative; filter: Alpha(opacity=50);"" nodeType=""a"" id=""" & url & """><IMG class=clsNoHand style=""filter:alpha(opacity=100);"" src=""" & imgSrc & """ width=""16"" height=""16"" align=""absmiddle""> <A title=""" & temp & """ target=""fraTopic"" href=""" & url & """>" & temp & "</A></LI>"
    Next

    Set objDir = Nothing
    Set objFSO = Nothing

    If strNav = "" Then strNav = "<li><font color=""#003366"">&nbsp;<i>(Empty)</i></font>"
    LoadFolder = strNav
  Else
    Set objDir = Nothing
    Set objFSO = Nothing
    Response.Write "Error Occured."
  End If
End Function
%>



