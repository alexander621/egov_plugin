<%
Function LoadFolder(vpath,ecpath)
  Dim sSql, oRst
  Dim objFSO, objDir, objFound, url, strNav, temp, pos, fullPath, blnFound, newPath, name, icon, bCanSeeSecure
  

  ' ONLY TRANLATE TO A LOCAL PATH IF A VIRTUAL PATH IS SPECIFIED
  If InStr(1, path, ":") < 1 Then
    fullPath = Server.MapPath(vpath)
  Else
    fullPath = path
  End If


  ' SET PERMISSION LEVEL EXPLICITLY FOR PUBLIC SITE
  bCanSeeSecure = True
  
 
  ' LOOP THRU DATABASE AND FILESYSTEM LISTING FOLDERS AND FILES
  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
  blnFound = objFSO.FolderExists(fullPath)
  If blnFound Then
	  Set objDir = objFSO.GetFolder(fullPath)
  else
  	response.write "File Path (fullpath) not found for: <b>" & fullPath & "</b>."
  End If


  strNav = ""
  If IsObject(objDir) And Err.Number = 0 Then
    sSQL = "EXEC ListFolders " & iorgid & ", 162, '" & ecpath & "'"
   ' response.write sSQL
   ' response.end

   
    Set oRst = Server.CreateObject("ADODB.Recordset")
    oRst.Open sSql, Application("DSN"),3,1

    If Not oRst.EOF Then
      For i = 1 to 6
	
	If Not oRST.EOF Then
		newpath = Server.URLEncode(oRst("FolderPath"))
		name = oRst("FolderName")

		' CHECK FOR INTERNAL FOLDER ACCESS
		
		response.write "<!--" & iorgid & request.Cookies("userid") & newpath & "-->" & vbcrlf 
		
		If HasAccess(iorgid,request.Cookies("userid"),oRst("FolderPath")) Then
		
			If Left(name,1) <> "_" Then   'if dir starts with underscore that means it is hidden
				  If Left(name,2) <> "z." Then 'if dir starts with z. that means it at the bottom of the list (sorting option)
					    If oRst("Secure") > 0 And bCanSeeSecure Then
					      icon = "images/locked_folder_closed.gif"
					    Else
					      icon = "images/folder_closed.gif"
					    End If
				  Else
					    name = Right(name,Len(name)-2)
					    icon = "images/help.gif"
				  End If
				  
				  strNav = strNav & "<LI><A title=""" & name & """ href=""docs/menu/home.asp?path="& newPath & """>" & name & "</A></LI>" & vbCrLf
			End If

		End If
		
		
	      oRst.MoveNext
      End If

      Next
      
      oRst.Close
    else
    	response.write "FolderName"
    End If
    Set oRst = Nothing


    Set objDir = Nothing
    Set objFSO = Nothing

    If strNav = "" Then strNav = "<li><font color=""#003366""><img src=""images/spacer.gif"" width=0 height=18 align=""absmiddle"">&nbsp;&nbsp;&nbsp;&nbsp;<i>(Empty)</i></font></li>"
    LoadFolder = strNav
  Else
    Set objDir = Nothing
    Set objFSO = Nothing
    Response.Write "Error Occured."
  End If
End Function


'-------------------------------------------------------------------------------------------------------
' FUNCTION HASACCESS(IORGID,IUSERID,STRVPATH)
'-------------------------------------------------------------------------------------------------------
Function HasAccess(iorgid,iuserid,strvpath)

	  On Error Resume Next

	  iReturnValue = False
	  response.write "<!--" & sSql & "-->" & vbcrlf
	  Set oCnn = Server.CreateObject("ADODB.Connection")
	  oCnn.Open Application("DSN")
	  sSql = "EXEC CHECKFOLDERACCESS '" & iorgid & "','" & iuserid & "','" & strvpath & "'"
	  Set rstAccess = oCnn.Execute(sSql)

	  If NOT rstAccess.EOF Then
		If rstAccess("folderid") >= 0 Then
			iReturnValue = True
		End If
	  End If

	  oCnn.Close
	  Set rstAccess = Nothing
	  Set oCnn = Nothing

	  HasAccess = iReturnValue

End Function

%>