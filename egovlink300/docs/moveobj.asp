<!-- #include file="../includes/common.asp" //-->
<html>
<head>
  <title>Move Object</title>
  <link href="global.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#ffffff">
  <%
  Dim fs, src, dest, typ, strType

  src = Request("s")
  dest = Request("d")
  typ = Request("t")
  name = Request("name")


'response.write src & "<br>" & dest
'response.end

  Set fs = Server.CreateObject("Scripting.FileSystemObject")
  
    'JS 03/25/02 ADD CODE HERE FOR PERMISION CHECKING
  If HasPermission("CanEditDocuments") Then
	If typ = "a" Then

		'response.write src
		'response.write dest
		'response.write "<br>"
		arrDest = split(dest,"/")
		fulldest = dest
		dest = ""
		'response.write UBOUND(arrDest)-1
		'response.write "<br>"
		For x = 0 to UBOUND(arrDest) - 1
			dest = dest & arrDest(x) & "/"
		next
		dest = left(dest,len(dest)-1)
		'response.write "'" & dest & "'"
		'response.write "<br>"
		'response.write name
		'response.end

		sSQL = "SELECT FolderID FROM DocumentFolders WHERE FolderPath='" & dest & "'"
		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSQL, Application("DSN") , 3, 1

		sSQL = "SELECT DocumentID,ParentFolderID,DocumentTitle FROM Documents WHERE DocumentURL='" & src & "'"
		Set oRs2 = Server.CreateObject("ADODB.Recordset")
		oRs2.Open sSQL, Application("DSN") , 3, 1
		
		sSQL = "UPDATE Documents SET DocumentURL='" & fulldest & "',ParentFolderID=" & oRs("FolderID") & " WHERE DocumentID=" & oRs2("DocumentID")
		set oCmd = server.createobject("adodb.connection")
		oCmd.open Application("DSN")
		'response.write sSQL
		'response.end
		oCmd.Execute(sSQL)
		
		oRs.close
		oRs2.close
		oCmd.close

		fs.MoveFile Server.MapPath(src), Server.MapPath(fulldest)
		strType = "Article"
	ElseIf typ = "c" Then
		strType = "Category"
		fulldest = dest
		sSQL = "SELECT FolderID FROM DocumentFolders WHERE FolderPath='" & dest & "'"
		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSQL, Application("DSN") , 3, 1
		

		sSQL = "SELECT FolderID FROM DocumentFolders WHERE FolderPath='" & src & "'"
		Set oRs3 = Server.CreateObject("ADODB.Recordset")
		oRs3.Open sSQL, Application("DSN") , 3, 1
		
		sSQL = "SELECT FolderID FROM DocumentFolders WHERE ParentFolderID='" & oRs3("FolderID") & "'"
		Set oRs2 = Server.CreateObject("ADODB.Recordset")
		oRs2.Open sSQL, Application("DSN") , 3, 1


		if not oRs2.EOF then%>
  			<font size="3"><b>Move <%= strType %></b></font>
  			<hr style="height:1px;"><br>
  			<b>&quot;<%= Replace(src,"/eCapture/pub/","") %>&quot;</b>&nbsp;&nbsp;was not moved because the folder being moved has sub-folders.  Create a new folder where these folders/documents are to be moved and move the subfolders individually.
			</body>
			</html>
		<%
			response.end
		else
			oRs2.Close
			Set oRs2 = Nothing

			'response.write oRs3("FolderID") & "<br>"

			sSQL = "UPDATE DocumentFolders SET FolderPath='" & dest & "/" & name & "',ParentFolderID=" & oRs("FolderID") & " WHERE FolderPath='" & src & "'"
			set oCmd = server.createobject("adodb.connection")
			oCmd.open Application("DSN")
			oCmd.Execute(sSQL)
	
			sSQL = "SELECT DocumentID,ParentFolderID,DocumentTitle FROM Documents WHERE ParentFolderID=" & oRs3("FolderID")
			Set oRs2 = Server.CreateObject("ADODB.Recordset")
			oRs2.Open sSQL, Application("DSN") , 3, 1
	
			do while not oRs2.eof
				arrFileName = Split(oRs2("DocumentTitle"),"                              ")
				sSQL = "UPDATE Documents SET DocumentURL='" & Replace(dest,"/eCapture/pub/","") & "/" & name & "/" & arrFileName(0) & "' WHERE DocumentID=" & oRs2("DocumentID")
				'response.write sSQL & "<br>"
				oCmd.Execute(sSQL)
				oRs2.movenext
			loop
			'response.end
	
			oCmd.Close
			oRs2.Close
			oRs.Close

			dest = dest & "/" & name
			'response.write src & "<br>"
			'response.write dest & "<br>"
			'response.end
			fs.MoveFolder Server.MapPath(src), Server.MapPath(dest)
		end if
	End If
  
    
  Set fs = Nothing
  Response.Write "<script language=""Javascript"">parent.fraToc.document.location.reload();</script>"
  %>

  <font size="3"><b>Move <%= strType %></b></font>
  <hr style="height:1px;"><br>
  <b>&quot;<%= Replace(src,"/eCapture/pub/","") %>&quot;</b>&nbsp;&nbsp;was succesfully moved to&nbsp;&nbsp;<b>&quot;<%= Replace(fulldest,"/eCapture/pub/","") %>&quot;</b>.
  <%Else
    ' User doesnt not have access
	Response.Write "<script language=""Javascript"">alert('You do not have permission to move documents.');</script>"
     Response.Write "<script language=""Javascript"">parent.fraTopic.document.location.href='main.asp';</script>"
   End If%>
</body>
</html>
