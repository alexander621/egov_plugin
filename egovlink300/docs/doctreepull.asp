<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="../include_top_functions.asp" //-->
<%
	Dim sTree

	sTree = "<ul>"

	' LIST ALL FOLDERS AND DOCUMENTS
	iFolderId = GetFolderId( iOrgId, "published_documents" )

	Dim iNodeCount, aFolders
	iNodeCount = 0
	ReDim aFolders(3)

	iFolderCount = ShowFoldersAndDocs( iOrgId, iFolderId, iNodeCount, True, 3 )

	sTree = sTree & "</ul>"

	response.write sTree


'-------------------------------------------------------------------------------------------------------
' Sub ShowFoldersAndDocs( iOrgId, iFolderId, iNodeCount )
'-------------------------------------------------------------------------------------------------------
Function ShowFoldersAndDocs( ByVal iOrgId, ByVal iFolderId, ByRef iNodeCount, ByVal bParentMatch, ByVal iFolderMatchLevel )
	Dim sSql, oRs, iFolderCount, iSubFolderCount, iMatchLevel, bMatch

	iFolderCount = 0 
	iSubFolderCount = 0
	bMatch = False 
	iMatchLevel = 0

	' Get the set of subfolders, loop and display
	sSql = "SELECT folderid, foldername FROM DocumentFolders WHERE orgid = " & iOrgId & " AND ParentFolderID = " & iFolderId & " ORDER BY foldername"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If UserHasDocFolderAccess( iOrgId, request.Cookies("userid"), oRs("folderid") ) Then
			iFolderCount = iFolderCount + 1
			iNodeCount = iNodeCount + 1
			sTree = sTree & "<li id=""foldheader""><strong> &nbsp;" & UCase(oRs("foldername")) '& "-" & iNodeCount & 
			sTree = sTree & "</strong></li>"
			sTree = sTree & "<ul id=""foldinglist"" name=""foldinglist"" style=""display:none;"" >"
			If bHavePath Then 
				iMatchLevel = iFolderMatchLevel + 1
				If iMatchLevel <= iPathSize And iFolderMatchLevel > 0 Then 
					If UCase(aPath(iMatchLevel)) = UCase(Trim(oRs("foldername"))) Then
						If bParentMatch Then 
							bMatch = True 
							Redim Preserve aFolders(iMatchLevel)
							aFolders(iMatchLevel) = iNodeCount
						End If 
					End If 
				End If 
			End If 
			iSubFolderCount = ShowFoldersAndDocs( iOrgId, oRs("folderid"), iNodeCount, bMatch, iMatchLevel )
			ShowDocumentsInFolder oRs("folderid"), iSubFolderCount, iNodeCount
			sTree = sTree & "</ul>"
		End If 
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 
	ShowFoldersAndDocs = iFolderCount
End Function  


'-------------------------------------------------------------------------------------------------------
' Sub ShowDocumentsInFolder( iFolderId, iFolderCount, iNodeCount )
'-------------------------------------------------------------------------------------------------------
Sub ShowDocumentsInFolder( iFolderId, iFolderCount, ByRef iNodeCount )
	Dim sSql, oRs, sDocumentSize

	' Get the set of files, loop and display
	sSql = "SELECT DocumentTitle, DocumentURL, documentsize FROM Documents WHERE ParentFolderID = " & iFolderId & " ORDER BY DocumentTitle"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			sFoldersPath = Replace( Request.ServerVariables("SERVER_NAME") & oRs("DocumentURL"), "/custom/pub", "" )
			If CLng(oRs("documentsize")) > CLng(1024) Then
				sDocumentSize = FormatNumber((oRs("documentsize") / 1024),0)  & " KB"
			Else
				sDocumentSize = FormatNumber(oRs("documentsize"),0) & " Bytes"
			End If
			sTree = sTree & "<li>"
			sTree = sTree & "<a class=""documentlist"" TARGET=""DOCUMENTS"" href=""" & "http://" & sFoldersPath  & """ > "
			'response.write vbcrlf & "<img src=""" & GetFileImageURL( Trim(oRs("DocumentTitle")) ) & """ border=""0"" /> " 
			iNodeCount = iNodeCount + 1
			sTree = sTree & UCase(Trim(oRs("DocumentTitle"))) & " ( " & sDocumentSize & " ) " '& "-" & iNodeCount
'			If iNodeCount = 62 Then
'				response.End 
'			End If 
			sTree = sTree & "</a></li>"
			oRs.MoveNext
		Loop
	Else
		iNodeCount = iNodeCount + 1
		If CLng(iFolderCount ) = CLng(0) Then 
			sTree = sTree & "<li class=""emptyfolder"" style=""font-style: italic;"">(EMPTY)" & "-" & iNodeCount & "</li>"
		End If 
	End If 
End Sub 


'-------------------------------------------------------------------------------------------------------
' SUB ENUMERATEDOCUMENTS(FOLDER)
'-------------------------------------------------------------------------------------------------------
Sub ENUMERATEDOCUMENTS(Folder,sPath,sVpath)

	' BUILD HYPERLINK BASE PATH
	sVirtualpath = replace(Folder.Path,sPath,"")
	sTempPath =  replace(Folder.Path,sPath,"")
	sVirtualpath = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & sVpath & replace(sVirtualpath,"\","/")

	' LIST CONTENTS OF FOLDER (SUBFOLDERS AND FILES)
	For Each SubFolder in Folder.SubFolders

		sVirtualpath2 =  replace(sVpath,"public_documents300","/public_documents300/custom/pub") & replace(sTempPath,"\","/") & "/" & SubFolder.Name

		' CHECK SECURITY FOR ACCESS
		If HasAccess(iorgid,request.Cookies("userid"),sVirtualpath2) Then

			' WRITE FOLDER INFORMATION
			response.write vbcrlf
			response.write "<li id=""foldheader""> " & SubFolder.Name & "</li>" & vbcrlf
			response.write "<ul id=""foldinglist"" name=""foldinglist"" style=""display:none;"" >" & vbcrlf

			' RECURSIVE CALL TO GET ANY SUBFOLDERS OF THE CURRENT FOLDER
			ENUMERATEDOCUMENTS Subfolder, sPath, sVpath

			' LIST FILES IN THE CURRENT FOLDER
			For each File in Subfolder.Files
				' GET FILE SIZE
				If File.Size > 1024 Then
					sFileSize = FormatNumber((File.Size / 1024),0)  & " KB"
				Else
					sFileSize =  FormatNumber(File.Size,0) & " Bytes"
				End If

				sHyperlink = sVirtualPath & "/" & Subfolder.Name & "/" & File.Name
				'response.write  "<li class=""" & GetListItemClassName( File.Name ) & """><a class=""documentlist"" TARGET=""DOCUMENTS"" href=""" & sHyperlink  & """ > " & UCASE(File.Name) & " ( " & sFileSize & ") </a></li>"
				response.write vbcrlf & "<li><a class=""documentlist"" TARGET=""DOCUMENTS"" href=""" & sHyperlink  & """ > " & UCASE(File.Name) & " ( " & sFileSize & ") </a></li>"
			Next

			' IF FOLDER CONTAINS NO FILES OR FOLDERS SHOW AS EMPTY
			If (Subfolder.Files.Count < 1) AND (Subfolder.SubFolders.Count < 1) Then
				'SET MARGIN TO -20 TO ALIGN SUB FOLDERS AND DOCUMENTS WITHIN A FOLDER PROPERLY.
				response.write  "<li class=""emptyfolder"" STYLE=""margin-left:-20px;""> (<i> EMPTY </i>) </li>"  & vbcrlf
			End IF

			' CLOSE FOLDER LIST TAG
			response.write "</ul>"
		Else 
			' Added to allow direct folder links to work while skipping restricted access folders
			response.write "<li id=""foldheader"" style=""display:none;"" ></li>" & vbcrlf
			response.write "<ul id=""foldinglist"" name=""foldinglist"" style=""display:none;"" ></ul>" & vbcrlf
			
			' RECURSIVE CALL TO GET ANY SUBFOLDERS OF THE CURRENT FOLDER
			ENUMERATEDOCUMENTS Subfolder, sPath, sVpath
		End If
	Next

End Sub


'-------------------------------------------------------------------------------------------------------
' Function UserHasDocFolderAccess( iOrgId, iUserId, iFolderId )
'-------------------------------------------------------------------------------------------------------
Function UserHasDocFolderAccess( iOrgId, iUserId, iFolderId )
	Dim rstAccess, sSql, iReturnValue

	On Error Resume Next

	iReturnValue = False

	Set oCnn = Server.CreateObject("ADODB.Connection")
	oCnn.Open Application("DSN")
	sSql = "EXEC CheckFolderIdAccess '" & iOrgId & "','" & iUserId & "','" & iFolderId & "'"
	'response.write sSql & "<br /><br />"
	Set rstAccess = oCnn.Execute(sSql)

	If Not rstAccess.EOF Then
		If rstAccess("folderid") = iFolderId Then
			iReturnValue = True
		End If
	End If

	oCnn.Close
	Set rstAccess = Nothing
	Set oCnn = Nothing

	UserHasDocFolderAccess = iReturnValue

End Function


'-------------------------------------------------------------------------------------------------------
' Function GetListItemClassName( sName )
'-------------------------------------------------------------------------------------------------------
Function GetListItemClassName( sName )

	sReturnValue = "msie"

	Select Case LCase(Right(Trim(sName),3))
		Case "doc"
            sReturnValue = "doc"
        Case "xls"
            sReturnValue = "xls"
        Case "ppt"
            sReturnValue = "ppt"
        Case "htm"
            sReturnValue = "htm"
        Case "pdf"
            sReturnValue = "pdf"
		Case "gif"
            sReturnValue = "gif"
        Case "jpg"
            sReturnValue = "jpg"
	End Select

	  GetListItemClassName = sReturnValue

End Function


'-------------------------------------------------------------------------------------------------------
' Function GetFileImageURL( sFileName )
'-------------------------------------------------------------------------------------------------------
Function GetFileImageURL( sFileName )

	sReturnValue = "./menu/images/"

	Select Case LCase(Right(Trim(sFileName),3))
		Case "doc"
            sReturnValue = sReturnValue & "msword.gif"
        Case "xls"
            sReturnValue = sReturnValue & "msexcel.gif"
        Case "ppt"
            sReturnValue = sReturnValue & "msppt.gif"
        Case "htm"
            sReturnValue = sReturnValue & "msie.gif"
        Case "pdf"
            sReturnValue = sReturnValue & "pdf.gif"
		Case "gif"
            sReturnValue = sReturnValue & "imageicon.gif"
        Case "jpg"
            sReturnValue = sReturnValue & "imageicon.gif"
		Case Else 
			sReturnValue = sReturnValue & "document.gif"
	End Select

	  GetFileImageURL = sReturnValue

End Function


'-------------------------------------------------------------------------------------------------------
' FUNCTION HASACCESS(IORGID,IUSERID,STRVPATH)
'-------------------------------------------------------------------------------------------------------
Function HasAccess(iorgid,iuserid,strvpath)

	  On Error Resume Next

	  iReturnValue = False

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


'-------------------------------------------------------------------------------------------------------
' Function GetFolderId( iOrgId, sFolder )
'-------------------------------------------------------------------------------------------------------
Function GetFolderId( iOrgId, sFolder )
	Dim sSql, oRs

	sSql = "SELECT folderid FROM DocumentFolders WHERE orgid = " & iOrgId & " AND FolderName = '" & sFolder & "'"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFolderId = oRs("folderid")
	Else
		GetFolderId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 

%>