<!-- #include file="../includes/common.asp" //-->
<!-- #include file="docscommon.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: adddocumentdo.asp
' AUTHOR: Steve Loar
' CREATED: 09/01/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description: Creates new documents of link and direct content types.
'
' MODIFICATION HISTORY
' 1.0   09/01/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSuccessFlag, blnOverwrite, strTitle, strContent, objFSO, oCmd, oFileNew, strSize, strDir
Dim objNewFile, blnNewTarget, strLink

'Add document by creating an HTML document with hyperlink to actual document
If request("txtMethod") = "link" Then 
	strTitle = request.form("txtURLTitle") & ".htm"

	'Determine if the "overwrite" option has been checked.
	If request("chkOverwrite") = "on" Then 
		blnOverwrite = True
	Else 
		blnOverwrite = False
	End If

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If objFSO.FileExists(Server.MapPath(request("txtTopic") & strTitle)) And Not blnOverwrite Then 
		strTitle = ""
	Else 
		'sPhysicalPath = replace(Server.MapPath(request("txtTopic")) & "\" & strTitle,"\custom\pub\custom\pub","\custom\pub") 'Account for error in converting virtual path to physical path
		'sPhysicalPath = "e:" & request("txtTopic") & strTitle
		sPhysicalPath = Application("DocumentsDrive") & request("txtTopic") & strTitle

		Set objNewFile = objFSO.CreateTextFile( sPhysicalPath, True)
		If request("openNew") = "on" Then 
			objNewFile.Write( "<html><body><script language=""javascript"">window.open('" & request("txtURL") & "');</script></body></html>" )
			blnNewTarget = 1
		Else 
			objNewFile.Write( "<META HTTP-EQUIV=refresh CONTENT=""0; URL=" & request.form("txtURL") & """>" )
			blnNewTarget = 0
		End If 
		objNewFile.Close

		'strDir  = request("txtTopic") & strTitle
		'strDir = Replace(request("txtTopic"), "egovlink300_docs", "public_documents300") & strTitle
		strDir = Replace(request("txtTopic"), Application("DocumentsRootDirectory"), "public_documents300") & strTitle
		strLink = request("txtURL")
		strSize = 0 ' This is a link so no size

		'---BEGIN: Update DB fields for Document(LINK) --------------------------------
		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
			.ActiveConnection = Application("DSN")
			.CommandText = "NewDocument"
			.CommandType = adCmdStoredProc
			.Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
			.Parameters.Append oCmd.CreateParameter("CreatorID", adInteger, adParamInput, 4, Session("UserID"))
			.Parameters.Append oCmd.CreateParameter("FolderPath", adVarChar, adParamInput, 300, strDir)
			.Parameters.Append oCmd.CreateParameter("LinkURL", adVarChar, adParamInput, 300, strLink)
			.Parameters.Append oCmd.CreateParameter("DocumentSize", adInteger, adParamInput, 4, strSize)
			.Parameters.Append oCmd.CreateParameter("LinkTargetsNew", adInteger, adParamInput, 4, blnNewTarget)
			.Execute
		End With
		Set oCmd = Nothing
		'---END: Update DB fields----------------------------------

		bSuccess = True 
	End If
	Set objFSO = Nothing
Else 
	'Create a new document with text supplied
	If request("txtTitle") <> "" And request("txtContent") <> "" Then
		If request("blnIsHTML") = "on" Then 
			strTitle   = request("txtTitle") & ".htm"
			strContent = request("txtContent")
		Else 
			strTitle   = request("txtTitle") & ".txt"
			strContent = dbsafe(request("txtContent"))
		End If 

		If request("chkOverwrite") = "on" Then 
			blnOverwrite = True
		Else 
			blnOverwrite = False
		End If

		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

		If objFSO.FileExists(Server.MapPath(request("txtTopic") & strTitle)) And Not blnOverwrite Then 
			response.write request("txtTopic")
			response.End

			strTitle = ""
		Else 

			'sPhysicalPath = Replace(Server.MapPath(request("txtTopic")) & "\" & strTitle,"\custom\pub\custom\pub","\custom\pub") 'Account for error in converting virtual path to physical path
			'sPhysicalPath = "e:" & request("txtTopic") & strTitle
			sPhysicalPath = Application("DocumentsDrive") & request("txtTopic") & strTitle

			Set objNewFile = objFSO.CreateTextFile(sPhysicalPath, True)
			objNewFile.Write(request("txtContent") )
			objNewFile.Close

			Set oFileNew = objFSO.GetFile(sPhysicalPath)
			strSize = oFileNew.Size ' Get the Size of the newly created text file
			Set oFileNew = Nothing

			strDir = Replace(request("txtTopic"), Application("DocumentsRootDirectory"), "public_documents300") & strTitle

			'---BEGIN: Update DB fields for Document(DIRECT CONTENT) --------------------------------
			Set oCmd = Server.CreateObject("ADODB.Command")
			With oCmd
				.ActiveConnection = Application("DSN")
				.CommandText = "NewDocument"
				.CommandType = adCmdStoredProc
				.Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
				.Parameters.Append oCmd.CreateParameter("CreatorID", adInteger, adParamInput, 4, Session("UserID"))
				.Parameters.Append oCmd.CreateParameter("FolderPath", adVarChar, adParamInput, 300, strDir)
				.Parameters.Append oCmd.CreateParameter("LinkURL", adVarChar, adParamInput, 300, null)
				.Parameters.Append oCmd.CreateParameter("DocumentSize", adInteger, adParamInput, 4, strSize)
				.Execute
			End With
			Set oCmd = Nothing
			'---END: Update DB fields----------------------------------

		End If 

		Set objFSO = Nothing

	End If 
End If 

response.redirect "adddocument.asp?sf=fa&path=" & request("txtTopic")


%>