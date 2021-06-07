<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: adduploadtodb.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description: Uploads new documents.
'
' MODIFICATION HISTORY
' 2.0   09/01/2010	Steve Loar - Modified for non-frameset version of documents
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim srtDir, strSize, sFileName, strPath, oCmd

'dtb_debug(session("orgid") & " adduploadtodb - Started: " & Now())
response.buffer = True

'SET VARIABLES
strDir = session("MyDoc")
'strDir = Replace(strDir, "egovlink300_docs", "public_documents300")
strDir = Replace(strDir, Application("DocumentsRootDirectory"), "public_documents300")

'dtb_debug(session("orgid") & " adduploadtodb - " & Now() & " - MyDoc: " & session("MyDoc"))

'It is necessary to pass these because the upload form submits to 
'file multipart/form-data file, so it is necessary to carry these
'values thru the querystring
sFileName   = request("filename")
'dtb_debug(session("orgid") & " adduploadtodb - " & Now() & " - Title: " & request("sFileName"))
'strMessage = request("Message")
'dtb_debug(session("orgid") & " adduploadtodb - " & Now() & " - Message: " & request("Message"))
strPath    = request("path")
'dtb_debug(session("orgid") & " adduploadtodb - " & Now() & " - Path: " & request("path"))

If request("size") <> "" Then 
	strSize	= request("size")
Else
	strSize = 0
End If 

'dtb_debug(session("orgid") & " adduploadtodb - " & Now() & " - Size: " & strSize)

'---BEGIN: Update DB fields for Document(DIRECT CONTENT) --------------------------------
' Clear out the old one if overwriting it.
Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
	.ActiveConnection = Application("DSN")
	.CommandText = "DelDocument"
	.CommandType = adCmdStoredProc
	.Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
	.Parameters.Append oCmd.CreateParameter("DocumentURL", adVarChar, adParamInput, 255, strDir)
	.Execute
End With
Set oCmd = Nothing

' Add the new one
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

'dtb_debug(session("orgid") & " adduploadtodb - " & Now() & " - Cmd Complete")

response.redirect "adddocument.asp?sf=su&filename=" & sFileName & "&path=" & strPath

%>

<html>
<head>
</head>
<body>

</body>
</html>

<%

'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void dtb_debug p_value
'--------------------------------------------------------------------------------------------------
Sub dtb_debug( ByVal p_value )
	Dim sSql 

	If session("orgid") = 5 Then 
		sSql = "INSERT INTO my_table_dtb ( notes ) VALUES ( '" & dbsafe(p_value) & "' )"
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql
	End If 

End Sub 

%>