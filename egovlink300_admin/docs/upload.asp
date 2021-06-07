<!-- #include file="../includes/common.asp" //-->

<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: upload.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description: Uploads new documents.
'
' MODIFICATION HISTORY
' 2.0   09/01/2010	Steve Loar - Modified for non-frameset version of documents
' 3.0	04/25/2013	Terry Foster - Using centralized upload sub
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim objUpload, strMessage, strFilename, blnSuccessful, temp, sFileSize, strPath, strFullPath

'response.write "Here in upload" & "<br /><br />"
Response.Buffer      = True
Response.Expires     = -1
Server.ScriptTimeout = 600  'in secs.  10 min.

blnSuccessful = False   

Set formFields = Server.CreateObject("Scripting.Dictionary")
call UploadFile(formFields, "[FORMPHYSICAL]txtTopic", false, true)

	strPath = formFields("txtTopic")
    	strFilename = formFields("filename")
    	sFileSize = formFields("filesize")
	session("MyDoc") = strPath & strFileName ' Used in Recording Database entry
	sSuccess = formFields("sSuccess")

'DONE UPLOADING...PROCESS DB AND RETURN
 If sSuccess = "true" Then
	response.redirect "adduploadtodb.asp?filename=" & strFileName & "&path=" & strPath & "&size=" & sFileSize
 Else
	'UPLOAD WAS NOT SUCCESFUL RETURN TO ADD PAGE WITH MESSAGE
	response.redirect "adddocument.asp?sf=" & sSuccess + "&path=" & strPath
 End If


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
