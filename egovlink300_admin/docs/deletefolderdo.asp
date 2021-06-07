<!-- #include file="../includes/common.asp" //-->
<!-- #include file="docscommon.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: deletefolderdo.asp
' AUTHOR: Steve Loar
' CREATED: 08/31/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description: Deletes document folders.
'
' MODIFICATION HISTORY
' 1.0   08/31/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSuccessFlag, sTargetFolder, iFolderId, sPhysicalPath, sFolderName, objDir, sDBFolder

sTargetFolder = request("path")
sFolderName = request("foldername")
iFolderId = CLng(request("folderid"))

'sPhysicalPath = "e:" & Replace(sTargetFolder, "/", "\") ' This is the physical folder we need to check and delete
sPhysicalPath = Application("DocumentsDrive") & Replace(sTargetFolder, "/", "\") ' This is the physical folder we need to check and delete

'sDBFolder = Replace(sTargetFolder, "egovlink300_docs", "public_documents300")
sDBFolder = Replace(sTargetFolder, Application("DocumentsRootDirectory"), "public_documents300")

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

If objFSO.FolderExists( sPhysicalPath ) Then

	' Delete the folder and everything below it.
	objFSO.DeleteFolder(sPhysicalPath)

	sSuccessFlag = "fod"

	' Delete any files even those in the sub folders
	sSql = "DELETE FROM documents WHERE orgid = " & session("orgid")
	sSql = sSql & " AND documenturl like '" & sDBFolder & "/%'"
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql

	' Delete any sub folders 
	sSql = "DELETE FROM DocumentFolders WHERE orgid = " & session("orgid")
	sSql = sSql & " AND FolderPath like '" & sDBFolder & "/%'"
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql

	' Delete the folder 
	sSql = "DELETE FROM DocumentFolders WHERE orgid = " & session("orgid")
	sSql = sSql & " AND FolderPath = '" & sDBFolder & "'"
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql

Else
	sSuccessFlag = "nf"

End If 

Set objFSO = Nothing 

response.write "sSuccessFlag = " & sSuccessFlag & "<br /><br />"

If sSuccessFlag= "nf" Then
	' folder not found so go back to the folder delete page
	response.redirect "deletefolder.asp?path=" & request("path") & "&foldername=" & sFolderName & "&folderid=" & iFolderId & "&sf=" & sSuccessFlag
Else
	' folder deleted so go to the default page 
	response.redirect "default.asp?foldername=" & sFolderName & "&sf=" & sSuccessFlag
End If 


%>