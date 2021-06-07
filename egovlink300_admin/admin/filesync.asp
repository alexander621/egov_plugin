<%
' END USED TO PREVENT ACCIDENTALLY RUNNING OF SCRIPT THRU ERRANT BROWSE OF THIS PAGE
	response.write "DOCUMENTS WERE NOT SYNCHRONIZED.  PLEASE DISABLE RESPONSE.END TO RUN SCRIPT."
	response.end

' WARNING - This script destroys the security on the file folders

' -------------------------------------------------------------------------------------------------
' BEGIN SCRIPT INFORMATION
'--------------------------------------------------------------------------------------------------
' AUTHOR:		JOHN STULLENBERGER
' DATE:			2/11/2005
' REVISION:		1.0
' DESCRIPTION:	SYNCHRONIZES SQL DATABASE WITH PHYSICAL FILES FOR BOARDSITE DOCUMENT MODULE.
' EXAMPLE.  HTTP://WWW.EGOVLINK.COM/st-helens/ADMIN/ADMIN/FILESYNC.ASP
' 
' -------------------------------------------------------------------------------------------------
' END SCRIPT INFORMATION
'--------------------------------------------------------------------------------------------------

' INITIALIZE VALUES AND OBJECTS
Const adExecuteNoRecords = 128
Const adCmdStoredProc = 4
Const adCmdText = 1
Const adInteger = 3
Const adVarChar = 200
Const adLongVarChar = 201
Const adDateTime = 135
Const adParamReturnValue = 4
Const adParamInput = 1
Const adParamOutput = 2
Const adOpenStatic = 3
Const adUseClient = 3
Const adLockReadOnly = 1
Const adStateOpen = 1
Dim sPath,sVirtualPath,iOrgID,iuserid,blnOverwrite,sDSN

blnOverwrite = True
'--------------------------------------------------------------------------------------------------
' BEGIN SCRIPT PARAMATERS
'--------------------------------------------------------------------------------------------------
sCity = "clarendonhills" ' CHANGE TO MATCH
iOrgID = "80" ' CHANGE TO MATCH
'--------------------------------------------------------------------------------------------------
' END SCRIPT PARAMATERS
'--------------------------------------------------------------------------------------------------
sPath ="C:\wwwroot\www.cityegov.com\egovlink300_admin\custom\pub\" & sCity
sVirtualPath = "/public_documents300/custom/pub/" & sCity
iuserid = "1364" ' IGNORE
sDSN = "Driver={SQL Server}; Server=ISPS0014; Database=egovlink300; UID=egovsa; PWD=egov_4303;"
Set FSO = CreateObject("Scripting.FileSystemObject")


response.write "<div style=""background-color:#e0e0e0;border: solid 1px #000000;padding:10px;FONT-FAMILY: Verdana,Tahoma,Arial;font-size:10px;"">"
response.flush


' APPENDING OR OVERWRITE
If blnOverwrite = True Then
	' DELETE EXISTING FILE AND FOLDER DATABASE ROWS SO SCRIPT CAN REBUILD FROM SCRATCH
	response.write "<p><b>Clearing existing database rows for orgid(" & iorgid & ")...</b><br>"
	ClearDocuments iorgid 
	response.write "<b>Done clearing database rows.</b><br></p>"
	response.flush
	
	' ADD ROOT PATH
	AddFolder iOrgID, iuserid, sVirtualPath 

Else

	' APPEND EXISTING TO EXISTING DATABASE ROWS
	response.write "<P><b>This option has not been coded.  Please contact a developer.</b><br></p>"
	response.flush
End If


' ENUMERATE FOLDERS AND FILES ADDING ENTRIES TO DATABASE
response.write "<b>Adding Files and Folders...</b><br>"
response.flush


' ENUMERATE FOLDERS AND FILES ADDING TO DATABASE
SyncDocuments FSO.GetFolder(sPath)

response.write "<b>Done Adding Files and Folders.</b>"
response.write "</div>"
response.flush



'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------------
' SUB SYNCDOCUMENTS(FOLDER)
'-------------------------------------------------------------------------------------------------------
Sub SyncDocuments(Folder)
	Dim strSize
   
	For Each Subfolder in Folder.SubFolders
		sSubPath = replace(Subfolder.Path,sPath,"")
		sSubPath = replace(sSubPath,"\","/")
		sSubPath = sVirtualPath & sSubPath
		response.write "<b>Adding Folder: </b>" & sSubPath & "<br />"
		AddFolder iOrgID, iuserid, sSubPath
		
		For each File in Subfolder.Files 
			sDocumentPath = sSubPath & "/" & File.Name 
			strSize = File.Size ' Get the Size of the file
			response.write  "<b>Adding File: </b>" &sDocumentPath & "<br />"
			response.flush
			AddDocument iOrgID, iUserID, sDocumentPath, strSize
		Next 
        
		SyncDocuments Subfolder
    Next
End Sub


'-------------------------------------------------------------------------------------------------------
' SUB ADDDOCUMENT(IORGID,IUSERID,SPATH)
'-------------------------------------------------------------------------------------------------------
Sub AddDocument( iOrgID, iUserID, sPath, strSize )

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = sDSN
		.CommandText = "NewDocument"
		.CommandType = adCmdStoredProc
		.Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, iOrgID)
		.Parameters.Append oCmd.CreateParameter("CreatorID", adInteger, adParamInput, 4, Session("UserID"))
		.Parameters.Append oCmd.CreateParameter("FolderPath", adVarChar, adParamInput, 300, sPath)
		.Parameters.Append oCmd.CreateParameter("LinkURL", adVarChar, adParamInput, 300, null)
		.Parameters.Append oCmd.CreateParameter("DocumentSize", adInteger, adParamInput, 4, strSize)
		.Execute
	End With

	Set oCmd = Nothing

End Sub


'-------------------------------------------------------------------------------------------------------
' SUB ADDFOLDER()
'-------------------------------------------------------------------------------------------------------
Sub AddFolder(iorgid,iuserid,sPath)

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = sDSN
		.CommandText = "NewFolder"
		.CommandType = adCmdStoredProc
		.Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, iorgid)
		.Parameters.Append oCmd.CreateParameter("CreatorID", adInteger, adParamInput, 4, iuserid)
		.Parameters.Append oCmd.CreateParameter("FolderPath", adVarChar, adParamInput, 300, sPath)
		.Execute
	End With
    
	Set oCmd = Nothing

End Sub


'-------------------------------------------------------------------------------------------------------
' SUB CLEARDOCUMENTS(IORGID)
'-------------------------------------------------------------------------------------------------------
Sub ClearDocuments( iorgid )
	Dim oCmd
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	' clear document annotations
	oCmd.CommandText = "DELETE FROM annotations WHERE ORGID = " & iorgid
	oCmd.Execute
	' Clear folder security
	oCmd.CommandText = "DELETE FROM featureaccess WHERE ORGID = " & iorgid
	oCmd.Execute
	' clear the documents
	oCmd.CommandText = "DELETE FROM DOCUMENTS WHERE ORGID = " & iorgid
	oCmd.Execute
	' Clear the folders
	oCmd.CommandText = "DELETE FROM DOCUMENTFOLDERS WHERE ORGID = " & iorgid
	oCmd.Execute

	Set oCmd = Nothing

End Sub
%>
