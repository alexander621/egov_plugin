<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: PDF_SAVE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/21/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  SAVE ATTACHMENT FOR REQUEST 
'
' MODIFICATION HISTORY
' 1.0   2/21/07	JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------


' CALL ROUTINE TO SAVE PDF
Call subSavePDF()


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------------
' SUB SUBSAVEPDF()
'------------------------------------------------------------------------------------------------------------
Sub subSavePDF()

' CREATE UPLOAD OBJECT
Set oUpload = Server.CreateObject("Dundas.Upload.2")
oUpload.MaxFileSize = (4096000 * 5) ' MAX SIZE OF UPLOAD SPECIFIED IN BYTES, (4096000 * 5) =  APPX. 20MB
oUpload.SaveToMemory

' GET VARIABLES
sDesc = oUpload.Form("pdfdesc")
sFilePath = oUpload.Files(0).OriginalPath
sFileName = LCASE(RIGHT(sFilePath,LEN(sFilePath) - instrrev(sFilePath,"\")))
sFileExt = LCASE(RIGHT(sFileName,LEN(sFileName) - instrrev(sFileName,".")))
iAdminUserId = session("UserID")
iOrgId = session("orgid")
sServerPath = server.mappath("../../") & "\custom\pub\" & session("virtualdirectory") & "\pdf_forms"

' CALL TO MAKE SURE ATTACHMENT FOLDER EXISTS - CREATE IF IT DOESNT
Call subPDFFolderCheck(sServerPath)


' SAVE FILE INFORMATION IN DATABASE
sServerFileName = FnStorePDFInfo(sFileName,sDesc,iAdminUserId)
sServerFileName = sServerFileName & "." & sFileExt 


' STORE FILE IN SERVER FILESYSTEM
If oUpload.FileExists( sServerPath & "\" & sServerFileName ) Then 
	' DELETE IF ALREADY EXISTS ON SERVER FILESYSTEM
	oUpload.FileDelete( sServerPath & "\" & sServerFileName )
End If

' SAVE FILE ON SERVER FILESYSTEM
oUpload.Files(0).SaveAs(  sServerPath & "\" & sServerFileName )


' CLEAN UP OBJECTS
Set oUpload = Nothing


' RETURN REQUEST PAGE
response.redirect("requestmergeforms_manage.asp?r=save")

End Sub


'------------------------------------------------------------------------------------------------------------
' FUNCTION FNSTOREPDFINFO(SPDF_NAME,SPDF_DESC,IADMINUSERID)
'------------------------------------------------------------------------------------------------------------
Function fnStorePDFInfo(sPDF_Name,sPDF_Desc,iadminuserid)

	Set oPDF = Server.CreateObject("ADODB.Recordset")
	oPDF.CursorLocation = 3
	oPDF.Open "SELECT pdfid,pdf_name,pdf_description,date_added,orgid,adminuserid FROM egov_action_request_pdfforms WHERE 1=2",Application("DSN"),1, 2

	' ADD DATABASE ROW
	oPDF.AddNew
	oPDF("pdf_name") = DBsafe(sPDF_Name)
	oPDF("pdf_description") = DBSafe(sPDF_Desc)
	oPDF("adminuserid") = iadminuserid
	oPDF("date_added") = Now() ' NEED TO UPDATE TO USE CITY'S TIME
	oPDF("orgid") = session("orgid")
	oPDF.Update
	iReturnValue = oPDF("pdfid")

	' CLOSE RECORDSET
	oPDF.Close
	Set oPDF = Nothing

	fnStorePDFInfo = iReturnValue

End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION DBSAFE( STRDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function


'----------------------------------------------------------------------------------------------------------------------
' SUB SUBPDFFOLDERCHECK(SFOLDERPATH)
'----------------------------------------------------------------------------------------------------------------------
Sub subPDFFolderCheck(sFolderPath)

	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	If oFSO.FolderExists(sFolderPath) <> True Then
		
		' CREATE ATTACHMENTS FOLDER
		Set oFolder = oFSO.CreateFolder(sFolderPath)
		Set oFolder = Nothing

	End If

	Set oFSO = Nothing

End Sub
%>
