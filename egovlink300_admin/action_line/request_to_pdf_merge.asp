<!--#include file="merge_field_functions.asp"-->
<%
response.buffer = True 

'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: REQUEST_TO_PDF_MERGE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/28/07 
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' DESCRIPTION:  CREATES PDF FILE CONTAINING WAIVER INFORMATION AND RESERVATION DETAIL INFORMATION
' FOR THE CONSUMER TO PRINT AND SIGN.
'
' MODIFICATION HISTORY
' 1.0   02/28/07   JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------


' INITIALIZE AND DECLARE VARIABLES
Dim oPDF, sSystem, sDB
sSystem = request("sys")
irequestid = request("irequestid")


' SET THE CONNECTION STRING
Select Case sSystem 
	Case "DEV"
		sDB = "Driver={SQL Server}; Server=L3SQL2; Database=egovlink300; UID=egovsa; PWD=egov_4303;"
	Case "QA"
		sDB = "Driver={SQL Server}; Server=L3SQL2; Database=egovlink400_QA_Test; UID=egovsa; PWD=egov_4303;"
	Case "LIVE"
		sDB = "Driver={SQL Server}; Server=ISPS0014; Database=egovlink300; UID=sa; PWD=;"
	Case Else
		sDB = "Driver={SQL Server}; Server=L3SQL2; Database=egovlink300; UID=egovsa; PWD=egov_4303;"
End Select 

Set oUser = Server.CreateObject("ADODB.Recordset")
sSQL = "SELECT userid FROM egov_action_request_view WHERE action_autoid='" & irequestid & "'"
oUser.Open sSQL,sDB, 3, 3
'response.write sSQL
if not oUser.EOF then iUserID = oUser("UserID")
oUser.Close

'Dynamically Determine the form letter to use
'sPath = Server.mappath("output_pdfs/Rubbish Letter.pdf")
sPath = Server.mappath("output_pdfs/Rubbish_Letter_2.pdf")

' CREATE THE PDF
Call FillForm(sPath)



'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' Sub FillForm(sPDFPath)
'--------------------------------------------------------------------------------------------------
Sub FillForm(sPDFPath)

	irequestid = request("irequestid")

	' CREATE PDF OBJECT
	Set oPDF = Server.CreateObject("APToolkit.Object")
	oDocument = oPDF.OpenOutputFile("MEMORY") 'CREATE THE OUTPUT INMEMORY

	' BUILD PDF DOCUMENT
	oPDF.OutputPageWidth = 612 ' 8.5 inches
	oPDF.OutputPageHeight = 792 ' 11 inches

	' ADD FORM
	r = oPDF.OpenInputFile(sPDFPath)


	' ADD DATA TO FORM
	'Call PopulateFormwithData(oPDF,irequestid)
	Call fnFillMergeFields(oPDF,irequestid,sAddText)
	oPDF.FlattenRemainingFormFields = True 
	r = oPDF.CopyForm(0, 0)
	'response.end

	' CLOSE PDF DOCUMENT
	oPDF.CloseOutputFile
	oDocument = oPDF.binaryImage 


	' STREAM PDF TO BROWSER
	response.expires = 0
	response.Clear
	response.ContentType = "application/pdf"
	response.AddHeader "Content-Type", "application/pdf"
	response.AddHeader "Content-Disposition", "inline;filename=FORMS.PDF"
	response.BinaryWrite oDocument  


	' DESTROY OBJECTS
	Set oPDF = Nothing
	Set oDocument = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' SUB POPULATEFORMWITHDATA(OPDF)
'--------------------------------------------------------------------------------------------------
Sub PopulateFormwithData(oPDF,irequestid)

	' POPULATE CITIZEN CONTACT INFORMATION INFORMATION

	' GET DATA FROM SQL
	Set oData = Server.CreateObject("ADODB.Recordset")
	'oData.Open "SELECT userfname + ' ' + userlname as full_name,useraddress as address,usercity as city,userstate as state,userzip as zip FROM egov_action_request_view where action_autoid='" & irequestid & "'","Provider=SQLOLEDB; Data Source=L3SQL2; User ID=egovsa; Password=egov_4303; Initial Catalog=egovlink300;", 3, 3
	'oData.Open "SELECT UserfName,UserlName,useraddress as UserStreetAddress,UserCity,UserState,UserZip FROM egov_action_request_view WHERE action_autoid='" & irequestid & "'","Provider=SQLOLEDB; Data Source=L3SQL2; User ID=egovsa; Password=egov_4303; Initial Catalog=egovlink300;", 3, 3
	oData.Open "SELECT userfname + ' ' + userlname as full_name,useraddress as UserStreetAddress,UserCity,UserState,UserZip FROM egov_action_request_view WHERE action_autoid='" & irequestid & "'",sDB, 3, 3
	
	' IF RECORDSET HAS DATA POPULATE FORM FIELDS WITH DATA
	If NOT oData.EOF Then

		' LOOP THRU EACH COLUMN FILL FORM FIELDS
		For Each oColumn in oData.Fields 
			' FILL MATCHING FORM FIELD WITH DATA
			r = oPDF.SetFormFieldData(oColumn.Name,oColumn.Value,1)
		Next
			
		' CLOSE CONNECTION
		oData.Close

	End If

	' DESTROY OBJECTS
	Set oData = Nothing


	
	' POPULATE REQUEST FORM PROMPTS BOTH PUBLIC AND INTERNAL

	' GET DATA FROM SQL
	Set oData = Server.CreateObject("ADODB.Recordset")
	sSQL = "SELECT submitted_request_form_field_name,submitted_request_field_response FROM   egov_submitted_request_fields INNER JOIN egov_submitted_request_field_responses ON egov_submitted_request_fields.submitted_request_field_id=egov_submitted_request_field_responses.submitted_request_field_id WHERE submitted_request_id='" & irequestid & "' and submitted_request_form_field_name IS NOT NULL"
	oData.Open sSQL,"Provider=SQLOLEDB; Data Source=L3SQL2; User ID=egovsa; Password=egov_4303; Initial Catalog=egovlink300;", 3, 1

	
	' IF RECORDSET HAS DATA POPULATE FORM FIELDS WITH DATA
	If NOT oData.EOF Then
		Do While NOT oData.EOF 
			r = oPDF.SetFormFieldData(oData("submitted_request_form_field_name"),oData("submitted_request_field_response"),1)
			oData.MoveNext
		Loop

		' CLOSE CONNECTION
		oData.Close

	End If

	' DESTROY OBJECTS
	Set oData = Nothing

End Sub



'--------------------------------------------------------------------------------------------------
' SUB GETDATA(ICOLUMNINDEX,IDATAINDEX)
'--------------------------------------------------------------------------------------------------
Function GetData(iColumnIndex,iDataIndex)
	
	sReturnValue = "UNKNOWN "

	' GET DATA FROM SQL
	Set oData = Server.CreateObject("ADODB.Recordset")
	oData.Open "SELECT * FROM statement_data where statementdataid='" & iDataIndex & "'","Provider=SQLOLEDB; Data Source=L3SQL2; User ID=egovsa; Password=egov_4303; Initial Catalog=egovlink300;", 3, 3
	
	' IF RECORDSET HAS DATA ADD TO FORM
	If NOT oData.EOF Then
		
		sReturnValue = Trim(oData("column" & iColumnIndex) & " ")
		oData.MoveNext

		' CLOSE CONNECTION
		oData.Close
	End If

	' DESTROY OBJECTS
	Set oData = Nothing

	' RETURN VALUE
	GetData = sReturnValue 

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION FORMATDATA(SVALUE,STYPE)
'--------------------------------------------------------------------------------------------------
Function FormatData(sValue,sType)
	
	' SET DEFAULT VALUE
	sReturnValue = sValue

	' FORMAT VALUE ACCORDING TO DATATYPE
	If trim(sValue) <> "" AND NOT isnull(sValue) Then
		
		Select Case trim(sType)

			Case "decimal"
				' DECIMAL WITH TWO SIGNIFICANT DIGITS
				sReturnValue = FormatNumber(sReturnValue,2)
			Case Else
				' DO NOTHING RETURN SUPPLIED VALUE

		End Select

	End If

	' RETURN VALUE
	FormatData = sReturnValue

End Function
%>
