<%
response.buffer = True 


'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: DISPLAY_WAIVER.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/6/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' DESCRIPTION:  CREATES PDF FILE CONTAINING WAIVER INFORMATION AND RESERVATION DETAIL INFORMATION
' FOR THE CONSUMER TO PRINT AND SIGN.
'
' MODIFICATION HISTORY
' 1.0   02/6/06   JOHN STULLENBERGER - INITIAL VERSION
' 2.0   03/2/06	  JOHN STULLENBERGER - MODIFIED TO ADD EXTERNAL PDFS
' LINK: HTTPS://SECURE.ECLINK.COM/EGOVLINK/DISPLAY_WAIVER.ASP?MASK=X9
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------


' INITIALIZE AND DECLARE VARIABLES
sMask = request("MASK")
If request("adminlink") = "true" Then
	sMask = GetAdditionalWaivers(sMask)
End If

' CHECK TO SEE IF ANY WAIVERS ARE REQUIRED\SELECTED
If TRIM(sMask) = "" Then 
	response.write "NO WAIVERS SELECTED OR REQUIRED!"
	response.end
End If

sBody = GetWaiverText( sMask ) ' GET WAIVER BODY

' PREPOPULATE DATA
If request("adminlink") = "true" Then
	' ADMIN SIDE ALL SQL
	sBody = GetReservationDetailAdmin(sBody)
Else
	' CLIENT SIDE COMBINATION SQL AND HTML FORM
	sBody = GetReservationDetail(sBody)
End If
	


arrBody = split(sBody,"[*NEWPAGE*]")
iPageCount = 1

' CREATE PDF OBJECT
response.end
Set oPDF = Server.CreateObject("APToolkit.Object")
oDocument = oPDF.OpenOutputFile("MEMORY") 'CREATE THE OUTPUT INMEMORY

' BUILD PDF DOCUMENT
oPDF.OutputPageWidth = 612 ' 8.5 inches
oPDF.OutputPageHeight = 792 ' 11 inches

' ADD TEXT TO DOCUMENT
For iLine = 1 to UBOUND(arrBody)
	
	' BUILD INDIVIDUAL PAGE TEXT
	sPageText = sPageText & arrBody(iLine) & VBCRLF

	' ADD TEXT TO PAGE
	oPDF.PrintMultilineText "Helvetica",10,72,720,468,648,sPageText,0
	
	' START NEW PAGE
	oPDF.NewPage
	sPageText = ""
	iPageCount = iPageCount + 1

Next


' HANDLE ADDING EXTERNAL PDF DOCUMENTS
If request("facilityid") = 9 Then
	' IF TW ADD FLOOR PLAN
	sPathtoAdd = server.mappath("pdfs/tw_floor_plan.pdf")
	oDocument = oPDF.MergeFile(sPathtoAdd, 0, 0)
End If

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



'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' FUNCTION GETWAIVERTEXT(IWAIVERID)
'--------------------------------------------------------------------------------------------------
Function GetWaiverText(sMask)
	
	sMask = replace(sMask,"X",",")

	sReturnValue = ""

	sSQLb = "Select body from egov_waivers where waiverid IN (" & sMask & ")"
	Set oWaiver = Server.CreateObject("ADODB.Recordset")
	'oWaiver.Open sSQLb, "Driver={SQL Server}; Server=ISPS0014; Database=egovlink300; UID=egovsa; PWD=egov_6814;", 3, 1
	oWaiver.Open sSQLb, "Provider=SQLOLEDB; Data Source=CO-SQL-03; User ID=egovsa; Password=Egov_6814; Initial Catalog=egovlink300;", 3, 1
	
	If Not oWaiver.EOF Then
		DO while NOT oWaiver.EOF 
			sReturnValue = sReturnValue & oWaiver("body") 
			oWaiver.MoveNext
		Loop
	End If

	GetWaiverText = sReturnValue
			 
End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION  GETRESERVATIONDETAIL(SBODY)
'--------------------------------------------------------------------------------------------------
Function  GetReservationDetail( ByVal sBody )

	' GET VALUES
	sAmount = request("amount")
	scheckindate = request("checkindate")
	scheckintime = request("checkintime")
	scheckouttime = request("checkouttime")
	scheckoutdate = request("checkoutdate")
	sorganization = request("custom_1")
	
	' GET CUSTOM VALUES
	For Each oField IN Request.Form
		If Left(oField,7) = "custom_" Then
			' GET VALUES
			arrValues = split(oField,"_")
			iFieldName = arrValues(1)

			' SET VALUES
			Select Case iFieldName

				Case "poc"
					spoc = request(oField)

				Case "attending"
					attending = request(oField)

				Case "organization"
					org = request(oField)

				Case Else

			End Select
	
		End If
	Next 
	
	' USER INFO
	'sSQL = "SELECT * FROM egov_users WHERE userid = '" & request("iuserid") & "'"
	sSql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, ISNULL(useraddress,'') AS useraddress, "
	sSql = sSql & "ISNULL(usercity,'') AS usercity, ISNULL(userstate,'') AS userstate, ISNULL(userzip,'') AS userzip, "
	sSql = sSql & "ISNULL(useremail,'') AS useremail, ISNULL(userhomephone,'') AS userhomephone, ISNULL(userbusinessname,'') AS userbusinessname, "
	sSql = sSql & "ISNULL(userworkphone,'') AS userworkphone, ISNULL(userfax, '') AS userfax FROM egov_users WHERE userid = '" & CLng(request("iuserid")) & "'"
	'response.write sSql & "<br /><br />"
	'response.flush

	Set oInfo = Server.CreateObject("ADODB.Recordset")
	'oInfo.Open sSQL, "Driver={SQL Server}; Server=ISPS0014; Database=egovlink300; UID=egovsa; PWD=egov_4303;" , 3, 1
	oInfo.Open sSql, "Provider=SQLOLEDB; Data Source=CO-SQL-03; User ID=egovsa; Password=Egov_6814; Initial Catalog=egovlink300;", 3, 1

	If NOT oInfo.EOF Then
		' USER FOUND SET VALUES
		sFirstName = oInfo("userfname")
		sLastName = oInfo("userlname")
		sAddress1 = oInfo("useraddress")
		sCity = oInfo("usercity")
		sState = oInfo("userstate")
		sZip = oInfo("userzip")
		sEmail = oInfo("useremail")
		sHomePhone = oInfo("userhomephone")
		sWorkPhone = oInfo("userworkphone")
		sBusinessName = oInfo("userbusinessname")
		sFax = oInfo("userfax")
	End If

	oInfo.Close 
	Set oInfo = Nothing


	' REPLACE WITH VALUES
	sBody = replace(sBody,"[*checkindate*]",scheckindate)
	sBody = replace(sBody,"[*checkintime*]",scheckintime)
	sBody = replace(sBody,"[*checkoutdate*]",scheckoutdate)
	sBody = replace(sBody,"[*checkouttime*]",scheckouttime)
	sBody = replace(sBody,"[*amount*]",sAmount)
	sBody = replace(sBody,"[*firstname*]",sFirstName)
	sBody = replace(sBody,"[*middle*]",sMiddle)
	sBody = replace(sBody,"[*lastname*]",sLastName)
	sBody = replace(sBody,"[*address1*]",sAddress1)
	sBody = replace(sBody,"[*address2*]",sAddress2)
	sBody = replace(sBody,"[*city*]",sCity)
	sBody = replace(sBody,"[*state*]",sState)
	sBody = replace(sBody,"[*zip*]",sZip)
	sBody = replace(sBody,"[*email*]",sEmail)
	sBody = replace(sBody,"[*pointofcontact*]",spoc)
	sBody = replace(sBody,"[*attending*]",attending)
	sBody = replace(sBody,"[*organization*]",org)

	GetReservationDetail = sBody

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETWAIVERTEXT(IWAIVERID)
'--------------------------------------------------------------------------------------------------
Function GetAdditionalWaivers( ByVal sMask )
	
	sReturnValue = sMask

	For Each oField IN Request.Form
		If Left(oField,11) = "chkwaivers_" Then
		    iWaiverCount = iWaiverCount + 1
			If sList = "" Then
				sList = sList & request(oField)
			Else
				sList = sList & "X" & request(oField)
			End If
			
		End If
	Next

	sReturnValue = sReturnValue & sList

	GetAdditionalWaivers = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION  GETRESERVATIONDETAILADMIN(SBODY)
'--------------------------------------------------------------------------------------------------
Function  GetReservationDetailAdmin( ByVal sBody )


	' GET INFORMATION FOR THIS RESERVATION
	sSQL = "Select * FROM egov_facilityschedule INNER JOIN egov_facility ON egov_facilityschedule.facilityid = egov_facility.facilityid where facilityscheduleid = '" & CLng(request("ReservationID")) & "'"

	Set oReservation = Server.CreateObject("ADODB.Recordset")
	'oReservation.Open sSQL, "Driver={SQL Server}; Server=ISPS0014; Database=egovlink300; UID=egovsa; PWD=egov_4303;", 3, 1
	oReservation.Open sSQL, "Provider=SQLOLEDB; Data Source=CO-SQL-03; User ID=egovsa; Password=Egov_6814; Initial Catalog=egovlink300;", 3, 1
	
	' IF RESERVATION HAS INFORMATION POPULATE VALUES
	If not oReservation.EOF Then
		samount = formatcurrency(oReservation("amount"),2)
		scheckindate = oReservation("checkindate") 
		scheckintime = oReservation("checkintime") 
		scheckoutdate = oReservation("checkoutdate") 
		scheckouttime = oReservation("checkouttime") 
		ilesseeid	 =  oReservation("lesseeid") 
	End If

	' CLEAN UP OBJECTS
	Set oReservation = Nothing

	
	' GET CUSTOM FIELD VALUES
	sSQL = "SELECT V.facilityvalueid, V.fieldid, V.fieldvalue, V.paymentid, F.fieldid AS Expr1, F.fieldname, F.fieldtype, F.facilityid, F.sequence, F.isrequired, F.fieldchoices "
	sSql = sSql & "FROM egov_facility_field_values V INNER JOIN egov_facility_fields F ON V.fieldid = F.fieldid "
	sSql = sSql & "WHERE (V.paymentid = '" & CLng(request("ReservationID")) & "') ORDER BY V.paymentid, V.fieldid"

	Set oFacilityDetails = Server.CreateObject("ADODB.Recordset")
	'oFacilityDetails.Open sSQL, "Driver={SQL Server}; Server=ISPS0014; Database=egovlink300; UID=egovsa; PWD=egov_4303;" , 3, 1
	oFacilityDetails.Open sSQL, "Provider=SQLOLEDB; Data Source=CO-SQL-03; User ID=egovsa; Password=Egov_6814; Initial Catalog=egovlink300;", 3, 1

	If NOT oFacilityDetails.EOF Then

		Do While NOT oFacilityDetails.EOF 
			
			Select Case oFacilityDetails("fieldname")
			Case "poc"
				spoc = oFacilityDetails("fieldvalue") 
			Case "organization"
				org =  oFacilityDetails("fieldvalue") 
			Case "purpose"
				spurpose =  oFacilityDetails("fieldvalue") 
			Case "attending"
				sattending =  oFacilityDetails("fieldvalue") 
			End Select
			
			oFacilityDetails.MoveNext
		Loop
	
	End If

	Set oFacilityDetails = Nothing

	GetFacilityFieldValues = sReturnValue



	
	' USER INFO
	'sSQL = "SELECT * FROM egov_users WHERE userid = '" & ilesseeid & "'"
	sSql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, ISNULL(useraddress,'') AS useraddress, "
	sSql = sSql & "ISNULL(usercity,'') AS usercity, ISNULL(userstate,'') AS userstate, ISNULL(userzip,'') AS userzip, "
	sSql = sSql & "ISNULL(useremail,'') AS useremail, ISNULL(userhomephone,'') AS userhomephone, ISNULL(userbusinessname,'') AS userbusinessname, "
	sSql = sSql & "ISNULL(userfax, '') AS userfax FROM egov_users WHERE userid = '" & CLng(request("iuserid")) & "'"

	Set oInfo = Server.CreateObject("ADODB.Recordset")
	'oInfo.Open sSQL, "Driver={SQL Server}; Server=ISPS0014; Database=egovlink300; UID=egovsa; PWD=egov_4303;" , 3, 1
	oInfo.Open sSQL, "Provider=SQLOLEDB; Data Source=CO-SQL-03; User ID=egovsa; Password=Egov_6814; Initial Catalog=egovlink300;", 3, 1

	If NOT oInfo.EOF Then
		' USER FOUND SET VALUES
		sFirstName = oInfo("userfname")
		sLastName = oInfo("userlname")
		sAddress1 = oInfo("useraddress")
		sCity = oInfo("usercity")
		sState = oInfo("userstate")
		sZip = oInfo("userzip")
		sEmail = oInfo("useremail")
		sHomePhone = oInfo("userhomephone")
		sWorkPhone = oInfo("userworkphone")
		sBusinessName = oInfo("userbusinessname")
		sFax = oInfo("userfax")
	End If

	Set oInfo = Nothing


	' REPLACE WITH VALUES
	sBody = replace(sBody,"[*checkindate*]",scheckindate)
	sBody = replace(sBody,"[*checkintime*]",scheckintime)
	sBody = replace(sBody,"[*checkoutdate*]",scheckoutdate)
	sBody = replace(sBody,"[*checkouttime*]",scheckouttime)
	sBody = replace(sBody,"[*amount*]",sAmount)
	sBody = replace(sBody,"[*firstname*]",sFirstName)
	sBody = replace(sBody,"[*middle*]",sMiddle)
	sBody = replace(sBody,"[*lastname*]",sLastName)
	sBody = replace(sBody,"[*address1*]",sAddress1)
	sBody = replace(sBody,"[*address2*]",sAddress2)
	sBody = replace(sBody,"[*city*]",sCity)
	sBody = replace(sBody,"[*state*]",sState)
	sBody = replace(sBody,"[*zip*]",sZip)
	sBody = replace(sBody,"[*email*]",sEmail)
	sBody = replace(sBody,"[*pointofcontact*]",spoc)
	sBody = replace(sBody,"[*attending*]",attending)
	sBody = replace(sBody,"[*organization*]",org)
	sBody = replace(sBody,"[*attending*]",sattending)
	sBody = replace(sBody,"[*purpose*]",spurpose)

	GetReservationDetailAdmin = sBody

End Function


%>
