<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalupdate.asp
' AUTHOR: Steve Loar
' CREATED: 08/19/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Updates Rentals. Called from rentaledit.asp
'
' MODIFICATION HISTORY
' 1.0   08/19/2009	Steve Loar - INITIAL VERSION
' 1.1	03/24/2011	Steve Loar - deactivated check added
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRentalId, sSql, sRentalName, sPublicCanView, sPublicCanReserve, sNeedsApproval, iLocationid
Dim sWidth, sLength, sCapacity, iSupervisorUserId, sDescription, sReceiptNotes, sImagesPlacement
Dim iMaxImages, iMaxDocuments, iMaxRentals, iRecreationCategoryId, sMessageFlag, sHasOffSeason
Dim sOffSeasonStartMonth, sOffSeasonStartDay, sOffSeasonEndMonth, sOffSeasonEndDay, sShortDescription
Dim sTerms, iMaxItems, sRentalItem, iAccountId, iMaxAvailable, iAmount, sResidentRentalPeriod
Dim sNonResidentRentalPeriod, iPriceTypeId, sOffSeasonEndYear, sNoCostToRent, sReservationsDuringSeason
Dim sNonResidentsWait, sNonresidentWaitDays, sIconImageUrl, iMaxAlertRows, x, sCheckUseHTMLonLong
Dim sDeactivatedCheck

iRentalId = CLng(request("rentalid"))

If request("rentalname") <> "" Then
	sRentalName = "'" & dbsafe(request("rentalname")) & "'"
Else
	sRentalName = "NULL"
End If 

If request("nocosttorent") = "on" Then
	sNoCostToRent = 1
Else
	sNoCostToRent = 0
End If 

If request("isdeactivated") = "on" Then
	sDeactivatedCheck = 1
	sPublicCanView = 0
	sPublicCanReserve = 0
Else
	sDeactivatedCheck = 0
	If request("publiccanview") = "on" Then
		sPublicCanView = 1
	Else
		sPublicCanView = 0
	End If 

	If request("publiccanreserve") = "on" Then
		sPublicCanReserve = 1
	Else
		sPublicCanReserve = 0
	End If 
End If 

If request("chkUseHTMLonLong") = "on" Then
	sCheckUseHTMLonLong = 1
Else
	sCheckUseHTMLonLong = 0
End If 

' we do not do this anymore
'If request("needsapproval") = "on" Then
'	sNeedsApproval = 1
'Else
	sNeedsApproval = 0
'End If 

If CLng(request("locationid")) > CLng(0) Then
	iLocationid = CLng(request("locationid"))
Else
	iLocationid = "NULL"
End If 

If request("width") <> "" Then
	sWidth = "'" & DBsafeWithHTML(request("width")) & "'"
Else
	sWidth = "NULL"
End If 

If request("length") <> "" Then
	sLength = "'" & DBsafeWithHTML(request("length")) & "'"
Else
	sLength = "NULL"
End If 

If request("capacity") <> "" Then
	sCapacity = "'" & DBsafeWithHTML(request("capacity")) & "'"
Else
	sCapacity = "NULL"
End If 

If CLng(request("supervisoruserid")) > CLng(0) Then
	iSupervisorUserId = CLng(request("supervisoruserid"))
Else
	iSupervisorUserId = "NULL"
End If 

If request("description") <> "" Then
	sDescription = "'" & DBsafeWithHTML(request("description")) & "'"
Else
	sDescription = "NULL"
End If 

If request("shortdescription") <> "" Then
	sShortDescription = "'" & DBsafeWithHTML(request("shortdescription")) & "'"
Else
	sShortDescription = "NULL"
End If 

If request("receiptnotes") <> "" Then
	sReceiptNotes = "'" & DBsafeWithHTML(request("receiptnotes")) & "'"
Else
	sReceiptNotes = "NULL"
End If 

If request("terms") <> "" Then
	sTerms = "'" & DBsafeWithHTML(request("terms")) & "'"
Else
	sTerms = "NULL"
End If 

If request("imagesplacement") <> "" Then
	sImagesPlacement = "'" & dbsafe(request("imagesplacement")) & "'"
Else
	sImagesPlacement = "NULL"
End If 

If request("hasoffseason") = "on" Then
	sHasOffSeason = 1
Else
	sHasOffSeason = 0
End If 

If request("offseasonstartmonth") <> "" Then
	sOffSeasonStartMonth = clng(request("offseasonstartmonth"))
Else
	sOffSeasonStartMonth = 1
End If 

If request("offseasonstartday") <> "" Then
	sOffSeasonStartDay = clng(request("offseasonstartday"))
Else
	sOffSeasonStartDay = 1
End If 

If request("offseasonendmonth") <> "" Then
	sOffSeasonEndMonth = clng(request("offseasonendmonth"))
Else
	sOffSeasonEndMonth = 1
End If 

If request("offseasonendday") <> "" Then
	sOffSeasonEndDay = clng(request("offseasonendday"))
Else
	sOffSeasonEndDay = 1
End If 

If request("offseasonendyear") <> "" Then
	sOffSeasonEndYear = clng(request("offseasonendyear"))
Else
	sOffSeasonEndYear = 0
End If 

If request("residentrentalperiod") <> "" Then
	sResidentRentalPeriod = request("residentrentalperiod")
Else
	sResidentRentalPeriod = "NULL"
End If 

If request("nonresidentrentalperiod") <> "" Then
	sNonResidentRentalPeriod = request("nonresidentrentalperiod")
Else
	sNonResidentRentalPeriod = "NULL"
End If 

If request("reservationsduringseason") = "on" Then
	sReservationsDuringSeason = "1"
Else
	sReservationsDuringSeason = "0"
End If 

If request("nonresidentswait") = "on" Then
	sNonResidentsWait = "1"
	If request("nonresidentwaitdays") <> "" Then 
		sNonresidentWaitDays = request("nonresidentwaitdays")
	Else
		sNonresidentWaitDays = "1"	' put a 1 day delay if they did not enter a delay.
	End If 
Else
	sNonResidentsWait = "0"
	sNonresidentWaitDays = "NULL"
End If 

If request("iconimageurl") <> "" Then 
	sIconImageUrl = "'" & dbsafe(request("iconimageurl")) & "'"
Else
	sIconImageUrl = "NULL"
End If 

If iRentalId > CLng(0) Then
	sMessageFlag = "u"

	' Update existing rental
	sSql = "UPDATE egov_rentals SET"
	sSql = sSql & " rentalname = " & sRentalName
	sSql = sSql & ", locationid = " & iLocationid
	sSql = sSql & ", width = " & sWidth
	sSql = sSql & ", length = " & sLength
	sSql = sSql & ", capacity = " & sCapacity
	sSql = sSql & ", imagesplacement = " & sImagesPlacement
	sSql = sSql & ", publiccanview = " & sPublicCanView
	sSql = sSql & ", publiccanreserve = " & sPublicCanReserve
	sSql = sSql & ", needsapproval = " & sNeedsApproval
	sSql = sSql & ", supervisoruserid = " & iSupervisorUserId
	sSql = sSql & ", receiptnotes = " & sReceiptNotes
	sSql = sSql & ", terms = " & sTerms
	sSql = sSql & ", shortdescription = " & sShortDescription
	sSql = sSql & ", description = " & sDescription
	sSql = sSql & ", usehtmlonlongdesc = " & sCheckUseHTMLonLong
	sSql = sSql & ", hasoffseason = " & sHasOffSeason
	sSql = sSql & ", offseasonstartmonth = " & sOffSeasonStartMonth
	sSql = sSql & ", offseasonstartday = " & sOffSeasonStartDay
	sSql = sSql & ", offseasonendmonth = " & sOffSeasonEndMonth
	sSql = sSql & ", offseasonendday = " & sOffSeasonEndDay
	sSql = sSql & ", offseasonendyear = " & sOffSeasonEndYear
	sSql = sSql & ", residentrentalperiod = " & sResidentRentalPeriod
	sSql = sSql & ", nonresidentrentalperiod = " & sNonResidentRentalPeriod
	sSql = sSql & ", reservationsduringseason = " & sReservationsDuringSeason
	sSql = sSql & ", nonresidentswait = " & sNonResidentsWait
	sSql = sSql & ", nonresidentwaitdays = " & sNonresidentWaitDays
	sSql = sSql & ", nocosttorent = " & sNoCostToRent
	sSql = sSql & ", iconimageurl = " & sIconImageUrl
	sSql = sSql & ", isdeactivated = " & sDeactivatedCheck
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND rentalid = " & iRentalId
	session("RentalsSQL") = sSql

	RunSQLStatement sSql
'	response.write sSql & "<br />"
'	response.end
	session("RentalsSQL") = ""

	' Clear out categories
	sSql = "DELETE FROM egov_rentals_to_categories WHERE rentalid = " & iRentalId
	RunSQLStatement sSql

	' Clear out rental images
	sSql = "DELETE FROM egov_rentalimages WHERE rentalid = " & iRentalId
	RunSQLStatement sSql

	' Clear out rental documents
	sSql = "DELETE FROM egov_rentaldocuments WHERE rentalid = " & iRentalId
	RunSQLStatement sSql

	' Clear out associated rentals
	sSql = "DELETE FROM egov_rentals_to_rentals WHERE rentalid = " & iRentalId
	RunSQLStatement sSql

	' Clear the available flags on the rental days
'	sSql = "UPDATE egov_rentaldays SET isavailabletopublic = 0 WHERE rentalid = " & iRentalId & " AND orgid = " & session("orgid")
'	RunSQLStatement sSql

	' Set the available flags on the rental days that are flagged
'	For Each iDayId In request("inseasondayid")
'		sSql = "UPDATE egov_rentaldays SET isavailabletopublic = 1 WHERE rentalid = " & iRentalId
'		sSql = sSql & " AND orgid = " & session("orgid") & " AND dayid = " & iDayId
'		RunSQLStatement sSql
'	Next 

'	For Each iDayId In request("offseasondayid")
'		sSql = "UPDATE egov_rentaldays SET isavailabletopublic = 1 WHERE rentalid = " & iRentalId
'		sSql = sSql & " AND orgid = " & session("orgid") & " AND dayid = " & iDayId
'		RunSQLStatement sSql
'	Next 

	' Clear out the Items
	sSql = "DELETE FROM egov_rentalitems WHERE rentalid = " & iRentalId
	RunSQLStatement sSql

	' Clear out the rental fees
	sSql = "DELETE FROM egov_rentalfees WHERE rentalid = " & iRentalId
	RunSQLStatement sSql

	' Clear out the alerts
	sSQl = "DELETE FROM egov_rentalalerts WHERE rentalid = " & iRentalId
	RunSQLStatement sSql

Else
	sMessageFlag = "n"

	' Create a new rental 
	sSql = "INSERT INTO egov_rentals ( orgid, rentalname, locationid, width, length, capacity, imagesplacement, "
	sSql = sSql & "publiccanview, publiccanreserve, needsapproval, supervisoruserid, receiptnotes, description, shortdescription, "
	sSql = sSql & "terms, residentrentalperiod, nonresidentrentalperiod, nocosttorent, iconimageurl, hasoffseason, offseasonstartmonth, "
	sSql = sSql & "offseasonstartday, offseasonendmonth, offseasonendday, offseasonendyear, usehtmlonlongdesc, isdeactivated ) VALUES ( "
	sSql = sSql & session("orgid") & ", " & sRentalName & ", " & iLocationid & ", " & sWidth & ", " & sLength & ", "
	sSql = sSql & sCapacity & ", " & sImagesPlacement & ", " & sPublicCanView & ", " & sPublicCanReserve & ", " 
	sSql = sSql & sNeedsApproval & ", " & iSupervisorUserId & ", " & sReceiptNotes & ", " & sDescription & ", " & sShortDescription & ", "
	sSql = sSql & sTerms & ", " & sResidentRentalPeriod & ", " & sNonResidentRentalPeriod & ", " & sNoCostToRent & ", " & sIconImageUrl
	sSql = sSql & ", 0, 1, 1, 1, 31, 0, " & sCheckUseHTMLonLong & ", " & sDeactivatedCheck & " )"
	session("RentalsSQL") = sSql

	iRentalId = RunInsertStatement( sSql )

	session("RentalsSQL") = ""

	' Set up initial in season schedule 
	SetUpInitialSchedule iRentalId, 0

	' Set up initial off season schedule 
	SetUpInitialSchedule iRentalId, 1
End If 

' Handle Categories
For Each iRecreationCategoryId In Request("recreationcategoryid")
	sSql = "INSERT INTO egov_rentals_to_categories ( rentalid, recreationcategoryid ) VALUES ( " 
	sSql = sSql & iRentalId & ", " & iRecreationCategoryId & " )"
	RunSQLStatement sSql
Next 

' Handle Images
iMaxImages = CLng(request("maximages"))
For x = 1 To iMaxImages
	' If they entered an image url save it
	If request("imageurl" & x) <> "" Then 
		sAltTag = "'" & dbsafe(request("alttag" & x)) & "'"

		sSql = "INSERT INTO egov_rentalimages ( rentalid, orgid, imageurl, alttag, displayorder ) VALUES ( "
		sSql = sSql & iRentalId & ", " & session("orgid") & ", '" & dbsafe(request("imageurl" & x)) & "', " & sAltTag & ", " & x & " )"
		RunSQLStatement sSql
	End If 
Next 

ReorderImageRows iRentalId

' Handle Documents
iMaxDocuments = CLng(request("maxdocuments"))
For x = 1 To iMaxDocuments
	' If they entered a document url and name save it, otherwise do not
	If request("documenturl" & x) <> "" And request("documenttitle" & x) <> "" Then 
		sSql = "INSERT INTO egov_rentaldocuments ( rentalid, orgid, documenturl, documenttitle ) VALUES ( "
		sSql = sSql & iRentalId & ", " & session("orgid") & ", '" & dbsafe(request("documenturl" & x)) & "', '" & dbsafe(request("documenttitle" & x)) & "' )"
		RunSQLStatement sSql
	End If 
Next 

' Handle Associated Rentals
'iMaxRentals = CLng(request("maxrentals"))
'For x = 1 To iMaxRentals
'	If CLng(request("associatedrentalid" & x)) > CLng(0) Then
'		sSql = "INSERT INTO egov_rentals_to_rentals ( rentalid, associatedrentalid ) VALUES ( "
'		sSql = sSql & iRentalId & ", " & request("associatedrentalid" & x) & " )"
'		RunSQLStatement sSql
'	End If 
'Next 

' Handle Items    
iMaxItems = CLng(request("maxitems"))
'response.write iMaxItems & "<br /><br />"
For x = 1 To iMaxItems
'	response.write request("rentalitem" & x) & "<br /><br />"
'	response.write request("maxavailable" & x) & "<br /><br />"
'	response.write request("amount" & x) & "<br /><br />"
	' If they entered an item, max available and amount then save it
	If request("rentalitem" & x) <> "" And request("maxavailable" & x) <> "" And request("amount" & x) <> "" Then 
		sRentalItem = "'" & dbsafe(request("rentalitem" & x)) & "'"
		If request("itemaccountid" & x) <> "" then
			iAccountId = request("itemaccountid" & x)
		Else
			iAccountId = "NULL"
		End If 
		iMaxAvailable = request("maxavailable" & x)
		iAmount = request("amount" & x)

		sSql = "INSERT INTO egov_rentalitems ( rentalid, orgid, rentalitem, accountid, maxavailable, amount ) VALUES ( "
		sSql = sSql & iRentalId & ", " & session("orgid") & ", " & sRentalItem & ", " & iAccountId & ", "
		sSql = sSql & iMaxAvailable & ", " & iAmount & " )"
		RunSQLStatement sSql
	End If 
Next 

' Create Rental Fees
For Each iPriceTypeId In request("pricetypeid")
	If request("prompt" & iPriceTypeId) <> "" Then
		sPrompt = "'" & dbsafe(request("prompt" & iPriceTypeId)) & "'"
	Else
		sPrompt = "NULL"
	End If 

	If request("amount" & iPriceTypeId) = "" Then 
		sAmount = "0.00"
	Else
		sAmount = request("amount" & iPriceTypeId)
	End If 

	If request("accountid" & iPriceTypeId) <> "" then
		iAccountId = request("accountid" & iPriceTypeId)
	Else
		iAccountId = "NULL"
	End If 

	sSql = "INSERT INTO egov_rentalfees ( rentalid, pricetypeid, orgid, accountid, amount, prompt ) VALUES ( "
	sSql = sSql & iRentalId & ", " & iPriceTypeId & ", " & session("orgid") & ", " & iAccountId & ", "
	sSql = sSql & sAmount & ", " & sPrompt & " )"
	RunSQLStatement sSql
Next 

' Create the alerts
iMaxAlertRows = CLng(request("maxalertrows"))
For x = 1 To iMaxAlertRows
	' See if the alert type has a user selected
	If request("userid" & x) <> "" Then 
		If CLng(request("userid" & x)) > CLng(0) Then 
			sSql = "INSERT INTO egov_rentalalerts ( orgid, rentalid, rentalalerttypeid, userid ) VALUES ( "
			sSql = sSql & session("orgid") & ", " & iRentalId & ", " & CLng(request("rentalalerttypeid" & x)) & ", " & CLng(request("userid" & x)) & " )"
			'response.write sSql & "<br />"
			RunSQLStatement sSql
		End If 
	End If 
Next 

session("RentalId") = iRentalId

' Take them to the edit page for this rental
response.redirect "rentaledit.asp?rentalid=" & iRentalId & "&s=" & sMessageFlag


'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' void ReorderImageRows( iRentalId )
'-------------------------------------------------------------------------------------------------
Sub ReorderImageRows( ByVal iRentalId )
	Dim iNewOrder, oRs

	iNewOrder = CLng(0)
	
	sSql = "SELECT imageid, imageurl, displayorder FROM egov_rentalimages "
	sSql = sSql & " WHERE rentalid = " & iRentalId & " ORDER BY displayorder, imageurl"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.CursorLocation = 3
	oRs.Open sSQL, Application("DSN"), 1, 3

	Do While Not oRs.EOF
		iNewOrder = iNewOrder + CLng(1)
		oRs("displayorder") = iNewOrder
		oRs.Update
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void SetUpInitialSchedule( iRentalId, iIsOffSeason )
'-------------------------------------------------------------------------------------------------
Sub SetUpInitialSchedule( ByVal iRentalId, ByVal iIsOffSeason )
	Dim sSql, x, sDayName

	For x = 1 To 7
		sDayName = "'" & WeekDayName(x) & "'"

		sSql = "INSERT INTO egov_rentaldays ( rentalid, orgid, isoffseason, weekdayname, dayofweek ) VALUES ( "
		sSql = sSql & iRentalId & ", " & session("orgid") & ", " & iIsOffSeason & ", " & sDayName & ", " & x & " )"

		RunSQLStatement sSql
	Next 

End Sub 



%>
