<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalcopydo.asp
' AUTHOR: Steve Loar
' CREATED: 08/26/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Copies Rentals. Called from rentalcopy.asp
'
' MODIFICATION HISTORY
' 1.0   08/26/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRentalId, sSql, oRs, sRentalName, sDescription, sPublicCanView, sPublicCanReserve, sNeedsApproval
Dim iLocationid, sWidth, sLength, sCapacity, iSupervisorUserId, sReceiptNotes, sImagesPlacement
Dim sHasOffSeason, sOffSeasonStartMonth, sOffSeasonStartDay, sOffSeasonEndMonth, sOffSeasonEndDay
Dim iMaxImages, iMaxDocuments, iMaxRentals, sButtonValue, sLoadMsg, sShortDescription, iNewRentalId
Dim iNewDayId, sIsAvailableToPublic, sOpeningHour, sOpeningMinute, sOpeningAmPm, sClosingHour
Dim sClosingMinute, sClosingAmPm, sClosingDay, sLatestStartHour, sLatestStartMinute, sLatestStartAmPm
Dim sPostBuffer, sPostBufferTimeTypeId, sMinimumRental, sMinimumRentalTimeTypeId, sIsOffSeason
Dim sTerms, sResidentRentalPeriod, sNonResidentRentalPeriod, sPrompt, sAmount, sIsOpen, iAccountId

iRentalId = CLng(request("rentalid"))

sSql = "SELECT rentalname, locationid, ISNULL(width,'') AS width, ISNULL(length,'') AS length, "
sSql = sSql & "ISNULL(capacity,'') AS capacity, imagesplacement, hasoffseason, offseasonstartmonth, "
sSql = sSql & "offseasonstartday, offseasonendmonth, offseasonendday, publiccanview, publiccanreserve, "
sSql = sSql & "needsapproval, ISNULL(supervisoruserid,0) AS supervisoruserid, ISNULL(receiptnotes,'') AS receiptnotes, "
sSql = sSql & "ISNULL(description,'') AS description, ISNULL(shortdescription,'') AS shortdescription, "
sSql = sSql & "ISNULL(terms,'') AS terms, ISNULL(residentrentalperiod,0) AS residentrentalperiod, "
sSql = sSql & "ISNULL(nonresidentrentalperiod,0) AS nonresidentrentalperiod "
sSql = sSql & "FROM egov_rentals WHERE orgid = " & session("orgid") & " AND rentalid = " & iRentalId
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

sRentalName = "'Copy of " & dbsafe(oRs("rentalname")) & "'"

If oRs("publiccanview") Then
	sPublicCanView = 1
Else
	sPublicCanView = 0
End If 

If oRs("publiccanreserve") Then
	sPublicCanReserve = 1
Else
	sPublicCanReserve = 0
End If 

If oRs("needsapproval") Then
	sNeedsApproval = 1
Else
	sNeedsApproval = 0
End If 

If CLng(oRs("locationid")) > CLng(0) Then
	iLocationid = CLng(oRs("locationid"))
Else
	iLocationid = "NULL"
End If 

If oRs("width") <> "" Then
	sWidth = "'" & DBsafeWithHTML(oRs("width")) & "'"
Else
	sWidth = "NULL"
End If 

If oRs("length") <> "" Then
	sLength = "'" & DBsafeWithHTML(oRs("length")) & "'"
Else
	sLength = "NULL"
End If 

If oRs("capacity") <> "" Then
	sCapacity = "'" & DBsafeWithHTML(oRs("capacity")) & "'"
Else
	sCapacity = "NULL"
End If 

If CLng(oRs("supervisoruserid")) > CLng(0) Then
	iSupervisorUserId = CLng(oRs("supervisoruserid"))
Else
	iSupervisorUserId = "NULL"
End If 

If oRs("description") <> "" Then
	sDescription = "'" & DBsafeWithHTML(oRs("description")) & "'"
Else
	sDescription = "NULL"
End If 

If oRs("shortdescription") <> "" Then
	sShortDescription = "'" & DBsafeWithHTML(oRs("shortdescription")) & "'"
Else
	sShortDescription = "NULL"
End If 

If oRs("receiptnotes") <> "" Then
	sReceiptNotes = "'" & DBsafeWithHTML(oRs("receiptnotes")) & "'"
Else
	sReceiptNotes = "NULL"
End If 

If oRs("terms") <> "" Then
	sTerms = "'" & DBsafeWithHTML(oRs("terms")) & "'"
Else
	sTerms = "NULL"
End If 

If oRs("imagesplacement") <> "" Then
	sImagesPlacement = "'" & dbsafe(oRs("imagesplacement")) & "'"
Else
	sImagesPlacement = "NULL"
End If 

If oRs("hasoffseason") Then
	sHasOffSeason = 1
Else
	sHasOffSeason = 0
End If 

If oRs("offseasonstartmonth") <> "" Then
	sOffSeasonStartMonth = clng(oRs("offseasonstartmonth"))
Else
	sOffSeasonStartMonth = 1
End If 

If oRs("offseasonstartday") <> "" Then
	sOffSeasonStartDay = clng(oRs("offseasonstartday"))
Else
	sOffSeasonStartDay = 1
End If 

If oRs("offseasonendmonth") <> "" Then
	sOffSeasonEndMonth = clng(oRs("offseasonendmonth"))
Else
	sOffSeasonEndMonth = 1
End If 

If oRs("offseasonendday") <> "" Then
	sOffSeasonEndDay = clng(oRs("offseasonendday"))
Else
	sOffSeasonEndDay = 1
End If 

If clng(oRs("residentrentalperiod")) > clng(0) Then
	sResidentRentalPeriod = oRs("residentrentalperiod")
Else
	sResidentRentalPeriod = "NULL"
End If 

If  clng(oRs("nonresidentrentalperiod")) > clng(0) Then
	sNonResidentRentalPeriod = oRs("nonresidentrentalperiod")
Else
	sNonResidentRentalPeriod = "NULL"
End If 


oRs.Close
Set oRs = Nothing 

' Create a new rental 
sSql = "INSERT INTO egov_rentals ( orgid, rentalname, locationid, width, length, capacity, imagesplacement, "
sSql = sSql & "publiccanview, publiccanreserve, needsapproval, supervisoruserid, receiptnotes, description, shortdescription, "
sSql = sSql & "hasoffseason, offseasonstartmonth, offseasonstartday, offseasonendmonth, offseasonendday, terms, "
sSql = sSql & "residentrentalperiod, nonresidentrentalperiod ) VALUES ( "
sSql = sSql & session("orgid") & ", " & sRentalName & ", " & iLocationid & ", " & sWidth & ", " & sLength & ", "
sSql = sSql & sCapacity & ", " & sImagesPlacement & ", " & sPublicCanView & ", " & sPublicCanReserve & ", " 
sSql = sSql & sNeedsApproval & ", " & iSupervisorUserId & ", " & sReceiptNotes & ", " & sDescription & ", " & sShortDescription & ", "
sSql = sSql & sHasOffSeason & ", " & sOffSeasonStartMonth & ", " & sOffSeasonStartDay & ", " & sOffSeasonEndMonth & ", "
sSql = sSql & sOffSeasonEndDay & ", " & sTerms & ", " & sResidentRentalPeriod & ", " & sNonResidentRentalPeriod & " )"
'response.write sSql & "<br /><br />"

iNewRentalId = RunInsertStatement( sSql )


' Copy the season day schedules
sSql = "SELECT dayid, orgid, weekdayname, dayofweek, isoffseason, isopen, isavailabletopublic, "
sSql = sSql & "ISNULL(openinghour,0) AS openinghour, ISNULL(openingminute,0) AS openingminute, "
sSql = sSql & "ISNULL(openingampm,'AM') AS openingampm, ISNULL(closinghour,0) AS closinghour, "
sSql = sSql & "ISNULL(closingminute,0) AS closingminute, ISNULL(closingampm,'PM') AS closingampm, "
sSql = sSql & "ISNULL(closingday,0) AS closingday, ISNULL(lateststarthour,0) AS lateststarthour, "
sSql = sSql & "ISNULL(lateststartminute,0) AS lateststartminute, ISNULL(lateststartampm,'PM') AS lateststartampm, "
sSql = sSql & "ISNULL(postbuffer,0) AS postbuffer, ISNULL(postbuffertimetypeid,0) AS postbuffertimetypeid, "
sSql = sSql & "ISNULL(minimumrental,0) AS minimumrental, ISNULL(minimumrentaltimetypeid,0) AS minimumrentaltimetypeid "
sSql = sSql & "FROM egov_rentaldays "
sSql = sSql & "WHERE rentalid = " & iRentalId
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

Do While Not oRs.EOF
	If oRs("isoffseason") Then 
		sIsOffSeason = 1
	Else
		sIsOffSeason = 0
	End If 

	If oRs("isavailabletopublic") Then 
		sIsAvailableToPublic = 1
	Else
		sIsAvailableToPublic = 0
	End If 

	If oRs("isopen") Then
		sIsOpen = "1"
	Else
		sIsOpen = "0"
	End If 

	If clng(oRs("openinghour")) > clng(0) then
		sOpeningHour = clng(oRs("openinghour"))
		sOpeningMinute = clng(oRs("openingminute"))
	Else
		sOpeningHour = "NULL"
		sOpeningMinute = "NULL"
	End If 

	sOpeningAmPm = "'" & oRs("openingampm") & "'"

	If clng(oRs("closinghour")) > clng(0) then
		sClosingHour = clng(oRs("closinghour"))
		sClosingMinute = clng(oRs("closingminute"))
	Else
		sClosingHour = "NULL"
		sClosingMinute = "NULL"
	End If 

	sClosingAmPm = "'" & oRs("closingampm") & "'"

	sClosingDay = oRs("closingday")

	If clng(oRs("lateststarthour")) > clng(0) then
		sLatestStartHour = clng(oRs("lateststarthour"))
		sLatestStartMinute = clng(oRs("lateststartminute"))
	Else
		sLatestStartHour = "NULL"
		sLatestStartMinute = "NULL"
	End If 

	sLatestStartAmPm = "'" & oRs("lateststartampm") & "'"

	If clng(oRs("postbuffer")) <> clng(0) Then
		sPostBuffer = clng(oRs("postbuffer"))
	Else
		sPostBuffer = "NULL"
	End If 

	sPostBufferTimeTypeId = oRs("postbuffertimetypeid")

	If clng(oRs("minimumrental")) > clng(0) Then
		sMinimumRental = clng(oRs("minimumrental"))
	Else
		sMinimumRental = "NULL"
	End If 

	sMinimumRentalTimeTypeId = oRs("minimumrentaltimetypeid")

	sSql = "INSERT INTO egov_rentaldays ( rentalid, orgid, weekdayname, dayofweek, isoffseason, isopen, isavailabletopublic, "
	sSql = sSql & "openinghour, openingminute, openingampm, closinghour, closingminute, closingampm, closingday, "
	sSql = sSql & "lateststarthour, lateststartminute, lateststartampm, postbuffer, postbuffertimetypeid, "
	sSql = sSql & "minimumrental, minimumrentaltimetypeid ) VALUES ( " & iNewRentalId & ", " & oRs("orgid") & ", '"
	sSql = sSql & oRs("weekdayname") & "', " & oRs("dayofweek") & ", " & sIsOffSeason & ", " & sIsOpen & ", " & sIsAvailableToPublic & ", "
	sSql = sSql & sOpeningHour & ", " & sOpeningMinute & ", " & sOpeningAmPm & ", " & sClosingHour & ", " 
	sSql = sSql & sClosingMinute & ", " & sClosingAmPm & ", " & sClosingDay & ", " & sLatestStartHour & ", "
	sSql = sSql & sLatestStartMinute & ", " & sLatestStartAmPm & ", " & sPostBuffer & ", " 
	sSql = sSql & sPostBufferTimeTypeId & ", " & sMinimumRental & ", " & sMinimumRentalTimeTypeId & " )"
	'response.write sSql & "<br /><br />"

	iNewDayId = RunInsertStatement( sSql )

	' Copy the day rates for each day pulled
	CopyRentalDayRates oRs("dayid"), iNewDayId, iNewRentalId
	oRs.MoveNext
Loop

oRs.Close
Set oRs = Nothing 


' Copy the Categories
sSql = "SELECT recreationcategoryid FROM egov_rentals_to_categories WHERE rentalid = " & iRentalId
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

Do While Not oRs.EOF
	sSql = "INSERT INTO egov_rentals_to_categories ( rentalid, recreationcategoryid ) VALUES ( " 
	sSql = sSql & iNewRentalId & ", " & oRs("recreationcategoryid") & " )"
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql
	oRs.MoveNext 
Loop

oRs.Close 
Set oRs = Nothing 


' Copy the Images
sSql = "SELECT imageurl, ISNULL(alttag,'') AS alttag, displayorder FROM egov_rentalimages WHERE rentalid = " & iRentalId
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

Do While Not oRs.EOF
	If oRs("alttag") <> "" Then
		sAltTag = "'" & dbsafe(oRs("alttag")) & "'"
	Else
		sAltTag = "NULL"
	End If 
	sSql = "INSERT INTO egov_rentalimages ( rentalid, orgid, imageurl, alttag, displayorder ) VALUES ( "
	sSql = sSql & iNewRentalId & ", " & session("orgid") & ", '" & dbsafe(oRs("imageurl")) & "', " & sAltTag & ", " & oRs("displayorder") & " )"
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql
	oRs.MoveNext 
Loop

oRs.Close 
Set oRs = Nothing 


' Copy the Documents
sSql = "SELECT orgid, documenturl, documenttitle FROM egov_rentaldocuments WHERE rentalid = " & iRentalId
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

Do While Not oRs.EOF
	sSql = "INSERT INTO egov_rentaldocuments ( rentalid, orgid, documenturl, documenttitle ) VALUES ( "
	sSql = sSql & iNewRentalId & ", " & oRs("orgid") & ", '" & dbsafe(oRs("documenturl")) & "', '" & dbsafe(oRs("documenttitle")) & "' )"
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql
	oRs.MoveNext 
Loop

oRs.Close 
Set oRs = Nothing 


' Copy the Associated Rentals
'sSql = "SELECT associatedrentalid FROM egov_rentals_to_rentals WHERE rentalid = " & iRentalId
'response.write sSql & "<br /><br />"

'Set oRs = Server.CreateObject("ADODB.Recordset")
'oRs.Open sSql, Application("DSN"), 3, 1

'Do While Not oRs.EOF
'	sSql = "INSERT INTO egov_rentals_to_rentals ( rentalid, associatedrentalid ) VALUES ( "
'	sSql = sSql & iNewRentalId & ", " & oRs("associatedrentalid") & " )"
'	'response.write sSql & "<br /><br />"
'	RunSQLStatement sSql
'	oRs.MoveNext 
'Loop

'oRs.Close 
'Set oRs = Nothing 

' Copy the Rental Items
sSql = "SELECT orgid, rentalitem, ISNULL(accountid,0) AS accountid, maxavailable, amount FROM egov_rentalitems WHERE rentalid = " & iRentalId
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

Do While Not oRs.EOF
	If CLng(oRs("accountid")) = CLng(0) Then
		iAccountId = "NULL"
	Else
		iAccountId = oRs("accountid")
	End If 
	sSql = "INSERT INTO egov_rentalitems ( rentalid, orgid, rentalitem, accountid, maxavailable, amount ) VALUES ( "
	sSql = sSql & iNewRentalId & ", " & oRs("orgid") & ", '" & dbsafe(oRs("rentalitem")) & "', " & iAccountId & ", "
	sSql = sSql & oRs("maxavailable") & ", " & oRs("amount") & " )"
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql
	oRs.MoveNext 
Loop

oRs.Close 
Set oRs = Nothing 


' Copy the Rental Fees
sSql = "SELECT pricetypeid, orgid, ISNULL(accountid,0) AS accountid, ISNULL(amount,0) AS amount, ISNULL(prompt,'') AS prompt "
sSql = sSql & "FROM egov_rentalfees WHERE rentalid = " & iRentalId
response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

Do While Not oRs.EOF
	If oRs("prompt") <> "" Then
		sPrompt = "'" & dbsafe(oRs("prompt")) & "'"
	Else
		sPrompt = "NULL"
	End If 

	If oRs("amount") <> "" Then
		sAmount = oRs("amount")
	Else
		sAmount = "NULL"
	End If

	If CLng(oRs("accountid")) > CLng(0) Then
		iAccountId = oRs("accountid")
	Else
		iAccountId = "NULL"
	End If

	sSql = "INSERT INTO egov_rentalfees (rentalid, pricetypeid, orgid, accountid, amount, prompt) VALUES ( "
	sSql = sSql & iNewRentalId & ", " & oRs("pricetypeid") & ", " & oRs("orgid") & ", " & iAccountId & ", "
	sSql = sSql & sAmount & ", " & sPrompt & " )"
	response.write sSql & "<br /><br />"
	RunSQLStatement sSql
	oRs.MoveNext 
Loop

oRs.Close 
Set oRs = Nothing 

' Copy the rental alerts
sSql = "SELECT rentalalerttypeid, userid FROM egov_rentalalerts WHERE rentalid = " & iRentalId
response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

Do While Not oRs.EOF
	sSql = "INSERT INTO egov_rentalalerts ( orgid, rentalid, rentalalerttypeid, userid ) VALUES ( "
	sSql = sSql & session("orgid") & ", " & iNewRentalId & ", " & oRs("rentalalerttypeid") & ", " & oRs("userid") & " )"
	'response.write sSql & "<br />"
	RunSQLStatement sSql
	oRs.MoveNext 
Loop 

oRs.Close
Set oRs = Nothing 

' Take them to the edit page for the new rental
response.redirect "rentaledit.asp?rentalid=" & iNewRentalId & "&s=c" 




'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Sub CopyRentalDayRates iDayId, iNewDayId, iNewRentalId
'-------------------------------------------------------------------------------------------------
Sub CopyRentalDayRates( ByVal iDayId, ByVal iNewDayId, ByVal iNewRentalId )
	Dim sSql, oRs, iStartHour, iStartMinute, iStartAMPM

	sSql = "SELECT pricetypeid, orgid, ISNULL(accountid,0) AS accountid, ratetypeid, amount, ISNULL(starthour,0) AS starthour, "
	sSql = sSql & " ISNULL(startminute,0) AS startminute, ISNULL(startampm,'') AS startampm "
	sSql = sSql & " FROM egov_rentaldayrates WHERE dayid = " & iDayId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		If clng(oRs("starthour")) > clng(0) then
			iStartHour = oRs("starthour")
			iStartMinute = oRs("startminute")
		Else
			iStartHour = "NULL"
			iStartMinute = "NULL"
		End If 

		If oRs("startampm") <> "" then
			iStartAMPM = "'" & oRs("startampm") & "'"
		Else
			iStartAMPM = "NULL"
		End If

		If CLng(oRs("accountid")) > CLng(0) Then
			iAccountId = oRs("accountid")
		Else
			iAccountId = "NULL"
		End If
		
		sSql = "INSERT INTO egov_rentaldayrates ( rentalid, dayid, pricetypeid, orgid, accountid, "
		sSql = sSql & "ratetypeid, amount, starthour, startminute, startampm ) VALUES ( " 
		sSql = sSql & iNewRentalId & ", " & iNewDayId & ", " & oRs("pricetypeid") & ", " & oRs("orgid") & ", "
		sSql = sSql & iAccountId & ", " & oRs("ratetypeid") & ", " & oRs("amount") & ", "
		sSql = sSql & iStartHour & ", " & iStartMinute & ", " & iStartAMPM & " )"
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql
		oRs.MoveNext 
	Loop

	oRs.Close 
	Set oRs = Nothing 

End Sub 


%>
