<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentaldaycopydo.asp
' AUTHOR: Steve Loar
' CREATED: 08/26/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Copies Rental Days. Called from rentaldaycopy.asp
'
' MODIFICATION HISTORY
' 1.0   08/26/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iSourceDayId, iTargetDayId, sSql, oRs, sIsAvailableToPublic, sOpeningHour, sOpeningMinute, sOpeningAmPm
Dim sClosingMinute, sClosingAmPm, sClosingDay, sLatestStartHour, sLatestStartMinute, sLatestStartAmPm
Dim sPostBuffer, sPostBufferTimeTypeId, sMinimumRental, sMinimumRentalTimeTypeId, iRentalId
Dim iPriceTypeId, iStartHour, iStartMinute, iStartAMPM, sClosingHour, sIsOpen

iSourceDayId = CLng(request("sourcedayid"))
iTargetDayId = CLng(request("targetdayid"))
iRentalId = CLng(request("rentalid"))

' Get the source Day information
sSql = "SELECT R.rentalid, R.rentalname, L.name AS locationname, D.weekdayname, D.isoffseason, D.isopen, D.isavailabletopublic, "
sSql = sSql & "ISNULL(D.openinghour,0) AS openinghour, ISNULL(D.openingminute,0) AS openingminute, "
sSql = sSql & "ISNULL(D.openingampm,'AM') AS openingampm, ISNULL(D.closinghour,0) AS closinghour, "
sSql = sSql & "ISNULL(D.closingminute,0) AS closingminute, ISNULL(D.closingampm,'PM') AS closingampm, "
sSql = sSql & "ISNULL(D.closingday,0) AS closingday, ISNULL(D.lateststarthour,0) AS lateststarthour, "
sSql = sSql & "ISNULL(D.lateststartminute,0) AS lateststartminute, ISNULL(D.lateststartampm,'PM') AS lateststartampm, "
sSql = sSql & "ISNULL(D.postbuffer,0) AS postbuffer, ISNULL(D.postbuffertimetypeid,0) AS postbuffertimetypeid, "
sSql = sSql & "ISNULL(D.minimumrental,0) AS minimumrental, ISNULL(D.minimumrentaltimetypeid,0) AS minimumrentaltimetypeid "
sSql = sSql & "FROM egov_rentaldays D, egov_rentals R, egov_class_location L "
sSql = sSql & "WHERE D.rentalid = R.rentalid AND R.locationid = L.locationid AND D.dayid = " & iSourceDayId & " AND R.orgid = " & session("orgid")
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

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

oRs.Close
Set oRs = Nothing 

' Update the target day
sSql = "UPDATE egov_rentaldays SET"
sSql = sSql & "  isavailabletopublic = " & sIsAvailableToPublic
sSql = sSql & ",  isopen = " & sIsOpen
sSql = sSql & ",  openinghour = " & sOpeningHour
sSql = sSql & ",  openingminute = " & sOpeningMinute
sSql = sSql & ",  openingampm = " & sOpeningAmPm
sSql = sSql & ",  closinghour = " & sClosingHour
sSql = sSql & ",  closingminute = " & sClosingMinute
sSql = sSql & ",  closingampm = " & sClosingAmPm
sSql = sSql & ",  closingday = " & sClosingDay
sSql = sSql & ",  lateststarthour = " & sLatestStartHour
sSql = sSql & ",  lateststartminute = " & sLatestStartMinute
sSql = sSql & ",  lateststartampm = " & sLatestStartAmPm
sSql = sSql & ",  postbuffer = " & sPostBuffer
sSql = sSql & ",  postbuffertimetypeid = " & sPostBufferTimeTypeId
sSql = sSql & ",  minimumrental = " & sMinimumRental
sSql = sSql & ",  minimumrentaltimetypeid = " & sMinimumRentalTimeTypeId
sSql = sSql & " WHERE dayid = " & iTargetDayId & " AND orgid = " & session("orgid") & " AND rentalid = " & iRentalId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Delete the Target Day rates
sSql = "DELETE FROM egov_rentaldayrates WHERE dayid = " & iTargetDayId & " AND orgid = " & session("orgid") & " AND rentalid = " & iRentalId
response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Get the Source Day Rates
sSql = "SELECT pricetypeid, ISNULL(accountid,0) AS accountid, ISNULL(ratetypeid,0) AS ratetypeid, ISNULL(amount,0) AS amount, "
sSql = sSql & " ISNULL(starthour,0) AS starthour, ISNULL(startminute,0) AS startminute, ISNULL(startampm,'AM') AS startampm "
sSql = sSql & " FROM egov_rentaldayrates "
sSql = sSql & " WHERE rentalid = " & iRentalId & " AND dayid = " & iSourceDayId 
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

' Add the new rates
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

	sSql ="INSERT INTO egov_rentaldayrates (rentalid, dayid, orgid, pricetypeid, accountid, ratetypeid, amount, starthour, "
	sSql = sSql & " startminute, startampm) VALUES ( " & iRentalId & ", " & iTargetDayId & ", " & session("orgid") & ", "
	sSql = sSql & oRs("pricetypeid") & ", " & oRs("accountid") & ", " & oRs("ratetypeid") & ", " & oRs("amount") & ", "
	sSql = sSql & iStartHour & ", " & iStartMinute & ", " & iStartAMPM & " )"
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql

	oRs.MoveNext 
Loop  

oRs.Close
Set oRs = Nothing 

' Return to the edit page
response.redirect "rentaldayedit.asp?dayid=" & iTargetDayId & "&s=c"


%>
