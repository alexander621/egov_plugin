<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentaldayupdate.asp
' AUTHOR: Steve Loar
' CREATED: 08/25/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Updates Rental Days. Called from rentaldayedit.asp
'
' MODIFICATION HISTORY
' 1.0   08/25/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iDayId, sSql, sIsAvailableToPublic, sOpeningHour, sOpeningMinute, sOpeningAmPm, sClosingHour
Dim sClosingMinute, sClosingAmPm, sClosingDay, sLatestStartHour, sLatestStartMinute, sLatestStartAmPm
Dim sPostBuffer, sPostBufferTimeTypeId, sMinimumRental, sMinimumRentalTimeTypeId, iRentalId
Dim iPriceTypeId, iStartHour, iStartMinute, iStartAMPM, sIsOpen, iAccountId

iDayId = CLng(request("dayid"))
iRentalId = CLng(request("rentalid"))

If request("isavailabletopublic") = "on" Then
	sIsAvailableToPublic = "1"
Else
	sIsAvailableToPublic = "0"
End If 

If request("isopen") = "on" Then
	sIsOpen = "1"
Else
	sIsOpen = "0"
End If 

sOpeningHour = CLng(request("openinghour"))

sOpeningMinute = CLng(request("openingminute"))

sOpeningAmPm = "'" & request("openingampm") & "'"

sClosingHour = CLng(request("closinghour"))

sClosingMinute = CLng(request("closingminute"))

sClosingAmPm = "'" & request("closingampm") & "'"

sClosingDay = CLng(request("closingday"))

sLatestStartHour = CLng(request("lateststarthour"))

sLatestStartMinute = CLng(request("lateststartminute"))

sLatestStartAmPm = "'" & request("lateststartampm") & "'"

If request("postbuffer") <> "" Then
	sPostBuffer = CLng(request("postbuffer"))
Else
	sPostBuffer = "NULL"
End If 

sPostBufferTimeTypeId = CLng(request("postbuffertimetypeid"))

If request("minimumrental") <> "" Then
	sMinimumRental = CLng(request("minimumrental"))
Else
	sMinimumRental = "NULL"
End If 

sMinimumRentalTimeTypeId = CLng(request("minimumrentaltimetypeid"))

sSql = "UPDATE egov_rentaldays SET"
sSql = sSql & "  isavailabletopublic = " & sIsAvailableToPublic
sSql = sSql & ", isopen = " & sIsOpen
sSql = sSql & ", openinghour = " & sOpeningHour
sSql = sSql & ", openingminute = " & sOpeningMinute
sSql = sSql & ", openingampm = " & sOpeningAmPm
sSql = sSql & ", closinghour = " & sClosingHour
sSql = sSql & ", closingminute = " & sClosingMinute
sSql = sSql & ", closingampm = " & sClosingAmPm
sSql = sSql & ", closingday = " & sClosingDay
sSql = sSql & ", lateststarthour = " & sLatestStartHour
sSql = sSql & ", lateststartminute = " & sLatestStartMinute
sSql = sSql & ", lateststartampm = " & sLatestStartAmPm
sSql = sSql & ", postbuffer = " & sPostBuffer
sSql = sSql & ", postbuffertimetypeid = " & sPostBufferTimeTypeId
sSql = sSql & ", minimumrental = " & sMinimumRental
sSql = sSql & ", minimumrentaltimetypeid = " & sMinimumRentalTimeTypeId
sSql = sSql & " WHERE dayid = " & iDayId & " AND orgid = " & session("orgid") & " AND rentalid = " & iRentalId
'response.write sSql & "<br />"
RunSQLStatement sSql

' Delete the current rates
sSql = "DELETE FROM egov_rentaldayrates WHERE dayid = " & iDayId & " AND orgid = " & session("orgid") & " AND rentalid = " & iRentalId
RunSQLStatement sSql

' Add the new rates
For Each iPriceTypeId In request("pricetypeid")
	If request("starthour" & iPriceTypeId) <> "" then
		iStartHour = request("starthour" & iPriceTypeId)
	Else
		iStartHour = "NULL"
	End If 
	If request("startminute" & iPriceTypeId) <> "" then
		iStartMinute = request("startminute" & iPriceTypeId)
	Else
		iStartMinute = "NULL"
	End If 
	If request("startampm" & iPriceTypeId) <> "" then
		iStartAMPM = "'" & request("startampm" & iPriceTypeId) & "'"
	Else
		iStartAMPM = "NULL"
	End If 
	If request("accountid" & iPriceTypeId) <> "" Then
		iAccountId = request("accountid" & iPriceTypeId)
	Else
		iAccountId = "NULL"
	End If 

	sSql ="INSERT INTO egov_rentaldayrates (rentalid, dayid, orgid, pricetypeid, accountid, ratetypeid, amount, starthour, "
	sSql = sSql & " startminute, startampm) VALUES ( " & iRentalId & ", " & iDayId & ", " & session("orgid") & ", "
	sSql = sSql & iPriceTypeId & ", " & iAccountId & ", " & request("ratetypeid" & iPriceTypeId) & ", "
	sSql = sSql & request("amount" & iPriceTypeId) & ", " & iStartHour & ", " & iStartMinute & ", " & iStartAMPM & " )"
	'response.write sSql & "<br />"
	RunSQLStatement sSql
Next 

' Return to the edit page
response.redirect "rentaldayedit.asp?dayid=" & iDayId & "&s=u"


%>
