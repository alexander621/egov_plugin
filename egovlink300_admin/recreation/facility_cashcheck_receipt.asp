<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: FACILITY_CASHCHECK_RECEIPT.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/14/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   02/14/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/06/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "reservations" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>

<html lang="en">
<head>
	<meta charset="UTF-8">

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="reservation.css" />

</head>

<body>
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN: PAGE CONTENT-->
<div style="padding:20px;">

<input type="button" class="facilitybutton" onclick="location.href='facility_calendar.asp?<%=request("backlink")%>';" value="Return to Reservation Calendar" /><br /><br />

<% DisplayReceipt sOrderID %>

</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>



<%
'--------------------------------------------------------------------------------------------------
' SUB DISPLAYRECIEPT()
'--------------------------------------------------------------------------------------------------
Sub DisplayReceipt( ByVal sOrderID )
		
	' CHECK FOR RECURRENCY
	If request("isrecursive") = "on" Then
		' RECURRENT RESERVATION PAYMENTS
		ProcessRecurrent()
	Else
		' STORE RESERVATION INFORMATION
		iFacilityPaymentID = StoreFacilityInformation( session("orgid") )

		' SINGLE RESERVATION
		UpdateFacilityPayment iFacilityPaymentID, "EGOVLINK_ADMIN", "EGOVLINK_ADMIN", "APPROVED", "APPROVED", request("paymenttype"), request("paymentlocation")
	End If

	' DISPLAY RECEIPT
	response.write "<p><div style=""border: 1px solid #000000; padding: 10px;""><p>Your purchase has been <b>approved</b>.<br /> You will receive a confirmation email containing this receipt.  It is also recommended that you print this page as proof of your purchase.<blockquote>"

	If OrgHasDisplay( session("orgid"), "facility receipt notes top" ) Then
		response.write vbcrlf & GetOrgDisplay( session("orgid"), "facility receipt notes top" )
	End If

	' TRANSACTION RESULT DETAILS
	response.write "<table>"
	response.write "<tr><td colspan=2><b>Transaction Details</b></td></tr>"
	response.write "<tr><td><font color=#000000>Amount Charged: </font></td><td> " & formatcurrency(request("amount"),2) & "</td></tr>"
	response.write "<tr><td><font color=#000000>Order Number:</font></td><td> " & iFacilityPaymentID & "F3000 </td></tr>"
	response.write "<tr><td><font color=#000000>Payment Type:</font></td><td> " & GetPaymentTypeName(request("paymenttype")) & " </td></tr>"
	response.write "<tr><td><font color=#000000>Payment Location:</font></td><td> " & GetPaymentLocationName(request("paymentlocation")) & " </td></tr>"
	response.write "</table>"


	' PRODUCT INFORMATION
	response.write "<p><table>"
	response.write "<tr><td colspan=2><b>Reservation Information</b></td></tr>"
	response.write "<tr><td>Product: </td><td>(" & request("ITEM_NUMBER")& ") " & request("ITEM_NAME") & " - " & request("LODGENAME") & "</td></tr>"
	response.write "<tr><td valign=top>Details: </td><td  valign=top>" & request("checkindate") & " " & request("checkintime") & " - " & request("checkoutdate") & " " & request("checkouttime") & "</td></tr>"
	If request("isrecursive") = "on" Then 
		response.write vbcrlf & "<tr><td valign=""top"">Recurs: </td><td  valign=""top"">" & request("wrecurrenttimepart") & "</td></tr>"
		response.write vbcrlf & "<tr><td valign=""top"">Ending By: </td><td  valign=""top"">" & request("recurrentenddate") & "</td></tr>"
	End If 
	response.write vbcrlf & "</table>"


	' CREDIT CARD INFORMATION	
	UserInfo(request("lesseeid"))

	If OrgHasDisplay( session("orgid"), "facility receipt notes bottom" ) Then
		response.write vbcrlf & GetOrgDisplay( session("orgid"), "facility receipt notes bottom" )
	End If 

	response.write vbcrlf & "</blockquote>"
	response.write vbcrlf & "</div>"

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION GETPAYMENTTYPENAME(IPAYMENTTYPEID)
'--------------------------------------------------------------------------------------------------
Function GetPaymentTypeName( ByVal ipaymenttypeid )
	Dim sReturnValue, sSql, oRs

		' SET DEFAULT RETURN VALUE
		sReturnValue = "UNKNOWN"

		' SELECT PAYMENT TYPE NAME
		sSql = "SELECT paymenttypename FROM egov_paymenttypes WHERE paymenttypeid = " & ipaymenttypeid 

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		' IF PAYMENT TYPE ID FOUND
		If Not oRs.EOF Then
			'SET RETURN VALUE TO PAYMENT NAME
			sReturnValue = oRs("paymenttypename")
		End If
		
		oRs.Close
		Set oRs = Nothing

		' RETURN VALUE
		GetPaymentTypeName = sReturnValue
			
End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETPAYMENTLOCATIONNAME(IPAYMENTLOCATIONID)
'--------------------------------------------------------------------------------------------------
Function GetPaymentLocationName( ByVal ipaymentlocationid )
	Dim oRs, sSql, sReturnValue

	' SET DEFAULT RETURN VALUE
	sReturnValue = "UNKNOWN"

	' SELECT PAYMENT TYPE NAME
	sSql = "SELECT paymentlocationname FROM egov_paymentlocations WHERE paymentlocationid = " & ipaymentlocationid 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	' IF PAYMENT TYPE ID FOUND
	If not oRs.EOF Then
		'SET RETURN VALUE TO PAYMENT NAME
		sReturnValue = oRs("paymentlocationname")
	End If
	
	' CLEAN UP OBJECT
	oRs.Close
	Set oRs = Nothing

	' RETURN VALUE
	GetPaymentLocationName = sReturnValue
		
End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION STOREFACILITYINFORMATION()
'------------------------------------------------------------------------------------------------------------
Function StoreFacilityInformation( ByVal iOrgID )
	Dim oCmd, iReturnValue
		
		iReturnValue = 0

		Set oCmd = Server.CreateObject("ADODB.Command")
		 With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "StoreFacilityInformation"
		.CommandType = 4

		' RESERVATION INFORMATION
		.Parameters.Append oCmd.CreateParameter("checkintime", 200, 1, 50, request("checkintime"))
		.Parameters.Append oCmd.CreateParameter("checkindate", 200, 1, 50, request("checkindate"))
		.Parameters.Append oCmd.CreateParameter("checkouttime", 200, 1, 50, request("checkouttime"))
		.Parameters.Append oCmd.CreateParameter("checkoutdate", 200, 1, 50, request("checkoutdate"))
		.Parameters.Append oCmd.CreateParameter("lesseeid", 3, 1, 4, request("lesseeid"))
		.Parameters.Append oCmd.CreateParameter("facilitytimepartid", 3, 1, 4, request("timepartid"))
		.Parameters.Append oCmd.CreateParameter("facilityid", 3, 1, 4, request("facilityid"))
		.Parameters.Append oCmd.CreateParameter("orgid", 3, 1, 4, iOrgID)
		.Parameters.Append oCmd.CreateParameter("amount", 3, 1, 4, request("amount"))
		.Parameters.Append oCmd.CreateParameter("internalnote", 200, 1, 1024, request("internalnote"))
		.Parameters.Append oCmd.CreateParameter("sessionid",3,1,4,session.sessionid)
		.Parameters.Append oCmd.CreateParameter("facilitypaymentid", 3, 2, 4)

		' CALL TO STORE VALUE INFORMATION
		.Execute

		iFacilityPaymentID = .Parameters("facilitypaymentid")
		
		If iFacilityPaymentID <> "" Then
			iReturnValue = iFacilityPaymentID
		End If

		' STORE FIELD VALUES
		StoreFacilityFieldValues iFacilityPaymentID

     End With

	 Set oCmd = Nothing

	StoreFacilityInformation = iReturnValue

End Function


'------------------------------------------------------------------------------------------------------------
' SUB STOREFACILITYFIELDVALUES(IFACILITYPAYMENTID)
'------------------------------------------------------------------------------------------------------------
Sub StoreFacilityFieldValues( ByVal iFacilityPaymentID )
		
		' LOOP THRU EACH OF THE FIELDS AND ENTER VALUES SUBMITTED
		For Each oField IN Request.Form
			
			If Left(oField,7) = "custom_" Then
				' GET VALUES
				arrValues = split(oField,"_")
				iFieldID = clng(arrValues(2))

				 Set oCmd = Server.CreateObject("ADODB.Command")
				 With oCmd
					.ActiveConnection = Application("DSN")
					.CommandText = "StoreFacilityFieldValues"
					.CommandType = 4
					
					' STORE VALUES
					.Parameters.Append oCmd.CreateParameter("iFieldID", 3, 1, 4, iFieldID)
					.Parameters.Append oCmd.CreateParameter("iFacilityPaymentID", 3, 1, 4, iFacilityPaymentID)
					.Parameters.Append oCmd.CreateParameter("sValue", 200, 1, 2000, request(oField))
					.Execute
				End With

				Set oCmd = Nothing
			End If
		Next

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB UPDATEFACILITYPAYMENT(IFACILITYPAYMENTID, SAUTHCODE, SPNREF, SRESULT, SRESPMSG, itype, ilocation)
'------------------------------------------------------------------------------------------------------------
Sub UpdateFacilityPayment( ByVal iFacilityPaymentId, ByVal sAuthCode, ByVal sPNRef, ByVal sResult, ByVal sRespMsg, ByVal itype, ByVal ilocation )
	Dim oCmd 

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "UpdateFacilityPayment"
		.CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iFacilityPaymentID", 3, 1, 4, iFacilityPaymentId)
		.Parameters.Append oCmd.CreateParameter("@sAuthCode", 200, 1, 50, sAuthCode)
		.Parameters.Append oCmd.CreateParameter("@sPNRef", 200, 1, 50, sPNRef)
		.Parameters.Append oCmd.CreateParameter("@sResult", 200, 1, 50, sResult)
		.Parameters.Append oCmd.CreateParameter("@sReplyMsg", 200, 1, 255, sRespMsg)
		.Parameters.Append oCmd.CreateParameter("@paymentlocation", 3, 1, 4, ilocation)
		.Parameters.Append oCmd.CreateParameter("@paymenttype", 3, 1, 4, itype)
		.Parameters.Append oCmd.CreateParameter("@status", 200, 1, 50, request("reservationstatus"))
		.Parameters.Append oCmd.CreateParameter("@adminid", 3, 1, 4, request.cookies("user")("userid"))
		.Execute
	End With

	Set oCmd = Nothing

End Sub


'------------------------------------------------------------------------------------------------------------
' Sub UserInfo(iUserID)
'------------------------------------------------------------------------------------------------------------
Sub UserInfo( ByVal iUserID )
	Dim sSql, oRs

	' SELECT ROW WITH THIS USER'S INFORMATION
	sSql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, ISNULL(useraddress,'') AS useraddress, "
	sSql = sSql & "ISNULL(usercity,'') AS usercity, ISNULL(userstate,'') AS userstate, ISNULL(userzip,'') AS userzip FROM egov_users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	' IF RECORDSET NOT EMPTY THEN DISPLAY USER INFORMATION
	If NOT oRs.EOF Then
		
		response.write "<p><table>"
		response.write "<tr><td colspan=2><b>User Information</b></td></tr>"
		response.write "<tr><td>Name: </td><td>" & oRs("userfname") & " " & oRs("userlname")  &  "</td></tr>"
		response.write "<tr><td>Address: </td><td>" & oRs("useraddress") & "</td></tr>"
		response.write "<tr><td>City: </td><td>" & oRs("usercity") & "</td></tr>"
		response.write "<tr><td>State: </td><td>" & oRs("userstate") & "</td></tr>"
		response.write "<tr><td>Zip: </td><td>" & oRs("userzip") & "</td></tr>"
		response.write "</table></p>"

	End If

	oRs.Close 
	Set oRs = Nothing
	
End Sub


'--------------------------------------------------------------------------------------------------
'  SUB SETWEEKLYDATES(SSTARTDATE,SENDDATE,WFREQUENCY,WDAYOFTHEWEEK)
'--------------------------------------------------------------------------------------------------
Sub SetWeeklyDates( ByVal sStartDate, ByVal sEndDate, ByVal wfrequency, ByVal wdayoftheweek )
	Dim dTemp, sStatus

	' GET NEXT DATE OF SPECIFIED DAY OF WEEK (COULD BE START DATE BUT NOT ALWAYS)
	dTemp = GetNextWeekDay( wdayoftheweek, sStartDate )

	' LOOP UNTIL END DATE
	Do While CDate(dTemp) <= CDate(sEndDate)
		' WRITE DAY OF WEEK
		
		' GET CURRENT STATUS OF TIMEPART AND DATE
		sStatus = GetTimePartStatus( request("facilityid"), request("timepartid"), dTemp )
		
		' NOT RESERVED OR ONHOLD THEN MARK AVAILABLE
		If sStatus = "OPEN" or sStatus = "CANCELLED" Then
			' STORE RESERVATION INFORMATION
			iFacilityPaymentID = StoreFacilityInformationRecurrent( session("orgid"), dTemp, dTemp, request("timepartid") )
			
			' UPDATE DATABASE
			UpdateFacilityPayment iFacilityPaymentID, "EGOVLINK_ADMIN","EGOVLINK_ADMIN","APPROVED","APPROVED",request("paymenttype"),request("paymentlocation")
		Else
			' DISPLAY NOT AVAILABLE
			response.write UCase("<font color=""red""><b>" & dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  " NOT AVAILABLE!</b></font><br>")
		End If

		' ADD SPECIFIED WEEK FREQUENCY
		dTemp = DateAdd("ww",wfrequency,dTemp)
	Loop

End Sub


'--------------------------------------------------------------------------------------------------
'  SUB SETMONTHLYDATES(SSTARTDATE,SENDDATE,MSERIES,MFREQUENCY,MDAYOFTHEWEEK)
'--------------------------------------------------------------------------------------------------
Sub SetMonthlyDates( ByVal sStartDate, ByVal sEndDate, ByVal mseries, ByVal mfrequency, ByVal mdayoftheweek )
	Dim dTemp, sStatus

	' GET NEXT DATE OF SPECIFIED DAY OF WEEK (COULD BE START DATE BUT NOT ALWAYS)
	dTemp = GetNextOrdinalDayMonth( mdayoftheweek, mseries, Month(sStartDate), Year(sStartDate) )

	' LOOP UNTIL END DATE
	Do While CDate(dTemp) <= CDate(sEndDate) 

		' IF SERIES IS GREATER THAN OR EQUAL THEN USE VALUE
		If Cdate(dTemp) >= Cdate(sStartDate)  Then

			' GET CURRENT STATUS OF TIMEPART AND DATE
			sStatus = GetTimePartStatus( request("facilityid"), request("timepartid"), dTemp ) 
			
			' NOT RESERVED OR ONHOLD THEN MARK AVAILABLE
			If sStatus = "OPEN" or sStatus = "CANCELLED" Then
				' STORE RESERVATION INFORMATION
				iFacilityPaymentID = StoreFacilityInformationRecurrent( session("orgid"), dTemp, dTemp, request("timepartid") )
				
				' UPDATE DATABASE
				UpdateFacilityPayment iFacilityPaymentID, "EGOVLINK_ADMIN", "EGOVLINK_ADMIN", "APPROVED", "APPROVED", request("paymenttype"), request("paymentlocation")
			Else	
				' DISPLAY NOT AVAILABLE
				response.write UCase("<font color=""red""><b>" & dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  " NOT AVAILABLE!</b></font><br>")
			End If

		End If

		' ADD SPECIFIED MONTHLY FREQUENCY
		dTemp = DateAdd("m",mfrequency,dTemp)

		' GET NEXT DATE OF SPECIFIED DAY OF WEEK AND SERIES
		dTemp = GetNextOrdinalDayMonth( mdayoftheweek, mseries, Month(dTemp), Year(dTemp) )
	Loop

End Sub


'--------------------------------------------------------------------------------------------------
'  SUB SETYEARLYDATES(SSTARTDATE,SENDDATE,YSERIES,YDAYOFTHEWEEK,YMONTH)
'--------------------------------------------------------------------------------------------------
Sub SetYearlyDates( ByVal sStartDate, ByVal sEndDate, ByVal yseries, ByVal ydayoftheweek, ByVal ymonth)

	' GET NEXT DATE OF SPECIFIED DAY OF WEEK (COULD BE START DATE BUT NOT ALWAYS)
	dTemp = GetNextOrdinalDayMonth( ydayoftheweek, yseries, ymonth, Year(sStartDate) )


	' LOOP UNTIL END DATE
	Do While cdate(dTemp) <= cdate(sEndDate)
		
		' IF SERIES IS GREATER THAN OR EQUAL THEN USE VALUE
		If Cdate(dTemp) >= Cdate(sStartDate) Then
		
			' GET CURRENT STATUS OF TIMEPART AND DATE
			sStatus = GetTimePartStatus( request("facilityid"), request("timepartid"), dTemp )
			
			' NOT RESERVED OR ONHOLD THEN MARK AVAILABLE
			If sStatus = "OPEN" or sStatus = "CANCELLED" Then
				' STORE RESERVATION INFORMATION
				iFacilityPaymentID = StoreFacilityInformationRecurrent( session("orgid"), dTemp, dTemp, request("timepartid") )
				
				' UPDATE DATABASE
				UpdateFacilityPayment iFacilityPaymentID, "EGOVLINK_ADMIN", "EGOVLINK_ADMIN", "APPROVED", "APPROVED", request("paymenttype"), request("paymentlocation")
			Else
				' DISPLAY NOT AVAILABLE
				response.write ucase("<font color=""red""><b>" & dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  " NOT AVAILABLE!</b></font><br>")
			End If

		End If

		' ADD SPECIFIED 1 TO YEAR
		dTemp = DateAdd("yyyy",1,dTemp)

		' GET NEXT DATE OF SPECIFIED DAY OF WEEK AND SERIES
		dTemp = GetNextOrdinalDayMonth( ydayoftheweek, yseries, Month(dTemp), Year(dTemp) )
	Loop

End Sub


'--------------------------------------------------------------------------------------------------
'  FUNCTION GETNEXTWEEKDAY(IWEEKDAY,DTEMPDATE)
'--------------------------------------------------------------------------------------------------
Function GetNextWeekDay( ByVal iWeekDay, ByVal dTempdate )
  
	' LOOP TO THE NEXT SPECIFIED DAY OF THE WEEK
	Do While Not  clng(WeekDay(dTempdate)) = clng(iWeekDay) 
		' ADD 1 DAY TO CURRENT DATE
		dTempdate = DateAdd("d",1,dTempdate)
	Loop

	' RETURN SPECIFIED DATE
	GetNextWeekDay = dTempdate

End Function



'--------------------------------------------------------------------------------------------------
'  FUNCTION GETNEXTORDINALDAYMONTH(IWEEKDAY,IPOS,IMONTH,IYEAR)
'--------------------------------------------------------------------------------------------------
Function GetNextOrdinalDayMonth( ByVal iWeekDay, ByVal ipos, ByVal iMonth, ByVal iYear )
	Dim iCount, dTemp, dReturnValue

	' INITIALIZE DATE VALUES
	dTemp = cdate(iMonth & "/1/" &iYear)
	dReturnValue = dTemp
	iCount = 0

	' PERFORM DATE LOOKUP BASED ON ORDINAL POSITION
	Select Case ipos

		Case 1
			' FIRST OCCURRENCE
			 Do While Not  clng(WeekDay(dTemp)) = clng(iWeekDay) 
				' ADD 1 DAY TO CURRENT DATE
				dTemp = DateAdd("d",1,dTemp)
			 Loop
			 dReturnValue = dTemp

		Case 2
			' SECOND OCCURRENCE
			 Do While iCount < 2
				' FOUND DAY OF WEEK MATCH
				If (clng(WeekDay(dTemp)) = clng(iWeekDay)) Then
					iCount = iCount + 1 ' ADD 1 TO OCCURENCE COUNT
					dReturnValue = dTemp
				End If
				' ADD 1 DAY TO CURRENT DATE
				dTemp = DateAdd("d",1,dTemp)
			 Loop

		Case 3
			' THIRD OCCURRENCE
			 Do While clng(iCount) < clng(3)
				' FOUND DAY OF WEEK MATCH
				If (clng(WeekDay(dTemp)) = clng(iWeekDay)) Then
					iCount = iCount + 1 ' ADD 1 TO OCCURENCE COUNT
					dReturnValue = dTemp
'					dtb_debug "WeekDay match: " & dReturnValue
				End If
				' ADD 1 DAY TO CURRENT DATE
				dTemp = DateAdd("d",1,dTemp)
			 Loop
			 
		Case 4
			' FOURTH OCCURRENCE
			 Do While iCount < 4
				' FOUND DAY OF WEEK MATCH
				If (clng(WeekDay(dTemp)) = clng(iWeekDay)) Then
					iCount = iCount + 1 ' ADD 1 TO OCCURENCE COUNT
					dReturnValue = dTemp
				End If
				' ADD 1 DAY TO CURRENT DATE
				dTemp = DateAdd("d",1,dTemp)
			 Loop
			 
		Case 5
			datNextMonth = dateAdd("m",1,dTemp)

			' LAST OCCURRENCE
			 Do While iCount < 5 AND (cdate(dtemp) < cdate(datNextMonth))
				' FOUND DAY OF WEEK MATCH
				If (clng(WeekDay(dTemp)) = clng(iWeekDay)) Then
					iCount = iCount + 1 ' ADD 1 TO OCCURENCE COUNT
					dReturnValue = dTemp
				End If
				' ADD 1 DAY TO CURRENT DATE
				dTemp = DateAdd("d",1,dTemp)
			 Loop
			
	End Select

	' RETURN DATE VALUE
	GetNextOrdinalDayMonth = dReturnValue
'	dtb_debug "GetNextOrdinalDayMonth returning: " & dReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' SUB PROCESSRECURRENT()
'--------------------------------------------------------------------------------------------------
Sub ProcessRecurrent()

	' RECURRENCE TEST
	sStartDate = request("checkindate")
	sEndDate = request("recurrentenddate")
	
	' WEEKLY VALUES
	wfrequency = request("wfrequencynumber")
	wdayoftheweek = request("wdayofweek")

	' MONTHLY VALUES
	mseries = request("mseries")
	mfrequency = request("mfrequencynumber")
	mdayoftheweek = request("mdayofweek")

	' YEARLY TEST VALUES
	yseries = request("yseries")
	ydayoftheweek = request("ydayofweek")
	ymonth = request("ymonth")


	' CALL ROUTINE TO CREATE SERIES
	Select Case request("wrecurrenttimepart")
		
		Case "daily"
			' CALL DAILY DATES
			SetDailyDates sStartDate, sEndDate
		
		Case "weekly"
			' CALL WEEKLY DATES
			SetWeeklyDates sStartDate, sEndDate, wfrequency, wdayoftheweek

		Case "monthly"
			' CALL MONTHLY DATES
			SetMonthlyDates sStartDate, sEndDate, mseries, mfrequency, mdayoftheweek 

		Case "yearly"
			' CALL YEARLY DATES
			SetYearlyDates sStartDate, sEndDate, yseries, ydayoftheweek, ymonth

	End Select

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION GETTIMEPARTSTATUS(IFACILITYID,ITIMEPARTID,STIMERANGE)
'--------------------------------------------------------------------------------------------------
Function GetTimePartStatus( ByVal iFacilityid, ByVal itimepartid, ByVal datDate )
	Dim sSql, oRs, sReturnValue

	sReturnValue = "OPEN"

	' GET STATUS OF THIS TIME PART FROM SQL IF AVAILABLE
	sSql = "SELECT DISTINCT status FROM egov_facilityschedule WHERE facilityid = " & iFacilityId & " AND facilitytimepartid = " & itimepartid & " AND checkindate = '" & datDate & "'"
	'response.write sSql & "<br><br>"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	' IF RESERVATION HAS BEEN MADE FOR THIS TIME PART GET ITS STATUS
	If Not oRs.EOF Then
		sReturnValue = oRs("status")
	End If

	oRs.Close
	Set oRs = Nothing
	
	' RETURN STATUS
	GetTimePartStatus = sReturnValue

End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION STOREFACILITYINFORMATIONRECURRENT(IORGID,CHECKINDATE,CHECKOUTDATE)
'------------------------------------------------------------------------------------------------------------
Function StoreFacilityInformationRecurrent( ByVal iOrgID, ByVal checkindate, ByVal checkoutdate, ByVal itimepartid )
	Dim iReturnValue, oCmd

	iReturnValue = 0

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "StoreFacilityInformationRecurrent"
		.CommandType = 4

		' RESERVATION INFORMATION
		.Parameters.Append oCmd.CreateParameter("checkintime", 200, 1, 50, request("checkintime"))
		.Parameters.Append oCmd.CreateParameter("checkindate", 200, 1, 50, checkindate)
		.Parameters.Append oCmd.CreateParameter("checkouttime", 200, 1, 50, request("checkouttime"))
		.Parameters.Append oCmd.CreateParameter("checkoutdate", 200, 1, 50, checkoutdate)
		.Parameters.Append oCmd.CreateParameter("lesseeid", 3, 1, 4, request("lesseeid"))
		.Parameters.Append oCmd.CreateParameter("facilitytimepartid", 3, 1, 4, itimepartid)
		.Parameters.Append oCmd.CreateParameter("facilityid", 3, 1, 4, request("facilityid"))
		.Parameters.Append oCmd.CreateParameter("orgid", 3, 1, 4, iOrgID)
		.Parameters.Append oCmd.CreateParameter("amount", 3, 1, 4, request("amount"))
		.Parameters.Append oCmd.CreateParameter("irecurrentid", 3, 1, 4, request("irecurrentid"))
		.Parameters.Append oCmd.CreateParameter("facilitypaymentid", 3, 2, 4)

		' CALL TO STORE VALUE INFORMATION
		.Execute

		iFacilityPaymentID = .Parameters("facilitypaymentid")

		If iFacilityPaymentID <> "" Then
			iReturnValue = iFacilityPaymentID
		End If

		' STORE FIELD VALUES
		StoreFacilityFieldValues( iFacilityPaymentID )

	End With

	Set oCmd = Nothing

	StoreFacilityInformationRecurrent = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' SUB SETDAILYDATES(SSTARTDATE,SENDDATE)
'--------------------------------------------------------------------------------------------------
Sub SetDailyDates( ByVal sStartDate, ByVal sEndDate )
	Dim sSql, oRs

	' CHECK FOR VALID ENDDATE
	If sEndDate = "" Then
		response.write  UCase("<font color=""red""><b>No end date specified. Press back and add end date.</b></font><br>")
		Exit Sub
	End If

	dTemp = sStartDate ' SET FLOATING DATE

	' LOOP UNTIL END DATE
	Do While CDate(dTemp) <= CDate(sEndDate)
		
		' FOR EACH TIME PART FOR THAT DAY SET TO CLOSED
		sSql = "SELECT facilityid, rateid, facilitytimepartid, beginhour, beginampm, endhour, endampm, weekday,description,rate "
		sSql = sSql & "FROM egov_facilitytimepart where facilityid = " & request("facilityid") & " and weekday = " & Weekday( dTemp )
		sSql = sSql & " ORDER BY weekday, description,beginampm, beginhour"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		' LOOP THRU ALL TIME PARTS FOR THE DAY
		If Not oRs.EOF Then

			Do While Not oRs.EOF 

					' GET CURRENT STATUS OF TIMEPART AND DATE
					sStatus = GetTimePartStatus(request("facilityid"),oRs("facilitytimepartid"),dTemp)
					
					' NOT RESERVED OR ONHOLD THEN MARK AVAILABLE
					If sStatus = "OPEN" or sStatus = "CANCELLED" Then

						' IF SETTING STATUS = CLOSED UPDATE ALL TIME PARTS AS BEING CLOSED
						If request("reservationstatus") = "CLOSED" Then

							' STORE RESERVATION INFORMATION
							iFacilityPaymentID = StoreFacilityInformationRecurrent( session("orgid"), dTemp, dTemp, oRs("facilitytimepartid") )
							
							' UPDATE DATABASE
							UpdateFacilityPayment iFacilityPaymentID, "EGOVLINK_ADMIN", "EGOVLINK_ADMIN", "APPROVED", "APPROVED", request("paymenttype"), request("paymentlocation")
				
						End If
					
						' IF SETTING STATUS <> CLOSED UPDATE ONLY THE CORRESPONDING TIME PART ON THE NEXT DAY
						If (request("reservationstatus") <> "CLOSED") AND (request("checkintime") = oRs("beginhour")&":"&oRs("beginampm") AND request("checkouttime") = oRs("endhour")&":"&oRs("endampm") ) Then

							' STORE RESERVATION INFORMATION
							iFacilityPaymentID = StoreFacilityInformationRecurrent( session("orgid"), dTemp, dTemp, oRs("facilitytimepartid") )
							
							' UPDATE DATABASE
							UpdateFacilityPayment iFacilityPaymentID, "EGOVLINK_ADMIN", "EGOVLINK_ADMIN", "APPROVED", "APPROVED", request("paymenttype"), request("paymentlocation")
				
						End If
					Else
						
						' DISPLAY NOT AVAILABLE
						response.write ucase("<font color=""red""><b>" & dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  " NOT AVAILABLE!</b></font><br>")
					
					End If
				
					oRs.MoveNext			
			Loop

		End If

		oRs.Close
		Set oRs = Nothing

		' ADD SPECIFIED WEEK FREQUENCY
		dTemp = DateAdd("d",1,dTemp)
	Loop

End Sub



%>


