<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->

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
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFacilityScheduleId, oOrganization

Set oOrganization = New classOrganization

iFacilityScheduleId = request("iFacilityScheduleId")
%>

<html>
<head>
	<title><%=oOrganization.GetOrgName()%> E-Gov Facility Rental Details</title>

	<link href="../global.css" rel="stylesheet" type="text/css">
	<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" type="text/css">
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />
</head>

<!--#Include file="../include_top.asp"-->

	<!--BEGIN:  USER REGISTRATION - USER MENU-->
<%	RegisteredUserDisplay( "../" ) %>
	<!--END:  USER REGISTRATION - USER MENU-->

<!--BEGIN: PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		<div id="receiptlinks">
			<img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.go(-1)">&nbsp;Back</a><span id="printbutton"><input type="button" onclick="javascript:window.print();" value="Print" /></span>
		</div>

		<h3><%=Session("sOrgName")%> Facility Rental Details</h3>

		<% DisplayReciept iFacilityScheduleId %>
	</div>
</div>

<%	Set oOrganization = Nothing %>

<!--END: PAGE CONTENT-->

<!--#Include file="../include_bottom.asp"-->  




<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' SUB DISPLAYRECIEPT( iFacilityScheduleId )
'--------------------------------------------------------------------------------------------------
Sub DisplayReciept( iFacilityScheduleId )
Dim sSql, oFacility 

	' Get payment information 
	sSql = "Select amount, checkindate, checkintime, checkoutdate, checkouttime, paymenttype, paymentlocation, "
	sSql = sSql & " P.datecreated, P.facilityid, facilityname, lesseeid from egov_facilityschedule P, egov_facility F"
	sSql = sSql & " where P.facilityid = F.facilityid and P.facilityscheduleid = " & iFacilityScheduleId
	
	Set oFacility = Server.CreateObject("ADODB.Recordset")
	oFacility.Open sSQL, Application("DSN"), 3, 1

	If Not oFacility.EOF Then 

		' Display the user information
		ShowUserInfo oFacility("lesseeid")

		' TRANSACTION RESULT DETAILS
		response.write vbcrlf & "<div class=""purchasereportshadow"">"
		response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
		response.write "<tr><th colspan=""2"" align=""left""><b>Transaction Details</b></th></tr>"
		response.write "<tr><td width=""20%"">Purchase Date:</td><td>" & DateValue(oFacility("datecreated")) & "</td></tr>"
		response.write "<tr><td>Payment Method:</td><td> " & GetPaymentTypeName(oFacility("paymenttype")) & " </td></tr>"
		response.write "<tr><td>Payment Location:</td><td> " & GetPaymentLocationName(oFacility("paymentlocation")) & " </td></tr>"
		response.write "<tr><td>Amount:</td><td> " & FormatCurrency(oFacility("amount"),2) & "</td></tr>"
		response.write "</table></div>"
		
		' PRODUCT INFORMATION
		response.write vbcrlf & "<div class=""purchasereportshadow"">"
		response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
		response.write "<tr><th colspan=""2"" align=""left""><b>Facility Rental Information</b></th></tr>"
		response.write "<tr><td width=""20%"">Order Number:</td><td> " & iFacilityScheduleId & "F3000 </td></tr>"
		response.write "<tr><td>Facility:</td><td>" & oFacility("facilityname") & "</td></tr>"
		response.write "<tr><td valign=""top"">Rental Period: </td><td>" & oFacility("checkindate") & " " & oFacility("checkintime") & " &ndash; " & oFacility("checkoutdate") & " " & oFacility("checkouttime")&  "</td></tr>"
		response.write "</table></div>"
	Else
		response.write "<p>No details are available.</p>"
	End If 

	oFacility.close
	Set oFacility = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION GETPAYMENTTYPENAME(IPAYMENTTYPEID)
'--------------------------------------------------------------------------------------------------
Function GetPaymentTypeName(ipaymenttypeid)

		' SET DEFAULT RETURN VALUE
		sReturnValue = "UNKNOWN"

		' SELECT PAYMENT TYPE NAME
		sSQL = "Select * FROM egov_paymenttypes where paymenttypeid = '" & ipaymenttypeid & "'"
		Set oPaymentType = Server.CreateObject("ADODB.Recordset")
		oPaymentType.Open sSQL, Application("DSN"), 3, 1

		' IF PAYMENT TYPE ID FOUND
		If not oPaymentType.EOF Then
			'SET RETURN VALUE TO PAYMENT NAME
			sReturnValue = oPaymentType("paymenttypename")
		End If
		
		' CLEAN UP OBJECT
		Set oPaymentType = Nothing

		' RETURN VALUE
		GetPaymentTypeName = sReturnValue
			
End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETPAYMENTLOCATIONNAME(IPAYMENTLOCATIONID)
'--------------------------------------------------------------------------------------------------
Function GetPaymentLocationName(ipaymentlocationid)

		' SET DEFAULT RETURN VALUE
		sReturnValue = "UNKNOWN"

		' SELECT PAYMENT TYPE NAME
		sSQL = "Select * FROM egov_paymentlocations where paymentlocationid = '" & ipaymentlocationid & "'"
		Set oPaymentLocation = Server.CreateObject("ADODB.Recordset")
		oPaymentLocation.Open sSQL, Application("DSN"), 3, 1

		' IF PAYMENT TYPE ID FOUND
		If not oPaymentLocation.EOF Then
			'SET RETURN VALUE TO PAYMENT NAME
			sReturnValue = oPaymentLocation("paymentlocationname")
		End If
		
		' CLEAN UP OBJECT
		Set oPaymentLocation = Nothing

		' RETURN VALUE
		GetPaymentLocationName = sReturnValue
			
End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION STOREFACILITYINFORMATION()
'------------------------------------------------------------------------------------------------------------
Function StoreFacilityInformation(iOrgID)
		
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
		.Parameters.Append oCmd.CreateParameter("facilitypaymentid", 3, 2, 4)

		' CALL TO STORE VALUE INFORMATION
		.Execute

		iFacilityPaymentID = .Parameters("facilitypaymentid")
		
		If iFacilityPaymentID <> "" Then
			iReturnValue = iFacilityPaymentID
		End If

		' STORE FIELD VALUES
		StoreFacilityFieldValues(iFacilityPaymentID)

     End With

	 Set oCmd = Nothing

	StoreFacilityInformation = iReturnValue

End Function


'------------------------------------------------------------------------------------------------------------
' SUB STOREFACILITYFIELDVALUES(IFACILITYPAYMENTID)
'------------------------------------------------------------------------------------------------------------
Sub StoreFacilityFieldValues(iFacilityPaymentID)
		
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
' SUB UPDATEFACILITYPAYMENT(IFACILITYPAYMENTID, SAUTHCODE, SPNREF, SRESULT, SRESPMSG)
'------------------------------------------------------------------------------------------------------------
Sub UpdateFacilityPayment(iFacilityPaymentId, sAuthCode, sPNRef, sResult, sRespMsg, itype, ilocation)
	
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
			.Parameters.Append oCmd.CreateParameter("@paymenttype", 200, 1, 4, itype)
			.Parameters.Append oCmd.CreateParameter("@paymentlocation", 200, 1, 4, ilocation)
			.Parameters.Append oCmd.CreateParameter("@status", 200, 1, 50, request("reservationstatus"))
			.Execute
	End With

	Set oCmd = Nothing

End Sub


'------------------------------------------------------------------------------------------------------------
' Sub UserInfo(iUserID)
'------------------------------------------------------------------------------------------------------------
Sub UserInfo(iUserID)

	' SELECT ROW WITH THIS USER'S INFORMATION
	sSQL = "SELECT * FROM egov_users WHERE userid = '" & iUserId & "'"
	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open sSQL, Application("DSN") , 3, 1

	' IF RECORDSET NOT EMPTY THEN DISPLAY USER INFORMATION
	If NOT oUser.EOF Then
		
		response.write "<P><table>"
		response.write "<tr><td colspan=2><b>User Information</b></td></tr>"
		response.write "<tr><td>Name: </td><td>" & oUser("userfname") & " " & oUser("userlname")  &  "</td></tr>"
		response.write "<tr><td>Address: </td><td>" & oUser("useraddress") & "</td></tr>"
		response.write "<tr><td>City: </td><td>" & oUser("usercity") & "</td></tr>"
		response.write "<tr><td>State: </td><td>" & oUser("userstate") & "</td></tr>"
		response.write "<tr><td>Zip: </td><td>" & oUser("userzip") & "</td></tr>"
		response.write "</table></p></blockquote></div></p>"

	Else

		' RECORDSET EMPTY DON'T DISPLAY ANY INFORMATION

	End If

	Set oUser = Nothing
	
End Sub


'--------------------------------------------------------------------------------------------------
'  SUB SETWEEKLYDATES(SSTARTDATE,SENDDATE,WFREQUENCY,WDAYOFTHEWEEK)
'--------------------------------------------------------------------------------------------------
Sub SetWeeklyDates(sStartDate,sEndDate,wfrequency,wdayoftheweek)

	' GET NEXT DATE OF SPECIFIED DAY OF WEEK (COULD BE START DATE BUT NOT ALWAYS)
	dTemp = GetNextWeekDay(wdayoftheweek,sStartDate)

	' LOOP UNTIL END DATE
	Do While cdate(dTemp) <= cdate(sEndDate)
		' WRITE DAY OF WEEK
		
		' GET CURRENT STATUS OF TIMEPART AND DATE
		sStatus = GetTimePartStatus(request("L"),request("selTimePartID"),dTemp)
		
		' NOT RESERVED OR ONHOLD THEN MARK AVAILABLE
		If sStatus = "OPEN" or sStatus = "CANCELLED" Then
			' STORE RESERVATION INFORMATION
			iFacilityPaymentID = StoreFacilityInformationRecurrent(iOrgId,dTemp,dTemp,request("timepartid"))
			
			' UPDATE DATABASE
			UpdateFacilityPayment iFacilityPaymentID, "EGOVLINK_ADMIN","EGOVLINK_ADMIN","APPROVED","APPROVED",request("paymenttype"),request("paymentlocation")
		Else
			' DISPLAY NOT AVAILABLE
			response.write ucase("<font color=red><B>" & dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  " NOT AVAILABLE!</b></font><BR>")
		End If

		' ADD SPECIFIED WEEK FREQUENCY
		dTemp = DateAdd("ww",wfrequency,dTemp)
	Loop

End Sub


'--------------------------------------------------------------------------------------------------
'  SUB SETMONTHLYDATES(SSTARTDATE,SENDDATE,MSERIES,MFREQUENCY,MDAYOFTHEWEEK)
'--------------------------------------------------------------------------------------------------
Sub SetMonthlyDates(sStartDate,sEndDate,mseries,mfrequency,mdayoftheweek)

	' GET NEXT DATE OF SPECIFIED DAY OF WEEK (COULD BE START DATE BUT NOT ALWAYS)
	dTemp = GetNextOrdinalDayMonth(mdayoftheweek,mseries,Month(sStartDate),Year(sStartDate))

	' LOOP UNTIL END DATE
	Do While cdate(dTemp) <= cdate(sEndDate) 

		' IF SERIES IS GREATER THAN OR EQUAL THEN USE VALUE
		If Cdate(dTemp) >= Cdate(sStartDate)  Then

			' GET CURRENT STATUS OF TIMEPART AND DATE
			sStatus = GetTimePartStatus(request("L"),request("selTimePartID"),dTemp)
			
			' NOT RESERVED OR ONHOLD THEN MARK AVAILABLE
			If sStatus = "OPEN" or sStatus = "CANCELLED" Then
				' STORE RESERVATION INFORMATION
				iFacilityPaymentID = StoreFacilityInformationRecurrent(iOrgId,dTemp,dTemp,request("timepartid"))
				
				' UPDATE DATABASE
				UpdateFacilityPayment iFacilityPaymentID, "EGOVLINK_ADMIN","EGOVLINK_ADMIN","APPROVED","APPROVED",request("paymenttype"),request("paymentlocation")
			Else	
				' DISPLAY NOT AVAILABLE
				response.write ucase("<font color=red><B>" & dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  " NOT AVAILABLE!</b></font><BR>")
			End If

		End If

		' ADD SPECIFIED MONTHLY FREQUENCY
		dTemp = DateAdd("m",mfrequency,dTemp)

		' GET NEXT DATE OF SPECIFIED DAY OF WEEK AND SERIES
		dTemp = GetNextOrdinalDayMonth(mdayoftheweek,mseries,Month(dTemp),Year(dTemp))
	Loop

End Sub


'--------------------------------------------------------------------------------------------------
'  SUB SETYEARLYDATES(SSTARTDATE,SENDDATE,YSERIES,YDAYOFTHEWEEK,YMONTH)
'--------------------------------------------------------------------------------------------------
Sub SetYearlyDates(sStartDate,sEndDate,yseries,ydayoftheweek,ymonth)

	' GET NEXT DATE OF SPECIFIED DAY OF WEEK (COULD BE START DATE BUT NOT ALWAYS)
	dTemp = GetNextOrdinalDayMonth(ydayoftheweek,yseries,ymonth,Year(sStartDate))


	' LOOP UNTIL END DATE
	Do While cdate(dTemp) <= cdate(sEndDate)
		
		' IF SERIES IS GREATER THAN OR EQUAL THEN USE VALUE
		If Cdate(dTemp) >= Cdate(sStartDate) Then
		
			' GET CURRENT STATUS OF TIMEPART AND DATE
			sStatus = GetTimePartStatus(request("L"),request("selTimePartID"),dTemp)
			
			' NOT RESERVED OR ONHOLD THEN MARK AVAILABLE
			If sStatus = "OPEN" or sStatus = "CANCELLED" Then
				' STORE RESERVATION INFORMATION
				iFacilityPaymentID = StoreFacilityInformationRecurrent(iOrgId,dTemp,dTemp,request("timepartid"))
				
				' UPDATE DATABASE
				UpdateFacilityPayment iFacilityPaymentID, "EGOVLINK_ADMIN","EGOVLINK_ADMIN","APPROVED","APPROVED",request("paymenttype"),request("paymentlocation")
			Else
				' DISPLAY NOT AVAILABLE
				response.write ucase("<font color=red><B>" & dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  " NOT AVAILABLE!</b></font><BR>")
			End If

		End If

		' ADD SPECIFIED 1 TO YEAR
		dTemp = DateAdd("yyyy",1,dTemp)

		' GET NEXT DATE OF SPECIFIED DAY OF WEEK AND SERIES
		dTemp = GetNextOrdinalDayMonth(ydayoftheweek,yseries,Month(dTemp),Year(dTemp))
	Loop

End Sub


'--------------------------------------------------------------------------------------------------
'  FUNCTION GETNEXTWEEKDAY(IWEEKDAY,DTEMPDATE)
'--------------------------------------------------------------------------------------------------
Function GetNextWeekDay(iWeekDay,dTempdate)
  
  ' LOOP TO THE NEXT SPECIFIED DAY OF THE WEEK
  Do While Not  clng(WeekDay(dTempdate)) = clng(iWeekDay) 
	' ADD 1 DAY TO CURRENT DATE
	dTempdate = DateAdd("d",1,dTempdate)
 Loop

  ' RETURN SPECIFIED DATE
  GetNextWeekDay=dTempdate

End Function



'--------------------------------------------------------------------------------------------------
'  FUNCTION GETNEXTORDINALDAYMONTH(IWEEKDAY,IPOS,IMONTH,IYEAR)
'--------------------------------------------------------------------------------------------------
Function GetNextOrdinalDayMonth(iWeekDay,ipos,iMonth,iYear)

	' INITIALIZE DATE VALUES
	dTemp = cdate(iMonth & "/1/" &iYear)
	dReturnValue = dTemp

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
				' ADD 1 DAY TO CURRENT DATE
				dTemp = DateAdd("d",1,dTemp)

				' FOUND DAY OF WEEK MATCH
				If (clng(WeekDay(dTemp)) = clng(iWeekDay)) Then
					iCount = iCount + 1 ' ADD 1 TO OCCURENCE COUNT
					dReturnValue = dTemp
				End If
			 Loop

			 

		Case 3
			
			' THIRD OCCURRENCE
			 Do While iCount < 3
				' ADD 1 DAY TO CURRENT DATE
				dTemp = DateAdd("d",1,dTemp)

				' FOUND DAY OF WEEK MATCH
				If (clng(WeekDay(dTemp)) = clng(iWeekDay)) Then
					iCount = iCount + 1 ' ADD 1 TO OCCURENCE COUNT
					dReturnValue = dTemp
				End If
			 Loop

			 

		Case 4

			' FOURTH OCCURRENCE
			 Do While iCount < 4
				' ADD 1 DAY TO CURRENT DATE
				dTemp = DateAdd("d",1,dTemp)

				' FOUND DAY OF WEEK MATCH
				If (clng(WeekDay(dTemp)) = clng(iWeekDay)) Then
					iCount = iCount + 1 ' ADD 1 TO OCCURENCE COUNT
					dReturnValue = dTemp
				End If
			 Loop

			 

		Case 5

			datNextMonth = dateAdd("m",1,dTemp)

			' LAST OCCURRENCE
			 Do While iCount < 5 AND (cdate(dtemp) < cdate(datNextMonth))
				' ADD 1 DAY TO CURRENT DATE
				dTemp = DateAdd("d",1,dTemp)
				
				' FOUND DAY OF WEEK MATCH
				If (clng(WeekDay(dTemp)) = clng(iWeekDay)) Then
					iCount = iCount + 1 ' ADD 1 TO OCCURENCE COUNT
					dReturnValue = dTemp
				End If
			 Loop

			

	End Select

	' RETURN DATE VALUE
	GetNextOrdinalDayMonth = dReturnValue

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
			Call SetDailyDates(sStartDate,sEndDate)
		
		Case "weekly"
			' CALL WEEKLY DATES
			Call SetWeeklyDates(sStartDate,sEndDate,wfrequency,wdayoftheweek)

		Case "monthly"
			' CALL MONTHLY DATES
			Call SetMonthlyDates(sStartDate,sEndDate,mseries,mfrequency,mdayoftheweek)

		Case "yearly"
			' CALL YEARLY DATES
			Call SetYearlyDates(sStartDate,sEndDate,yseries,ydayoftheweek,ymonth)

	End Select

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION GETTIMEPARTSTATUS(IFACILITYID,ITIMEPARTID,STIMERANGE)
'--------------------------------------------------------------------------------------------------
Function GetTimePartStatus(iFacilityid,itimepartid,datDate)

	sReturnValue= "OPEN"

	' GET STATUS OF THIS TIME PART FROM SQL IF AVAILABLE
	sSQL = "Select distinct status FROM egov_facilityschedule where facilityid = '" & iFacilityId & "' AND facilitytimepartid='" & itimepartid & "' AND checkindate='" & datDate & "'"
	Set oStatus = Server.CreateObject("ADODB.Recordset")
	oStatus.Open sSQL, Application("DSN"), 3, 1
	
	' IF RESERVATION HAS BEEN MADE FOR THIS TIME PART GET ITS STATUS
	If not oStatus.EOF Then
		sReturnValue = oStatus("status")
	End If

	' CLEAN UP OBJECTS
	Set oStatus = Nothing
	
	' RETURN STATUS
	GetTimePartStatus = sReturnValue

End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION STOREFACILITYINFORMATIONRECURRENT(IORGID,CHECKINDATE,CHECKOUTDATE)
'------------------------------------------------------------------------------------------------------------
Function StoreFacilityInformationRecurrent(iOrgID,checkindate,checkoutdate,itimepartid)
		
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
		StoreFacilityFieldValues(iFacilityPaymentID)

     End With

	 Set oCmd = Nothing

	StoreFacilityInformationRecurrent = iReturnValue
	 

End Function


'--------------------------------------------------------------------------------------------------
' SUB SETDAILYDATES(SSTARTDATE,SENDDATE)
'--------------------------------------------------------------------------------------------------
Sub SetDailyDates(sStartDate,sEndDate)

	' CHECK FOR VALID ENDDATE
	If sEndDate = "" Then
		response.write  ucase("<font color=red><B>No end date specified. Press back and add end date.</b></font><BR>")
		Exit Sub
	End If

	dTemp = sStartDate ' SET FLOATING DATE

	' LOOP UNTIL END DATE
	Do While cdate(dTemp) <= cdate(sEndDate)
		
				' FOR EACH TIME PART FOR THAT DAY SET TO CLOSED
				sSQL = "Select facilityid, rateid, facilitytimepartid, beginhour, beginampm, endhour, endampm, weekday,description,rate,description from egov_facilitytimepart where facilityid = '" & request("facilityid") & "' and weekday = '" & weekday( dTemp ) &"'  order by weekday, description,beginampm, beginhour"

				Set oAvail = Server.CreateObject("ADODB.Recordset")
				oAvail.Open sSQL, Application("DSN"), 3, 1

				' LOOP THRU ALL TIME PARTS FOR THE DAY
				If NOT oAvail.EOF Then

					Do While NOT oAvail.EOF 

							' GET CURRENT STATUS OF TIMEPART AND DATE
							sStatus = GetTimePartStatus(request("facilityid"),oAvail("facilitytimepartid"),dTemp)
							
							' NOT RESERVED OR ONHOLD THEN MARK AVAILABLE
							If sStatus = "OPEN" or sStatus = "CANCELLED" Then

								' IF SETTING STATUS = CLOSED UPDATE ALL TIME PARTS AS BEING CLOSED
								If request("reservationstatus") = "CLOSED" Then

									' STORE RESERVATION INFORMATION
									iFacilityPaymentID = StoreFacilityInformationRecurrent(iOrgId,dTemp,dTemp,oAvail("facilitytimepartid"))
									
									' UPDATE DATABASE
									UpdateFacilityPayment iFacilityPaymentID, "EGOVLINK_ADMIN","EGOVLINK_ADMIN","APPROVED","APPROVED",request("paymenttype"),request("paymentlocation")
						
								End If
							
								response.write (request("reservationstatus") <> "CLOSED") & (request("checkintime") & oAvail("beginhour")&":"&oAvail("beginampm") & request("checkouttime") & oAvail("endhour")&":"&oAvail("endampm") ) & "<BR>"

								' IF SETTING STATUS <> CLOSED UPDATE ONLY THE CORRESPONDING TIME PART ON THE NEXT DAY
								If (request("reservationstatus") <> "CLOSED") AND (request("checkintime") = oAvail("beginhour")&":"&oAvail("beginampm") AND request("checkouttime") = oAvail("endhour")&":"&oAvail("endampm") ) Then

									' STORE RESERVATION INFORMATION
									iFacilityPaymentID = StoreFacilityInformationRecurrent(iOrgId,dTemp,dTemp,oAvail("facilitytimepartid"))
									
									' UPDATE DATABASE
									UpdateFacilityPayment iFacilityPaymentID, "EGOVLINK_ADMIN","EGOVLINK_ADMIN","APPROVED","APPROVED",request("paymenttype"),request("paymentlocation")
						
								End If
							

							Else
								
								' DISPLAY NOT AVAILABLE
								response.write ucase("<font color=red><B>" & dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  " NOT AVAILABLE!</b></font><BR>")
							
							End If
						
							oAvail.MoveNext			
					Loop

				End If

				Set oAvail = Nothing

		' ADD SPECIFIED WEEK FREQUENCY
		dTemp = DateAdd("d",1,dTemp)
	Loop

End Sub


'--------------------------------------------------------------------------------------------------
' Sub ShowUserInfo( iUserId )
'--------------------------------------------------------------------------------------------------
Sub ShowUserInfo( iUserId )
	Dim oCmd, sResidentDesc, sUserType

'	sUserType = GetUserResidentType(iUserid)
	' If they are not one of these (R, N), we have to figure which they are
'	If sUserType <> "R" And sUserType <> "N" Then
		' This leaves E and B - See if they are a resident, also
'		sUserType = GetResidentTypeByAddress(iUserid, Session("OrgID"))
'	End If 

'	sResidentDesc = GetResidentTypeDesc(sUserType)

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserInfoList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iUserId", 3, 1, 4, iUserId)
	    Set oUser = .Execute
	End With

	response.write vbcrlf & "<div class=""purchasereportshadow"">"
	response.write vbcrlf & "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
	response.write vbcrlf & "<tr><th colspan=""2"" align=""left"">Purchaser Contact Information</th></tr>"
	response.write vbcrlf & "<tr><td width=""20%"" valign=""top"">Name:</td><td>" & oUser("userfname") & " " & oUser("userlname")
'	response.write "<br /><strong>" & sResidentDesc & "</strong>"
	response.write "</td></tr>"
	response.write vbcrlf & "<tr><td>Email:</td><td>" & oUser("useremail") & "</td></tr>"
	response.write vbcrlf & "<tr><td>Phone:</td><td>" & FormatPhone(oUser("userhomephone")) & "</td></tr>"
	response.write vbcrlf & "<tr><td valign=""top"">Address:</td><td>" & oUser("useraddress") & "<br />" 
	If oUser("useraddress2") = "" Then 
		response.write oUser("useraddress2") & "<br />" 
	End If 
	response.write oUser("usercity") & ", " & oUser("userstate") & " " & oUser("userzip") & "</td></tr>"
	response.write vbcrlf & "</table></div>"

	oUser.close
	Set oUser = Nothing
	Set oCmd = Nothing
	
End Sub 


'--------------------------------------------------------------------------------------------------
' Function FormatPhone( Number )
'--------------------------------------------------------------------------------------------------
Function FormatPhone( Number )
	If Len(Number) = 10 Then
		FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
	Else
		FormatPhone = Number
	End If
End Function


%>


