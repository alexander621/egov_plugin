<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: FACILITY_RESERVATION_PAGE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/13/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   02/13/06   JOHN STULLENBERGER - INITIAL VERSION
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

	<script src="../scripts/jquery-1.9.1.min.js"></script>
	
	<script>
	  <!--//

		function openwindow( sURL ) {
			window.open( sURL );
		}

		function ContinuePurchase() {
			document.frmReservation.target = '_self';

			// CREDIT CARD PROCESSING
			if ($("#paymenttype").val() == '1') {
				// SUBMIT FORM-NO REAL TIME PROCESSING
				//document.frmReservation.submit();
				// CHANGE FORM'S ACTION URL AND SUBMIT
				document.frmReservation.action = 'facility_cashcheck_receipt.asp';
				document.frmReservation.submit();
			}

			// CASH PROCESSING
			if ($("#paymenttype").val() == '2'){
				// CHANGE FORM'S ACTION URL AND SUBMIT
				document.frmReservation.action = 'facility_cashcheck_receipt.asp';
				document.frmReservation.submit();
			}

			// CHECK PROCESSING
			if ($("#paymenttype").val() == '3'){
				// CHANGE FORM'S ACTION URL AND SUBMIT
				document.frmReservation.action = 'facility_cashcheck_receipt.asp';
				document.frmReservation.submit();
			}
		}

		function view_waivers(surl) {
			// CHANGE FORM'S ACTION URL AND SUBMIT
			document.frmReservation.action = surl;
			document.frmReservation.target = '_NEW';
			document.frmReservation.submit();
		}

	//-->
	</script>

</head>
<body>
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN: PAGE CONTENT-->
<div style="padding:20px;">
<form name="frmReservation" action="facility_cashcheck_receipt.asp" method="post">

	<!--BEGIN: REQUIRED INFORMATION-->
	<input type="hidden" name="ITEM_NAME" value="FACILITY RESERVATION" />
	<input type="hidden" name="ITEM_NUMBER" value="F300" />
	<input type="hidden" name="iPAYMENT_MODULE" value="3" />
	<input type="hidden" name="LODGENAME" value="<%=request("LodgeName")%>" />
	<input type="hidden" name="amount" value="<%=request("amounttotal")%>" />
	<input type="hidden" name="checkintime" value="<%=request("checkintime")%>" />
	<input type="hidden" name="checkindate" value="<%=request("checkindate")%>" />
	<input type="hidden" name="checkouttime" value="<%=request("checkouttime")%>" />
	<input type="hidden" name="checkoutdate" value="<%=request("checkoutdate")%>" />
	<input type="hidden" name="lesseeid" value="<%=request("userid")%>" />
	<input type="hidden" name="timepartid" value="<%=request("selTimePartID")%>" />
	<input type="hidden" name="facilityid" value="<%=request("L")%>" />
	<input type="hidden" name="iuserid" value="<%=request("userid")%>" />
	<input type="hidden" name="internalnote" value="<%=request("internalnote")%>" />
	<input type="hidden" name="backlink" value="<%=request("backlink")%>" />

	<%
	' GET ALL CUSTOM INFORMATION FROM REQUEST OBJECT
	For Each oField IN Request.Form
		If Left(oField,7) = "custom_" Then
			response.write "<input type=""hidden"" name=""" & oField & """ value=""" & request(oField) & """ />"
		End If
	Next
	%>

	<!--RECURRENT INFORMATION-->
	<input type="hidden" name="recurrentenddate" value="<%=request("recurrentenddate")%>" />
	<input type="hidden" name="wfrequencynumber" value="<%=request("wfrequencynumber")%>" />
	<input type="hidden" name="wdayofweek" value="<%=request("wdayofweek")%>" />
	<input type="hidden" name="mseries" value="<%=request("mseries")%>" />
	<input type="hidden" name="mfrequencynumber" value="<%=request("mfrequencynumber")%>" />
	<input type="hidden" name="mdayofweek" value="<%=request("mdayofweek")%>" />
	<input type="hidden" name="yseries" value="<%=request("mseries")%>" />
	<input type="hidden" name="ydayofweek" value="<%=request("ydayofweek")%>" />
	<input type="hidden" name="ymonth" value="<%=request("ymonth")%>" />
	<input type="hidden" name="wrecurrenttimepart" value="<%=request("wrecurrenttimepart")%>" />
	<input type="hidden" name="isrecursive" value="<%=request("isrecursive")%>" />
	<!--END: REQUIRED INFORMATION-->


<!--BEGIN: RESERVATION DETAILS-->
<div class="reserveformtitle">Reservation Details</div>
<div class="reserveforminputarea">
<%
' STANDARD DETAILS
GetReservationDetails

' RECURRENCE INFORMATION
If request("isrecursive") = "on" Then
	ProcessRecurrent()
End If
 %>
</div>
<!--END: RESERVATION DETAILS-->


<!--BEGIN: TERM/CONDITIONS AND WAIVER DOWNLOAD-->
<div class="reserveformtitle">Terms/Conditions and Waiver Downloads</div>
<div class="reserveforminputarea">


<b>Download the following form(s) to print, sign, and bring with you when picking up the key:</b>
<p><% ListWaivers %></p>
</div>
<!--END: TERM/CONDITIONS AND WAIVER DOWNLOAD-->


<!--BEGIN: TOTAL COSTS-->
<div class="reserveformtitle">Totals Costs</div>
<div class="reserveforminputarea"><% GetPaymentDetails %></div>
<!--END: TOTAL COSTS-->


<!--BEGIN: MAKE PURCHASE-->
<div class="reserveformtitle">Make Purchase</div>
<div class="reserveforminputarea">
	<table border="0" cellpadding="5" cellspacing="0" >
			<tr><td> 
			<b>Payment Type: </b>
			<select id="paymenttype" name="paymenttype">
				<option value="1">CreditCard</option>
				<option value="2">Check</option>
				<option value="3">Cash</option>
			</select>
			</td>
			<td>
				<b>Payment Location: </b>
				<select id="paymentlocation" name="paymentlocation">
					<option value="1">Walk In</option>
					<option value="2">Phone Call</option>
				</select>
			</td>
			<td>
				<b>Status: </b>
				<select id="reservationstatus" name="reservationstatus">
					<option value="RESERVED">RESERVED</option>
					<option value="ONHOLD">ON HOLD</option>
					<option value="CLOSED">CLOSED</option>
					<option value="CALL">CALL TO RESERVE</option>
					<option value="HOLIDAY">HOLIDAY BLOCK</option>
				</select>
			</td>
			<td width="200">
				<input type="button" class="facilitybutton" name="continue" value="Continue with Purchase" onclick="ContinuePurchase();" />
			</td></tr>
	</table>
</div>
<!--END: MAKE PURCHASE-->


</form>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' DISPLAYTERMS(IFACILITYID) -- this is never called
'--------------------------------------------------------------------------------------------------
Sub DisplayTerms( ByVal iFacilityId )

		' SELECT ALL TERMS FOR FACILITY
		sSql = "Select * FROM egov_recreation_Terms where facilityid = '" & iFacilityId & "' order by displayorder"
		Set oTerms = Server.CreateObject("ADODB.Recordset")
		oTerms.Open sSql, Application("DSN"), 3, 1

		' IF TERMS FOUND DISPLAY
		If not oTerms.EOF Then
		
			' LOOP THRU TERMS DISPLAYING THEM
			Do While NOT oTerms.EOF
				response.write "<P><input type=checkbox name=term> " & oTerms("termdescription") & "</p>"
				oTerms.MoveNext
			Loop
		
		Else

			' DO NOTHING - NO TERMS FOUND
			
		End If

End Sub 


'--------------------------------------------------------------------------------------------------
' GetReservationDetails
'--------------------------------------------------------------------------------------------------
Sub GetReservationDetails()

	' DISPLAY CHECK IN/OUT DATETIMES AND FACILITY NAME
	response.write "<br><b>Facility Name: " & request("LodgeName") & "</b><BR><BR>"
	
	'response.write "<b>Check In Time</b>: " & request("checkintime") & "<BR>"
	'response.write "<b>Check In Date</b>: " & request("checkindate") & "<BR>"
	'response.write "<b>Check Out Time</b>: " & request("checkouttime") & "<BR>"
	'response.write "<b>Check Out Date</b>: " & request("checkoutdate") & "<BR>"
	
	response.write "<b>Check In:</b> " & request("checkindate") & " at " & request("checkintime") & "<BR>"
	response.write "<b>Check Out:</b> " & request("checkoutdate") & " at " & request("checkouttime") & "<BR>"
	
	If request("isrecursive") = "on" Then 
		response.write "<br><b>Recurs</b>: " & request("wrecurrenttimepart") & "<br />"
		response.write "<b>Ending By</b>: " & request("recurrentenddate") & "<br />"
	End If 
	
	response.write "<br><b>Internal Note:</b><br>" & request("internalnote") & "<BR>"

End Sub 


'--------------------------------------------------------------------------------------------------
' GetPaymentDetails
'--------------------------------------------------------------------------------------------------
Sub GetPaymentDetails()

	response.write vbcrlf & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" id=""paymentdisplay"">"
	
	' DISPLAY COST OF FACILITY
	response.write vbcrlf & "<tr><td class=""amountlabel"" nowrap=""nowrap"">Facility Reservation Cost:</td><td class=""amountdisplay"">" & FormatNumber(request("reservetotal"),2,,,0) & "</td></tr>"
	
	' IF KEYDEPOSIT CHECKED DISPLAY KEY DEPOSIT CHARGE
	If LCase(request("keydeposit")) = "on" Then
		response.write vbcrlf & "<tr><td class=""amountlabel"" nowrap=""nowrap"" nowrap=""nowrap"">Key Deposit Charge:</td><td class=""amountdisplay"">" & FormatNumber(request("reservetotal"),2,,,0) &  "</td></tr>"
	End If
	
	' DISPLAY TOTAL COST OF RESERVATION
	response.write vbcrlf & "<tr><td class=""amountlabel"" nowrap=""nowrap"">Total:</td><td class=""amountdisplay"">" & FormatNumber(request("amounttotal"),2,,,0) & "</td></tr>"
	response.write vbcrlf & "</table>"
End Sub 


'--------------------------------------------------------------------------------------------------
' ListWaivers
'--------------------------------------------------------------------------------------------------
Sub ListWaivers()
	
	' SET INITIAL VALUES
	iWaiverCount = 0
	sList = ""

	' LOOP THRU EACH ITEM IN THE FORM COLLECTION
	For Each oField IN Request.Form
		' IF ITEM IS A WAIVER THEN ADD TO LIST
		If Left(oField,11) = "chkwaivers_" Then
		    iWaiverCount = iWaiverCount + 1 ' INCREASE WAIVER COUNT
			
			' IF ITEM IS NOT EMPTY ADD TO LIST
			If sList = "" Then
				' ADD FIRST WAIVER TO LIST
				sList = sList & request(oField)
			Else
				' SEPARATE SUBSEQUENT WAIVERS WITH X WHEN ADDING TO LIST
				sList = sList & "X" & request(oField)
			End If
			
		End If
	Next

	' IF WAIVER COUNT GREATER THAN ZERO THEN THERE WERE WAIVERS FOUND SO DISPLAY
	If iWaiverCount <> 0 Then
		' DISPLAY WAIVER LIST WITH LINK TO OPEN THE WAIVERS AS PDF IN NEW WINDOW
		response.write "<P><input class=""reserveformbutton"" style=""width:300px;text-align:center;"" type=button value=""Click to download PDF forms"" onClick=""view_waivers('display_waiver.aspx?MASK=" & sList & "');"" ></P>"
	Else
		' NO WAIVERS FOUND - INFORM USER
		response.write "<P> No waivers required. </p>"
	End If
	
End Sub


'--------------------------------------------------------------------------------------------------
' SetDailyDates sStartDate, sEndDate
'--------------------------------------------------------------------------------------------------
Sub SetDailyDates( ByVal sStartDate, ByVal sEndDate )


	' CHECK FOR ENDDATE
	If sEndDate = "" Then
		response.write  UCase("<font color=""red""><b>No end date specified. Press back and add end date.</b></font><br>")
		Exit Sub
	End If

	dTemp = sStartDate

	' LOOP UNTIL END DATE
	Do While CDate(dTemp) <= CDate(sEndDate)
		
		
		' GET CURRENT STATUS OF TIMEPART AND DATE
		sStatus = GetTimePartStatus(request("L"),request("selTimePartID"),dTemp)
		
		' NOT RESERVED OR ONHOLD THEN MARK AVAILABLE
		If sStatus = "OPEN" or sStatus = "CANCELLED" Then
			' DISPLAY AVAILABLE
			'DEBUG CODE: response.write ucase("<font color=GREEN><B>" & dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  " IS RESERVABLE!</b></font><BR>")
		Else
			' DISPLAY NOT AVAILABLE
			response.write ucase("<font color=red><B>" & dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  " NOT AVAILABLE!</b></font><BR>")
		End If

		' ADD SPECIFIED WEEK FREQUENCY
		dTemp = DateAdd("d",1,dTemp)
	Loop

End Sub


'--------------------------------------------------------------------------------------------------
'  SetWeeklyDates(sStartDate,sEndDate,wfrequency,wdayoftheweek
'--------------------------------------------------------------------------------------------------
Sub SetWeeklyDates( ByVal sStartDate, ByVal sEndDate, ByVal wfrequency, ByVal wdayoftheweek )

	' GET NEXT DATE OF SPECIFIED DAY OF WEEK (COULD BE START DATE BUT NOT ALWAYS)
	dTemp = GetNextWeekDay( wdayoftheweek, sStartDate )

	' LOOP UNTIL END DATE
	Do While CDate(dTemp) <= CDate(sEndDate)
		' WRITE DAY OF WEEK
		
		' GET CURRENT STATUS OF TIMEPART AND DATE
		sStatus = GetTimePartStatus(request("L"),request("selTimePartID"),dTemp)
		
		' NOT RESERVED OR ONHOLD THEN MARK AVAILABLE
		If sStatus = "OPEN" or sStatus = "CANCELLED" Then
			' DISPLAY AVAILABLE
			'DEBUG CODE: response.write ucase("<font color=GREEN><B>" & dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  " IS RESERVABLE!</b></font><BR>")
			' INSERT ROW INTO DATABASE
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
Sub SetMonthlyDates( ByVal sStartDate, ByVal sEndDate, ByVal mseries, ByVal mfrequency, ByVal mdayoftheweek )
	' This seems to only check availability of the desired reservation dates

	' GET NEXT DATE OF SPECIFIED DAY OF WEEK (COULD BE START DATE BUT NOT ALWAYS)
	dTemp = GetNextOrdinalDayMonth( mdayoftheweek, mseries, Month(sStartDate), Year(sStartDate) )

	' LOOP UNTIL END DATE
	Do While cdate(dTemp) <= cdate(sEndDate) 

		' IF SERIES IS GREATER THAN OR EQUAL THEN USE VALUE
		If Cdate(dTemp) >= Cdate(sStartDate)  Then

			' GET CURRENT STATUS OF TIMEPART AND DATE
			sStatus = GetTimePartStatus(request("L"),request("selTimePartID"),dTemp)
			
			' NOT RESERVED OR ONHOLD THEN MARK AVAILABLE
			If sStatus = "OPEN" or sStatus = "CANCELLED" Then
				' DISPLAY AVAILABLE
				' DEBUG CODE: response.write dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  "<BR>"
				' INSERT ROW INTO DATABASE
			Else
				' DISPLAY NOT AVAILABLE
				response.write UCase("<font color=""red""><b>" & dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  " NOT AVAILABLE!</b></font><br />")
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
Sub SetYearlyDates( ByVal sStartDate, ByVal sEndDate, ByVal yseries, ByVal ydayoftheweek, ByVal ymonth )

	' GET NEXT DATE OF SPECIFIED DAY OF WEEK (COULD BE START DATE BUT NOT ALWAYS)
	dTemp = GetNextOrdinalDayMonth( ydayoftheweek, yseries, ymonth, Year(sStartDate) )


	' LOOP UNTIL END DATE
	Do While CDate(dTemp) <= CDate(sEndDate)
		
		' IF SERIES IS GREATER THAN OR EQUAL THEN USE VALUE
		If Cdate(dTemp) >= Cdate(sStartDate) Then
		
			' GET CURRENT STATUS OF TIMEPART AND DATE
			sStatus = GetTimePartStatus( request("L"), request("selTimePartID"), dTemp )
			
			' NOT RESERVED OR ONHOLD THEN MARK AVAILABLE
			If sStatus = "OPEN" or sStatus = "CANCELLED" Then
				' DISPLAY AVAILABLE
				' DEBUG CODE: 	response.write dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  "<BR>"
				' INSERT ROW INTO DATABASE
			Else
				' DISPLAY NOT AVAILABLE
				response.write UCase("<font color=""red""><b>" & dTemp & "-" & WeekDayName(WeekDay(dTemp)) &  " NOT AVAILABLE!</b></font><BR>")
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
	Do While Not clng(WeekDay(dTempdate)) = clng(iWeekDay) 
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
	dTemp = CDate(iMonth & "/1/" &iYear)
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

End Function


'--------------------------------------------------------------------------------------------------
' SUB PROCESSRECURRENT()
'--------------------------------------------------------------------------------------------------
Sub ProcessRecurrent()
	Dim sStartDate, sEndDate, wfrequency, wdayoftheweek, mseries, mfrequency, mdayoftheweek, yseries, ydayoftheweek, ymonth

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
			
			' STORE RECURRENCE INFORMATION
			storerecurrent sStartDate, sEndDate, 0, 0, 0, 0, 4

		Case "weekly"
			' CALL WEEKLY DATES
			SetWeeklyDates sStartDate, sEndDate, wfrequency, wdayoftheweek

			' STORE RECURRENCE INFORMATION
			storerecurrent sStartDate,sEndDate, wfrequency, 0, wdayoftheweek, 0, 1 
		Case "monthly"
			' CALL MONTHLY DATES
			SetMonthlyDates sStartDate, sEndDate, mseries, mfrequency, mdayoftheweek
			
			' STORE RECURRENCE INFORMATION
			storerecurrent sStartDate, sEndDate, mfrequency, mseries, mdayoftheweek, 0, 2
		Case "yearly"
			' CALL YEARLY DATES
			 SetYearlyDates sStartDate, sEndDate, yseries, ydayoftheweek, ymonth

			' STORE RECURRENCE INFORMATION
			storerecurrent sStartDate, sEndDate, 0, ydayoftheweek, yseries, ymonth, 3 
	End Select

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION GETTIMEPARTSTATUS(IFACILITYID,ITIMEPARTID,STIMERANGE)
'--------------------------------------------------------------------------------------------------
Function GetTimePartStatus( ByVal iFacilityid, ByVal itimepartid, ByVal datDate )
	Dim sSql, oRs, sReturnValue

	sReturnValue = "OPEN"

	' GET STATUS OF THIS TIME PART FROM SQL IF AVAILABLE
	sSql = "SELECT DISTINCT status FROM egov_facilityschedule WHERE facilityid = " & iFacilityId & " AND facilitytimepartid = " & itimepartid & " AND checkindate='" & datDate & "'"

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


'--------------------------------------------------------------------------------------------------
' storeRecurrent sStartDate, sEndDate, ifrequency, iordinal, idayofweek, imonth, idatepart
'--------------------------------------------------------------------------------------------------
Sub storeRecurrent( ByVal sStartDate, ByVal sEndDate, ByVal ifrequency, ByVal iordinal, ByVal idayofweek, ByVal imonth, ByVal idatepart )
	Dim sSql

	iReturnValue = 0

	' BUILD INSERT STATEMENT
	sSql = "INSERT INTO egov_facilityrecurrence ( startdate, enddate, frequency, ordinal, dayofweek, month, datepart ) VALUES ( '"
	sSql = sSql & sStartDate & "','" & sEndDate & "','" & ifrequency & "','" & iordinal & "','" & idayofweek & "','" & imonth & "','" & idatepart & "' )" 	

	iReturnValue = RunInsertStatement( sSql )

	' WRITE HIDDEN FIELD WITH ADDED RECURRENT ROWID 
	response.write "<input type=""hidden"" name=""irecurrentid"" value=""" & iReturnValue & """ />"

End Sub 


%>


