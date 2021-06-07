<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: FACILITY_DATE_EDIT.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/22/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   02/22/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1	05/11/07	Steve Loar - Adding in the header changes for the drop down menu.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sFacilityName, sCheckInDate, sCurrentCheckInTime, sCheckOutDate, sCurrentCheckOutTime, blnRecurrence, iUserID  
Dim sStatus, ioccurrenceid, itimepartid, ifacilityid, iReservationID

sLevel = "../" ' Override of value from common.asp

iReservationID = CLng(request("iReservationID"))

' IF POST OPERATION PROCESS UPDATE REQUEST
If request.servervariables("REQUEST_METHOD") = "POST" Then
	UpdateRecord request("iReservationID"), request("checkintime"), request("checkouttime") 
End If


' GET RESERVATION INFORMATION
GetReservationDetails iReservationID, sFacilityName, sCheckInDate, sCurrentCheckInTime, sCheckOutDate, sCurrentCheckOutTime, blnRecurrence, itimepartid, ifacilityid, sStatus, iUserID, ioccurrenceid 

%>

<html lang="en">
	<meta charset="UTF-8">
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="querytool.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="reservation.css" />

	<script src="../scripts/jquery-1.7.2.min.js"></script>

	<script src="../scripts/easyform.js"></script>

	<script>
		

// TIMEPART ARRAY INFORMATION
<% 	BuildJavascriptTimePartArray ifacilityid,Year(sCheckInDate),Month(sCheckInDate),Day(sCheckInDate) %>

	function test( opt, itimepartcount ) {
		// INITIALIZE VARIABLES
		var blnTimesOk = true;
	   
	   // LOOP THRU ALL TIME PARTS
	   for (var intLoop = 0; intLoop < opt.length; intLoop++) {
		
			// IF TIME PART CHECKED COMPARE AGAIN ALL OTHER CHECKED TIME PARTS
			if ((opt[intLoop].checked)) {
				
				// LOOP THRU ALL ALL CHECK TIME PARTS AND COMPARE
				for (var intLoop2 = 0; intLoop2 < opt.length; intLoop2++) {

					// IF CHECKED COMPARE
					if ((opt[intLoop2].checked)) {
						blnTimesOk = CheckDateOverLap(timeparts[intLoop][0],timeparts[intLoop][1],timeparts[intLoop2][0],timeparts[intLoop2][1]);
						if (blnTimesOk != true){break;}
					}
				}
			}
		}

		// IF ERROR DISPLAY MESSAGE TO THE USER
		if(blnTimesOk !=  true){
			
			if (opt[itimepartcount].checked) {
				// CLEAR OFFENDING TIME ENTRY
				opt[itimepartcount].checked = false;
			}
			
			alert('This time would overlap one or more of the currently selected times.  It cannot be added until you review and correct.');
		}
		else
		{

			// UPDATE END DATE IF TIME PART SELECTED EXPANDS TWO DAYS
			// LOOP THRU ALL TIME PARTS
			for (var intLoop3 = 0; intLoop3 < opt.length; intLoop3++) {
			
				var blnFound = false;

				// IF TIME PART CHECKED COMPARE AGAIN ALL OTHER CHECKED TIME PARTS
				// IF CHECKED
				if (opt[intLoop3].checked) {
					// IF SELECTED IS OVERLAP DATE
					if (timeparts[intLoop3][3] == '1'){
						//INCREASE DATE BY ONE DAY
						var datNewDate = new Date(frmAvail.checkoutdate.value);
						datNewDate.setDate(datNewDate.getDate() + 1);
						frmAvail.checkoutdate.value = (datNewDate.getMonth() + 1) + '/' + datNewDate.getDate() + '/' + datNewDate.getYear();
						blnFound = true;
					}
				}

			}

			// CLEAR ANY UNCHECKED TIME PARTS THAT SPAN TWO DAYS
			if (blnFound == false){
				frmAvail.checkoutdate.value = frmAvail.checkindate.value;
			}
		}
	}

	function CheckDateOverLap( datDateOneStart, datDateOneEnd, datDateTwoStart, datDateTwoEnd ) {
		
		var blnReturn = true;

		// DOES DATE TWO START DURING DATE ONE RANGE
		if ((datDateTwoStart > datDateOneStart) && (datDateTwoStart < datDateOneEnd))
			{blnReturn = false;}

		// DOES DATE TWO END DURING DATE ONE RANGE
		if ((datDateTwoEnd > datDateOneStart) && (datDateTwoEnd < datDateOneEnd))
			{blnReturn = false;}

		// DOES DATE ONE START DURING DATE TWO RANGE
		if ((datDateOneStart > datDateTwoStart) && (datDateOneStart < datDateTwoEnd))
			{blnReturn = false;}

		// DOES DATE ONE END DURING DATE TWO RANGE
		if ((datDateOneEnd > datDateTwoStart) && (datDateOneEnd < datDateTwoEnd))
			{blnReturn = false;}

		return blnReturn;
	}
</script>

</head>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div style="padding:20px;">

<p><a href="FACILITY_CALENDAR.ASP?L=<%=request("l")%>" class="linkbutton"><< Calendar</a></p>

<form name="frmAvail" action="facility_date_edit.asp?iReservationID=<%= iReservationID %>" method="post">
	<input type="hidden" name="L" value="<%= ifacilityid %>" />
	<input type="hidden" name="LodgeName" value="<%= sFacilityName %>" />
	<input type="hidden" name="selTimePartID" value="<%= itimepartid %>" />
	<input type="hidden" name="checkindate"  value="<%= sCheckInDate %>" />
	<input type="hidden" name="checkoutdate" value="<%= sCheckOutDate %>" />

<!--BEGIN: FACILITY-->
<div class="reserveformtitle">Facility</div>

	<div class="reserveforminputarea">
		<strong>Facility:</strong> <%=sFacilityName%><br>

		<%= GetFacilityFieldValues( iReservationID ) %>

	</div>
<!--END: FACILITY-->


<!--SELECT DATES-->
<div class="reserveformtitle">Select Date/Time</div>
<div class="reserveforminputarea">
	<p><font class="reserveforminstructions">Instructions: Select the Check-In and Check-Out times for your reservation.</font></p>

	<!--DRAW AVAILABILITY-->
	<p><% DrawAvailability ifacilityid,itimepartid,Year(sCheckInDate),Month(sCheckInDate),Day(sCheckInDate) %></p>


<!--DRAW DATE/TIME SELECTION-->
<table >
	<tr><td class="reservationformlabel">Exact Arrival Time:</td><td>
	<% sCheckInTime = GetCheckInTime( itimepartid )%>
	<% sCheckOutTime = GetCheckOutTime( itimepartid )%>

	<select name="checkintime">
	<%
		' 1 AM TO 11:30 AM
		blnIsAvailable = False
		blnFoundEndTime = False

		
		If Trim(sCheckOutTime) = "1:00:AM" Then
			sRealTime = "1:00:AM"
			sCheckOutTime = "12:30:AM"
		End If
		
		' BUILD TIME STRING
		For i=1 to 11

			sTime = i  & ":00:AM"
			' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
			If Trim(sCheckInTime) = Trim(sTime) Then
				' TURN ON WRITING AVAILABLE TIMES
				blnIsAvailable = True
			End If
			
			If Trim(sCurrentCheckInTime) = Trim(sTime) Then
				sSelected = " selected=""selected"""
			Else
				sSelected = ""
			End If
			
			' DISPLAY ONSCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & " value=""" & sTime & """>" & sTime & "</option>"
			End If 

			
			' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
			If blnFoundEndTime = True Then
				blnIsAvailable = False
			End If


			sTime = i  & ":30:AM"
			' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
			If Trim(sCheckInTime) = Trim(sTime) Then
				' TURN ON WRITING AVAILABLE TIMES
				blnIsAvailable = True
			End If

			If Trim(sCurrentCheckInTime) = Trim(sTime) Then
				sSelected = " selected=""selected"""
			Else
				sSelected = ""
			End If
			
			' DISPLAY ONSCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & " value=""" & sTime & """>" & sTime & "</option>"
			End If 

			' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
			If blnFoundEndTime = True Then
				blnIsAvailable = False
			End If
		Next
		

		' NOON 
		' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
		If Trim(sCheckInTime) = Trim("12:00:PM") Then
			' TURN ON WRITING AVAILABLE TIMES
			blnIsAvailable = TRUE
		End If

		' IS THIS THE SELECTED TIME?
		If Trim(sCurrentCheckInTime) = Trim("12:00:PM") Then
				sSelected = " selected=""selected"""
			Else
				sSelected = ""
		End If

		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & " value=""12:00:PM"">12:00:PM</option>"
		End If

		' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
		If blnFoundEndTime = True Then
			blnIsAvailable = False
		End If

		' 12:30 PM
		' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
		If Trim(sCheckInTime) = Trim("12:30:PM") Then
			' TURN ON WRITING AVAILABLE TIMES
			blnIsAvailable = TRUE
		End If
		
		' IS THIS THE SELECTED TIME?
		If Trim(sCurrentCheckInTime) = Trim("12:30:PM") Then
			sSelected = " selected=""selected"""
			blnIsAvailable = TRUE
		Else
			sSelected = ""
		End If

		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & " value=""12:30:PM"">12:30:PM</option>"
		End If
		
		' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
		If blnFoundEndTime = True Then
			blnIsAvailable = False
		End If
		

		' 1 PM TO 11:30 PM
		For i= 1 to 11
			
			sTime = i  & ":00:PM"
			' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
			If Trim(sCheckInTime) = Trim(sTime) Then
				' TURN ON WRITING AVAILABLE TIMES
				blnIsAvailable = TRUE
			End If
		
			If Trim(sCurrentCheckInTime) = Trim(sTime) Then
				sSelected = " selected=""selected"""
			Else
				sSelected = ""
			End If

			' DISPLAY ON SCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & " value=""" & sTime & """>" & sTime & "</option>"
			End If
			
			' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
			If blnFoundEndTime = True Then
				blnIsAvailable = False
			End If
				
			

			sTime = i  & ":30:PM"
			' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
			If Trim(sCheckInTime) = Trim(sTime) Then
				' TURN ON WRITING AVAILABLE TIMES
				blnIsAvailable = True
			End If

			If Trim(sCurrentCheckInTime) = Trim(sTime) Then
				sSelected = " selected=""selected"""
			Else
				sSelected = ""
			End If

			' DISPLAY ON SCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & " value=""" & sTime & """>" & sTime & "</option>"
			End If

			' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
			If blnFoundEndTime = True Then
				blnIsAvailable = False
			End If

		Next
		
		' MIDNIGHT 
		
		' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
		If Trim(sCheckInTime) = Trim("12:00:AM") Then
			' TURN ON WRITING AVAILABLE TIMES
			blnIsAvailable = True
		End If
		
		' IS THIS THE SELECTED TIME?
		If Trim(sCurrentCheckInTime) = Trim("12:00:AM") Then
				sSelected = " selected=""selected"""
				blnIsAvailable = TRUE
			Else
				sSelected = ""
		End If

		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & " value=""12:00:AM"">12:00:AM</option>"
		End If

		' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
		If blnFoundEndTime = True Then
			blnIsAvailable = False
		End If

		' 12:30 AM

		' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
		If Trim(sCheckInTime) = Trim("12:30:AM") Then
			' TURN ON WRITING AVAILABLE TIMES
			blnIsAvailable = True
		End If

		' IS THIS THE SELECTED TIME
		If Trim(sCurrentCheckInTime) = Trim("12:30:AM") Then
			sSelected = " selected=""selected"""
			blnIsAvailable = TRUE
		Else
			sSelected = ""
		End If

		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & " value=""12:30:AM"">12:30:AM</option>"
		End If

		' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
		If blnFoundEndTime = True Then
			blnIsAvailable = False
		End If

		' IS THIS THE SELECTED TIME
		If Trim(sCurrentCheckInTime) = Trim("1:00:AM") Then
			sSelected = " selected=""selected"""
		Else
			sSelected = ""
		End If

		If sRealTime = "1:00:AM" Then
			response.write "<option" & sSelected & " value=""1:00:AM"">1:00:AM</option>"
		End If
		
		
		%>
	</select>

	</td></tr>

	<tr><td class="reservationformlabel">Exact Departure Time:</td><td>
	<select name="checkouttime">
	<%
		' 1 AM TO 11:30 AM
		blnIsAvailable = False
		blnFoundEndTime = False

		
		If Trim(sCheckOutTime) = "1:00:AM" Then
			sRealTime = "1:00:AM"
			sCheckOutTime = "12:30:AM"
		End If
		
		' BUILD TIME STRING
		For i=1 to 11

			sTime = i  & ":00:AM"
			' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
			If Trim(sCheckInTime) = Trim(sTime) Then
				' TURN ON WRITING AVAILABLE TIMES
				blnIsAvailable = True
			End If
			
			If Trim(sCurrentCheckOutTime) = Trim(sTime) Then
				sSelected = " selected=""selected"""
			Else
				sSelected = ""
			End If
			
			' DISPLAY ONSCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & " value=""" & sTime & """>" & sTime & "</option>"
			End If 

			
			' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
			If blnFoundEndTime = True Then
				blnIsAvailable = False
			End If


			sTime = i  & ":30:AM"
			' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
			If Trim(sCheckInTime) = Trim(sTime) Then
				' TURN ON WRITING AVAILABLE TIMES
				blnIsAvailable = True
			End If

			If Trim(sCurrentCheckOutTime) = Trim(sTime) Then
				sSelected = " selected=""selected"""
			Else
				sSelected = ""
			End If
			
			' DISPLAY ONSCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & " value=""" & sTime & """>" & sTime & "</option>"
			End If 

			' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
			If blnFoundEndTime = True Then
				blnIsAvailable = False
			End If
		Next
		

		' NOON 
		' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
		If Trim(sCheckInTime) = Trim("12:00:PM") Then
			' TURN ON WRITING AVAILABLE TIMES
			blnIsAvailable = TRUE
		End If

		' IS THIS THE SELECTED TIME?
		If Trim(sCurrentCheckOutTime) = Trim("12:00:PM") Then
				sSelected = " selected=""selected"""
			Else
				sSelected = ""
		End If

		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & " value=""12:00:PM"">12:00:PM</option>"
		End If

		' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
		If blnFoundEndTime = True Then
			blnIsAvailable = False
		End If

		' 12:30 PM
		' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
		If Trim(sCheckInTime) = Trim("12:30:PM") Then
			' TURN ON WRITING AVAILABLE TIMES
			blnIsAvailable = TRUE
		End If
		
		' IS THIS THE SELECTED TIME?
		If Trim(sCurrentCheckOutTime) = Trim("12:30:PM") Then
			sSelected = " selected=""selected"""
			blnIsAvailable = TRUE
		Else
			sSelected = ""
		End If

		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & " value=""12:30:PM"">12:30:PM" & "</option>"
		End If
		
		' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
		If blnFoundEndTime = True Then
			blnIsAvailable = False
		End If
		

		' 1 PM TO 11:30 PM
		For i= 1 to 11
			
			sTime = i  & ":00:PM"
			' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
			If Trim(sCheckInTime) = Trim(sTime) Then
				' TURN ON WRITING AVAILABLE TIMES
				blnIsAvailable = TRUE
			End If
		
			If Trim(sCurrentCheckOutTime) = Trim(sTime) Then
				sSelected = " selected=""selected"""
			Else
				sSelected = ""
			End If

			' DISPLAY ON SCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & " value=""" & sTime & """>" & sTime & "</option>"
			End If
			
			' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
			If blnFoundEndTime = True Then
				blnIsAvailable = False
			End If
				
			

			sTime = i  & ":30:PM"
			' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
			If Trim(sCheckInTime) = Trim(sTime) Then
				' TURN ON WRITING AVAILABLE TIMES
				blnIsAvailable = True
			End If

			If Trim(sCurrentCheckOutTime) = Trim(sTime) Then
				sSelected = " selected=""selected"""

			Else
				sSelected = ""
			End If

			' DISPLAY ON SCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & " value=""" & sTime & """>" & sTime & "</option>"
			End If

			' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
			If blnFoundEndTime = True Then
				blnIsAvailable = False
			End If

		Next

		
		' MIDNIGHT 
		
		' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
		If Trim(sCheckInTime) = Trim("12:00:AM") Then
			' TURN ON WRITING AVAILABLE TIMES
			blnIsAvailable = True
		End If
		
		' IS THIS THE SELECTED TIME?
		If Trim(sCurrentCheckOutTime) = Trim("12:00:AM") Then
				sSelected = " selected=""selected"""
				blnIsAvailable = TRUE
			Else
				sSelected = ""
		End If

		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & " value=""12:00:AM"">12:00:AM" & "</option>"
		End If

		' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
		If blnFoundEndTime = True Then
			blnIsAvailable = False
		End If

		' 12:30 AM

		' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
		If Trim(sCheckInTime) = Trim("12:30:AM") Then
			' TURN ON WRITING AVAILABLE TIMES
			blnIsAvailable = True
		End If

		' IS THIS THE SELECTED TIME
		If Trim(sCurrentCheckOutTime) = Trim("12:30:AM") Then
			sSelected = " selected=""selected"""
			blnIsAvailable = TRUE
		Else
			sSelected = ""
		End If

		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & " value=""12:30:AM"">12:30:AM" & "</option>"
		End If

		' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
		If blnFoundEndTime = True Then
			blnIsAvailable = False
		End If

		' IS THIS THE SELECTED TIME
		If Trim(sCurrentCheckOutTime) = Trim("1:00:AM") Then
			sSelected = " selected=""selected"""
		Else
			sSelected = ""
		End If

		If sRealTime = "1:00:AM" Then
			response.write "<option" & sSelected & " value=""1:00:AM"">1:00:AM" & "</option>"
		End If
		
		
		%>
	</select>

	</td></tr>
	</table>
</p>
</div>


<!--BEGIN: RECURRENCE-->
<% If blnRecurrence Then %>
	<div class="reserveformtitle">Recurrent Reservation </div>
	<div class="reserveforminputarea">
		First Recurrance:  1/18/2006 8:00:00 AM <br>
		 Note:  Every 1 week(s) starting on Wednesday until 3/8/2006 
	</div>
<% End If %>
<!--END: RECURRENCE-->


<!--BEGIN: LESSEE-->
<div class="reserveformtitle">Lessee</div>
<div class="reserveforminputarea">
	<% DisplayUserInfo iUserID %>
</div>
<!--END: LESSEE-->


<!--BEGIN: SAVE BUTTON-->
<input align="center" style="text-align:center;" class="facilitybutton" type="submit" value="SAVE RESERVATION" />
<!--END: SAVE BUTTON-->


</form>
<!--END: PAGE CONTENT-->


<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>



<%
'--------------------------------------------------------------------------------------------------
'  DRAWAVAILABILITY(IFACILITYID)
'--------------------------------------------------------------------------------------------------
Sub DrawAvailability( ByVal ifacilityid, ByVal itempartid, ByVal iYear, ByVal iMonth, ByVal iDay )
	Dim sSql, oRs, iTimePartCount, sChecked, sTimeRange

	sSql = "SELECT facilityid, rateid, facilitytimepartid, beginhour, beginampm, endhour, endampm, weekday,description,rate "
	sSql = sSql & "FROM egov_facilitytimepart WHERE facilityid = " & ifacilityid & " AND weekday = " & Weekday( iMonth & "/" & iDay & "/" & iYear )
	sSql = sSql & " ORDER BY weekday, description,beginampm, beginhour"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	iTimePartCount = 0

	If Not oRs.EOF Then
		
		response.write "<fieldset style=""padding:5px;"">"
		response.write "<legend><b>Available Time(s)</b></legend>"

		Do While Not oRs.EOF 
			sChecked = ""
			If clng(itimepartid) = clng(oRs("facilitytimepartid")) Then
				sChecked = " checked=""checked"""
			End If
			
			irateid = oRs("rateid")
			
			sTimeRange = " " & oRs("beginhour") & " " & oRs("beginampm") & "-" & oRs("endhour") & " " & oRs("endampm") & " - " & oRs("description") & GetTimePartStatusName( iFacilityid, itimepartid, sTimeRange, iTimePartCount )
			response.write "<input type=""checkbox"" name=""timeparts"" value=""" & iTimePartCount & """" & sChecked & " class=""reserveformcheckbox"" style=""" & GetTimePartStatusColor( iFacilityid, itimepartid, sTimeRange, iTimePartCount ) &""" onClick=""test(this.form.timeparts,this.value);"" > " & sTimeRange & "<br>"
			iTimePartCount = iTimePartCount  + 1

			oRs.MoveNext
		Loop

		response.write "</fieldset>"

	End If

	oRs.Close
	Set oRs = Nothing 
		
End Sub


'--------------------------------------------------------------------------------------------------
'  BuildJavascriptTimePartArray(IFACILITYID)
'--------------------------------------------------------------------------------------------------
Sub BuildJavascriptTimePartArray( ByVal ifacilityid, ByVal iYear, ByVal iMonth, ByVal iDay )
	Dim sSql, oRs

	sSql = "SELECT rate, facilityid, rateid, facilitytimepartid, beginhour, beginampm, endhour, endampm, weekday, description, rate "
	sSql = sSql & "FROM egov_facilitytimepart WHERE facilityid =  " & ifacilityid & " AND weekday = " & Weekday( iMonth & "/" & iDay & "/" & iYear )
	sSql = sSql & " ORDER BY weekday, description,beginampm, beginhour"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	iArrayCount = 0

	If Not oRs.EOF Then
	
		response.write "var timeparts = new Array(" & oRs.recordcount - 1 & ");" & vbcrlf
		
		Do While Not oRs.EOF 
			arrStart = Split(oRs("beginhour"),":")
			arrEnd = Split(oRs("endhour"),":")
			sStartHour = GetMilitaryTime(clng(arrStart(0)),clng(arrEnd(0)),oRs("beginampm"),oRs("endampm"),0)
			sStartMinute = clng(arrStart(1))
			sEndHour = GetMilitaryTime(clng(arrStart(0)),clng(arrEnd(0)),oRs("beginampm"),oRs("endampm"),1)
			sEndMinute = clng(arrEnd(1))
			response.write "timeparts[" & iArrayCount & "] = new Array(4);" & vbcrlf
			response.write "timeparts[" & iArrayCount & "][0] = new Date(" & iYear &"," & iMonth & "," & iDay & "," & sStartHour & "," & sStartMinute & ",0);" & vbcrlf
			response.write "timeparts[" & iArrayCount & "][1] = new Date(" & iYear &"," & iMonth & "," & iDay & "," & sEndHour & "," & sEndMinute & ",0);" & vbcrlf
			response.write "timeparts[" & iArrayCount & "][2] = '" & oRs("rate") & "';" & vbcrlf
			' HANDLE TIME IF IT JUMPS TO NEXT DATE
			If sEndHour > 24 Then
				response.write "timeparts[" & iArrayCount & "][3] = '1';" & vbcrlf
			Else
				response.write "timeparts[" & iArrayCount & "][3] = '0';" & vbcrlf ' NEXT DAY
			End If
			iArrayCount = iArrayCount + 1

			oRs.MoveNext
		Loop

	End If

	oRs.Close
	Set oRs = Nothing 
		
End Sub


'--------------------------------------------------------------------------------------------------
'  FUNCTION GETMILITARYTIME(IHOUR,IENDHOUR,SBEGINAMPM,SENDAMPM,ISTARTOREND)
'--------------------------------------------------------------------------------------------------
Function GetMilitaryTime( ByVal iHour, ByVal iEndHour, ByVal sBeginAMPM, ByVal sEndAMPM, ByVal iStartorEnd )
	Dim iReturnValue
	
	' SET DEFAULT RETURN VALUE
	If iStartorEnd = 0 Then
		iReturnValue = iHour
		iTempHour = iHour
		sTempAM = sBeginAMPM
	Else
		iReturnValue = iEndHour
		iTempHour = iEndHour
		sTempAM = sEndAMPM
	End If

	' NON-MIDNIGHT AND NON-NOON HOURS
	If (iTempHour < 12) And (UCase(sTempAM)="AM") Then
		iReturnValue = iTempHour
	Else
		iReturnValue = iTempHour + 12
	End If

	' NOON
	If iTempHour = 12 And (UCase(sTempAM)="PM") Then
		iReturnValue = 12
	End If 

	' MIDNIGHT
	If iTempHour = 12 And (UCase(sTempAM)="AM") Then
		iReturnValue = 0
	End If 

	' SEE IF END TIME CROSSES MIDNIGHT
	If ihour > iEndHour And UCase(sBeginAMPM)="AM" And UCase(sEndAMPM)="AM" And iStartorEnd = 1  Then
		iReturnValue = iTempHour + 24	
	End If


	' RETURN VALUE 
	GetMilitaryTime = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETTIMEPARTSTATUS(IFACILITYID,ITIMEPARTID,STIMERANGE)
'--------------------------------------------------------------------------------------------------
Function GetTimePartStatusColor( ByVal iFacilityid, ByVal itimepartid, ByVal sTimeRange, ByVal iDayofWeek )
	Dim sReturnValue

	sReturnValue = sTimeRange

	iStatus = 1

	Select Case iStatus

		Case 1
		' OPEN
		sReturnValue = "background-color:green;"
		
		Case 2
		' RESERVED
		sReturnValue = "background-color:red;"

		Case 3
		' ON HOLD
		sReturnValue = ";background-color:yellow;"

	End Select

	 GetTimePartStatusColor = sReturnValue

End Function



'--------------------------------------------------------------------------------------------------
' FUNCTION GETTIMEPARTSTATUS(IFACILITYID,ITIMEPARTID,STIMERANGE)
'--------------------------------------------------------------------------------------------------
Function GetTimePartStatusName( ByVal iFacilityid, ByVal itimepartid, ByVal sTimeRange, ByVal iDayofWeek )
	Dim sReturnValue

	sReturnValue = sTimeRange

	iStatus = 1

	Select Case iStatus

		Case 1
		' OPEN
		sReturnValue = "<font style=""color:green;""> (OPEN)</font>"
		
		Case 2
		' RESERVED
		sReturnValue = "<font style=""color:red;""> (RESERVED)</font>"

		Case 3
		' ON HOLD
		sReturnValue = "<font style=""color:yellow;""> (ON HOLD)</font>"

	End Select

	 GetTimePartStatusName = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETCHECKINTIME(ITIMEPARTID)
'--------------------------------------------------------------------------------------------------
Function GetCheckInTime( ByVal itimepartid )
	Dim sSql, oRs, sReturnValue

	sReturnValue = "UNKNOWN"

	sSql = "SELECT beginhour, beginampm FROM egov_facilitytimepart WHERE facilitytimepartid = " & itimepartid 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		sReturnValue = oRs("beginhour") & ":" &  oRs("beginampm")
	End If

	oRs.Close
	Set oRs = Nothing
	
	' RETURN STATUS
	GetCheckInTime = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETCHECKOUTTIME(ITIMEPARTID)
'--------------------------------------------------------------------------------------------------
Function GetCheckOutTime( ByVal itimepartid )
	Dim sSql, oRs, sReturnValue

	sReturnValue = "UNKNOWN"

	sSql = "SELECT endhour, endampm FROM egov_facilitytimepart WHERE facilitytimepartid = " & itimepartid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		sReturnValue = oRs("endhour") & ":" &  oRs("endampm")
	End If

	oRs.Close
	Set oRs = Nothing
	
	' RETURN STATUS
	GetCheckOutTime = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' GetReservationDetails( iReservationID, sFacilityName, sCheckInDate, sCurrentCheckInTime, sCheckOutDate, sCurrentCheckOutTime, blnRecurrence, itimepartid, ifacilityid, sStatus, iUserID, ioccurrenceid )
'--------------------------------------------------------------------------------------------------
Sub GetReservationDetails( ByVal iReservationID, ByRef sFacilityName, ByRef sCheckInDate, ByRef sCurrentCheckInTime, ByRef sCheckOutDate, ByRef sCurrentCheckOutTime, ByRef blnRecurrence, ByRef itimepartid, ByRef ifacilityid, ByRef sStatus, ByRef iUserID, ByRef ioccurrenceid )
	Dim sSql, oRs

	' GET INFORMATION FOR THIS RESERVATION
	'sSql = "SELECT * FROM egov_facilityschedule S INNER JOIN egov_facility F ON S.facilityid = F.facilityid "
	'sSql = sSql & "WHERE S.facilityscheduleid = " & iReservationID 
	sSql = "SELECT S.facilitytimepartid, S.facilityid, facilityname, checkindate, checkintime, checkoutdate, checkouttime, "
	sSql = sSql & "isrecurrent, status, lesseeid, facilityrecurrenceid, ISNULL(internalnote,'') AS internalnote "
	sSql = sSql & "FROM egov_facilityschedule S INNER JOIN egov_facility F ON S.facilityid = F.facilityid "
	sSql = sSql & "WHERE S.facilityscheduleid = " & CLng(iReservationID)

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	' IF RESERVATION HAS INFORMATION POPULATE VALUES
	If Not oRs.EOF Then
		sFacilityName = oRs("facilityname") 
		sCheckInDate = oRs("checkindate") 
		sCurrentCheckInTime = oRs("checkintime") 
		sCheckOutDate = oRs("checkoutdate") 
		sCurrentCheckOutTime = oRs("checkouttime") 
		blnRecurrence = oRs("isrecurrent")
		itimepartid = oRs("facilitytimepartid")
		ifacilityid = oRs("facilityid")
		sStatus = oRs("status") 
		iUserID = oRs("lesseeid")
		ioccurrenceid = oRs("facilityrecurrenceid")
	End If

	oRs.Close
	Set oRs = Nothing
	
End Sub


'------------------------------------------------------------------------------------------------------------
' FUNCTION GETFACILITYFIELDVALUES(IFACILITYPAYMENTID)
'------------------------------------------------------------------------------------------------------------
Function GetFacilityFieldValues( ByVal iFacilityPaymentID )
	Dim sSql, oRs, sReturnValue
		
	sReturnValue = ""

	sSql = "SELECT V.facilityvalueid, V.fieldid, V.fieldvalue, V.paymentid, F.fieldprompt, F.fieldtype, F.facilityid, F.sequence, F.isrequired, F.fieldchoices "
	sSql = sSql & "FROM egov_facility_field_values V INNER JOIN egov_facility_fields F ON V.fieldid = F.fieldid WHERE V.paymentid = " & iFacilityPaymentID
	sSql = sSql & " ORDER BY V.paymentid, V.fieldid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		sReturnValue = sReturnValue & oRs("fieldprompt") & " : " & oRs("fieldvalue") & "<br>" 			
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing

	GetFacilityFieldValues = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYUSERINFO(IUSERID)
'--------------------------------------------------------------------------------------------------
Sub DisplayUserInfo( ByVal iUserID )
	Dim sSql, oRs

	' SELECT ROW WITH THIS USER'S INFORMATION
	'sSql = "SELECT * FROM egov_users WHERE userid = " & iUserId
	sSql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, ISNULL(useraddress,'') AS useraddress, "
	sSql = sSql & "ISNULL(useremail,'') AS useremail, ISNULL(usercity,'') AS usercity, ISNULL(userstate,'') AS userstate, "
	sSql = sSql & "ISNULL(userzip,'') AS userzip FROM egov_users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	' IF RECORDSET NOT EMPTY THEN DISPLAY USER INFORMATION
	If Not oRs.EOF Then
		
		response.write "<table>"
		response.write "<tr><td>Name: </td><td>" & oRs("userfname") & " " & oRs("userlname")  &  "</td></tr>"
		response.write "<tr><td>Address: </td><td>" & oRs("useraddress") & "</td></tr>"
		response.write "<tr><td>Email: </td><td>" & oRs("useremail") & "</td></tr>"
		response.write "<tr><td>City: </td><td>" & oRs("usercity") & "</td></tr>"
		response.write "<tr><td>State: </td><td>" & oRs("userstate") & "</td></tr>"
		response.write "<tr><td>Zip: </td><td>" & oRs("userzip") & "</td></tr>"
		response.write "</table>"

	End If

	oRs.Close
	Set oRs = Nothing
	
End Sub


'--------------------------------------------------------------------------------------------------
' SUB UPDATERECORD(IRESERVATIONID)
'--------------------------------------------------------------------------------------------------
Sub UpdateRecord( ByVal iReservationId, ByVal sCheckInTime, ByVal sCheckOutTime )
	Dim sSql

	' UPDATE RECORD WITH NEW TIME INFORMATION
	sSql = "UPDATE EGOV_FACILITYSCHEDULE SET "
	sSql = sSql & "checkintime = '" & sCheckInTime & "', "
	sSql = sSql & "checkouttime = '" & sCheckOutTime & "' WHERE facilityscheduleid = " & CLng(iReservationId)

	RunSQLStatement sSql

'	Set oUpdate = Server.CreateObject("ADODB.Recordset")
'	oUpdate.Open sSql, Application("DSN") , 3, 1
'	Set oUpdate = Nothing

End Sub




%>
