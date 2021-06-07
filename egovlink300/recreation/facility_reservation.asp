<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="facility_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: FACILITY_RESERVATION.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  01/18/06  John Stullenberger - INITIAL VERSION
' 1.1  10/10/08  David Boyer - Added "isFacilityAvail" check
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim selected_date, lcl_facility_avail, re, matches, lcl_TP, lcl_L

' Handle SQL Intrusions gracefully
If Not IsDate(request("D")) Then 
  	response.redirect "facility_list.asp"
Else
	selected_date = CDate(request("D"))
End If 

Set re = New RegExp
re.Pattern = "^\d+$"

' Check that the time part is a number'
lcl_TP = request("TP")
Set matches = re.Execute(lcl_TP)
If matches.Count < 1 Then
	response.redirect("facility_list.asp")
End If 

' Check that the facility id is a number'
lcl_L = request("L")
Set matches = re.Execute(lcl_L)
If matches.Count < 1 Then
	response.redirect("facility_list.asp")
End If 

'Check to see if this facility has been reserved while user has been filling out reservation form for same facility.
 If request("success") <> "NA" Then 
    lcl_facility_avail = isFacilityAvail("", request("D"), request("D"), request("TP"), request("L"), request("D"))

    If Not lcl_facility_avail Then 
       response.redirect "facility_availability.asp?L=" & request("L") & "&Y=" & year(request("D")) & "&M=" & month(request("D")) & "&success=NA"
    End If 
 End If 

 %>
<html lang="en">
<head>
	<meta charset="UTF-8">
<%
  if iorgid = 7 then
     lcl_title = sOrgName
  else
     lcl_title = "E-Gov Services " & sOrgName
  end if

  response.write "<title>" & lcl_title & "</title>" & vbcrlf
%>
	<link rel="stylesheet" href="../css/styles.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" />
	<link rel="stylesheet" href="facility_styles.css" />

	<script src="../scripts/jquery-1.7.2.min.js"></script>
	<script src="../scripts/modules.js"></script>
	<script src="../scripts/easyform.js"></script>
	
	<script>
	<!--

		function reloadpage() {
			var datdate = '<%=request("D")%>';
			var iFacility = $("#selfacility").val();
			location.href = 'facility_reservation.asp?Y=' + <%=Year(request("D"))%> + '&M=' + <%=Month(request("D"))%> + '&L=' + iFacility + '&D=' + datdate;
		}

	// TIMEPART ARRAY INFORMATION
	<% 	BuildJavascriptTimePartArray request("L"),Year(request("D")),Month(request("D")),Day(request("D")),request("TP") %>

		function test(opt,itimepartcount) {
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
				// CALCULATE TOTAL
				calculatetotal(opt);

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
							var datNewDate = new Date(document.getElementById("frmAvail").checkoutdate.value);
							datNewDate.setDate(datNewDate.getDate() + 1);
							document.getElementById("frmAvail").checkoutdate.value = (datNewDate.getMonth() + 1) + '/' + datNewDate.getDate() + '/' + datNewDate.getYear();
							blnFound = true;
						}
					}

				}

				// CLEAR ANY UNCHECKED TIME PARTS THAT SPAN TWO DAYS
				if (blnFound == false){
					document.getElementById("frmAvail").checkoutdate.value = document.getElementById("frmAvail").checkindate.value;
				}
			}
		}

		function CheckDateOverLap(datDateOneStart,datDateOneEnd,datDateTwoStart,datDateTwoEnd)
		{
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

		function calculatetotal(opt)
		{
			// RESET TOTAL TO ZERO
			var curTotal = 0;
			
		   // IN CASE OF MULTIPLE TIME PARTS
		   if (parseInt(opt.length) > 0) {

			   // LOOP THRU ALL TIME PARTS CHECKED
			   for (var intLoop = 0; intLoop < opt.length; intLoop++) {
					// IF CHECKED ADD TO TOTAL
					if ((opt[intLoop].checked)) {
						curTotal = parseInt(curTotal) + parseInt(timeparts[intLoop][2]);
					}
			   }
		   }
		   else
		   {
			   //IN CASE OF SINGLE TIME PART
			   if (opt.checked)
			   {
				curTotal = parseInt(timeparts[0][2]);
				
				}
			}

			// RESERVATION TIME TOTAL
			$("#reservetotal").val( formatCurrency(curTotal) );


		   // TOTAL COST
		   $("#amounttotal").val( formatCurrency(curTotal) );
		}

		function formatCurrency(num) 
		{
			num = num.toString().replace(/\$|\,/g,'');
			if(isNaN(num))
			num = "0";
			sign = (num == (num = Math.abs(num)));
			num = Math.floor(num*100+0.50000000001);
			cents = num%100;
			num = Math.floor(num/100).toString();
			if(cents<10)
			cents = "0" + cents;
			for (var i = 0; i < Math.floor((num.length-(1+i))/3); i++)
			num = num.substring(0,num.length-(4*i+3))+','+
			num.substring(num.length-(4*i+3));
			return (((sign)?'':'-') + num + '.' + cents);
		}

		function formatMoney(num) 
		{
			// same as formatCurrency, but with a $ sign
			num = num.toString().replace(/\$|\,/g,'');
			if(isNaN(num))
			num = "0";
			sign = (num == (num = Math.abs(num)));
			num = Math.floor(num*100+0.50000000001);
			cents = num%100;
			num = Math.floor(num/100).toString();
			if(cents<10)
			cents = "0" + cents;
			for (var i = 0; i < Math.floor((num.length-(1+i))/3); i++)
			num = num.substring(0,num.length-(4*i+3))+','+
			num.substring(num.length-(4*i+3));
			return (((sign)?'':'-') + '$' + num + '.' + cents);
		}

		function process_form()
		{
			// CHECK FOR ERRORS

			// UPDATE VALUES
			//document.getElementById("frmAvail").LodgeName.value = document.getElementById("frmAvail").selfacility.options[document.getElementById("frmAvail").selfacility.selectedIndex].text;
			$("#LodgeName").val( $('#selfacility option:selected').text() );

			// SUBMIT FORM
			document.getElementById("frmAvail").submit();
		}

	//-->
	</script>

</head>

<!--#Include file="../include_top.asp"-->

<%	RegisteredUserDisplay( "../" ) 


' CAPTURE CURRENT PATH
Session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()
Session("RedirectLang") = "Return to Facility Reservation"

Dim sFirstName,sLastName,sAddress,sCity,sState,sZip,sEmail,sHomePhone,sWorkPhone,sFax,sBusinessName, iCitizenUserId

SetUserInformation()
%>

<!--BEGIN PAGE CONTENT-->
<a href="facility_availability.asp?L=<%=request("L")%>&Y=<%=Year(request("D"))%>&M=<%=Month(request("D"))%>" class="linkbutton"><< Calendar</a>
<%
'CHECK TO SEE IF USER/ADDRESS IS BLOCKED, ONLY BLOCK FOR MONTGOMERY (ORG 26)
if iOrgId = 26 or iOrgId = 37 then
	sSQL = "SELECT u.userid " _
			& " FROM egov_Users u " _
			& " INNER JOIN egov_users u2 ON ((u2.useraddress = u.useraddress AND u2.usercity = u.usercity AND u2.userstate = u.userstate " _
										& " AND u2.userzip = u.userzip) OR u2.useremail = u.useremail)  " _
										& " AND u2.orgid = u.orgid " _
			& " WHERE u2.userID = '" & Track_DBSafe(request.cookies("userid")) & "' AND u.FacilityABUSE = 1 "
	Set oBlock = Server.CreateObject("ADODB.RecordSet")
	oBlock.Open sSQL, Application("DSN"), 3, 1
	if not oBlock.EOF then
		response.write "<p>We are not able to make this reservation at this time.  Please call 513-891-2424 for more information.</p>" 
		response.end
	end if
	oBlock.Close
	Set oBlock = Nothing
end if
%>
<form name="frmAvail" id="frmAvail" action="facility_reserve_summary.asp" method="post">
  <input type="hidden" id="L" name="L" value="<%=request("L")%>" />
  <input type="hidden" id="D" name="D" value="<%=request("D")%>" />
  <input type="hidden" id=="LodgeName" name="LodgeName" value="" />
  <input type="hidden" id="selTimePartID" name="selTimePartID" value="<%=request("TP")%>" />
  <input type="hidden" id="selfacility" name="selfacility" value="<%=request("L")%>" />

<%
  if request("success") <> "" then
     lcl_message = getSuccessMessage(request("success"))

     if lcl_message <> "" then
        response.write "<div align=""right"" style=""width: 750px"">" & lcl_message & "</div>" & vbcrlf
     end if
  end if
%>
<!--BEGIN:  USER REGISTRATION - USER MENU-->
<%
	If sOrgRegistration Then 
		If request.cookies("userid") <> "" And request.cookies("userid") <> "-1" Then 
			' they are logged in'
			iCitizenUserId = request.cookies("userid")
			'response.write "<!-- requires address check: " & OrgHasFeature( iOrgId, "requires address check" ) & " -->"
			If OrgHasFeature( iOrgId, "requires address check" ) Then
				' check if user has an address populated. CitizenAddressIsMissing is in include_top_functions.asp
				'Response.write "<!-- CitizenAddressIsMissing: " & CitizenAddressIsMissing( iCitizenUserId ) & " -->"
			 	If CitizenAddressIsMissing( iCitizenUserId ) Then 
					' if they are missing an address, take them to a page to enter it.
					session("RedirectPage") = "recreation/facility_reservation.asp?L=" & request("L") & "&TP=" & request("TP") & "&D=" & request("D")
					response.redirect "../getuseraddress.asp?userid=" & iCitizenUserId
				End If 
			End If 
			response.write "<p></p>" & vbcrlf
		Else 
			'Added this to make the login work like classes and events - Steve Loar 5/19/2006
			session("LoginDisplayMsg") = "Please sign in first and then we'll send you right along."
			response.redirect "../user_login.asp"
		End If 
	Else 
		response.write "<!--REGISTRATION OR LOGIN-->" & vbcrlf
		response.write "<p>" & vbcrlf
		response.write "<div class=""reserveformtitle"">Contact Information</div>" & vbcrlf
		response.write "<div class=""reserveforminputarea"">" & vbcrlf
		response.write "<p><font class=""reserveforminstructions"">"
		response.write "You need to sign in or register now to complete your purchase."
		If CLng(iOrgID) = CLng(26) Then 
			response.write "As a result of changes to the City of Montgomery"
			response.write " website you will need to re-register.  We apologize for the inconvenience, however this is a one-time process and once you register, you will not need to register again."
		End If 
		response.write "</font></p>" & vbcrlf
		response.write "<p>" & vbcrlf
		response.write "<input onClick=""location.href='../user_login.asp';"" value=""Login"" class=""reserveformbutton"" style=""width:75px;text-align:center;"" type=""button"" />" & vbcrlf
		response.write " or " & vbcrlf
		response.write "<input value=""Register Now!"" class=""reserveformbutton"" style=""width:150px;text-align:center;"" type=""button"" onClick=""location.href='../register.asp';"" />" & vbcrlf
		response.write "</p>" & vbcrlf
		response.write "</div>" & vbcrlf
		response.write "</p>" & vbcrlf
	End If 
%>
<!--END:  USER REGISTRATION - USER MENU-->

<!--BEGIN: SELECT FACILITY-->
<div class="reserveformtitle">Facility</div>
<div class="reserveforminputarea">
<!-- <p><font class="reserveforminstructions">Instructions: Select the facility for your reservation.</font></p> -->
<p>
	<% 
	ifacilityid = request("L")
	If ifacilityid = "" Then
  		ifacilityid = 0
	End If

	datCheckInDate  = request("D")
	datCheckOutDate = request("D") ' CHECK BASED ON TIMEPARTID
	itimepartid     = request("TP")

	'DrawSelectFacility ifacilityid
	response.write "<strong>" & getFacilityName( iFacilityId ) & "</strong>" ' in facility_global_functions.asp

	%>
	<input type="hidden" id="checkindate" name="checkindate"  value="<%=datCheckInDate%>" />
	<input type="hidden" id="checkoutdate" name="checkoutdate" value="<%=datCheckOutDate%>" />
</p>
</div>
<!--END: SELECT FACILITY-->

<!--BEGIN: SELECT DATES-->
<div class="reserveformtitle">Select Date/Time</div>
<div class="reserveforminputarea">
<p>
	<strong>Selected Date: </strong><%= selected_date %>
</p>
<p><font class="reserveforminstructions">Instructions: Select the Check-In and Check-Out times for your reservation.</font></p>
<!--END: SELECT DATES-->

<!--BEGIN: DRAW AVAILABILITY-->
<p><% DrawAvailability ifacilityid, itimepartid, Year(request("D")), Month(request("D")), Day(request("D")) %></p>
<!--END: DRAW AVAILABILITY-->

<%
If OrgHasDisplay( iOrgID, "facility arrival message" ) Then
  	response.write GetOrgDisplay( iOrgID, "facility arrival message" )
End If 
%>

<!--DRAW DATE/TIME SELECTION-->
<table>
  <tr>
      <td class="reservationformlabel">Exact Arrival Time:</td>
      <td>
         	<% sCheckInTime = GetCheckInTime( itimepartid )%>
         	<% sCheckOutTime = GetCheckOutTime( itimepartid )%>
         	<select name="checkintime">
	<%
		' 1 AM TO 11:30 AM
		blnIsAvailable = False
		For i = 1 to 11

			' BUILD TIME STRING
			sTime = i  & ":00:AM"
			If trim(sCheckInTime) = trim(sTime) Then
				sSelected = " selected=""selected"""
				blnIsAvailable = TRUE
			Else
				sSelected = ""
			End If
			' DISPLAY ONSCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & ">" & sTime & "</option>"
			End If

			' CHECK TO SEE IF WE HAVE PASSED CHECKOUT TIME
			If trim(sCheckOutTime) = trim(sTime) Then
				' TURN OFF WRITING AVAILABLE TIMES
				blnIsAvailable = False
			End If

			' BUILD TIME STRING
			sTime = i  & ":30:AM"
			If trim(sCheckInTime) = trim(sTime) Then
				sSelected = " selected=""selected"""
				blnIsAvailable = TRUE
			Else
				sSelected = ""
			End If

			' DISPLAY ONSCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & ">" & sTime & "</option>"
			End If

			' CHECK TO SEE IF WE HAVE PASSED CHECKOUT TIME
			If trim(sCheckOutTime) = trim(sTime) Then
				' TURN OFF WRITING AVAILABLE TIMES
				blnIsAvailable = False
			End If
		Next
		

		' NOON 
		If trim(sCheckInTime) = trim("12:00:PM") Then
				sSelected = " selected=""selected"""
				blnIsAvailable = TRUE
		Else
				sSelected = ""
		End If
		
		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & ">12:00:PM" & "</option>"
		End If

		' CHECK TO SEE IF WE HAVE PASSED CHECKOUT TIME
		If trim(sCheckOutTime) = trim("12:00:PM") Then
			' TURN OFF WRITING AVAILABLE TIMES
			blnIsAvailable = False
		End If

		' 12:30 PM
		If trim(sCheckInTime) = trim("12:30:PM") Then
			sSelected = " selected=""selected"""
			blnIsAvailable = TRUE
		Else
			sSelected = ""
		End If
		
		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & ">12:30:PM" & "</option>"
		End If

		' CHECK TO SEE IF WE HAVE PASSED CHECKOUT TIME
		If Trim(sCheckOutTime) = Trim("12:30:PM") Then
			' TURN OFF WRITING AVAILABLE TIMES
			blnIsAvailable = False
		End If

		
		' 1 PM TO 11:30PM
		For i= 1 to 11

			sTime = i  & ":00:PM"
			If trim(sCheckInTime) = trim(sTime) Then
				sSelected = " selected=""selected"""
				blnIsAvailable = TRUE
			Else
				sSelected = ""
			End If
			' DISPLAY ONSCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & ">" & sTime & "</option>"
			End If

			' CHECK TO SEE IF WE HAVE PASSED CHECKOUT TIME
			If trim(sCheckOutTime) = trim(sTime) Then
				' TURN OFF WRITING AVAILABLE TIMES
				blnIsAvailable = False
			End If

			sTime = i  & ":30:PM"
			If trim(sCheckInTime) = trim(sTime) Then
				sSelected = " selected=""selected"""
				blnIsAvailable = TRUE
			Else
				sSelected = ""
			End If
			' DISPLAY ONSCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & ">" & sTime & "</option>"
			End If

			' CHECK TO SEE IF WE HAVE PASSED CHECKOUT TIME
			If trim(sCheckOutTime) = trim(sTime) Then
				' TURN OFF WRITING AVAILABLE TIMES
				blnIsAvailable = False
			End If
		Next

		' MIDNIGHT
		If trim(sCheckInTime) = trim("12:00:AM") Then
				sSelected = " selected=""selected"""
				blnIsAvailable = TRUE
		Else
				sSelected = ""
		End If
		
		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & ">12:00:AM" & "</option>"
		End If

		' CHECK TO SEE IF WE HAVE PASSED CHECKOUT TIME
		If trim(sCheckOutTime) = trim("12:00:AM") Then
			' TURN OFF WRITING AVAILABLE TIMES
			blnIsAvailable = False
		End If

		' 12:30 AM
		If trim(sCheckInTime) = trim("12:30:AM") Then
			sSelected = " selected=""selected"""
			blnIsAvailable = TRUE
		Else
			sSelected = ""
		End If

		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & ">12:30:AM" & "</option>"
		End If

		' CHECK TO SEE IF WE HAVE PASSED CHECKOUT TIME
		If trim(sCheckOutTime) = trim("12:30:AM") Then
			' TURN OFF WRITING AVAILABLE TIMES
			blnIsAvailable = False
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

		If trim(sCheckOutTime) = "1:00:AM" Then
			sRealTime = "1:00:AM"
			sCheckOutTime = "12:30:AM"
		End If
		
		' BUILD TIME STRING
		For i=1 to 11

			sTime = i  & ":00:AM"
			' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
			If trim(sCheckInTime) = trim(sTime) Then
				' TURN ON WRITING AVAILABLE TIMES
				blnIsAvailable = True
			End If
			
			If trim(sCheckOutTime) = trim(sTime) Then
				sSelected = " selected=""selected"""
				blnFoundEndTime = True
			Else
				sSelected = ""
			End If
			
			' DISPLAY ONSCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & ">" & sTime & "</option>"
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

			If trim(sCheckOutTime) = trim(sTime) Then
				sSelected = " selected=""selected"""
				blnFoundEndTime = True
			Else
				sSelected = ""
			End If
			
			' DISPLAY ONSCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & ">" & sTime & "</option>"
			End If 

			' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
			If blnFoundEndTime = True Then
				blnIsAvailable = False
			End If
		Next
		

		' NOON 
		' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
		If trim(sCheckInTime) = trim("12:00:PM") Then
			' TURN ON WRITING AVAILABLE TIMES
			blnIsAvailable = TRUE
		End If

		' IS THIS THE SELECTED TIME?
		If trim(sCheckOutTime) = trim("12:00:PM") Then
				sSelected = " selected=""selected"""
				blnFoundEndTime = True
			Else
				sSelected = ""
		End If

		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & ">12:00:PM" & "</option>"
		End If

		' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
		If blnFoundEndTime = True Then
			blnIsAvailable = False
		End If

		' 12:30 PM
		' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
		If trim(sCheckInTime) = trim("12:30:PM") Then
			' TURN ON WRITING AVAILABLE TIMES
			blnIsAvailable = TRUE
		End If
		
		' IS THIS THE SELECTED TIME?
		If trim(sCheckOutTime) = trim("12:30:PM") Then
			sSelected = " selected=""selected"""
			blnIsAvailable = TRUE
		Else
			sSelected = ""
		End If

		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & ">12:30:PM" & "</option>"
		End If
		
		' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
		If blnFoundEndTime = True Then
			blnIsAvailable = False
		End If
		

		' 1 PM TO 11:30 PM
		For i= 1 to 11
			
			sTime = i  & ":00:PM"
			' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
			If trim(sCheckInTime) = trim(sTime) Then
				' TURN ON WRITING AVAILABLE TIMES
				blnIsAvailable = TRUE
			End If
		
			If trim(sCheckOutTime) = trim(sTime) Then
				sSelected = " selected=""selected"""
				blnFoundEndTime = True
			Else
				sSelected = ""
			End If

			' DISPLAY ON SCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & ">" & sTime & "</option>"
			End If
			
			' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
			If blnFoundEndTime = True Then
				blnIsAvailable = False
			End If
				
			

			sTime = i  & ":30:PM"
			' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
			If trim(sCheckInTime) = trim(sTime) Then
				' TURN ON WRITING AVAILABLE TIMES
				blnIsAvailable = True
			End If

			If trim(sCheckOutTime) = trim(sTime) Then
				sSelected = " selected=""selected"""
				blnFoundEndTime = True
			Else
				sSelected = ""
			End If

			' DISPLAY ON SCREEN
			If blnIsAvailable = True Then
				response.write "<option" & sSelected & ">" & sTime & "</option>"
			End If

			' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
			If blnFoundEndTime = True Then
				blnIsAvailable = False
			End If
				

		Next

		
		' MIDNIGHT 
		
		' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
		If trim(sCheckInTime) = trim("12:00:PM") Then
			' TURN ON WRITING AVAILABLE TIMES
			blnIsAvailable = True
		End If
		
		' IS THIS THE SELECTED TIME?
		If trim(sCheckOutTime) = trim("12:00:PM") Then
				sSelected = " selected=""selected"""
				blnIsAvailable = TRUE
			Else
				sSelected = ""
		End If

		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & ">12:00:AM" & "</option>"
		End If

		' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
		If blnFoundEndTime = True Then
			blnIsAvailable = False
		End If

		' 12:30 AM

		' CHECK TO SEE IF WE HAVE PASSED CHECKIN TIME
		If trim(sCheckInTime) = trim("12:30:PM") Then
			' TURN ON WRITING AVAILABLE TIMES
			blnIsAvailable = True
		End If

		' IS THIS THE SELECTED TIME
		If trim(sCheckOutTime) = trim("12:30:PM") Then
			sSelected = " selected=""selected"""
			blnIsAvailable = TRUE
		Else
			sSelected = ""
		End If

		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option" & sSelected & ">12:30:AM" & "</option>"
		End If

		' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
		If blnFoundEndTime = True Then
			blnIsAvailable = False
		End If

		
		If sRealTime = "1:00:AM" Then
			response.write "<option" & sSelected & ">1:00:AM" & "</option>"
		End If
		
		
		%>
	</select>	

	</td></tr>
	<tr><td class="reservationformlabel">Reservation Cost:</td><td><input id="reservetotal" name="reservetotal" style="text-align:right;" readonly type="text" value="$0.00" /></td></tr>
	</table>

<%	If OrgHasDisplay( iOrgid, "facility deposit message" ) Then %>	
		<!--KEY CHARGE-->
		<table>
			<tr><td colspan="2"><!--<input  name="keydeposit" type=checkbox checked onClick="calculatetotal(frmAvail.timeparts);">Pay key deposit charge online now (equal to reservation cost).<br />-->
				<!--<strong>Important!</strong> <font color=red> <i>You will be required to pay a key deposit charge equal to the cost of the reservation or $100.00, whichever is greater at the time you pick up the key from the City Hall.</i> </font>-->
				<%= GetOrgDisplay( iOrgid, "facility deposit message" ) %>
			</td></tr>
		</table>
<%	End If  %>

	<!--TOTAL COST-->
	<table>
		<tr><td class="reservationformlabel">Total Amount:</td><td><input id="amounttotal" name="amounttotal" style="text-align:right;" readonly type="text" value="0.00" /></td></tr>
	</table>



</p>
</div>

<!--RESERVATION PURPOSE/EVENT INFORMATION-->
<div class="reserveformtitle">Reservation Purpose/Event Information</div>
<div class="reserveforminputarea">
<P><font class="reserveforminstructions">Please provide the following to allow us to better serve your needs.</font></p>
<%DrawFacilityFields(ifacilityid)%>


<!--WAIVERS-->
<% DrawWaivers ifacilityid  %>

<P><strong>Note: You will be presented with summary of reservation costs and terms/conditions before the payment will be processed.<br /></strong></P>
</div>


<!--CONTINUE BUTTON-->
<%
If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then 
	sDisabled = ""
	sMsg = ""
Else
	sDisabled = "DISABLED"
	sMsg = "<br /><strong>*You need to <a class=""none"" href=""../user_login.asp"">Login</a> or <a class=""none"" href=""../register.asp"">Register Now!</a> to continue with this reservation."
End IF


 'If the facility is no longer available to reserve then hide the "Continue" button.
  if request("success") = "NA" then
     if lcl_message <> "" then
        response.write "<div style=""width: 750px"">" & lcl_message & "</div><br />" & vbcrlf
     end if

     response.write "<input type=""button"" value=""Return to Calendar"" style=""text-align:center;"" class=""reserveformbutton button"" onclick=""location.href='facility_availability.asp?L=" & request("L") & "&Y=" & year(request("D")) & "&M=" & month(request("D")) & "'"" />" & vbcrlf
  else
     response.write "<input type=""button"" value=""CONTINUE WITH RESERVATION"" style=""text-align:center;"" class=""facilitybutton"" onclick=""if (validateForm('frmAvail')) {process_form();}"" " & sDisabled & " />" & vbcrlf
  end if
%>

<%=sMsg%>
</form>
<!--END: PAGE CONTENT-->


<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--UPDATE COSTS-->
<SCRIPT>
	calculatetotal( document.getElementById("frmAvail").timeparts);
</SCRIPT>

<!--#Include file="../include_bottom.asp"-->  


<%
'--------------------------------------------------------------------------------------------------
'  DrawSelectFacility ifacilityid
'--------------------------------------------------------------------------------------------------
Sub DrawSelectFacility( ByVal ifacilityid )
	Dim sSql, oRs, sSelected
	
	If ifacilityid = "" Then
		ifacilityid = 0
	End If

	' GET SELECT CATEGORY ROW
	sSql = "SELECT facilityid, facilityname FROM egov_facility WHERE isviewable = 1 AND orgid = " & iorgid & " ORDER BY facilityname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

    ' LOOP THRU LIST OF AVAILABLE FACILITIES AND DISPLAY TO USER
    Response.Write "<font class=""reservationformlabel"">Facility:</font> <select class=""reservationformselect facilitylist"" onChange=""reloadpage();"" id=""selfacility"" name=""selfacility"">"

    Do While Not oRs.EOF
		sSelected = ""

		If CLng(ifacilityid) = CLng(oRs("facilityid")) Then
			sSelected = " selected=""selected"""
		End If
		
		Response.Write vbcrlf & "<option" & sSelected & " value=""" & oRs("facilityid") & """>" & oRs("facilityname") & "</option>"

		oRs.MoveNext
	Loop

    Response.Write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
'  SUB DRAWAVAILABILITY(IFACILITYID)
'--------------------------------------------------------------------------------------------------
Sub DrawAvailability( ByVal ifacilityid, ByVal itempartid, ByVal iYear, ByVal iMonth, ByVal iDay )
	Dim sSql, oRs, sChecked, iTimePartCount

	sSql = "SELECT facilityid, rateid, facilitytimepartid, beginhour, beginampm, endhour, endampm, weekday, description, rate "
	sSql = sSql & "FROM egov_facilitytimepart WHERE facilityid = " & ifacilityid & " AND weekday = '" & weekday( iMonth & "/" & iDay & "/" & iYear ) &"' "
	sSql = sSql & "ORDER BY weekday, description, beginampm, beginhour"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	iTimePartCount = 0

	' IF THERE ARE TIMEPARTS FOR THIS DAY LOOP THRU THEM
	If Not oRs.EOF Then
		
		' BEGIN BUILDING LIST OF TIMEPARTS
		response.write "<fieldset style=""padding:5px;"">"
		response.write "<legend><strong>Available Time(s)</strong></legend>"

		' LOOP THRU ALL TIME PARTS
		Do While Not oRs.EOF 
			
			' CHECK SELECTED TIMEPART
			sChecked = ""
			If CLng(itimepartid) = CLng(oRs("facilitytimepartid")) Then
				sChecked = " checked=""checked"""
			End If
			
			' GET STATUS OF CURRENT TIMEPART ROW
			sStatus = GetTimePartStatus(iFacilityid,oRs("facilitytimepartid"),iMonth & "/" & iDay & "/" & iYear)
			'response.write sStatus
			' IF STATUS IS RESERVED/ONHOLD MARK AS READONLY
			If sStatus = "RESERVED" or sStatus = "ONHOLD" Then
				sREADONLY = "disabled"
			Else
				sREADONLY = ""
			End If

			' BUILD TIMEPART DISPLAY STRING
			'sTimeRange = formatcurrency(oRs("rate"),2) & " - " & oRs("beginhour") & " " & oRs("beginampm") & "-" & oRs("endhour") & " " & oRs("endampm") & " - " & oRs("description") & GetTimePartStatusName(iFacilityid,oRs("facilitytimepartid"),sStatus)
			sTimeRange =  FormatNumber( GetFacilityRate( oRs("rateid"), ifacilityid ), 2 ) & " - " & oRs("beginhour") & " " & oRs("beginampm") & "-" & oRs("endhour") & " " & oRs("endampm") & " - " & oRs("description") & GetTimePartStatusName( sStatus )

			' ONLY DISPLAY TIMEPART TIME/STATUS INFORMATION FOR THE ONE SELECTED
			If CLng(itimepartid) = CLng(oRs("facilitytimepartid")) Then
  				response.write "<input " & sREADONLY & " type=""checkbox"" id=""timeparts"" name=""timeparts"" value=""" & iTimePartCount & """ " & sChecked & " class=""reserveformcheckbox"" style=""" & GetTimePartStatusColor(sStatus) &""" onClick=""test(this.form.timeparts,this.value);"" /> " & sTimeRange & "<br />"
			End If

			iTimePartCount = iTimePartCount  + 1

			' MOVE TO NEXT TIMEPART
			oRs.MoveNext
		Loop

		response.write "</fieldset>"

	End If

	oRs.Close
	Set oRs = Nothing 
		
End Sub


'--------------------------------------------------------------------------------------------------
'  BuildJavascriptTimePartArray( ifacilityid, iYear, iMonth, iDay, itimepartid
'--------------------------------------------------------------------------------------------------
Sub BuildJavascriptTimePartArray( ByVal ifacilityid, ByVal iYear, ByVal iMonth, ByVal iDay, ByVal itimepartid )
	Dim sSql, oRs, arrStart, arrEnd, sStartHour, sStartMinute, sEndHour, sEndMinute, iArrayCount

	sSql = "SELECT rate, facilityid, rateid, facilitytimepartid, beginhour, beginampm, endhour, endampm, weekday, description, rate "
	sSql = sSql & "FROM egov_facilitytimepart where facilityid = " & ifacilityid & " AND weekday = " & Weekday( iMonth & "/" & iDay & "/" & iYear )
	sSql = sSql & " ORDER BY weekday, description,beginampm, beginhour"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 3
	iArrayCount = 0

	If Not oRs.EOF Then
	
		response.write "var timeparts = new Array(" & oRs.recordcount - 1 & ");" & vbcrlf
		
		Do While Not oRs.EOF 
			
			If clng(itimepartid) = clng(oRs("facilitytimepartid")) Then
			
				arrStart = Split(oRs("beginhour"),":")
				arrEnd = Split(oRs("endhour"),":")
				sStartHour = GetMilitaryTime(clng(arrStart(0)),clng(arrEnd(0)),oRs("beginampm"),oRs("endampm"),0)
				sStartMinute = clng(arrStart(1))
				sEndHour = GetMilitaryTime(clng(arrStart(0)),clng(arrEnd(0)),oRs("beginampm"),oRs("endampm"),1)
				sEndMinute = clng(arrEnd(1))
				response.write "timeparts[" & iArrayCount & "] = new Array(4);" & vbcrlf
				response.write "timeparts[" & iArrayCount & "][0] = new Date(" & iYear &"," & iMonth & "," & iDay & "," & sStartHour & "," & sStartMinute & ",0);" & vbcrlf
				response.write "timeparts[" & iArrayCount & "][1] = new Date(" & iYear &"," & iMonth & "," & iDay & "," & sEndHour & "," & sEndMinute & ",0);" & vbcrlf
				response.write "timeparts[" & iArrayCount & "][2] = '" & GetFacilityRate( oRs("rateid"), ifacilityid ) & "';" & vbcrlf
				' HANDLE TIME IF IT JUMPS TO NEXT DATE
				If sEndHour > 24 Then
					response.write "timeparts[" & iArrayCount & "][3] = '1';" & vbcrlf
				Else
					response.write "timeparts[" & iArrayCount & "][3] = '0';" & vbcrlf ' NEXT DAY
				End If

				iArrayCount = iArrayCount + 1

			End If
			
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
	Dim iReturnValue, iTempHour, sTempAM
	
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
	If (iTempHour < 12) And (UCase(sTempAM) = "AM") Then
		iReturnValue = iTempHour
	Else
		iReturnValue = iTempHour + 12
	End If

	' NOON
	If iTempHour = 12 And (UCase(sTempAM) = "PM") Then
		iReturnValue = 12
	End If 

	' MIDNIGHT
	If iTempHour = 12 And (UCase(sTempAM) = "AM") Then
		iReturnValue = 0
	End If 

	' SEE IF END TIME CROSSES MIDNIGHT
	If ihour > iEndHour And UCase(sBeginAMPM) = "AM" And UCase(sEndAMPM) = "AM" And iStartorEnd = 1 Then
		iReturnValue = iTempHour + 24	
	End If

	' RETURN VALUE 
	GetMilitaryTime = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETTIMEPARTSTATUS(IFACILITYID,ITIMEPARTID,STIMERANGE)
'--------------------------------------------------------------------------------------------------
Function GetTimePartStatusColor( ByVal sStatus )
	Dim sReturnValue

	Select Case sStatus

		Case "OPEN","CANCELLED"
			' OPEN
			sReturnValue = "background-color:green;"
		
		Case  "RESERVED", "ONHOLD"
			' RESERVED
			sReturnValue = "background-color:red;"

	End Select

	 GetTimePartStatusColor = sReturnValue

End Function



'--------------------------------------------------------------------------------------------------
' FUNCTION GETTIMEPARTSTATUS(IFACILITYID,ITIMEPARTID,STIMERANGE)
'--------------------------------------------------------------------------------------------------
Function GetTimePartStatusName( ByVal sStatus )
	Dim sReturnValue

	Select Case sStatus

		Case "OPEN"
		' OPEN
		sReturnValue = " <font style=""color:green;"">(OPEN)</font>"
		
		Case "CANCELLED"
		' OPEN
		sReturnValue = " <font style=""color:green;"">(OPEN)</font>"

		Case "RESERVED","ONHOLD"
		' RESERVED
		sReturnValue = " <font style=""color:red;"">(RESERVED)</font>"

	End Select

	 GetTimePartStatusName = sReturnValue

End Function



'--------------------------------------------------------------------------------------------------
' DRAWTIMEPARTS(IFACILITYID,IDAYOFWEEK,SCELLDATE)
'--------------------------------------------------------------------------------------------------
Sub DrawWaivers( ByVal ifacilityid )
	Dim sSql, oRs, iCount

	sSql = "SELECT F.waiverid, W.description, W.isrequired FROM egov_facilitywaivers F, egov_waivers W "
	sSql = sSql & "WHERE F.waiverid = W.waiverid AND F.facilityid = " & ifacilityid & " ORDER BY W.isrequired, W.name"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	iCount = 0
	If Not oRs.EOF Then
		iCount = iCount + 1
		response.write "<p><strong>Waivers/Release Forms:</strong></p><br />"

		Do While NOT oRs.EOF 
			If oRs("isrequired") Then
				response.write "<input type=""hidden"" name=""chkwaivers_" & iCount & """  value=""" & oRs("waiverid") & """ />"
			Else
				response.write "<input name=""chkwaivers_" & iCount & """ type=""checkbox"" value=""" & oRs("waiverid") & """ />" & oRs("description") & "<br />"  
			End If
			oRs.MoveNext
		Loop

	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' SUB DRAWFACILITYFIELDS(IFACILITYID)
'--------------------------------------------------------------------------------------------------
Sub DrawFacilityFields( ByVal iFacilityID )
	Dim sSql, oRs, sRequired, sHeight, arrAnswers

	sSql = "SELECT * FROM egov_facility_fields WHERE facilityid = " & ifacilityID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		
		response.write "<P>"
		response.write "<TABLE>"

		Do While NOT oRs.EOF

			' CHECK IF IT IS REQUIRED
			If  oRs("isrequired") Then
				sRequired = "<span style=""color: #ff0000"">*&nbsp;</span>"
				response.write "<input type=""hidden"" name=""ef:custom_" & oRs("fieldname") & "_" & oRs("fieldid") & "-text/" & oRs("validation") & """ value=""" & oRs("fieldprompt") & """>"
			Else
				sRequired = " "
			End If

			' SET HEIGHT FOR INPUT BOX BASED ON FIELD TYPE, 1=STANDARD, 2=SIMULATED TEXT AREA
			If  oRs("fieldtype") = 2 Then
				sHeight = "HEIGHT: 100px;"
			Else
				sHeight = " " 
			End If
		
			response.write "<tr bgColor=""#e0e0e0"">"
			response.write "<td valign=""top"" class=reservationformlabel align=""right"">" & sRequired
			response.write oRs("fieldprompt")
			response.write ": </td>"

			response.write "<td style=""font-family:Arial; font-size:8pt; color:#000000"" align=""left"">"
			
			Select Case oRs("fieldtype")

				Case "1"
					' TEXT BOX
					response.write "<INPUT name=""custom_" & oRs("fieldname") & "_" & oRs("fieldid") & """ type=""text"" style=""FONT-SIZE: 8pt; WIDTH: 300px; " & sHeight & " FONT-FAMILY: Arial"" >"
				Case "2"
					' PSEUDO TEXT AREA
					response.write "<textarea name=""custom_" & oRs("fieldname") & "_" & oRs("fieldid") & """ style=""FONT-SIZE: 8pt; WIDTH: 300px; " & sHeight & " FONT-FAMILY: Arial"" ></textarea>"
				Case "3"
					' SELECT BOX
					arrAnswers = Split(oRs("fieldchoices"),"@@")
			
					response.write "<select name=""custom_" & oRs("fieldname") & "_" & oRs("fieldid") & """ >"
					For alist = 0 to UBound(arrAnswers)
						response.write "<option value=""" & arrAnswers(alist) & """>" & arrAnswers(alist) & "</option>" 
					Next
					response.write "</select>"

				Case Else
					' UNKNOWN TYPE DONT PROCESS
					response.write "INPUT TYPE ERROR. PLEASE CONSULT SETUP."
			
			End Select
			response.write "</td></tr>"

			oRs.MoveNext
		Loop

		response.write "</table>"
		response.write "</p>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' SUB SETUSERINFORMATION()
'--------------------------------------------------------------------------------------------------
Sub SetUserInformation()
	Dim sSql, oRs, iUserID

	If sOrgRegistration Then 
		If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then
			
			iUserID = request.cookies("userid")
		
			sSql = "SELECT userfname, userlname, ISNULL(useraddress,'') AS useraddress, ISNULL(usercity,'') AS usercity, "
			sSql = sSql & "ISNULL(userstate,'') AS userstate, ISNULL(userzip,'') AS userzip, ISNULL(useremail,'') AS useremail, "
			sSql = sSql & "ISNULL(userhomephone,'') AS userhomephone, ISNULL(userworkphone,'') AS userworkphone, "
			sSql = sSql & "ISNULL(userbusinessname,'') AS userbusinessname, ISNULL(userfax,'') AS userfax "
			sSql = sSql & "FROM egov_users WHERE userid = " & iUserID

			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open sSql, Application("DSN"), 3, 1

			If Not oRs.EOF Then
				' USER FOUND SET VALUES
				sFirstName = oRs("userfname")
				sLastName = oRs("userlname")
				sAddress = oRs("useraddress")
				sCity = oRs("usercity")
				sState = oRs("userstate")
				sZip = oRs("userzip")
				sEmail = oRs("useremail")
				sHomePhone = oRs("userhomephone")
				sWorkPhone = oRs("userworkphone")
				sBusinessName = oRs("userbusinessname")
				sFax = oRs("userfax")
			Else
				' USER NOT FOUND SET VALUES TO EMPTY
				sFirstName = ""
				sLastName = ""
				sAddress = ""
				sCity = ""
				sState = ""
				sZip = ""
				sEmail = ""
				sHomePhone = ""
				sWorkPhone = ""
				sFax = ""
				sBusinessName = ""
			End If
			
			oRs.Close
			Set oRs = Nothing

		End If
	End If

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION GETTIMEPARTSTATUS(IFACILITYID,ITIMEPARTID,STIMERANGE)
'--------------------------------------------------------------------------------------------------
Function GetTimePartStatus( ByVal iFacilityid, ByVal itimepartid, ByVal datDate )
	Dim sSql, oRs, sReturnValue

	sReturnValue= "OPEN"

	' GET STATUS OF THIS TIME PART FROM SQL IF AVAILABLE
	sSql = "SELECT DISTINCT status FROM egov_facilityschedule WHERE facilityid =  " & iFacilityId 
	sSql = sSql & " AND facilitytimepartid = " & itimepartid & " AND checkindate = '" & datDate & "' AND status <> 'CANCELLED'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
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
' FUNCTION GETCHECKINTIME(ITIMEPARTID)
'--------------------------------------------------------------------------------------------------
Function GetCheckInTime( ByVal itimepartid )
	Dim sSql, oRs, sReturnValue

	sReturnValue= "UNKNOWN"

	sSql = "SELECT beginhour, beginampm FROM egov_facilitytimepart WHERE facilitytimepartid = " & itimepartid
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
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

	sReturnValue= "UNKNOWN"

	sSql = "SELECT endhour,endampm FROM egov_facilitytimepart WHERE facilitytimepartid = " & itimepartid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		sReturnValue = oRs("endhour") & ":" &  oRs("endampm")
	End If

	oRs.Close
	Set oRs = Nothing
	
	' RETURN STATUS
	GetCheckOutTime = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETFACILITYRATE(IRATEID)
'--------------------------------------------------------------------------------------------------
Function GetFacilityRate( ByVal irateid, ByVal iFacilityId )
	Dim sSql, oRs
	
	sUserType = GetUserResidentType(request.cookies("userid"))
	'response.write "sUserType: " & sUserType & "<br />"
	
	'sSql = "SELECT amount, pricetype FROM egov_facility_rate_to_pricetype R, dbo.egov_price_types P WHERE R.pricetypeid = P.pricetypeid AND R.rateid = " & irateid
	sSql = "SELECT R.amount, P.pricetype FROM egov_facility_rate_to_pricetype R, dbo.egov_price_types P, egov_facility F "
	sSql = sSql & "WHERE R.pricetypeid = P.pricetypeid AND P.pricetypegroupid = F.pricetypegroupid AND F.facilityid = " & iFacilityId
	sSql = sSql & " AND R.rateid = " & irateid & " ORDER BY P.displayorder"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	Do While Not oRs.EOF 
		If sUserType = oRs("pricetype") OR oRs("pricetype") = "E" Then
			sReturnValue = oRs("amount")
			Exit Do
		End If
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing 

	GetFacilityRate = sReturnValue

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetUserResidentType(iUserId)
'--------------------------------------------------------------------------------------------------
Function GetUserResidentType( ByVal iUserId )
	Dim oCmd

	If iUserid = "" Then
		GetUserResidentType = ""
	Else
		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
			.ActiveConnection = Application("DSN")
		    .CommandText = "GetUserResidentType"
		    .CommandType = 4
			.Parameters.Append oCmd.CreateParameter("@iUserid", 3, 1, 4, iUserId)
			.Parameters.Append oCmd.CreateParameter("@ResidentType", 129, 2, 1)
		    .Execute
		End With
		
		GetUserResidentType = oCmd.Parameters("@ResidentType").Value

		Set oCmd = Nothing

		' if you are not a  resident then you get non-resident pricing
		If IsNull(GetUserResidentType) Or GetUserResidentType = "" Or GetUserResidentType <> "R" Then
			GetUserResidentType = "N"
		End if
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' showFacilityName iFacilityId
'--------------------------------------------------------------------------------------------------
Sub showFacilityName( ByVal iFacilityId )
	Dim sSql, oRs

	sSql = "SELECT facilityname FROM egov_facility WHERE facilityid = " & iFacilityId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		response.write "<strong>" &  oRs("facilityname") & "</strong>"
	End If

	oRs.Close
	Set oRs = Nothing

End Sub 



%>

