<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: facility_reservation.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/17/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/06/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearchName, sResults, iUserId, reserveTotal, amountTotal

sLevel = "../" ' Override of value from common.asp

reserveTotal = CDbl(0)
amountTotal = CDbl(0)

If Not UserHasPermission( Session("UserId"), "reservations" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

If request("userid") <> "" Then
	iUserId = request("userid")
Else
	iUserId = GetFirstUserId()
End If

' See if a search term was passed
If request("searchname") <> "" Then 
	sSearchName = request("searchname")
Else
	sSearchName = ""
End If 

If request("results") <> "" Then
	sResults = request("results")
Else
	sResults = ""
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

	<script src="../scripts/easyform.js"></script>
	<script src="../scripts/ajaxLib.js"></script>

	<script>
	<!--
		function doCalendar(sfieldname) {
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?fn=' + sfieldname + '&p=1", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function reloadpage()
		{
			var datdate = '<%=request("D")%>';
			var iFacility = document.frmAvail.selfacility.options[document.frmAvail.selfacility.selectedIndex].value;
			location.href='facility_reservation.asp?Y=' + <%=Year(request("D"))%> + '&M=' + <%=Month(request("D"))%> + '&L=' + iFacility + '&D=' + datdate;
		}

		// TIMEPART ARRAY INFORMATION
		<% 	BuildJavascriptTimePartArray request("L"),Year(request("D")),Month(request("D")),Day(request("D")) %>

		function test(opt,itimepartcount) 
		{
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
		   if (parseInt(opt.length) > 0) 
		   {
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
		   //frmAvail.reservetotal.value =formatMoney(getCheckedValue(frmAvail.rcost));
		   $("#reservetotal").val( formatMoney( $("input:radio[name='rcost']:checked").val() ));

		   // TOTAL COST
		   //frmAvail.amounttotal.value = formatMoney(getCheckedValue(frmAvail.rcost));
		   $("#amounttotal").val( formatMoney( $("input:radio[name='rcost']:checked").val() ));

		}

		function recalculatetotal()
		{
			// RESERVATION TIME TOTAL
			$("#reservetotal").val( formatMoney( $("input:radio[name='rcost']:checked").val() ));


			// TOTAL COST
			$("#amounttotal").val( formatMoney( $("input:radio[name='rcost']:checked").val() ));

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
				num = num.substring(0,num.length-(4*i+3))+','+ num.substring(num.length-(4*i+3));
			return (((sign)?'':'-') + '$' + num + '.' + cents);
		}

		function formatMoney(num) 
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
				num = num.substring(0,num.length-(4*i+3)) + num.substring(num.length-(4*i+3));
			return (((sign)?'':'-') + num + '.' + cents);
			
			//num = num.substring(0,num.length-(4*i+3))+','+ num.substring(num.length-(4*i+3));
		}

		function process_form()
		{
			// CHECK FOR ERRORS

			if ( $("#amounttotal").val() == "" ) {
				alert("The total amount must contain a value.");
				$("#amounttotal").focus();
				return false;
			}
			else {
				// check for numeric values
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec($("#amounttotal").val());
				if ( Ok ) {
					// check that it is > 0
					var totalAmount = parseFloat($("#amounttotal").val());
					if ( totalAmount < 0) {
						alert("The total amount must be 0 or greater.");
						$("#amounttotal").focus();
						return false;
					}
				}
				else {
					alert("The total amount must contain a numeric value.");
					$("#amounttotal").focus();
					return false;
				}
			}


			// If recurrent dates are selected, check that there is an end date
			if (document.frmAvail.isrecursive.checked)
			{
				// CHECK RECURRENCE ENDDATE
				if (document.frmAvail.recurrentenddate.value == "")
				{
					alert('You have selected to make this a recurrent reservation. \nPlease enter a recurrent end date to continue.');
					document.frmAvail.recurrentenddate.focus();
					return;
				}

				// CHECK WEEKLY FREQUENYCY
				if (document.frmAvail.wrecurrenttimepart[1].checked)
				{
					if (trim(document.frmAvail.wfrequencynumber.value) == ''){
						alert('Weekly frequency is required but was not entered!');
						document.frmAvail.wfrequencynumber.focus();
						return;
					}
				}


				// CHECK MONTHLY FREQUENCY 
				if (document.frmAvail.wrecurrenttimepart[2].checked)
				{
					if (trim(document.frmAvail.mfrequencynumber.value) == ''){
						alert('Monthly frequency is required but was not entered!');
						document.frmAvail.mfrequencynumber.focus();
						return;
					}
				}
			}

			// UPDATE VALUES - already has the name in it
			//document.frmAvail.LodgeName.value = document.frmAvail.selfacility.options[document.frmAvail.selfacility.selectedIndex].text;

			// SUBMIT FORM
			frmAvail.submit();
		}

		function trim(str)
		{
		   return str.replace(/^\s*|\s*$/g,'');
		}

		function doUserPickFetch()
		{
			if ($("#searchname").val() != "")
			{
				// Try to get a drop down of citizen names
				doAjax('getcitizenpicks.asp', 'searchname=' + $("#searchname").val(), 'UpdateApplicants', 'get', '0');
			}
			else
			{
				$("#applicant").html("<input type='hidden' name='userid' id='userid' value='0' />");
				$("input.userbutton").hide();
			}
		}

		function UpdateApplicants( sResult )
		{
			//alert("Back");
			$("#applicant").html(sResult);
			if (sResult.substr(0,6) == "<label")
			{
				$("input.userbutton").show();
				// ajax call to get pricing
				//doAjax('getcitizenprice.asp', 'userid=' + $("#userid").val() + '&facilityid=' + $("#L").val() , 'setPricePicks', 'get', '0');
				setPricePicks( );
			}
			else
				$("input.userbutton").hide();
		}

		function changeUserPricePick()
		{
			alert($("#userid").val());
			// ajax call to get pricing
			//doAjax('getcitizenprice.asp', 'userid=' + $("#userid").val() + '&facilityid=' + $("#L").val() , 'setPricePicks', 'get', '0');
		}

		function setPricePicks( ) {
			
			
			var rateCount = parseInt($("#ratecount").val());
			if ( rateCount > 1 )
			{
				//alert($("#userid").val());
				var selectedUser = $("#userid option:selected").text();
				//alert(selectedUser);
				var startAt = selectedUser.indexOf("(") + 1;
				var endAt = selectedUser.indexOf(")");
				//var length = endAt - startAt;
				var userResidency = selectedUser.substring( startAt, endAt).toLowerCase().replace(/ /g, '');
				//alert(userResidency);
				var radiobtn;

				for ( var i=1; i <= rateCount; i++)
				{
					// do something here
					if ( userResidency === $("#pricetype" + i).val()) {
						console.log($("#pricetype" + i).val());
						radiobtn = document.getElementById("rate" + i);
						radiobtn.checked = true;
						$("#amounttotal").val($("input[name='rcost']:checked").val());
						$("#reservetotal").val($("#amounttotal").val());
					}
				}
			}
		}

		function SearchCitizens( )
		{
			var iSearchStart = $("#searchstart").val();
			var optiontext;
			var optionchanged;
			//alert(document.frmAvail.searchname.value);
			var searchtext = $("#searchname").val();
			var searchchanged = searchtext.toLowerCase();
			
			iSearchStart = parseInt(iSearchStart) + 1;

			for (x=iSearchStart; x < document.frmAvail.userid.length ; x++)
			{
				optiontext = document.frmAvail.userid.options[x].text;
				optionchanged = optiontext.toLowerCase();
				if (optionchanged.indexOf(searchchanged) != -1)
				{
					document.frmAvail.userid.selectedIndex = x;
					document.frmAvail.results.value = 'Possible Match Found.';
					$('#searchresults').html('Possible Match Found.');
					document.frmAvail.searchstart.value = x;
					//$("#userid").change();
					return;
				}
			}
			document.frmAvail.results.value = 'No Match Found.';
			$('#searchresults').html('No Match Found.');
			document.frmAvail.searchstart.value = -1;
		}

		function ClearSearch()
		{
			document.frmAvail.searchstart.value = -1;
		}

		function UserPick()
		{
			$("#searchname").val("");
			$("#results").val('');
			$("#searchresults").html('');
			$("#searchstart").val(-1);
			location.href='facility_reservation.asp?L=<%=request("L")%>&TP=<%=request("TP")%>&D=<%=request("D")%>&userid=' + $("#userid").val();
		}

		function checkrecursive()
		{
			// If recursive is unchecked, then turn off the radio buttons, blank out the end date.
			if (document.frmAvail.isrecursive.checked == false)
			{
				//alert('unchecked');
				document.frmAvail.recurrentenddate.value = "";
				for (i = 0; i < document.frmAvail.wrecurrenttimepart.length; i++)
				{
					if (document.frmAvail.wrecurrenttimepart[i].checked == true)
					{
						document.frmAvail.wrecurrenttimepart[i].checked = false;
						break; 
					}
				}
			}
		}

		function getCheckedValue(radioObj) 
		{
			if(!radioObj)
				return "";
			var radioLength = radioObj.length;
			if(radioLength == undefined)
				if(radioObj.checked)
					return radioObj.value;
				else
					return "";
			for(var i = 0; i < radioLength; i++) {
				if(radioObj[i].checked) {
					return radioObj[i].value;
				}
			}
			return "";
		}



	//-->
	</script>

</head>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div style="padding:20px;">

<%
' CAPTURE CURRENT PATH
Session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString() '& "&userid=" & request("userid") & "&searchname=" & request("searchname") & "&results=" & request("results") & "&searchstart=" & request("searchstart")
Session("RedirectLang") = "Return to Facility Reservation"

Dim sFirstName, sLastName, sAddress, sCity, sState, sZip, sEmail, sHomePhone, sWorkPhone, sFax, sBusinessName

SetUserInformation sFirstName, sLastName, sAddress, sCity, sState, sZip, sEmail, sHomePhone, sWorkPhone, sBusinessName, sFax

%>


<!--BEGIN PAGE CONTENT-->
<div class="returnnav">
	<a class="linkbutton" href="facility_calendar.asp?L=<%=request("L")%>&Y=<%=Year(request("D"))%>&M=<%=Month(request("D"))%>"><< Calendar</a>
</div>

<form name="frmAvail" action="facility_reserve_summary.asp" method="post">
	<input type="hidden" id="L" name="L" value="<%=request("L")%>" />
	<input type="hidden" name="LodgeName" value="<% =GetFacilityName( request("L") ) %>" />
	<input type="hidden" id="selTimePartID" name="selTimePartID" value="<%=request("TP")%>" />


	<!--BEGIN: USER MENU-->
	<div class="reserveformtitle">Contact Information</div>
	<div class="reserveforminputarea">
		<div class="reserveforminstructions"><strong>Instructions:</strong> Search for a name then select one from the resulting list. If their name is not on the list click New User to create their account. You can also review their current contact information by clicking View/Edit.</div>

		<!-- Citizen Search -->
		<div>
			<label for="searchname">Name Search: </label><input type="text" id="searchname" name="searchname" value="" size="25" maxlength="50" onkeypress="if(event.keyCode=='13'){doUserPickFetch();return false;}" />
			<input type="button" class="facilitybutton" value="Search" onclick="doUserPickFetch();" />
			<input type="hidden" id="results" name="results" value="<%=sResults%>" />
			<input type="hidden" id="searchstart" name="searchstart" value="-1" />
			<span id="searchresults"><%=sResults%></span>
			<!-- <br /><div id="searchtip">(last name, first name)</div> -->
		</div>
		<div>
			<span id="applicant">
				
				<input type="hidden" value="0" name="userid" id="userid" />
				<!-- <select id="userid" name="userid" onchange="UserPick();" > 
				<select id="userid" name="userid"> -->
					<% 'ShowUserDropDown iUserId %>
				<!-- </select>-->
			</span>
			 &nbsp; 
			<input class="facilitybutton userbutton" type="button" value="Edit/View" onClick="location.href='../dirs/update_citizen.asp?userid=' + document.frmAvail.userid.options[document.frmAvail.userid.selectedIndex].value;" /> &nbsp; <input onClick="location.href='../dirs/register_citizen.asp';" class="facilitybutton userbutton" type="button" value="New User" />
		</div>
	</div>
	<!--END: USER MENU-->


	<!--SELECT FACILITY-->
	<div class="reserveformtitle">Facility</div>
	<div class="reserveforminputarea">
	<!--<p><font class="reserveforminstructions">Instructions: Select the facility for your reservation.</font></p>-->
	<p>
<% 
	ifacilityid = request("L")
	If ifacilityid = "" Then
		ifacilityid = GetFirstFacility()
	End If

	response.write "<strong>" & GetFacilityName( ifacilityid ) & "</strong>"

	datCheckInDate = request("D")
	datCheckOutDate = request("D") ' CHECK BASED ON TIMEPARTID
	itimepartid = request("TP")

	'DrawSelectFacility ifacilityid 
	response.write vbCrLf & "<input type=""hidden"" name=""selfacility"" value=""" & ifacilityid & """ />"

	Dim irateid
	%>
	<input type="hidden" name="checkindate"  value="<%=datCheckInDate%>">
	<input type="hidden" name="checkoutdate" value="<%=datCheckOutDate%>">
	<input type="hidden" name="backlink" value="L=<%=request("L")%>&Y=<%=Year(request("D"))%>&M=<%=Month(request("D"))%>">

	</p>
	</div>

	<!--SELECT DATES-->
	<div class="reserveformtitle">Select Date/Time</div>
	<div class="reserveforminputarea">
	<p><font class="reserveforminstructions">Instructions: Select the Check-In and Check-Out times for your reservation.</font></p>

	<p><strong>Selected Date:</strong> <%=WeekDayName(Weekday(request("D"))) & ", " & request("D")%></p>

	<!--DRAW AVAILABILITY-->
	<p><% Call DrawAvailability(ifacilityid,itimepartid,Year(request("D")),Month(request("D")),Day(request("D"))) %></P>

	<%
	If OrgHasDisplay( Session("OrgID"), "facility arrival message" ) Then
		response.write GetOrgDisplay( Session("OrgID"), "facility arrival message" )
	End If 
	%>

	<!--DRAW DATE/TIME SELECTION-->
	<table >
		<tr><td class="reservationformlabel">Exact Arrival Time:</td><td>
		<% sCheckInTime = GetCheckInTime(itimepartid)%>
		<% sCheckOutTime = GetCheckOutTime(itimepartid)%>
		<select name="checkintime">
		<%
		' 1 AM TO 11:30 AM
		blnIsAvailable = False
		'response.write sCheckOutTime
		'response.end

		If trim(sCheckOutTime) = "1:00:AM" Then
			sRealTime = "1:00:AM"
			sCheckOutTime = "12:30:AM"
		End If

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
				response.write "<option " & sSelected & " >" & sTime & "</option>"
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
			response.write "<option  " & sSelected & " >12:00:PM" & "</option>"
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
			response.write "<option  " & sSelected & " >12:30:PM" & "</option>"
		End If

		' CHECK TO SEE IF WE HAVE PASSED CHECKOUT TIME
		If trim(sCheckOutTime) = trim("12:30:PM") Then
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
				response.write "<option  " & sSelected & ">" & sTime & "</option>"
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
				response.write "<option  " & sSelected & ">" & sTime & "</option>"
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
			response.write "<option  " & sSelected & " >12:00:AM" & "</option>"
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
			response.write "<option  " & sSelected & " >12:30:AM" & "</option>"
		End If

		' CHECK TO SEE IF WE HAVE PASSED CHECKOUT TIME
		If trim(sCheckOutTime) = trim("12:30:AM") Then
			' TURN OFF WRITING AVAILABLE TIMES
			blnIsAvailable = False
		End If


		
		If sRealTime = "1:00:AM" Then
			response.write "<option  " & sSelected & " >1:00:AM" & "</option>"
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
				response.write "<option " & sSelected & " >" & sTime & "</option>"
			End If 

			
			' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
			If blnFoundEndTime = True Then
				blnIsAvailable = False
			End If


			sTime = i  & ":30:AM"
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
				response.write "<option " & sSelected & " >" & sTime & "</option>"
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
			response.write "<option  " & sSelected & " >12:00:PM" & "</option>"
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
			response.write "<option  " & sSelected & " >12:30:PM" & "</option>"
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
				response.write "<option  " & sSelected & ">" & sTime & "</option>"
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
				response.write "<option  " & sSelected & ">" & sTime & "</option>"
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
			response.write "<option  " & sSelected & " >12:00:AM" & "</option>"
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
			sSelected = " selected=""selected"" "
			blnIsAvailable = TRUE
		Else
			sSelected = ""
		End If

		' DISPLAY ONSCREEN
		If blnIsAvailable = True Then
			response.write "<option  " & sSelected & " >12:30:AM" & "</option>"
		End If

		' IF FOUND CHECKOUT TIME STOP DISPLAYING AVAILABLE TIMES
		If blnFoundEndTime = True Then
			blnIsAvailable = False
		End If

		If sRealTime = "1:00:AM" Then
			response.write "<option  " & sSelected & " >1:00:AM" & "</option>"
		End If
		
		%>
	</select>

	</td></tr>
	<tr><td class="reservationformlabel" valign="top">Reservation Cost:</td>

	<td valign="top">
<% 
		reserveTotal = DrawFacilityRates( irateid, ifacilityid )
		amountTotal = amountTotal + reserveTotal
%>
		<br>
		<!--END: AVAILABLE PRICING-->
	
		<input id="reservetotal" name="reservetotal" style="text-align:right;" type="text" value="<%= reserveTotal %>" />
	
	</td></tr>
	</table>
	
	<!--KEY CHARGE-->
	<table>
		<tr><td>
			<!--<input  name="keydeposit" type=checkbox checked onClick="calculatetotal(frmAvail.timeparts);">Pay key deposit charge online now.<br>-->
			<!-- <b>Important!</b> <font color=red> <i>You will be required to pay a key deposit charge equal to the cost of the reservation or $100.00, whichever is greater at the time you pick up the key from the City Hall.</i> </font>-->
			<%= GetOrgDisplay( session("orgid"), "facility deposit message" ) %>
		</td></tr>
	</table>

	<!--TOTAL COST-->
	<table>
		<tr>
			<td class="reservationformlabel" >Total Amount:</td>
			<td><input id="amounttotal" name="amounttotal" style="text-align:right;" type="text" value="<%= FormatNumber(amountTotal,2,,,0) %>" /></td>
		</tr>
	</table>



	</p>
	</div>


	<!--BEGIN: Schedule Recurrent Time -->
	<div class="reserveformtitle">Schedule Recurrent Time </div>
	<div class="reserveforminputarea">
	<p>To make a reservation recurrent: check the <strong>Reservation is Recurrent</strong> checkbox below, enter an <strong>End By</strong> 
	date, then select from the radio buttons, that follow, to set how the reservation is to repeat.</p>
	<p>To remove the recurrent options, uncheck the <strong>Reservation is Recurrent</strong> checkbox.</p>
	<% DisplayRecurrentOptions %>
	</div>
	<!--END: SCHEDULE RECURRENT TIME  -->


	<!--RESERVATION PURPOSE/EVENT INFORMATION-->
	<div class="reserveformtitle">Reservation Purpose/Event Information</div>
	<div class="reserveforminputarea">
	<p><font class="reserveforminstructions">Please provide the following to allow us to better serve your needs.</font></p>
	<% DrawFacilityFields( ifacilityid ) %>


	<!--BEGIN: ADMIN COMMENT FIELD-->
	<P><b>City Internal Note (512 max number of characters):</b><br><textarea maxlength="1024" id="internalnote" name="internalnote"></textarea></p>
	<!--END: AMDIN COMMENT FIELD-->




	<!--WAIVERS-->
	<% DrawWaivers ifacilityid %>
	</div>


	<!--CONTINUE BUTTON-->
	<b>* You will be presented with summary of reservation costs and terms/conditions before the payment will be processed.<br /></b> <br />
	<input value="CONTINUE WITH RESERVATION" class="facilitybutton" type="button" onclick="if (validateForm('frmAvail')) {process_form();}" />

</form>
<!--END: PAGE CONTENT-->


<!--SPACING CODE-->
<p><br>&nbsp;<br>&nbsp;</p>
<!--SPACING CODE-->

<!--UPDATE COSTS-->
<script>
	calculatetotal( frmAvail.timeparts );
</script>

</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>



<%
'--------------------------------------------------------------------------------------------------
'  DRAWSELECTFACILITY
'--------------------------------------------------------------------------------------------------
Sub DrawSelectFacility( ByVal ifacilityid )
	Dim sSql, oRs
	
	If ifacilityid = "" Then
		ifacilityid = 0
	End If

	' GET SELECT CATEGORY ROW
	sSql = "SELECT * FROM egov_facility WHERE isviewable = 1 and orgid = " & session("orgid") & " ORDER BY facilityname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 3, 1

    ' LOOP THRU LIST OF AVAILABLE FACILITIES AND DISPLAY TO USER
    Response.Write("<font class=""reservationformlabel"">Facility:</font> <select class=""reservationformselect"" onChange=""reloadpage();"" name=""selfacility"" class=""facilitylist>"">")
    Do While Not oRs.EOF
		sSelected = ""

		If clng(ifacilityid) = clng(oRs("facilityid")) Then
			sSelected = "SELECTED"
		End If
		
		Response.Write("<option " & sSelected & " value=""" & oRs("facilityid") & """>" & oRs("facilityname") & "</option>" & vbCrLf)
		oRs.MoveNext
	Loop
    Response.Write("</select>" & vbCrLf)

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
'  DRAWAVAILABILITY(IFACILITYID)
'--------------------------------------------------------------------------------------------------
Sub DrawAvailability( ByVal ifacilityid, ByVal itempartid, ByVal iYear, ByVal iMonth, ByVal iDay)
	Dim sSql, oRs, iTimePartCount

	sSql = "SELECT facilityid, rateid, facilitytimepartid, beginhour, beginampm, endhour, endampm, weekday, description, rate "
	sSql = sSql & "FROM egov_facilitytimepart "
	sSql = sSql & "WHERE facilityid = " & ifacilityid & " AND weekday = '" & Weekday( iMonth & "/" & iDay & "/" & iYear ) &"' "
	sSql = sSql & "ORDER BY weekday, description, beginampm, beginhour"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	iTimePartCount = 0

	If Not oRs.EOF Then
		
		response.write "<fieldset style=""padding:5px;"">"
		response.write "<legend><b>Available Time(s)</b></legend>"

		Do While Not oRs.EOF 
			sChecked = ""
			If clng(itimepartid) = clng(oRs("facilitytimepartid")) Then
				sChecked = " checked=""checked"" "
				irateid = oRs("rateid")
			End If
			
			sTimeRange = " " & oRs("beginhour") & " " & oRs("beginampm") & "-" & oRs("endhour") & " " & oRs("endampm") & " - " & oRs("description") & GetTimePartStatusName(iFacilityid,itimepartid,sTimeRange,iTimePartCount)
			response.write "<input type=""checkbox"" name=""timeparts"" value=""" & iTimePartCount & """   " & sChecked & "  class=""reserveformcheckbox"" style=""" & GetTimePartStatusColor(iFacilityid,itimepartid,sTimeRange,iTimePartCount) &""" onClick=""test(this.form.timeparts,this.value);"" > " & sTimeRange & "<br>"
			iTimePartCount = iTimePartCount  + 1
			oRs.MoveNext
		Loop

		response.write "</fieldset>"

	End If

	oRs.Close
	Set oRs = Nothing 
		
End Sub


'--------------------------------------------------------------------------------------------------
'  BuildJavascriptTimePartArray(ifacilityid,iYear,iMonth,iDay)
'--------------------------------------------------------------------------------------------------
Sub BuildJavascriptTimePartArray( ByVal ifacilityid, ByVal iYear, ByVal iMonth, ByVal iDay )
	Dim sSql, oRs

	sSql = "Select rate,facilityid, rateid, facilitytimepartid, beginhour, beginampm, endhour, endampm, weekday,description,rate from egov_facilitytimepart where facilityid = '" & ifacilityid & "' and weekday = '" & weekday( iMonth & "/" & iDay & "/" & iYear ) &"' order by weekday, description,beginampm, beginhour"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 3

	iArrayCount = 0

	If Not oRs.EOF Then
	
		response.write "var timeparts = new Array(" & oRs.recordcount - 1 & ");" & vbcrlf
		
		Do While Not oRs.EOF 
			arrStart = split(oRs("beginhour"),":")
			arrEnd = split(oRs("endhour"),":")
			sStartHour = GetMilitaryTime(clng(arrStart(0)),clng(arrEnd(0)),oRs("beginampm"),oRs("endampm"),0)
			sStartMinute = clng(arrStart(1))
			sEndHour = GetMilitaryTime(clng(arrStart(0)),clng(arrEnd(0)),oRs("beginampm"),oRs("endampm"),1)
			sEndMinute = clng(arrEnd(1))
			response.write "timeparts[" & iArrayCount & "] = new Array(4);" & vbcrlf
			response.write "timeparts[" & iArrayCount & "][0] = new Date(" & iYear &"," & iMonth & "," & iDay & "," & sStartHour & "," & sStartMinute & ",0);" & vbcrlf
			response.write "timeparts[" & iArrayCount & "][1] = new Date(" & iYear &"," & iMonth & "," & iDay & "," & sEndHour & "," & sEndMinute & ",0);" & vbcrlf
			response.write "timeparts[" & iArrayCount & "][2] = '" & GetFacilityRate(oRs("rateid")) & "';" & vbcrlf
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
	If (iTempHour < 12) AND (UCASE(sTempAM)="AM") Then
		iReturnValue = iTempHour
	Else
		iReturnValue = iTempHour + 12
	End If

	' NOON
	If iTempHour = 12 and (UCASE(sTempAM)="PM") Then
		iReturnValue = 12
	End If 

	' MIDNIGHT
	If iTempHour = 12 and (UCASE(sTempAM)="AM") Then
		iReturnValue = 0
	End If 

	' SEE IF END TIME CROSSES MIDNIGHT
	If ihour > iEndHour and UCASE(sBeginAMPM)="AM" AND UCASE(sEndAMPM)="AM" and iStartorEnd = 1  Then
		iReturnValue = iTempHour + 24	
	End If


	' RETURN VALUE 
	GetMilitaryTime = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETTIMEPARTSTATUS(IFACILITYID,ITIMEPARTID,STIMERANGE)
'--------------------------------------------------------------------------------------------------
Function GetTimePartStatusColor( ByVal iFacilityid, ByVal itimepartid, ByVal sTimeRange, ByVal iDayofWeek )

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
' DrawWaivers ifacilityid
'--------------------------------------------------------------------------------------------------
Sub DrawWaivers( ByVal ifacilityid )
	Dim sSql, oRs 

	sSql = "SELECT * FROM egov_facilitywaivers INNER JOIN egov_waivers ON egov_facilitywaivers.waiverid = egov_waivers.waiverid WHERE egov_facilitywaivers.facilityid = " & ifacilityid & " ORDER BY isrequired, name"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	iCount = 0
	If Not oRs.EOF Then
		iCount = iCount + 1
		response.write "<p><b>Optional Waivers:</b><br />"

		Do While Not oRs.EOF 
			If oRs("isrequired") Then
				response.write "<input type=hidden name=""chkwaivers_" & iCount & """ value=""" & oRs("waiverid") & """>"
			Else
				response.write "<input name=""chkwaivers_" & iCount & """ type=checkbox value=""" & oRs("waiverid") & """>" & oRs("description") & " <B><small>(OPTIONAL)</small></b><br>"
			End If

			oRs.MoveNext
		Loop

	End If

	oRs.Close
	Set oRs = Nothing

End Sub

'--------------------------------------------------------------------------------------------------
'  DRAWFACILITYFIELDS(IFACILITYID)
'--------------------------------------------------------------------------------------------------
Sub DrawFacilityFields( ByVal iFacilityID )
	Dim sSql, oRs, sHeight

	sSql = "SELECT * FROM egov_facility_fields WHERE facilityid = " & ifacilityID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		
		response.write "<p>"
		response.write "<table>"

		Do While Not oRs.EOF
			' CHECK IF IT IS REQUIRED
			If  oRs("isrequired") Then
				sRequired = "<SPAN style=""COLOR: #ff0000"">*&nbsp;</SPAN>"
				response.write "<input type=""hidden"" name=""ef:custom_" & oRs("fieldname") & "_" & oRs("fieldid") & "-text/" & oRs("validation") & """ value=""" & oRs("fieldprompt") & """>"
			Else
				sRequired = " "
			End If

			' SET HEIGHT FOR INPUT BOX BASED ON FIELD TYPE, 1=STANDARD, 2=SIMULATED TEXT AREA
			If  oRs("fieldtype") = 2 Then
				sHeight = "HEIGHT: 100px; "
			Else
				sHeight = "" 
			End If
		
			response.write "<tr bgColor=""#e0e0e0"">"
			response.write "<td valign=""top"" class=""reservationformlabel"" align=""right"">" & sRequired
			response.write oRs("fieldprompt")
			response.write ": </td>"

			response.write "<td style=""font-family:Arial; font-size:8pt; color:#000000"" align=""left"">"
			
			Select Case oRs("fieldtype")

				Case "1"
					' TEXT BOX
					response.write "<input name=""custom_" & oRs("fieldname") & "_" & oRs("fieldid") & """ type=""text"" style=""FONT-SIZE: 8pt; WIDTH: 300px; " & sHeight & " FONT-FAMILY: Arial"" />"
				Case "2"
					' PSEUDO TEXT AREA
					'response.write "<input name=""custom_" & oRs("fieldname") & "_" & oRs("fieldid") & """ type=""text"" style=""FONT-SIZE: 8pt; WIDTH: 300px; " & sHeight & " FONT-FAMILY: Arial"" >"
					response.write "<textarea name=""custom_" & oRs("fieldname") & "_" & oRs("fieldid") & """ style=""FONT-SIZE: 8pt; WIDTH: 100%; " & sHeight & "FONT-FAMILY: Arial""></textarea>"
				Case "3"
					' SELECT BOX
					arrAnswers = split(oRs("fieldchoices"),"@@")
			
					response.write "<select name=""custom_" & oRs("fieldname") & "_" & oRs("fieldid") & """ >"
					For alist = 0 to ubound(arrAnswers)
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
Sub SetUserInformation( ByRef sFirstName, ByRef sLastName, ByRef sAddress, ByRef sCity, ByRef sState, ByRef sZip, ByRef sEmail, ByRef sHomePhone, ByRef sWorkPhone, ByRef sBusinessName, ByRef sFax )
	Dim sSql, oRs

	If sOrgRegistration Then 
		If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then
			
			iUserID = request.cookies("userid")
		
			sSql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, ISNULL(useraddress,'') AS useraddress, "
			sSql = sSql & "ISNULL(usercity,'') AS usercity, ISNULL(userstate,'') AS userstate, ISNULL(userzip,'') AS userzip, "
			sSql = sSql & "ISNULL(useremail,'') AS useremail, ISNULL(userhomephone,'') AS userhomephone, ISNULL(userworkphone,'') AS userworkphone, "
			sSql = sSql & "ISNULL(userbusinessname,'') AS userbusinessname, ISNULL(userfax,'') AS userfax FROM egov_users WHERE userid = " & iUserID

			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open sSql, Application("DSN"), 0, 1

			If NOT oRs.EOF Then
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
' Function ShowUserInfo( iUserId, sUserType, sResidentDesc )
'--------------------------------------------------------------------------------------------------
Function ShowUserInfo( ByVal iUserId, ByVal sUserType, ByVal sResidentDesc )
	Dim sSql, oRs
	ShowUserInfo = ""

	sSql = "SELECT userfname, userlname, useraddress, useraddress2, usercity, userstate, userzip, usercountry, useremail, userhomephone, "
	sSql = sSql & "userworkphone, userfax, userbusinessname, userpassword, userregistered, residenttype "
	sSql = sSql & "FROM egov_users WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	' sUserType = oRs("residenttype")

	'ShowUserInfo = "<table border=" & Chr(34) & "0" & Chr(34) & " cellpadding=" & Chr(34) & "5" & Chr(34) & " cellspacing=" & Chr(34) & "0" & Chr(34) & ">"
	ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Name:</td><td width=""60%"">" & oRs("userfname") & " " & oRs("userlname") & "&nbsp;&nbsp;&nbsp;<strong>" & sResidentDesc & "</strong></td></tr>"
	ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Email:</td><td>" & oRs("useremail") & "</td></tr>"
	ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Phone:</td><td>" & oRs("userhomephone") & "</td></tr>"
	ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Address:</td><td>" & oRs("useraddress") & "<br />" 
	If oRs("useraddress2") = "" Then 
		ShowUserInfo = ShowUserInfo & oRs("useraddress2") & "<br />" 
	End If 
	ShowUserInfo = ShowUserInfo & oRs("usercity") & ", " & oRs("userstate") & " " & oRs("userzip") & "</td></tr>"
	ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Business:</td><td>" & oRs("userbusinessname") & "</td></tr>"
	'ShowUserInfo = ShowUserInfo & "<tr><td>&nbsp;</td><td>" & oRs("usercity") & ", " & oRs("userstate") & " " & oRs("userzip") & "</p>"
	'ShowUserInfo = ShowUserInfo & "</table>"

	oRs.close
	Set oRs = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' ShowUserDropDown(iUserId)
'--------------------------------------------------------------------------------------------------
Sub ShowUserDropDown( ByVal iUserId )
	Dim sSql, oCmd, oRs

'	Set oCmd = Server.CreateObject("ADODB.Command")
'	With oCmd
'		.ActiveConnection = Application("DSN")
'	    .CommandText = "GetEgovUserWithAddressList"
'	    .CommandType = 4
'		.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, Session("OrgID"))
'	    Set oRs = .Execute
'	End With

	sSql = "SELECT U.userid, ISNULL(U.userfname,'') AS userfname, ISNULL(U.userlname,'') AS userlname, ISNULL(U.useraddress,'') AS useraddress, "
	sSql = sSql & "residenttypename = CASE WHEN R.description IS NULL THEN '' ELSE R.description END "
	sSql = sSql & "FROM egov_users U LEFT OUTER JOIN egov_poolpassresidenttypes R ON U.residenttype = R.resident_type AND U.orgid = R.orgid "
	sSql = sSql & "WHERE U.isdeleted = 0 AND U.userregistered = 1 AND U.headofhousehold = 1 "
	sSql = sSql & "AND U.userfname IS NOT NULL AND userlname IS NOT NULL "
	sSql = sSql & "AND U.userfname != '' AND userlname != '' AND U.orgid = " & Session("OrgID")
	sSql = sSql & " ORDER BY U.userlname, U.userfname, U.residenttype"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("userid") & """"
		If CLng(iUserId) = CLng(oRs("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("userlname") & ", " & oRs("userfname") 
		If oRs("residenttypename") <> "" Then
			response.write " (" & oRs("residenttypename") & ")"
		End If 
		response.write " &ndash; " & oRs("useraddress") & "</option>"

		oRs.MoveNext
	Loop 
		
	oRs.Close
	Set oRs = Nothing
'	Set oCmd = Nothing

End Sub  


'--------------------------------------------------------------------------------------------------
' GetFirstUserId( )
'--------------------------------------------------------------------------------------------------
Function GetFirstUserId()
	Dim sSql, oRs

	sSql = "SELECT  TOP 1 userid FROM egov_users WHERE orgid = " & Session("OrgID") 
	sSql = sSql & " ORDER BY userlname, userfname, userid"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	GetFirstUserId = oRs("userid")

	oRs.close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' DisplayRecurrentOptions
'--------------------------------------------------------------------------------------------------
Sub DisplayRecurrentOptions

	' DISPLAY ON/OFF RECURRENCE
	response.write "<input type=""checkbox"" name=""isrecursive"" onclick='checkrecursive();'><strong>Reservation is Recurrent.</strong> &nbsp; &nbsp;  End By: <input type=""text"" name=""recurrentenddate""> <a href=""javascript:doCalendar('recurrentenddate');"" title=""Select from calendar""><img border=""0"" src=""../images/calsmall.gif"" alt=""Select from calendar"" /></a>"

	' DAILY OPTIONS
	response.write "<hr />"
	response.write "<strong>Reservation Repeats: </strong><br /><br />"
	response.write "<input onClick=""document.frmAvail.isrecursive.checked = true;"" type=""radio"" name=""wrecurrenttimepart"" value=""daily""><strong>Daily</strong> (Uses dates and times above as starting point.)<br />"
	response.write "Every day ending on above End By date."

	' WEEKLY OPTIONS
	response.write "<hr />"
	response.write "<input onClick=""document.frmAvail.isrecursive.checked = true;"" type=""radio"" name=""wrecurrenttimepart"" value=""weekly""><strong>Weekly</strong> (Uses dates and times above as starting point.)<br />"
	response.write "Every <input type=""text"" name=""wfrequencynumber"" value=""1"" style=""text-align:right;"" size=""3""> week(s) starting on: "
	response.write "<select name=""wdayofweek"">"
	For d = 1 to 7
		response.write "<option value=""" & d & """"
		If Weekday(request("D")) = d Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & WeekDayName(d) & " </option>"
	Next
	response.write "</select>"

	' MONTHLY OPTIONS
	response.write "<hr />"
	response.write "<input onClick=""document.frmAvail.isrecursive.checked = true;"" type=""radio"" name=""wrecurrenttimepart"" value=""monthly""><strong>Monthly</strong> (Uses dates and times above as starting point)<br />"
	response.write "The "
	response.write "<select name=""mseries"">"
	response.write "<option value=""1"">First </option>"
	response.write "<option value=""2"">Second </option>"
	response.write "<option value=""3"">Third </option>"
	response.write "<option value=""4"">Fourth </option>"
	response.write "<option value=""5"">Last </option>"
	response.write "</select>"
	response.write "  "
	response.write "<select name=""mdayofweek"">"
	For d = 1 to 7
		response.write "<option value=""" & d & """"
		If Weekday(request("D")) = d Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & WeekDayName(d) & " </option>"
	Next
	response.write "</select>"
	response.write " of every <input type=""text"" name=""mfrequencynumber"" value=""1"" style=""text-align:right;"" size=""3""> Month(s) "

	' YEARLY OPTIONS
	response.write "<hr />"
	response.write "<input onClick=""document.frmAvail.isrecursive.checked = true;"" type=""radio"" name=""wrecurrenttimepart"" value=""yearly""><strong>Yearly</strong> (Uses dates and times above as starting point)<br />"
	response.write "The "
	response.write "<select name=""yseries"">"
	response.write "<option value=""1"">First </option>"
	response.write "<option value=""2"">Second </option>"
	response.write "<option value=""3"">Third </option>"
	response.write "<option value=""4"">Fourth </option>"
	response.write "<option value=""5"">Last </option>"
	response.write "</select>"
	response.write "  "
	response.write "<select name=""ydayofweek"">"
	For d = 1 to 7
		response.write "<option value=""" & d & """"
		If Weekday(request("D")) = d Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & WeekDayName(d) & " </option>"
	Next
	response.write "</select>"
	response.write " of every "
	response.write "<select name=""ymonth"">"
	For d = 1 to 12
		response.write "<option value=""" & d & """" 
		If Month(request("D")) = d Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & MonthName(d) & " </option>"
	Next
	response.write "</select>"

End Sub


'--------------------------------------------------------------------------------------------------
' GETCHECKINTIME(ITIMEPARTID)
'--------------------------------------------------------------------------------------------------
Function GetCheckInTime( ByVal itimepartid )
	Dim sSql, oRs

	sReturnValue= "UNKNOWN"

	sSql = "SELECT beginhour, beginampm FROM egov_facilitytimepart WHERE facilitytimepartid = " & itimepartid 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		sReturnValue = oRs("beginhour") & ":" &  oRs("beginampm")
	End If

	oRs.Close
	Set oRs = Nothing
	
	' RETURN STATUS
	GetCheckInTime = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' GETCHECKOUTTIME(ITIMEPARTID)
'--------------------------------------------------------------------------------------------------
Function GetCheckOutTime( ByVal itimepartid )
	Dim sSql, oRs

	sReturnValue= "UNKNOWN"

	sSql = "SELECT endhour,endampm FROM egov_facilitytimepart WHERE facilitytimepartid = " & itimepartid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		sReturnValue = oRs("endhour") & ":" &  oRs("endampm")
	End If

	oRs.Close
	Set oRs = Nothing
	
	' RETURN STATUS
	GetCheckOutTime = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETFIRSTFACILITY()
'--------------------------------------------------------------------------------------------------
Function GetFirstFacility()
	Dim sSql, oRs

	iReturnValue= "0"

	sSql = "SELECT TOP 1 * FROM egov_facility WHERE isviewable = 1 AND orgid = " & session("orgid") & " ORDER BY facilityname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		iReturnValue = oRs("facilityid") 
	End If

	oRs.Close
	Set oRs = Nothing
	
	' RETURN STATUS
	GetFirstFacility = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' string = GetFacilityName( iFacilityId )
'--------------------------------------------------------------------------------------------------
Function GetFacilityName( ByVal iFacilityId )
	Dim sSql, oRs

	sSql = "SELECT facilityname FROM egov_facility WHERE facilityid = " & CLng(iFacilityId) & " AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		GetFacilityName = oRs("facilityname") 
	Else
		GetFacilityName = ""
	End If

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' FUNCTION ISHOLIDAY(SDATE)
'--------------------------------------------------------------------------------------------------
Function IsHoliday( ByVal sDate)
	Dim sSql, oRs

	blnReturnValue = False

	sSql = "SELECT * FROM egov_holidays where orgid = " & session("orgid") & " AND holidaymonth = " & Month(sDate) & " AND holidayday = " & Day(sDate)

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		blnReturnValue = True
	End If

	oRs.Close
	Set oRs = Nothing

	 IsHoliday = blnReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETFACILITYRATE(IRATEID)
'--------------------------------------------------------------------------------------------------
Function GetFacilityRate( ByVal irateid)
	Dim sSql, sUserType, oRs
	
	sUserType = GetUserResidentType(request.cookies("userid"))
	
	sSql = "SELECT amount, pricetype FROM egov_facility_rate_to_pricetype INNER JOIN dbo.egov_price_types ON dbo.egov_facility_rate_to_pricetype.pricetypeid = dbo.egov_price_types.pricetypeid where rateid = '" & irateid & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		Do While Not oRs.EOF 
			
			If sUserType = oRs("pricetype") Or oRs("pricetype") = "E" Or sUserType = "B" Or sUserType = "E" Then
				sReturnValue = oRs("amount")
				Exit Do
			End If
			oRs.MoveNext
		Loop
	End If

	oRs.Close
	Set oRs = Nothing 

	GetFacilityRate = sReturnValue

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetUserResidentType(iUserId)
'--------------------------------------------------------------------------------------------------
Function GetUserResidentType( ByVal iUserId )
	Dim oType, sResType
	sResType = ""

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

		If IsNull(GetUserResidentType) Or GetUserResidentType = "" Then
			GetUserResidentType = "N"
		End if
	End If 

End Function 



'--------------------------------------------------------------------------------------------------
' DrawFacilityRates( iRateId, iFacilityId )
'--------------------------------------------------------------------------------------------------
Function DrawFacilityRates( ByVal iRateId, ByVal iFacilityId )
	Dim sChecked, sSql, oRs, sUserType, selectedAmount, bNeedsCheck, iRateCount

	amount = "0.00"
	bNeedsCheck = True 
	iRateCount = 0

	sSql = "SELECT P.pricetype, P.pricetypename, R.amount "
	sSql = sSql & "FROM egov_facility_rate_to_pricetype R, egov_price_types P, egov_facility F "
	sSql = sSql & "WHERE R.pricetypeid = P.pricetypeid AND P.pricetypegroupid = F.pricetypegroupid "
	sSql = sSql & "AND F.facilityid = " & iFacilityId & " AND R.rateid = " & iRateId
	sSql = sSql & " ORDER BY P.displayorder"

	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	'sUserType = GetUserResidentType( iUserId )
	'response.write "sUserType: " & sUserType & "<br />"

	' DISPLAYS RATES
	If Not oRs.EOF Then
		
		response.write "<table>"

		Do While Not oRs.EOF
			iRateCount = iRateCount + 1

			' check the first rate on page load
			If iRateCount = 1 Then 
				sChecked = " checked=""checked"""
				selectedAmount = FormatNumber(oRs("amount"),2,,,0)
			Else
				sChecked = ""
			End If 

''			If bNeedsCheck And ( sUserType = oRs("pricetype") Or oRs("pricetype") = "E" Or sUserType = "B" Or sUserType = "E" ) Then
''				sChecked = " checked=""checked"""
''				selectedAmount = FormatNumber(oRs("amount"), 2)
''				bNeedsCheck = False 
''			Else
''				sChecked = ""
''			End If

			response.write "<tr><td>" & oRs("pricetypename") & ":</td><td><input" & sChecked & " onClick=""recalculatetotal();"" type=""radio"" name=""rcost"" value=""" & FormatNumber(oRs("amount"),2,,,0) & """ id=""rate" & iRateCount & """ />" & FormatNumber(oRs("amount"),2,,,0) 
			response.write "<input type=""hidden"" id=""pricetype" & iRateCount & """ name=""pricetype" & iRateCount & """ value="""  & LCase(Replace(oRs("pricetypename"), " ","")) & """ />"
			response.write "</td></tr>"
			oRs.MoveNext
		Loop
		
		response.write "</table>"
		response.write vbcrlf & "<input type=""hidden"" id=""ratecount"" name=""ratecount"" value=""" & iRateCount & """ />"
	End If

	oRs.Close
	Set oRs = Nothing

	DrawFacilityRates = selectedAmount

End Function   



%>
