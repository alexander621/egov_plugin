<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentaltimeselection.asp
' AUTHOR: Steve Loar
' CREATED: 10/12/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Select the time, reservation type and if required, a renter.
'
' MODIFICATION HISTORY
' 1.0   10/12/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iReservationTempId, iRentalId, bHasData, bHasHours, sSelectedDate, bOffSeasonFlag, sTitle
Dim bIsAllDayOnly, sStartTimeLabel, sEndTimeLabel, sIsAllDay, sMessage, bInitial, sResidentType
Dim iStartHour, iStartMinute, sStartAmPm, iEndHour, iEndMinute, sEndAmPm, iCitizenUserId
Dim iArrivalHour, iArrivalMinute, sArrivalAmPm, iDepartureHour, iDepartureMinute, sDepartureAmPm
Dim iMaxCharges, iIncludePriceTypeId, sLatestStart, bOkToDisplay, dNonResidentStartDate, sLessThanDate
Dim bFutureProblem, iReservationTypeId, sNoCostPhrase, bHasUsers, bIsPublicUsers, iTotalRows

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "create simple reservations", sLevel	' In common.asp

iReservationTempId = CLng(request("rti"))

If request("pk") <> "" Then
	bInitial = True 
Else 
	bInitial = False 
End If 


iStartHour = "12"
iStartMinute = "00"
sStartAmPm = "AM"
iEndHour = "12"
iEndMinute = "00"
sEndAmPm = "AM"
bFutureProblem = False 

' still need to confirm that the data is there, and if not take them away from this page.
bHasData = SetPageVariables( iReservationTempId )

If bHasData = False  Then 
	' Take them somewhere safe, as their data is gone.
	response.redirect "rentalcategoryselection.asp"
End If 

If request("reservationtypeid") <> "" Then 
	iReservationTypeId = CLng(request("reservationtypeid"))
Else
	iReservationTypeId = GetFirstReservationTypeInList( )
End If

If RentalHasNoCosts( iRentalId ) Then
	sNoCostPhrase = "<strong>There is no cost to rent this.</strong>"
Else
	sNoCostPhrase = ""
End If 

If CLng(iReservationTypeId) > CLng(0) Then 
	GetUserFlagsFromReservationTypeId iReservationTypeId, bHasUsers, bIsPublicUsers
Else
	bHasUsers = False 
	bIsPublicUsers = False 
End If 

If bHasUsers Then 
	If request("rentaluserid") <> "" Then 
		iRentalUserid = CLng(request("rentaluserid"))
	Else
		iRentalUserid = CLng(0)
	End If 
Else
	iRentalUserid = CLng(0)
End If 

If request("searchname") <> "" Then
	sSearchName = request("searchname")
Else
	sSearchName = ""
End If 

sMessage = request("msg")
If sMessage = "short" Then
	sLoadMsg = "doShortConfirm();"
End If 
If sMessage = "buffer" Then
	sLoadMsg = "doBufferConfirm();"
End If
If sMessage = "buffershort" Then
	sLoadMsg = "doBufferShortConfirm();"
End If
If sMessage = "shortnoconfirm" Then
	sLoadMsg = "displayScreenMsg('Warning: The duration is for less than the allowed minimum time.');"
End If 
If sMessage = "buffernoconfirm" Then
	sLoadMsg = "displayScreenMsg('Warning: There is a conflict with the buffering between reservations.');"
End If 
If sMessage = "buffershortnoconfirm" Then
	sLoadMsg = "displayScreenMsg('Warning: The duration is less than allowed and there is a conflict with the buffering.');"
End If 
If sMessage = "conflict" Then
	sLoadMsg = "displayScreenMsg('There is a conflict with an existing reservation.');"
End If 
If sMessage = "closed" Then
	sLoadMsg = "displayScreenMsg('The rental is not open, or the time requested is beyond operating hours.');"
End If 
If sMessage = "nouser" Then
	sLoadMsg = "displayScreenMsg('This type of reservation requires the selection of a person to complete.');"
End If 
If sMessage = "OK" Then
	sLoadMsg = "displayScreenMsg('The selected time checks out fine for this reservation.');"
End If 


%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

	<script language="Javascript">
	<!--

		function checkTime()
		{
			document.frmRentalTime.submit();
		}

		function doShortConfirm()
		{
			if (confirm("You have selected a time interval that is less than the allowed minimum.\nDo you wish to continue?"))
			{
				document.frmDateSelection.submit();
			}
		}

		
		function doBufferConfirm()
		{
			if (confirm("There is a conflict with the buffering between reservations.\nDo you wish to continue?"))
			{
				document.frmDateSelection.submit();
			}
		}

		function doBufferShortConfirm()
		{
			if (confirm("The duration is less than allowed and there is a conflict with the buffering.\nDo you wish to continue?"))
			{
				document.frmDateSelection.submit();
			}
		}

		function displayScreenMsg( iMsg ) 
		{
			if(iMsg!="") 
			{
				$("screenMsg").innerHTML = iMsg;
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("screenMsg").innerHTML = "";
		}

		function goBack()
		{
			document.frmBack.submit();
		}

		function loader()
		{
			<%=sLoadMsg%>
		}

		function newUser()
		{
			//location.href='../dirs/register_citizen.asp';
			var myRand = parseInt(Math.random() * 99999999 );
			eval('window.open("../dirs/register_citizen.asp?rand=' + myRand + '", "_picker", "width=800,height=800,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=10,top=10")');
		}

		function EditApplicant()
		{
			var strPickedUserId = document.frmDateSelection.rentaluserid.options[document.frmDateSelection.rentaluserid.selectedIndex].value;
			var myRand = parseInt(Math.random() * 99999999 );
			eval('window.open("rentaluseredit.asp?userid=' + strPickedUserId + '&rand=' + myRand + '", "_picker", "width=800,height=800,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=10,top=10")');
		}

		function doNameSearchChange()
		{
			if ($("searchname").value != "")
			{
				//document.frmDateSelection.searchname.value = $("searchname").value;
				//if ($("rentaluserid").value != '')
				//{
				//	document.frmDateSelection.rentaluserid.value = $("rentaluserid").value;
				//}
				//document.frmDateSelection.reservationtypeid.value = $("reservationtypeid").value;
				doUserPickChange();
			}
			else
			{
				alert('Please enter a name before searching.');
				$("searchname").focus();
			}
		}

		function doUserPickChange()
		{
			var iReservationTypeId = $("reservationtypeid").value;
			//document.frmSearchReturn.reservationtypeid.value = $("reservationtypeid").value;
			if ($("searchname") != '' || $("rentaluserid") != '0')
			{
				// Fire off job to get the rental type 
			doAjax('getreservationtype.asp', 'reservationtypeid=' + iReservationTypeId , 'changePickers', 'get', '0');
			}
		}

		function changePickers( sReturn )
		{
			//alert( sReturn);
			
			if (sReturn == 'public')
			{
				if ($("searchname").value != "")
				{
					// Try to get a drop down of citizen names
					doAjax('getcitizenpicks.asp', 'searchname=' + $("searchname").value, 'UpdateApplicants', 'get', '0');
				}
				else
				{
					$("applicant").innerHTML = "<input type='hidden' name='rentaluserid' id='rentaluserid' value='0' />Search for a name then select one from the resulting list.";
					$("edituserbtn").style.visibility = 'hidden';
				}
			}
			else
			{
				if (sReturn == 'admin')
				{
					if ($("searchname").value != "")
					{
						// Try to get a drop down of citizen names
						doAjax('getadminpicks.asp', 'searchname=' + $("searchname").value, 'UpdateAdminApplicants', 'get', '0');
					}
					else
					{
						$("applicant").innerHTML = "<input type='hidden' name='rentaluserid' id='rentaluserid' value='0' />Search for a name then select one from the resulting list.";
						$("edituserbtn").style.visibility = 'hidden';
					}
				}
				else
				{
					// for anything else, blank out the picks and put things back to nothing 
					$("applicant").innerHTML = "<input type='hidden' name='rentaluserid' id='rentaluserid' value='0' />This reservation type does not need a renter.";
					$("edituserbtn").style.visibility = 'hidden';
				}
			}
		}

		function UpdateAdminApplicants( sResult )
		{
			//alert(sResult);
			$("applicant").innerHTML = sResult;
			$("edituserbtn").style.visibility = 'hidden';
			//document.frmSearchReturn.rentaluserid.value = $("rentaluserid").value;
		}

		function UpdateApplicants( sResult )
		{
			//alert("Back");
			$("applicant").innerHTML = sResult;
			if (sResult.substr(0,6) == "Select")
			{
				$("edituserbtn").style.visibility = 'visible';
				//document.frmSearchReturn.rentaluserid.value = $("rentaluserid").value;
			}
			else
				$("edituserbtn").style.visibility = 'hidden';
		}

		function CheckDates()
		{
			
			// Check that the end time is later than start time
			var dtStart = new Date($("startdate1").value + " " + $("starthour1").value + ":" + $("startminute1").value + " " + $("startampm1").value);
			var dtEnd;
			if ($("endday1").value == "0")
			{
				dtEnd = new Date($("startdate1").value + " " + $("endhour1").value + ":" + $("endminute1").value + " " + $("endampm1").value);
				//alert(dtEnd);
			}
			else
			{
				dtEnd = new Date($("startdate1").value + " " + $("endhour1").value + ":" + $("endminute1").value + " " + $("endampm1").value);
				dtEnd.setDate(dtEnd.getDate()+1);
				//alert(dtEnd);
			}
			var difference_in_milliseconds = dtEnd - dtStart;
			if (difference_in_milliseconds <= 0)
			{
				alert("One of the end times is not after the start time. Please correct this and try again.");
				$("endhour1").focus();
				return;
			}
			//alert("OK");
			//return;
			// Now bundle the dates and times and send off to check routine via AJAX
			var sParameter = 'rentalid=' + encodeURIComponent($("rentalid").value);
			sParameter += '&maxrows=1' + encodeURIComponent($("maxrows").value);
			sParameter += '&rti=' + encodeURIComponent($("rti").value);
			sParameter += '&reservationtypeid=0';
			sParameter += '&rentaluserid=' + encodeURIComponent($("rentaluserid").value);
			sParameter += '&startdate1=' + encodeURIComponent($("startdate1").value);
			sParameter += '&starthour1=' + encodeURIComponent($("starthour1").value);
			sParameter += '&startminute1=' + encodeURIComponent($("startminute1").value);
			sParameter += '&startampm1=' + encodeURIComponent($("startampm1").value);
			sParameter += '&endhour1=' + encodeURIComponent($("endhour1").value);
			sParameter += '&endminute1=' + encodeURIComponent($("endminute1").value);
			sParameter += '&endampm1=' + encodeURIComponent($("endampm1").value);
			sParameter += '&endday1=' + encodeURIComponent($("endday1").value);

			// Fire off job to check dates and times
			doAjax('checkselecteddates.asp', sParameter , 'checkReturn', 'post', '0');
		}

		function checkReturn( sReturn )
		{
			//alert(sReturn);
			//document.frmSearchReturn.rentaluserid.value = $("rentaluserid").value;
			//document.frmSearchReturn.searchname.value = $("searchname").value;
			//document.frmSearchReturn.reservationtypeid.value = $("reservationtypeid").value;
			//alert( document.frmSearchReturn.rentaluserid.value );
			document.frmDateSelection.action = "rentaltimeselection.asp";
			if (sReturn != 'short' && sReturn != 'buffer' && sReturn != 'buffershort')
			{
				document.frmDateSelection.msg.value = sReturn;
			}
			else
			{
				document.frmDateSelection.msg.value = sReturn + 'noconfirm';
			}
			document.frmDateSelection.submit();
		}

		function WaitAndValidate()
		{
			// This causes the reservation to have a 0 to 3 sec wait before checking availability to try and break ties
			var ReturnTime = Math.floor(Math.random()*3000);
			setTimeout("Validate()", ReturnTime);
		}

		function Validate()
		{
			// Check that all end times are later than start times
			var dtStart = new Date($("startdate1").value + " " + $("starthour1").value + ":" + $("startminute1").value + " " + $("startampm1").value);
			var dtEnd;
			if ($("endday1").value == "0")
			{
				dtEnd = new Date($("startdate1").value + " " + $("endhour1").value + ":" + $("endminute1").value + " " + $("endampm1").value);
				//alert(dtEnd);
			}
			else
			{
				dtEnd = new Date($("startdate1").value + " " + $("endhour1").value + ":" + $("endminute1").value + " " + $("endampm1").value);
				dtEnd.setDate(dtEnd.getDate()+1);
				//alert(dtEnd);
			}
			var difference_in_milliseconds = dtEnd - dtStart;
			if (difference_in_milliseconds <= 0)
			{
				alert("One of the end times is not after the start time. Please correct this and try again.");
				$("endhour1" + i).focus();
				return;
			}

			//alert("OK");
			//return;
			// Now bundle the dates and times and send off to check routine via AJAX
			var sParameter = 'rentalid=' + encodeURIComponent($("rentalid").value);
			sParameter += '&maxrows=' + encodeURIComponent($("maxrows").value);
			sParameter += '&rti=' + encodeURIComponent($("rti").value);
			sParameter += '&reservationtypeid=' + encodeURIComponent($("reservationtypeid").value);
			sParameter += '&rentaluserid=' + encodeURIComponent($("rentaluserid").value);
			sParameter += '&startdate1=' + encodeURIComponent($("startdate1").value);
			sParameter += '&starthour1=' + encodeURIComponent($("starthour1").value);
			sParameter += '&startminute1=' + encodeURIComponent($("startminute1").value);
			sParameter += '&startampm1=' + encodeURIComponent($("startampm1").value);
			sParameter += '&endhour1=' + encodeURIComponent($("endhour1").value);
			sParameter += '&endminute1=' + encodeURIComponent($("endminute1").value);
			sParameter += '&endampm1=' + encodeURIComponent($("endampm1").value);
			sParameter += '&endday1=' + encodeURIComponent($("endday1").value);

			// Fire off job to check dates and times
			doAjax('checkselecteddates.asp', sParameter , 'validateReturn', 'post', '0');
		}

		function validateReturn( sReturn )
		{
			//alert(sReturn);
			//document.frmSearchReturn.rentaluserid.value = $("rentaluserid").value;
			//document.frmSearchReturn.searchname.value = $("searchname").value;
			//document.frmSearchReturn.reservationtypeid.value = $("reservationtypeid").value;
			if (sReturn != 'OK')
			{
				document.frmDateSelection.action = "rentaltimeselection.asp";
				document.frmDateSelection.msg.value = sReturn;
			}

			document.frmDateSelection.submit();
		}


	//-->
	</script>

</head>

<body onload="loader();">

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Make Simple Reservations</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			
			<span id="screenMsg">&nbsp;</span>

			<p>
				<input type="button" class="button" value="<< Back" onclick="goBack();" />
			</p>

			<form name="frmDateSelection" method="post" action="rentalreservationmake.asp">
				<input type="hidden" id="rid" name="rid" value="0" />
				<input type="text" id="rentalid" name="rentalid" value="<%=iRentalId%>" />
				<input type="hidden" id="rti" name="rti" value="<%=iReservationTempId%>" />
				<input type="hidden" name="msg" value="<%=sMessage%>" />

				<%	'ShowRentalsDetails iRentalId				%>
				<% ShowRentalNameAndLocation iRentalId %><br /><br />

				<table id="reservationtempinfo" cellpadding="0" cellspacing="1" border="0">
					<tr>
						<td class="labelcolumn"><strong>Reservation Type:</strong></td>
						<td class="pickcolumn" align="left"><% ShowRentalReservationTypes iReservationTypeId %></td>
						<td class="labelcolumn2">&nbsp;</td>
						<td><%=sNoCostPhrase%></td>
					</tr>
					<tr>
						<td class="labelcolumn"><strong>Name Is Like:</strong></td><td colspan="3"><input type="text" id="searchname" name="searchname" value="<%=sSearchName%>" size="25" maxlength="25" onkeypress="if(event.keyCode=='13'){doNameSearchChange();return false;}" /> 
							<input type="button" class="button" value="Search for a Name" onclick="doNameSearchChange();" /> <input type="button" class="button" value="New Public User" onclick="newUser();" />
						</td>
					</tr>
					<tr>
						<td colspan="4">
							<span id="applicant">
								<%	
									If bHasUsers Then 
										If iRentalUserid > CLng(0) Then 
											If bIsPublicUsers Then 
												' Show registered user picks
												ShowCitizenPicks iRentalUserid, sSearchName
											Else
												' Show admin picks
												ShowAdminPicks iRentalUserid, sSearchName
											End If 
										Else	%>
											<input type="hidden" value="0" name="rentaluserid" id="rentaluserid" />Search for a name then select one from the resulting list.
								<%		End If 
									Else	%>
										<input type="hidden" value="0" name="rentaluserid" id="rentaluserid" />This reservation type does not need a renter.
								<%	End If	%>
							</span> <input type="button" class="button" id="edituserbtn" value="Edit User" onclick="EditApplicant();" />
						</td>
					</tr>
				</table>

				<p>
					<table id="reservationtempdates" cellpadding="0" cellspacing="0" border="0">
						<tr><th class="firstcell">Date</th><th>Start Time</th><th>End Time</th><th class="lastcell">Available</th></tr>
<%						'Pull the wanted dates list here
						'sPeriodTypeSelector = GetSelectedPeriodTypeId( "anytime" )
						iTotalRows = ShowRentalAvailabilityDetails( iReservationTempId, iRentalId, "anytime" )
%>			
					</table>
					<input type="hidden" id="maxrows" name="maxrows" value="<%=iTotalRows%>" />
				</p>
				<p>
					<input type="button" class="button" name="checkbutton" id="checkbutton" value="Check Times" onclick="CheckDates()" />&nbsp;
					<input type="button" class="button" name="continuebutton" id="continuebutton" value="Check and Reserve" onclick="WaitAndValidate()" />
				</p>

			</form>
		</div>
	</div>

	<form name="frmBack" method="post" action="rentalavailability.asp">
		<input type="hidden" id="rti" name="rti" value="<%=iReservationTempId%>" />
	</form>


	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' boolean SetPageVariables( iReservationTempId )
'--------------------------------------------------------------------------------------------------
Function SetPageVariables( ByVal iReservationTempId )
	Dim sSql, oRs

	sSql = "SELECT rentalid, requestedstartdate, ISNULL(requestedstarthour,1) AS requestedstarthour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(requestedstartminute,0),2) AS requestedstartminute, "
	sSql = sSql & " ISNULL(requestedstartampm,'PM') AS requestedstartampm, ISNULL(requestedendhour,2) AS requestedendhour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(requestedendminute,0),2) AS requestedendminute, "
	sSql = sSql & " ISNULL(requestedendampm,'PM') AS requestedendampm, ISNULL(requestedendday,0) AS requestedendday "
	sSql = sSql & " FROM egov_rentalreservationstemp "
	sSql = sSql & " WHERE reservationtempid = " & iReservationTempId
	sSql = sSql & " AND orgid = " & session("orgid")
	'response.write sSql & "<br /><br />"
	'response.end


	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iRentalId = CLng(oRs("rentalid"))
		sSelectedDate = oRs("requestedstartdate")
'		iStartHour = oRs("requestedstarthour")
'		iStartMinute = oRs("requestedstartminute")
'		sStartAmPm = oRs("requestedstartampm")
'		iEndHour = oRs("requestedendhour")
'		iEndMinute = oRs("requestedendminute")
'		sEndAmPm = oRs("requestedendampm")
		SetPageVariables = True 
	Else
		SetPageVariables = False 
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowRentalsDetails iRentalId
'--------------------------------------------------------------------------------------------------
Sub ShowRentalsDetails( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT R.rentalid, R.rentalname, L.name AS locationname, ISNULL(R.width,'') AS width, ISNULL(R.length,'') AS length, "
	sSql = sSql & "ISNULL(R.capacity,'') AS capacity, R.publiccanreserve, "
	sSql = sSql & "ISNULL(R.shortdescription,'') AS description "
	sSql = sSql & "FROM egov_rentals R, egov_class_location L "
	sSql = sSql & "WHERE R.publiccanview = 1 AND R.locationid = L.locationId "
	sSql = sSql & "AND R.rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write "<table class=""availablerentals"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
		response.write "<tr><td class=""spacerrow"">&nbsp;</td></tr>"
		response.write "<tr>"
		response.write "<td valign=""top"" align=""left"" class=""availabledescription"">"

		response.write "<p><span class=""availableschedulerentalname"">"
		If oRs("locationname")  <> "" Then 
			response.write oRs("locationname") & " &ndash; " 
		End If 
		response.write oRs("rentalname")
		response.write "</span></p>"

		response.write "<p>" & oRs("description") & "</p>"

		If oRs("locationname")  <> "" Or oRs("width") <> "" Or oRs("capacity") <> "" Then 
			response.write vbcrlf & "<p>"
			If oRs("width") <> "" Then 
				response.write "<strong>Dimensions: </strong>" & oRs("width") & " x " & oRs("length") & "<br />"
			End If 
			If oRs("capacity") <> "" Then 
				response.write "<strong>Capacity: </strong>" & oRs("capacity") & "<br />"
			End If 
			response.write vbcrlf & "</p>"
		End If 

'		DisplayRentalDocuments iRentalId

		response.write "</td>"
		response.write "</tr>"
		response.write "</table>"
		oRs.MoveNext 
	Loop

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowRentalReservationTypes iReservationTypeId
'--------------------------------------------------------------------------------------------------
Sub ShowRentalReservationTypes( ByVal iReservationTypeId )
	Dim oRs, sSql

	sSql = "SELECT reservationtypeid, reservationtype FROM egov_rentalreservationtypes "
	sSql = sSql & "WHERE displayindropdown = 1 AND orgid = " & session("orgid") & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write "<select id=""reservationtypeid"" name=""reservationtypeid"" onchange=""doUserPickChange();"">"
	Do While Not oRs.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oRs("reservationtypeid") & """ "
		If CLng(oRs("reservationtypeid")) = CLng(iReservationTypeId) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("reservationtype") & "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetFirstReservationTypeInList( )
'--------------------------------------------------------------------------------------------------
Function GetFirstReservationTypeInList( )
	Dim oRs, sSql

	sSql = "SELECT reservationtypeid FROM egov_rentalreservationtypes "
	sSql = sSql & "WHERE displayindropdown = 1 AND orgid = " & session("orgid") & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFirstReservationTypeInList = oRs("reservationtypeid")
	Else
		GetFirstReservationTypeInList = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void GetUserFlagsFromReservationTypeId iReservationTypeId, bHasUsers, bIsPublicUsers
'--------------------------------------------------------------------------------------------------
Sub GetUserFlagsFromReservationTypeId( ByVal iReservationTypeId, ByRef bHasUsers, ByRef bIsPublicUsers )
	Dim oRs, sSql

	sSql = "SELECT hasusers, haspublicusers FROM egov_rentalreservationtypes "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " AND reservationtypeid = " & iReservationTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("hasusers") Then
			bHasUsers = True 
		Else
			bHasUsers = False 
		End If 
		If oRs("haspublicusers") Then
			bIsPublicUsers = True 
		Else
			bIsPublicUsers = False 
		End If 
	Else
		bHasUsers = False 
		bIsPublicUsers = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer ShowRentalAvailabilityDetails( iReservationTempId, iRentalId, sPeriodTypeSelector )
'--------------------------------------------------------------------------------------------------
function ShowRentalAvailabilityDetails( ByVal iReservationTempId, ByVal iRentalId, ByVal sPeriodTypeSelector )
	Dim sSQL, oRs, iRowCount, sAmPm, sReservationStartTime, sReservationEndTime, iEndDay, bOffSeasonFlag
	Dim bIsAllDayOnly, sDisabledOption
	Dim aWantedDates(1,0)

	iRowCount = 0

 'Get the temp dates
	 sSQL = "SELECT reservationstarttime, "
  sSQL = sSQL & " reservationendtime, "
  sSQL = sSQL & " endday "
  sSQL = sSQL & " FROM egov_rentalreservationdatestemp "
	 sSQL = sSQL & " WHERE reservationtempid = " & iReservationTempId
	 sSQL = sSQL & " ORDER BY reservationstarttime"

	 set oRs = Server.CreateObject("ADODB.Recordset")
	 oRs.Open sSQL, Application("DSN"), 0, 1

	 if not oRs.eof then
		   iRowCount = 1

   	'Build the date time row
   		sReservationStartTime = oRs("reservationstarttime")
   		sReservationEndTime   = oRs("reservationendtime")
   		iEndDay               = oRs("endday")
   		bOffSeasonFlag        = GetOffSeasonFlag( iRentalid, DateValue(sReservationStartTime) )

   		if iEndDay = "1" then
  	   		sReservationEndTime = CStr(DateAdd("d", 1, CDate(sReservationEndTime)))
     end if

   		if RentalIsAllDay( iRentalid, bOffSeasonFlag, Weekday(DateValue(sReservationStartTime)) ) then
			     bIsAllDayOnly   = true
   			  sDisabledOption = "disabled"
   		else
			     bIsAllDayOnly   = false 
   			  sDisabledOption = ""
   		end if

   	'If this is for all day, or the rental is only available for all day reservations then we need the opening and closing times on that day
   		if sPeriodTypeSelector = "allday" OR bIsAllDayOnly then
     			SetHoursToOpenAndClose iRentalId, DateValue(sReservationStartTime), sReservationStartTime, sReservationEndTime, iEndDay
   		end if

   		response.write "  <tr class=""dateline"">"
   		response.write "      <td class=""firstcell"">" & vbcrlf
     response.write "          <span id=""pickeddate"">" & DateValue(oRs("reservationstarttime")) & vbcrlf
   		response.write "          <input type=""hidden"" name=""startdate" & iRowCount & """ id=""startdate" & iRowCount & """ value=""" & DateValue(oRs("reservationstarttime")) & """ />" & vbcrlf
   		response.write "          </span>" & vbcrlf
     response.write "      </td>" & vbcrlf
   		response.write "      <td align=""center"">"  & vbcrlf
                             		ShowHourPicks "starthour" & iRowCount, GetHourFromDateTime( sReservationStartTime, sAmPm ), sDisabledOption  'In rentalsguifunctions.asp
   		response.write            ":" & vbcrlf
                               ShowMinutePicks "startminute" & iRowCount, Minute(sReservationStartTime), sDisabledOption  'In rentalsguifunctions.asp
   		response.write            " " & vbcrlf
                               ShowAmPmPicks "startampm" & iRowCount, sAmPm, sDisabledOption  'In rentalsguifunctions.asp
   		response.write "      </td>" & vbcrlf
   		response.write "      <td align=""center"">" & vbcrlf
                               ShowHourPicks "endhour" & iRowCount, GetHourFromDateTime( sReservationEndTime, sAmPm ), sDisabledOption  'In rentalsguifunctions.asp
   		response.write            ":" & vbcrlf
                               ShowMinutePicks "endminute" & iRowCount, Minute(sReservationEndTime), sDisabledOption  'In rentalsguifunctions.asp
   		response.write            " " & vbcrlf
                               ShowAmPmPicks "endampm" & iRowCount, sAmPm, sDisabledOption  'In rentalsguifunctions.asp
   		response.write            " " & vbcrlf
                               ShowSameNextDayPick "endday" & iRowCount, iEndDay, sDisabledOption  'In rentalsguifunctions.asp
   		response.write "      </td>" & vbcrlf

   	'Get the availability flag on that date and time
   		aWantedDates(0,0) = oRs("reservationstarttime")
   		aWantedDates(1,0) = oRs("reservationendtime")

   		response.write "      <td class=""lastcell"" align=""center"">" & vbcrlf
                             		ShowRentalAvailabilityFlag iRentalId, aWantedDates, sPeriodTypeSelector, False 
   		response.write "      </td>" & vbcrlf
   		response.write "  </tr>" & vbcrlf
    	response.write "  <tr>" & vbcrlf

    'Get rental details for that date
   		response.write "      <td class=""firstcell"" colspan=""2"">"
   		response.write            WeekDayName(Weekday(DateValue(sReservationStartTime)))
   		response.write            " &ndash; " & GetRentalSeason( iRentalId, DateValue(sReservationStartTime) )
   		response.write            GetRentalHours( iRentalId, DateValue(sReservationStartTime) )
   		response.write "      </td>" & vbcrlf

   	'Get the other reservations, etc for this date
   		response.write "      <td class=""lastcell"" colspan=""2"" valign=""top"">Also happening on this date" & vbcrlf
                             		ShowOtherReservationsForDate iRentalId, DateValue(sReservationStartTime)
   		response.write "      </td>" & vbcrlf
   		response.write "  </tr>" & vbcrlf

   	'The seperator Row
   		response.write "  <tr><td colspan=""4"" class=""tempseparator"">&nbsp;</td></tr>" & vbcrf
		
  end if
	
 	oRs.close
	 set oRs = nothing 

 	ShowRentalAvailabilityDetails = iRowCount

end function
%>