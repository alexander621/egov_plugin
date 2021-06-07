<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="facility_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: FACILITY_RESERVATION_EDIT.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/15/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   02/15/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1   03/10/2006 Steve Loar - Added Edit of Facility Field Values
' 1.2	04/18/2007	Steve Loar - Added the dropdown navagation
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
if Request.ServerVariables("REQUEST_METHOD") = "POST" and request("abuseuserid") <> "" then 
	If LCase(request("isabusive")) = "on" Then 
		sAbuseFlag = "1"
	Else
		sAbuseFlag = "0"
	End If 
	sAbuseNote = DBSafe(request("abusenote"))
	if request("abuseuserid") <> "" then
		sAbuseUserID = request("abuseuserid")
	else
		sAbuseUserID = "0"
	end if
	sSQL = "UPDATE egov_users SET facilityabuse = '" & sAbuseFlag & "', facilityabusenote = '" & sAbuseNote & "' WHERE userid = '" & sAbuseUserID & "'"
	RunSQLStatement sSql
end if
sLevel = "../" ' Override of value from common.asp

Dim sFacilityName, sCheckInDate, sCheckInTime, sCheckOutDate, sCheckOutTime, blnRecurrence, iUserID  
Dim sStatus, irecurrentid, sinternalnote, iFacilityid

iReservationID = request("iReservationID")

' Set the redirect page so we can return here from recreation_details_update.asp with all the correct parameters
Session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()

' GET RESERVATION INFORMATION
GetReservationDetails iReservationID, sFacilityName, iFacilityid, sCheckInDate, sCheckInTime, sCheckOutDate, blnRecurrence, sStatus, iUserID, irecurrentid, sinternalnote

%>

<html lang="en">
<head>
	<meta charset="UTF-8">

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="querytool.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="reservation.css" />

	<script src="../scripts/jquery-1.7.2.min.js"></script>

	<script>
	<!--//

		function checkstatus() {
			if (document.frmstatus.selstatus.options[document.frmstatus.selstatus.selectedIndex].value == 'CANCELLED'){
				location.href='facility_reservation_edit.asp?ireservationid=<%=request("ireservationid")%>&C=TRUE';
			}
		}

		function checkcancel() {
			if (document.frmstatus.selstatus.options[document.frmstatus.selstatus.selectedIndex].value == 'CANCELLED')
			{
				if (document.frmstatus.sCancelReason.value != '')
				{
					document.frmstatus.submit();
				}
				else
				{
					alert('You must enter a valid cancel reason!');
				}
			}
			else
			{
				document.frmstatus.submit();
			}
		}
    
		function view_waivers( )
		{
			var waiverForms = "";
			// loop through the selected waivers and build a mask field to pass
			$checkedCheckboxes = $("input:checkbox[name=chkwaivers]:checked");
			$checkedCheckboxes.each(function () {
				if (waiverForms != "")
				{
					waiverForms += ",";
				}
				waiverForms +=  $(this).val();
			});

			if ($("#requiredwaivers").val() != "")
			{
				if (waiverForms != "")
				{
					waiverForms += ",";
				}
				waiverForms += $("#requiredwaivers").val();
			}

			// debugging line
			//alert( "selected waiverForms: " + waiverForms );

			if (waiverForms != "")
			{
				// CHANGE FORM'S ACTION URL AND SUBMIT
				document.frmwaivers.action = "display_reservation_waiver.aspx?mask=" + waiverForms;
				document.frmwaivers.target = '_NEW';
				document.frmwaivers.submit();
			}
			else
			{
				alert( "Please select a waiver to display." );
			}
		}

	//-->
	</script>
</head>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div style="padding:20px;">

<!--BEGIN PAGE CONTENT-->
<p><a href="FACILITY_CALENDAR.ASP?L=<%=iFacilityid%>&Y=<%=Year(sCheckInDate)%>&M=<%=Month(sCheckInDate)%>" class="linkbutton"><< Calendar</a></p>


<!--BEGIN: FACILITY-->
<div class="reserveformtitle">Facility</div>
<div class="reserveforminputarea">

<form name="ReservationValues" action="recreation_details_update.asp" method="post">
	<input type="hidden" name="facilityscheduleid" value="<%= iReservationID %>" />

	<table>
		<tr><td align="right">Facility: </td><td><strong><%=sFacilityName%></strong></td></tr>
		<%= GetFacilityFieldValues( iReservationID ) %>
	</table>

	<!--BEGIN: ADMIN COMMENT FIELD-->
	<p><b>City Internal Note (512 max number of characters):</b><br><textarea maxlength="1024" id="internalnote" name="internalnote"><%=sinternalnote%></textarea></p>
	<!--END: AMDIN COMMENT FIELD-->

	<input type="submit" value="Save Changes" class="facilitybutton" />
</form>

</div>
<!--END: FACILITY-->


<!--BEGIN: RESERVATION DATES-->
<div class="reserveformtitle">Reservation Date(s) </div>
<div class="reserveforminputarea">
	<table>
		<tr><td align="right">Check In:</td><td>  <%=sCheckInDate & " " & sCheckInTime%></td></tr>
		<tr><td align="right">Expected Arrival Time:</td><td> <%=sCheckInTime%></td></tr>
		<tr><td align="right">Check Out:</td><td>  <%=sCheckOutDate & " " & sCheckOutTime%></td></tr>
		<tr><td align="right">Expected Departure Time:</td><td> <%=sCheckOutTime%></td></tr>
		<tr><td>&nbsp;</td><td><input onClick="location.href='facility_date_edit.asp?ireservationid=<%=iReservationID%>';" type="button" class="facilitybutton" value="Change Dates" /></td></tr>
	</table>
</div>
<!--END: RESERVATION DATES-->


<!--BEGIN: RECURRENCE-->
<%
If blnRecurrence Then%>
	<div class="reserveformtitle">Recurrent Reservation</div>

	<div class="reserveforminputarea">
		<% DisplayRecurrentInformation irecurrentid %>
	</div>

<%End If%>
<!--END: RECURRENCE-->


<!--BEGIN: LESSEE-->
<div class="reserveformtitle">Lessee</div>
<div class="reserveforminputarea">
	<% DisplayUserInfo iUserID %>
</div>
<!--END: LESSEE-->

<!--BEGIN: WAIVER-->
<form name="frmwaivers" action="#" method="post">
	<input type="hidden" name="adminlink" value="true" />
	<input type="hidden" name="reservationid" value="<%=iReservationID%>" />
	<div class="reserveformtitle">Waiver Downloads</div>
	<div class="reserveforminputarea">
		<p><b>Download the following form(s) to print, sign, and bring with you when picking up the key:</b></p>
		<p><% GetWaiverMask( ifacilityid ) %></p>
	</div>
</form>
<!--END: WAIVER-->


<!--BEGIN: RESERVATION STATUS -->
<form name="frmstatus" action="reservation_status_update.asp" method="post">
<input type="hidden" name="ireservationid" value="<%=iReservationID%>" />
<input type="hidden" name="ioccurrenceid" value="<%=irecurrentid%>" />
<div class="reserveformtitle">Reservation Status </div>
<div class="reserveforminputarea">
<p>
<table>
	<tr>
		<td align=right>Status:</td>
		<td><%DrawStatusSelect(sStatus)%></td>
	</tr>


<% If blnRecurrence = "True" Then %>
	<tr>
		<td align="right">Change Applies To:</td>
		<td>
			<select name="applies">
				<option value="single">SINGLE RESERVATION</option>
				<option value="all">ALL OCCURENCES OF THIS RESERVATION</option>
			</select>
		</td>
	</tr>
<%Else%>
	<input type="hidden" value="single" name="applies" />
<%End If%>
	
	<%	If UCase(request("C")) = "TRUE" Then %>
			<tr><td>Cancel Reason:</td><td><input maxlength="1024" type="textbox" value="" name="sCancelReason" style="width:350px;" /></td></tr>
	<%  End If %>
	
	<tr>
		<td align="right" colspan="2"><input onClick="checkcancel();" type="button" value="Save Changes" class="facilitybutton" /></td>
	</tr>
</table>
</div>
</form>
<!--END: RESERVATION STATUS -->


<!--END: PAGE CONTENT-->

</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>



<%
'--------------------------------------------------------------------------------------------------
' GetReservationDetails( iReservationID, sFacilityName, iFacilityid, sCheckInDate, sCheckInTime, sCheckOutDate, blnRecurrence, sStatus, iUserID, irecurrentid, sinternalnote )
'--------------------------------------------------------------------------------------------------
Sub GetReservationDetails( ByVal iReservationID, ByRef sFacilityName, ByRef iFacilityid, ByRef sCheckInDate, ByRef sCheckInTime, ByRef sCheckOutDate, ByRef blnRecurrence, ByRef sStatus, ByRef iUserID, ByRef irecurrentid, ByRef sinternalnote )
	Dim sSql, oRs

	' GET INFORMATION FOR THIS RESERVATION
	sSql = "SELECT S.facilityid, facilityname, checkindate, checkintime, checkoutdate, checkouttime, "
	sSql = sSql & "isrecurrent, status, lesseeid, facilityrecurrenceid, ISNULL(internalnote,'') AS internalnote "
	sSql = sSql & "FROM egov_facilityschedule S INNER JOIN egov_facility F ON S.facilityid = F.facilityid "
	sSql = sSql & "WHERE S.facilityscheduleid = " & CLng(iReservationID)

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	' IF RESERVATION HAS INFORMATION POPULATE VALUES
	If Not oRs.EOF Then
		sFacilityName = oRs("facilityname") 
		iFacilityid = oRs("facilityid") 
		sCheckInDate = oRs("checkindate") 
		sCheckInTime = oRs("checkintime") 
		sCheckOutDate = oRs("checkoutdate") 
		sCheckOutTime = oRs("checkouttime") 
		blnRecurrence = oRs("isrecurrent")
		sStatus = oRs("status") 
		iUserID = oRs("lesseeid")
		irecurrentid = oRs("facilityrecurrenceid")
		sinternalnote = oRs("internalnote")
	End If

	oRs.Close
	Set oRs = Nothing
	
End Sub


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYUSERINFO(IUSERID)
'--------------------------------------------------------------------------------------------------
Sub DisplayUserInfo( ByVal iUserID )
	Dim sSql, oRs

	' SELECT ROW WITH THIS USER'S INFORMATION
	sSql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, ISNULL(residenttype,'N') AS residenttype, ISNULL(useraddress,'') AS useraddress, "
	sSql = sSql & "ISNULL(useremail,'') AS useremail, ISNULL(usercity,'') AS usercity, ISNULL(userstate,'') AS userstate, "
	sSql = sSql & "ISNULL(userzip,'') AS userzip, CAST(ISNULL(facilityabuse,0) as bit) as facilityabuse, FacilityAbuseNote FROM egov_users WHERE userid = " & iUserId 
	'response.write sSql & "<br><br>"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	' IF RECORDSET NOT EMPTY THEN DISPLAY USER INFORMATION
	If Not oRs.EOF Then
		
		response.write "<table>"
		response.write "<tr><td>Name: </td><td>" & oRs("userfname") & " " & oRs("userlname")
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;<input onClick=""location.href='../dirs/update_citizen.asp?userid=" & iUserId & "';"" type=""button"" class=""facilitybutton"" value=""Edit User"" />"
		Response.write "</td></tr>"
		response.write "<tr><td>Residency: </td><td>" & GetResidencyDescription( Session("OrgID"),  oRs("residenttype") ) & "</td></tr>"
		response.write "<tr><td>Address: </td><td>" & oRs("useraddress") & "</td></tr>"
		response.write "<tr><td>Email: </td><td>" & oRs("useremail") & "</td></tr>"
		response.write "<tr><td>City: </td><td>" & oRs("usercity") & "</td></tr>"
		response.write "<tr><td>State: </td><td>" & oRs("userstate") & "</td></tr>"
		response.write "<tr><td>Zip: </td><td>" & oRs("userzip") & "</td></tr>"
		response.write "</table>"

		response.write "<form action=""#"" method=""POST"">"
     	response.write "<p class=""flagselection"">" & vbcrlf
	 	sAbuseChecked = ""
	 	if oRs("FacilityAbuse") then sAbuseChecked = "checked"
	 	response.write "<input type=""hidden"" name=""abuseuserid"" id=""abuseuserid"" value=""" & iUserId & """ />"
     	response.write "<input type=""checkbox"" name=""isabusive"" id=""isabusive"" " & sAbuseChecked & " /> Facility/Rental Abuser" & vbcrlf
	 	response.write "<br />"
	 	response.write "Abuse Notes: <br />"
	 	response.write "<textarea name=""abusenote"" id=""abusenote"" style=""width:740px;height:25px"">" & oRs("FacilityAbuseNote") & "</textarea>"
     	response.write "</p>" & vbcrlf
		response.write "<input type=""submit"" value=""Save Changes"" class=""facilitybutton"" />"
     	response.write "</form>" & vbcrlf
	End If

	oRs.Close
	Set oRs = Nothing
	
End Sub


'--------------------------------------------------------------------------------------------------
'  DRAWSELECTFACILITY()
'--------------------------------------------------------------------------------------------------
Sub DrawStatusSelect( ByVal sStatus )
	Dim sSql, oRs, sSelected
	
	' GET LIST OF STATUSES
	sSql = "SELECT facilitystatusname from egov_facilitystatus"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

    ' LOOP THRU LIST OF AVAILABLE FACILITIES AND DISPLAY TO USER
    Response.Write("<select  name=""selstatus"" onChange=""checkstatus();"">")
   
	Do While Not oRs.EOF
		sSelected = ""
		
		' IF LOOP ROW STATUS MATCHES DATABASE STATUS AND PAGE IS NOT CURRENTLY MARKED FOR EDITTING WITH STATUS OF CANCELLED
		If sStatus = oRs("facilitystatusname") And UCase(request("C")) <> "TRUE" Then
			sSelected = " selected=""selected"""
		End If

		' IF LOOP ROW STATUS IS CANCELLED AND PAGE IS CURRENTLY MARKED FOR EDITTING WITH STATUS OF CANCELLED
		If oRs("facilitystatusname") = "CANCELLED" And UCase(request("C")) = "TRUE" Then
			sSelected = " selected=""selected"""
		End If
		
		Response.Write("<option" & sSelected & " value=""" & oRs("facilitystatusname") & """>" & oRs("facilitystatusname") & "</option>" & vbCrLf)
		
		oRs.MoveNext
	Loop
    
	Response.Write("</select>" & vbCrLf)

	oRs.Close
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------------------------------------
' FUNCTION GETFACILITYFIELDVALUES(IFACILITYPAYMENTID)
'------------------------------------------------------------------------------------------------------------
Function GetFacilityFieldValues( ByVal iFacilityPaymentID )
	Dim sHeight, sSql, oRs, sReturnValue

	sReturnValue = ""

	sSql = "SELECT V.facilityvalueid, V.fieldid, V.fieldvalue, V.paymentid, F.fieldprompt, "
	sSql = sSql & "F.fieldtype, F.facilityid, F.sequence, F.isrequired, F.fieldchoices FROM egov_facility_field_values V "
	sSql = sSql & "INNER JOIN egov_facility_fields F ON V.fieldid = F.fieldid WHERE V.paymentid =  " & iFacilityPaymentID
	sSql = sSql & " ORDER BY V.paymentid, V.fieldid"
	'response.write sSql & "<br><br>"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then

		sReturnValue = sReturnValue & vbcrlf & "<input type=""hidden"" name=""operation"" value=""update"" />"
		'sReturnValue = sReturnValue & vbcrlf & "<input type=""hidden"" name=""paymentid"" value=""" & iFacilityPaymentID & """/>"

		Do While Not oRs.EOF 
			
			sReturnValue = sReturnValue & vbcrlf & "<tr><td align=""right"" valign=""top"">" & oRs("fieldprompt") & " : </td><td>"
			If  oRs("fieldtype") = 2 Then
				sReturnValue = sReturnValue & "<textarea name=""field_" & oRs("facilityvalueid") & """ value="""""" style=""FONT-SIZE: 8pt; WIDTH: 300px; HEIGHT: 100px; FONT-FAMILY: Arial"">"  & oRs("fieldvalue") & "</textarea>"
			Else 
				sReturnValue = sReturnValue & "<input type=""text""  name=""field_" & oRs("facilityvalueid") & """ value="""  & oRs("fieldvalue") & """ style=""FONT-SIZE: 8pt; WIDTH: 300px; " & sHeight & " FONT-FAMILY: Arial"" />"
			End If 
			sReturnValue = sReturnValue & "</td></tr>"
			oRs.MoveNext
		Loop
	
	Else
		oRs.Close
		' GET EMPTY DATA COLUMNS
		sSql = "SELECT fieldid, fieldprompt FROM egov_facility_fields WHERE facilityid = " & iFacilityid
		
		oRs.Open sSql, Application("DSN"), 3, 1
		
		If Not oRs.EOF Then

			sReturnValue = sReturnValue & vbcrlf & "<input type=""hidden"" name=""operation"" value=""input"" />"
			'sReturnValue = sReturnValue & vbcrlf & "<input type=""hidden"" name=""paymentid"" value=""" & iFacilityPaymentID & """ />"

			Do While Not oRs.EOF 
				' SET HEIGHT FOR INPUT BOX BASED ON FIELD TYPE, 1=STANDARD, 2=SIMULATED TEXT AREA
				sReturnValue = sReturnValue & vbcrlf & "<tr><td align=""right"" valign=""top"">" & oRs("fieldprompt") & " : </td><td>"
				If  oRs("fieldtype") = 2 Then
					sReturnValue = sReturnValue & "<textarea name=""field_" & oRs("fieldid") & """ value="""""" style=""FONT-SIZE: 8pt; WIDTH: 300px; HEIGHT: 100px; FONT-FAMILY: Arial""></textarea>"
				Else 
					sReturnValue = sReturnValue & "<input type=""text""  name=""field_" & oRs("fieldid") & """ value="""""" style=""FONT-SIZE: 8pt; WIDTH: 300px; FONT-FAMILY: Arial"" />"
				End If 
				sReturnValue = sReturnValue & "</td></tr>"
				oRs.MoveNext
			Loop

		End If
	
	End If

	oRs.Close
	Set oRs = Nothing

	GetFacilityFieldValues = sReturnValue

End Function


'------------------------------------------------------------------------------------------------------------
' DISPLAYRECURRENTINFORMATION(IRECURRENTID)
'------------------------------------------------------------------------------------------------------------
Sub DisplayRecurrentInformation( ByVal irecurrentid )
	Dim sSql, oRs

	' GET RECURRENCT DETAILS
	sSql = "SELECT frequency, dayofweek, datepart, enddate, ordinal, month, startdate AS FirstDate FROM egov_facilityrecurrence WHERE recurrentid = " & irecurrentid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write "First Recurrance: " & oRs("FirstDate") & "<br>"

		' GET RECURRENCE TIME FRAME
		Select Case oRs("datepart")

			Case 1
				response.write "Note: Every " & oRs("frequency") & " week(s) starting on " &  Weekdayname(oRs("dayofweek")) & " until " & oRs("enddate") & "."
			Case 2
				response.write "Note: The " & GetOrdinalName(oRs("ordinal")) & " " &  Weekdayname(oRs("dayofweek")) & " of every " & oRs("frequency") & " month(s) until " & oRs("enddate") & "."
			Case 3
				response.write "Hote: The " & GetOrdinalName(oRs("ordinal")) & " " &  Weekdayname(oRs("dayofweek")) & " of every " & MonthName(oRs("month")) & " until " & oRs("enddate") & "."
			Case 4
				response.write "Note: Every day until " & oRs("enddate") & "."
			Case Else

		End Select 

	End If

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------------------------------------
' FUNCTION GETORDINALNAME(IVALUE)
'------------------------------------------------------------------------------------------------------------
Function GetOrdinalName( ByVal iValue )
	Dim sReturnValue

	sReturnValue = "UNKNOWN"

	Select Case iValue

		Case 1
			sReturnValue = "first "

		Case 2
			sReturnValue = "second"

		Case 3
			sReturnValue = "third"

		Case 4
			sReturnValue = "fourth"

		Case 5
			sReturnValue = "last"

		Case Else

	End Select

	GetOrdinalName = sReturnValue

End Function



'--------------------------------------------------------------------------------------------------
' GETWAIVERMASK(IFACILITYID)
'--------------------------------------------------------------------------------------------------
Sub GetWaiverMask( ByVal ifacilityid )
	Dim sSql, oRs, sRequiredWaivers, iWaiverCount

	sRequiredWaivers = ""
	iWaiverCount = 0

	sSql = "SELECT W.waiverid, description, isrequired FROM egov_facilitywaivers F "
	sSql = sSql & "INNER JOIN egov_waivers W ON F.waiverid = W.waiverid "
	sSql = sSql & "WHERE F.facilityid = " & ifacilityid & " ORDER BY isrequired, name"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	iCount = 0
	sMask = ""

	If Not oRs.EOF Then
		iCount = iCount + 1
		response.write "<p><b>Optional Waivers:</b><br />"

		Do While Not oRs.EOF 
			iWaiverCount = iWaiverCount + 1
			If oRs("isrequired") Then
				'response.write "<input type=""hidden"" name=""chkwaivers_" & iCount & """ value=""" & oRs("waiverid") & """ />"
				If sRequiredWaivers <> "" Then 
					sRequiredWaivers = sRequiredWaivers & ","
				End If 
				sRequiredWaivers = sRequiredWaivers & oRs("waiverid")
			Else
				response.write "<input type=""checkbox"" name=""chkwaivers"" value=""" & oRs("waiverid") & """ />" & oRs("description") & " <b><small>(OPTIONAL)</small></b><br />"
			End If
			
			oRs.MoveNext
		Loop
		
		response.write "<input type=""hidden"" name=""requiredwaivers"" id=""requiredwaivers"" value=""" & sRequiredWaivers & """ />"

		If clng(iWaiverCount) > clng(0)  Then
			response.write "<p><input class=""facilitybutton"" type=""button"" value=""Click to download PDF forms"" onClick=""view_waivers( );"" /></p>"
		End If 
	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' SUB LISTWAIVERS()
'--------------------------------------------------------------------------------------------------
Sub ListWaivers()

	' DISPLAY WAIVER LIST WITH LINK TO OPEN THE WAIVERS AS PDF IN NEW WINDOW
	response.write "<p><input class=""facilitybutton"" type=""button"" value=""Click to download PDF forms"" onClick=""view_waivers('display_waiver.asp');"" /></p>"

End Sub

%>
