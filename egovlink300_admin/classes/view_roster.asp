<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: view_roster.asp
' AUTHOR: John Stullenberger
' CREATED: 01/17/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0	01/17/2006	John Stullenberger - Initial Version
' 1.1	10/17/2006	Steve Loar - Security, Header and nav changed
' 1.3	06/11/2008	David Boyer - Added "Download Roster" (excel download)
' 1.4	01/13/2009	David Boyer - Added "Download Team Roster" (excel download)
' 1.5	05/29/2009	Steve Loar - Restricted team roster download query to just active enrollments
' 1.6	08/13/2010	Steve Loar - Breaking into pagable roster list for large enrollments (>1000)
' 2.0	05/09/2011	Steve Loar - Cleaned up SELECT queries to be SQL Server 2008 Compatible
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iClassID, iTimeID, iClassListID, iSentCount, lcl_orghasfeature_residency_verification
Dim lcl_orghasfeature_customreports_classesevents_teamroster, lcl_orghasfeature_custom_registration_craigco
Dim lcl_userhaspermission_customreports_classesevents_teamroster, lcl_onload, lcl_success
Dim toDate, fromDate
		intMaxEnroll = 0
		intEnrolled = 0
		intWaitlist = 0


sLevel = "../"  'Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "registration" ) Then 
	response.redirect sLevel & "permissiondenied.asp"
End If 

iClassID     = CLng(request("classid"))
iTimeID      = CLng(request("timeid"))
iClassListID = request("classlistid")
iSentCount   = request("sentcount")

If Request("fromDate") <> "" Then 
	fromDate = Request("fromDate")
Else
	fromDate = ""
End If 

If Request("toDate") <> "" Then 
	toDate = Request("toDate")
Else
	toDate = ""
End If 

session("RedirectPage") = GetCurrentURL()
session("RedirectLang") = "Return To Class Roster"

'Check for org features
lcl_orghasfeature_residency_verification                 = orghasfeature("residency verification")
lcl_orghasfeature_customreports_classesevents_teamroster = orghasfeature("customreports_classesevents_teamroster")
lcl_orghasfeature_custom_registration_craigco            = orghasfeature("custom_registration_craigco")
bHasRegistrationDateFilter = OrgHasFeature( "registration date filter" )		' in common.asp

'Check for user permissions
lcl_userhaspermission_customreports_classesevents_teamroster = userhaspermission(session("userid"),"customreports_classesevents_teamroster")

'Check for a screen message
lcl_onload  = ""
lcl_success = request("success")

If lcl_success <> "" Then 
	lcl_msg = setupScreenMsg(lcl_success)

	If lcl_success = "SS" And iSentCount <> "" Then 
		lcl_msg = iSentCount & "&nbsp;" & lcl_msg
	End If 

	lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
End If  

%>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>E-Gov Administration Console {Roster}</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../recreation/facility.css" />
	<link rel="stylesheet" href="classes.css" />

	<script src="../scripts/jquery-1.7.2.min.js"></script>
	
	<script src="../scripts/formvalidation_msgdisplay.js"></script>
	<script src="../scripts/getdates.js"></script>
	<script src="../scripts/isvaliddate.js"></script>
	<script src="tablesort.js"></script>
	<script src="../scripts/layers.js"></script>
	<script src="../scripts/ajaxLib.js"></script>

	<script>
	<!--

		function GoToPrint( )
		{
			document.frmprint.submit();
		}

		function GoToAttendance( )
		{
			document.frmAttendanceSheet.submit();
		}

		function confirm_drop(iClassListID, srostername)
		{
			if (confirm("Are you sure you want to drop " + srostername + "?"))
				{ 
					// DELETE HAS BEEN VERIFIED
					location.href='drop_registrant.asp?classid=<%=request("classid")%>&timeid=<%=request("timeid")%>&iclasslistid=' + iClassListID;
				}
		}

		function confirm_move()
		{
			//var sclassname = document.frmrosterlist.newtimeid.options[document.frmrosterlist.newtimeid.selectedIndex].text;
			var sclassname = $("#newtimeid option:selected").text();
			if (confirm("Are you sure you want to move the selected registrants to " + sclassname + "?"))
			{ 
				document.frmrosterlist.action="move_registrants_cgi.asp";
				document.frmrosterlist.submit();
			}
		}

		function confirm_copy()
		{
			var okToCopy = false;
			var attendeeCount = parseInt(0);
			var className = $("#classtimeid option:selected").text();
			if (confirm("Are you sure you want to copy the selected registrants to " + className + "?")) { 
				//alert($("#classtimeid").val());
				// see if a class is selected
				var timeId = parseInt($("#classtimeid").val());
				if (timeId > 0) {
					okToCopy = true;
				}
				else {
					inlineMsg("copyattendeesbtn",'<strong>Copy Failed: </strong>You need to select a class to copy attendees to.',8,"copyattendeesbtn");
					return false;
				}

				// see if any attendees are selected
				var rosterCount = parseInt( $("#rostercount").val());

				if (rosterCount > 0) {
					for (var i = 1; i <= rosterCount; i++) {
						if ($('#classlistid' + i).is(':checked')) {
							attendeeCount++;
							break;
						}
					};
				}

				if (attendeeCount == 0) {
					okToCopy = false;
					inlineMsg("copyattendeesbtn",'<strong>Copy Failed: </strong>You need to select some attendees to copy.',8,"copyattendeesbtn");
					return false;
				}
				else {
					okToCopy = true;
				}


				if (okToCopy) {
					document.frmrosterlist.action="copy_registrants.asp";
					document.frmrosterlist.submit();
				}
			}
		}

		function pullSeasonClasses()
		{
			// pull new classes for the selected season
			var seasonId = $("#classseasonid").val();
			//alert(seasonId);
			var request = $.ajax({
  				url: "getseasonclasspicks.asp",
  				type: "POST",
  				data: { classseasonid : seasonId },
  				dataType: "html"
			});
 
			request.done(function( msg ) {
  				$( "#classpicks" ).html( msg );
			});
 
			request.fail(function( jqXHR, textStatus ) {
  				alert( "Class request failed: " + textStatus );
			});
		}

		function ViewCart()
		{
			location.href='class_cart.asp';
		}

		function ChangeWaiverOnFile( iClassListId )
		{
			// Fire off the waiver change code without any return handler
			doAjax('setwaiveronfile.asp', 'classlistid=' + iClassListId, '', 'get', '0');
		}

		function RosterDownload( p_classid, p_timeid, fromDate, toDate ) 
		{
			location.href = 'export_roster.asp?classid=' + p_classid + '&timeid=' + p_timeid + '&fromdate=' + fromDate + '&todate=' + toDate;
		}

		function openCustomReports(p_report) 
		{
			w = 900;
			h = 500;
			t = (screen.availHeight/2)-(h/2);
			l = (screen.availWidth/2)-(w/2);
			eval('window.open("../customreports/customreports.asp?cr='+p_report+'&export=Y", "_customreports", "width='+w+',height='+h+',toolbar=0,statusbar=0,scrollbars=1,resizable=1,menubar=0,left=' + l + ',top=' + t + '")');
		}

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			document.getElementById("screenMsg").innerHTML = "";
		}

		function doCalendar( sField ) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmrosterlist", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function applyFilter()
		{
			var okToSubmit = true;
			if ($("#fromDate").val() == '') 
			{
	            $("#fromDate").focus();
				inlineMsg("fromDate",'<strong>Missing Field: </strong>Please enter a From Date before filtering.',8,"fromDate");
				okToSubmit = false;
	        }
			else
			{
				if (! isValidDate($("#fromDate").val()))
				{
					inlineMsg("fromDate",'<strong>Invalid Date: </strong>Please enter a valid date.',8,"fromDate");
					okToSubmit = false;
				}
			}

			if ($("#toDate").val() == '') 
			{
	            $("#toDate").focus();
				inlineMsg("toDate",'<strong>Missing Field: </strong>Please enter an To Date before filtering.',8,"toDate");
				okToSubmit = false;
	        }
			else
			{
				if (! isValidDate($("#toDate").val()))
				{
					inlineMsg("toDate",'<strong>Invalid Date: </strong>Please enter a valid date.',8,"toDate");
					okToSubmit = false;
				}
			}

			if ( okToSubmit == true ) 
			{
				document.frmrosterlist.action = 'view_roster.asp';
				document.frmrosterlist.submit();
			}
		}

		function clearFilter()
		{
			$("#fromDate").val('');
			$("#toDate").val('');
			document.frmrosterlist.action = 'view_roster.asp';
			document.frmrosterlist.submit();
		}

		function toggleAllChecks()
		{
			var toggleVal = $("#checkallbox").is(':checked');
			var rosterCount = parseInt( $("#rostercount").val());

			if (rosterCount > 0) {
				for (var i = 1; i <= rosterCount; i++) {
					$('#classlistid' + i).prop('checked', toggleVal);
				};
			}
		}

	//-->
	</script>

</head>
<body onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<%
'BEGIN: Page Content
response.write "<div id=""content"">"
response.write "  <div id=""centercontent"">"

If CartHasItems() Then 
	response.write "<div id=""topbuttons"">"
	response.write "  <input type=""button"" name=""viewcart"" class=""button"" value=""View Cart"" onclick=""ViewCart();"" />"
	response.write "</div>"
End If 

'BEGIN: Page Title
response.write "<div>"
response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""600"">"
response.write "  <tr>"
response.write "      <td><font size=""+1""><strong>Recreation: Class Roster</strong></font></td>"
response.write "      <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>"
response.write "  </tr>"
response.write "</table>"
response.write "<input type=""button"" name=""returnButton"" id=""returnButton"" value=""<< Back"" class=""button"" onclick=""location.href='class_offerings.asp?classid=" & iClassID & "';"" />"
response.write "</div>"
'END: Page Title

'BEGIN: Class List
response.write "<form name=""frmrosterlist"" id=""frmrosterlist"" action=""move_registrants_cgi.asp"" method=""post"">"
response.write "  <input type=""hidden"" name=""classid"" value=""" & iClassID & """ />"
response.write "  <input type=""hidden"" name=""timeid"" value=""" & iTimeID & """ />"
response.write "<div>"

' Display the class details
DisplayClassDetails iClassID, iTimeID, toDate, fromDate

response.write "</div>"
response.write "<div>"

DisplayClassEventsRoster iClassID, iTimeID, toDate, fromDate

'MOVE STUDENT
response.write vbcrlf & "<div id=""move"" style=""padding:0 0 10px 0;margin:0 0 10px 0;"">"
DisplayMove iClassID
response.write vbcrlf & "</div>"

' Copy to another class as a waitlist
response.write vbcrlf & "<div id=""copyattendees"">"
DisplayCopyTo iClassID, iTimeID
response.write vbcrlf & "</div>"

response.write "</div>"
response.write "</form>"
'END: Class List

response.write "  </div>"
response.write "</div>"
'END: Page Content
		if intMaxEnroll > intEnrolled and intWaitlist > 0 and Session("OrgID") = "60" then
			response.write "<script>window.onload = function () {alert('There is now one or more openings in this class which has a waitlist.  Please contact participants on the waitlist to enroll in the class.');}</script>"
		end if
%>

<!--#Include file="../admin_footer.asp"-->

<%
'Print Roster
response.write "<form name=""frmprint"" id=""frmprint"" action=""print_roster.asp"" method=""post"">"
response.write "  <input type=""hidden"" value=""" & iClassID & "," & iTimeID & """ name=""print_" & iClassID & iTimeID & """ />"
response.write "</form>"

response.write "<form name=""frmAttendanceSheet"" id=""frmAttendanceSheet"" action=""attendance_sheet.asp"" method=""post"">"
response.write "  <input type=""hidden"" value=""" & request("classid") & "," & request("timeid") & """ name=""attendance_" & request("classid") & request("timeid") & """ />"
response.write "  <input type=""hidden"" name=""classid"" value=""" & request("classid") & """ />"
response.write "  <input type=""hidden"" name=""timeid"" value=""" & request("timeid") & """ />"
response.write "</form>"

response.write "</body>"
response.write "</html>"
response.flush 


'------------------------------------------------------------------------------
' void DisplayRegistrationStarts iClassId
'------------------------------------------------------------------------------
Sub DisplayRegistrationStarts( ByVal iClassId )
	Dim sSql, oRs

	sSql = "SELECT registrationstartdate, pricetypename FROM egov_class_pricetype_price C, egov_price_types P "
	sSql = sSql & "WHERE classid = " & iClassId & " AND C.pricetypeid = P.pricetypeid AND registrationstartdate IS NOT NULL "
	sSql = sSql & "ORDER BY displayorder "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write "<tr><td class=""classdetaillabel"" valign=""top"">Registration Starts:</td><td class=""classdetailvalue"">"
		Do While Not oRs.EOF
			response.write oRs("pricetypename") & ": " & FormatDateTime(oRs("registrationstartdate"),1) & "<br />"
			oRs.MoveNext
		Loop 
		response.write "</td></tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void DisplayClassEventsRoster iclassid, itimeid, toDate, fromDate
'------------------------------------------------------------------------------
Sub DisplayClassEventsRoster( ByVal iclassid, ByVal itimeid, ByVal toDate, ByVal fromDate )
	Dim iWaitlistCount, sSql, oRs, sWhereClause, iCounter, sSignupDate

	iWaitlistCount = 0
	iCounter = 0

	If toDate <> "" And fromDate <> "" Then
		sWhereClause = " AND signupdate >= '" & fromDate & "' AND signupdate < '" & DateAdd( "d", 1, CDate(toDate) ) & "' "
	Else
		sWhereClause = ""
	End If 

	' GET ALL CLASS\EVENTS FOR ORG
	sSql = "SELECT * FROM egov_class_roster "
	sSql = sSql & "WHERE classid = " & iclassid & " AND classtimeid = " & itimeid & " "
	sSql = sSql & sWhereClause
	sSql = sSql & " ORDER BY status, signupdate, userlname, userfname"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then

		' IF REGISTRATION REQUIRED SHOW REGISTERED USERS
		If oRs("optionid") = "1" Then

			' DRAW TABLE WITH CLASSES LISTED
			response.write vbcrlf & "<table cellpadding=""5"" cellspacing=""0"" border=""0"" class=""rostertable style-alternate sortable"" id=""rostertable"">"
			
			' HEADER ROW
			response.write vbcrlf & "<tr>"
			response.write "<th><input type=""checkbox"" id=""checkallbox"" onchange=""toggleAllChecks();"" /></th>"
			response.write "<th class=""sortable"">Student Name</th>"
			response.write "<th class=""sortable"">Signup Date</th>"
			response.write "<th>Age</th>"
			response.write "<th>Residency</th>"
			response.write "<th>Waivers<br />On File</th>"
			response.write "<th class=""sortable"">Contact<br />Information</th>"
			response.write "<th class=""sortable"">Emergency<br />Contact</th>"
			response.write "<th class=""sortable"">Status</th>"
			response.write "<th>Evaluation<br />Sent</th>"
			response.write "<th>&nbsp;</th>"
			response.write "</tr>"

			' LOOP THRU AND DISPLAY CLASS ROSTER
			Do While Not oRs.EOF
				iCounter = iCounter + 1
				response.write vbcrlf & "<tr>"
				response.write "<td><input type=""checkbox"" class=""listcheck"" value=""" & oRs("classlistid") & """ id=""classlistid" & iCounter & """ name=""classlistid"" /></td>"
				response.write "<td nowrap=""nowrap"">" 
				If oRs("isdeleted") Then
					' cannot edit deleted users here
					response.write oRs("lastname") & ", " & oRs("firstname") & " (deleted)"
				Else 
					if not isnull(oRs("familymemberuserid")) and not isnull(oRs("familyid")) then
						If CLng(oRs("familymemberuserid")) = CLng(oRs("familyid")) Then
							response.write "<a href=""../dirs/update_citizen.asp?userid=" & oRs("userid") & """>"
						Else
							response.write "<a href=""../dirs/manage_family_member.asp?u=" & oRs("familymemberuserid") & "&iReturn=-1"">"
						End If 
						If oRs("lastname") <> "" Or oRs("firstname") <> "" Then 
							response.write oRs("lastname") & ", " & oRs("firstname") 
						Else
							response.write "Missing Name (Click to Update)"
						End If 
					End If 
					response.write "</a>"
				End If
				response.write "</td>"
				
				sSignupDate = ""
				If Not IsNull( oRs("signupdate") ) Then
					sSignupDate = month(oRs("signupdate")) & "/" & day(oRs("signupdate")) & "/" & year(oRs("signupdate")) & "<br />"
					sSignupDate = sSignupDate & GetTimeFormat(oRs("signupdate"))
				End If 
				response.write "<td align=""center"">" & sSignupDate & "</td>"

				iAge = GetCitizenAge( oRs("birthdate") )
				If iAge >= 18 Then 
					iAge = "Adult"
				End If 
				response.write "<td align=""center"">" & iAge & "</td>"

				If oRs("residenttype") <> "R" Then
					response.write "<td>" & oRs("description") & "</td>"
				Else
					if lcl_orghasfeature_residency_verification then
						If Not oRs("residencyverified") Then 
							response.write "<td>(not verified)</td>"
						Else
							response.write "<td>&nbsp;</td>"
						End If 
					Else
						response.write "<td>&nbsp;</td>"
					End If 
				End If 

				' Waiver on File
				response.write "<td align=""center""><input type=""checkbox"" class=""listcheck"" name=""waivercheck"" value=""" & oRs("classlistid") & """"
				If oRs("waiveronfile") Then
					response.write " checked=""checked"" "
				End If 
				response.write " onClick='ChangeWaiverOnFile(" & oRs("classlistid") & ");' /></td>"
				

				'response.write "<td nowrap=""nowrap"">" & FormatPhone(oRs("userhomephone")) & "</td>"
				response.write "<td nowrap=""nowrap"" valign=""top"">" & GetRosterPhone( oRs("familymemberuserid") ) 
				If oRs("useremail") <> "" Then 
					response.write "<br /><a href=""mailto:" & oRs("useremail") & """>" & oRs("useremail") & "</a>"
				End If 
				response.write "</td>"

				response.write "<td valign=""top"" nowrap=""nowrap"">" 
				ShowEmergencyContactInfo oRs("familymemberuserid")
				response.write"</td>"

				If oRs("status") = "WAITLIST" Then
					iWaitlistCount = iWaitlistCount + 1
					response.write "<td nowrap=""nowrap"">" &  oRs("status") & " (" & iWaitlistCount & ")" & "</td>"
				Else 
					response.write "<td nowrap=""nowrap"">" &  oRs("status") 
					If oRs("isdropin") Then
						response.write "<br />(" & oRs("dropindate") & ")"
					End If 
					response.write "</td>"
				End If 

				response.write "<td>" & oRs("evaluationsent") & "</td>"

				response.write "<td nowrap=""nowrap"">"
				response.write "<input type=""button"" name=""viewReceipt"" id=""viewReceipt"" value=""Receipt"" class=""button"" onclick=""location.href='view_receipt.asp?iPaymentId=" & oRs("paymentid") & "'"" />"

				If (UCase(oRs("status")) = "ACTIVE" Or UCase(oRs("status")) = "DROPIN") Then 
 					'Only active participants and Drop Ins can drop (Drop Ins to correct mistakes)
					response.write "&nbsp;<input type=""button"" name=""dropUser"" id=""dropUser"" value=""Drop"" class=""button"" onclick=""location.href='drop_registrant_form.asp?classid=" & iclassid & "&timeid=" & itimeid & "&classlistid=" & oRs("classlistid") & "&iqty=1'"" />"
				End If 

				If UCase(oRs("status")) = "WAITLIST" Then 
				   response.write "&nbsp;<input type=""button"" name=""removebtn"" id=""removebtn"" value=""Remove"" class=""button"" onclick=""location.href='waitlist_removal_form.asp?classid=" & iclassid & "&timeid=" & itimeid & "&classlistid=" & oRs("classlistid") & "'"" />"

 					'Only those on the waitlist can become active
				   response.write "&nbsp;<input type=""button"" name=""activatebtn"" id=""activatebtn"" value=""Activate"" class=""button"" onclick=""location.href='change_status_form.asp?classid=" & iclassid & "&timeid=" & itimeid & "&classlistid=" & oRs("classlistid") & "'"" />"
				End If 

				response.write "</td>"
				response.write "</tr>"
				response.flush 

				oRs.MoveNext
			Loop 

		Else
			'NON-REGISTERED USERS SHOW PAID USERS
				
			'DRAW TABLE WITH CLASSES LISTED
			response.write "<table cellpadding=""5"" cellspacing=""0"" border=""0"" class=""rostertable style-alternate sortable-onload-2"" width=""100%"" id=""rostertable"">"
			
			'HEADER ROW
			response.write vbcrlf & "<tr>"
			response.write "<th>&nbsp;</th>"
			response.write "<th class=""sortable"">Payee Name</th>"
			response.write "<th>Age</th>"
			response.write "<th>Residency</th>"
			response.write "<th>Waivers<br />On File</th>"
			response.write "<th class=""sortable"">Contact<br />Information</th>"
			response.write "<th class=""sortable"">Emergency<br />Contact</th>"
			response.write "<th class=""sortable"">Status</th>"
			response.write "<th class=""sortable"">Qty</th>"
			response.write "<th>Evaluation<br />Sent</th>"
			response.write "<th>&nbsp;</th>"
			response.write "</tr>"


			'LOOP THRU AND DISPLAY CLASS ROSTER
			Do While Not oRs.EOF
				response.write vbcrlf & "<tr>"
				response.write "<td><input type =""checkbox"" class=""listcheck"" value=""" & oRs("classlistid") & """ id=""classlistid" & iCounter & """ name=""classlistid""></td>"
				response.write "<td nowrap=""nowrap"">"
				response.write "<a href=""../dirs/update_citizen.asp?userid=" & oRs("userid") & """>"
				response.write oRs("userlname") & ", " & oRs("userfname") & "</a>"
				response.write "</td>"

				iAge = GetCitizenAge( oRs("birthdate") )
				If iAge >= 18 Then 
			  		iAge = "Adult"
				End If 
				response.write "<td align=""center"">" & iAge & "</td>"

				If oRs("residenttype") <> "R" Then
  					response.write "<td>" & oRs("description") & "</td>"
				Else
					If lcl_orghasfeature_residency_verification Then 
						If Not oRs("residencyverified") Then 
							response.write "<td>(not verified)</td>"
						Else
							response.write "<td>&nbsp;</td>"
						End If 
					Else
						response.write "<td>&nbsp;</td>"
					End If 
				End If 

				'Waiver on File
				response.write "<td align=""center"">"
				response.write "<input type=""checkbox"" class=""listcheck"" name=""waivercheck"" value=""" & oRs("classlistid") & """"

				If oRs("waiveronfile") Then
    				response.write " checked=""checked"" "
 				End If 

				response.write " onClick='ChangeWaiverOnFile(" & oRs("classlistid") & ");' /></td>"

				response.write "<td nowrap=""nowrap"" valign=""top"">" & GetRosterPhone( oRs("userid") ) 
				If oRs("useremail") <> "" Then 
					response.write "<br /><a href=""mailto:" & oRs("useremail") & """>" & oRs("useremail") & "</a>"
				End If 
				response.write "</td>"

				response.write "<td valign=""top"" nowrap=""nowrap"">"
				ShowEmergencyContactInfo oRs("userid")
				response.write "</td>"

				If oRs("status") = "WAITLIST" Then
					iWaitlistCount = iWaitlistCount + 1
					response.write "<td nowrap=""nowrap"">" &  oRs("status") & " (" & iWaitlistCount & ")" & "</td>"
				Else 
					response.write "<td nowrap=""nowrap"">" &  oRs("status") 
					If oRs("isdropin") Then
						response.write "<br />(" & oRs("dropindate") & ")"
					End If 
					response.write "</td>"
				End If 

				response.write "<td align=""center"">" & oRs("quantity") & "</td>"
				response.write "<td>" & oRs("evaluationsent") & "</td>"
				response.write "<td nowrap=""nowrap""><a href=""view_receipt.asp?iPaymentId=" & oRs("paymentid") & """ >Receipt</a>"'" | <a href=""drop_registrant_form.asp?classid=" & iclassid & "&timeid=" & itimeid & "&classlistid=" & oRs("classlistid") & "&iqty=" & oRs("quantity") & """>Drop</a> | <a href=""change_status_form.asp?classid=" & iclassid & "&timeid=" & itimeid & "&classlistid=" & oRs("classlistid") & """>Status</a></td>"
				If oRs("status") = "ACTIVE" Then
					' Only active participants can drop
					response.write "| <a href=""drop_registrant_form.asp?classid=" & iclassid & "&timeid=" & itimeid & "&classlistid=" & oRs("classlistid") & "&iqty=" & oRs("quantity") & """ >Drop</a>"
				End If 
				If oRs("status") = "WAITLIST" Then
					response.write "| <a href=""waitlist_removal_form.asp?classid=" & iclassid & "&timeid=" & itimeid & "&classlistid=" & oRs("classlistid") & """ >Remove</a>"
					' Only those on the waitlist can become active
					response.write "| <a href=""change_status_form.asp?classid=" & iclassid & "&timeid=" & itimeid & "&classlistid=" & oRs("classlistid") & """>Activate</a> "
				End If 
				response.write "</td>"
				response.write "</tr>"

				oRs.MoveNext
			Loop 

		End If

		'ClOSE TABLE AND FREE OBJECTS
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "<input type=""hidden"" id=""rostercount"" name=""rostercount"" value=""" & iCounter & """ />"

		'MOVE STUDENT
		'response.write "<div id=""move"" style=""padding:0 0 10px 0;margin:0 0 10px 0;"">"
		'DisplayMove iclassid
		'response.write "</div>"
	
	Else
		' NO CLASS\EVENTS WERE FOUND
		response.write "<font style=""font-size:10px;"" color=""red""><b>No purchases/registrations have been made for this activity.</b></font>"
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
' void DisplayClassDetails classid, timeid, toDate, fromDate
'------------------------------------------------------------------------------
 Sub DisplayClassDetails( ByVal iClassid, ByVal iTimeid, ByVal toDate, ByVal fromDate )
	Dim sSql, oRs, arrDetails, arrDetailLabels

	'Initialize values
	arrDetails = Array("registrationenddate","startdate","enddate","evaluationdate","minage","maxage")
	arrDetailLabels = Array("Registration Ends","Start Date","End Date","Evaluation Date","Minimum Age","Maximum Age")

	'Get selected class information
	sSql = "SELECT egov_class.classid, egov_class_time.timeid, classname, classseasonid, locationid, registrationenddate, "
	sSql = sSql & "evaluationdate, minage, maxage, egov_class_time.instructorid, "
	sSql = sSql & "ISNULL(egov_class.startdate,0) AS startdate, ISNULL(egov_class.enddate,0) AS enddate, "
	sSql = sSql & "ISNULL(egov_class.imgurl,'EMPTY') AS imgurl, (firstname + ' ' + lastname) as Instructor "
	sSql = sSql & "FROM egov_class "
	sSql = sSql & "LEFT JOIN egov_class_time ON egov_class.classid = egov_class_time.classid "
	sSql = sSql & "LEFT JOIN egov_class_instructor ON egov_class_time.instructorid = egov_class_instructor.instructorid "
	sSql = sSql & "WHERE egov_class.classid = " & iClassid & " AND egov_class_time.timeid = " & iTimeid
	sSql = sSql & " ORDER BY noenddate DESC, startdate"
	'response.write sSql
	'response.end

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

    ' DISPLAY ITEM INFORMATION
    If Not oRs.EOF Then
		'WRITE TITLE
		response.write vbcrlf & "<h3>" &  oRs("classname") & " &nbsp; ( " & GetActivityNo( iTimeid ) & " )</h3><br />"

		response.write vbcrlf & "<table border=""0"" cellspacing=""0"" cellpadding=""0"">"
		response.write vbcrlf & "<tr>"
		response.write "<td nowrap=""nowrap"">"

		'Roster Print
		response.write "<input type=""button"" class=""button"" name=""rosterprint"" value=""Printable Roster"" onclick=""GoToPrint( );"" />&nbsp;"

		'Roster Mail button
		response.write "<input type=""button"" class=""button"" name=""rosteremail"" value=""Send Mail To Attendees"" onclick=""javascript:location.href='class_rosteremail.asp?classid=" & iClassid & "&timeid=" & iTimeid & "'"" />&nbsp;"

		'Registration
		response.write "<input type=""button"" class=""button"" name=""register"" value=""Registration"" onclick=""javascript:location.href='class_signup.asp?classid=" & iClassid & "&timeid=" & iTimeid & "'"" />&nbsp;"

		'Attendance Sheet
		response.write "<input type=""button"" class=""button"" name=""attendance"" value=""Attendance Sheet"" onclick=""GoToAttendance( );"" />&nbsp;"

		'Download Roster
		response.write "<input type=""button"" class=""button"" name=""export_roster"" value=""Download Roster"" onclick=""RosterDownload('" & iClassid & "','" & iTimeid & "', '" & fromDate & "', '" & toDate & "' )"" />"
		
		'Download Team Roster
		 if  lcl_orghasfeature_customreports_classesevents_teamroster _
   			AND lcl_userhaspermission_customreports_classesevents_teamroster _
   			AND lcl_orghasfeature_custom_registration_craigco then
       		lcl_teamroster_exists = checkTeamRosterExists(oRs("classid"), oRs("timeid"))

       if lcl_teamroster_exists then
       			session("CR_CLASSEVENTS_TEAMROSTER") = buildTeamRosterQuery(oRs("classid"), oRs("timeid"))
	 	   	   response.write "&nbsp;<input type=""button"" class=""button"" name=""export_roster"" value=""Download Team Roster"" onclick=""openCustomReports('CLASSEVENTS_TEAMROSTER')"" />" & vbcrlf
       end if
 		end if

		response.write "<td>"
		response.write "</tr>"

		response.write vbcrlf & "</table>"

		'Display Details
		response.write vbcrlf & "<div align=""left"">"
		response.write vbcrlf & "<fieldset><legend><strong>Details&nbsp;</strong></legend>"
		response.write vbcrlf & "<table style=""width:400px;align=left; margin: 5px 5px 5px 5px;"">"

		'Show the Season
		response.write vbcrlf & "<tr>"
		response.write "<td class=""classdetaillabel"">Season: </td>"
		response.write "<td class=""classdetailvalue"">" & GetSeasonName( oRs("classseasonid") ) & "</td>"
		response.write "</tr>"

		'Show the Location
		response.write vbcrlf & "<tr>"
		response.write "<td class=""classdetaillabel"">Location: </td>"
		response.write "<td class=""classdetailvalue"">" & GetLocationName( oRs("locationid") ) & "</td>"
		response.write "</tr>"

		'Show the Registration Start Dstes
		DisplayRegistrationStarts iClassid
		
		' DISPLAY DETAILS VALUE PAIR
		For d = 0 to UBound(arrDetails)
			If Trim(oRs(arrDetails(d))) <> "" And Not IsNull( oRs(arrDetails(d)) ) Then

				' IF DATE THEN FORMAT
				If IsDate(oRs(arrDetails(d))) Then
					' FORMAT DATE
					sValue = FormatDateTime(oRs(arrDetails(d)),1)
				Else
					' DISPLAY STORED VALUE UNFORMATTED
					sValue = oRs(arrDetails(d))
				End If

				response.write vbcrlf & "<tr>"
				response.write "<td class=""classdetaillabel"">" & arrDetailLabels(d) & ": </td>"
				response.write "<td class=""classdetailvalue"">" & sValue & "</td>"
				response.write "</tr>"
			
			End If
		Next

		'DISPLAY INSTRUCTOR
		If Trim(oRs("Instructor")) <> "" AND NOT IsNull(oRs("Instructor")) Then
			response.write vbcrlf & "<tr>"
			response.write "<td class=""classdetaillabel"" >Instructor: </td>"
			response.write "<td><a href=""instructor_info.asp?iID=" & oRs("instructorid")& """ target=_NEW>" & oRs("Instructor") & "</a></td>"
			response.write "</tr>"
		End If

		'Display Waiver Links
		response.write vbcrlf & "<tr>"
		response.write "<td class=""classdetaillabel"" >Waivers: </td>"
		response.write "<td>" 
  		ShowClassWaiverLinks iClassid 
		response.write "</td>"
		response.write "</tr>"

		' DISPLAY TIMES
		response.write vbcrlf & "<br />"
		response.write vbcrlf & "</table>"
		
		' Availability
		response.write vbcrlf & "<p><strong>Availability:</strong><br />"
		DisplayClassActivities iClassid, iTimeid, False   ' In class_global_functions.asp
		response.write "</p>"


		response.write vbcrlf & "<input type=""button"" class=""button"" name=""fixenrollment"" value=""Correct Participant Count"" onclick=""javascript:location.href='class_fixenrollment.asp?classid=" & iClassid & "&timeid=" & iTimeid & "'"" /><br /><br />"

		' Put registration date filter here
		If bHasRegistrationDateFilter Then 
			response.write vbcrlf & "<strong>Filter by Registration Date</strong><br />"
			response.write vbcrlf & "From: "
			response.write "<input type=""text"" id=""fromDate"" name=""fromDate"" value=""" & fromDate & """ size=""10"" maxlength=""10"" />"
			response.write "<a href=""javascript:void doCalendar('fromDate');""><img src=""../images/calendar.gif"" border=""0"" /></a>"
			response.write "&nbsp; To: "
			response.write "<input type=""text"" id=""toDate"" name=""toDate"" value=""" & toDate & """ size=""10"" maxlength=""10"" />"
			response.write "<a href=""javascript:void doCalendar('toDate');""><img src=""../images/calendar.gif"" border=""0"" /></a>"
			response.write "&nbsp;"
			' Date range pick - in common.asp
			DrawDateChoices "Date", 0

			response.write "&nbsp;&nbsp;<input type=""button"" class=""button"" value=""Apply Filter"" onclick=""applyFilter();"" />"
			response.write "&nbsp;&nbsp;<input type=""button"" class=""button"" value=""Clear Filter"" onclick=""clearFilter();"" />"
		End If 

		response.write vbcrlf & "</fieldset></div>"

	End If

	' CLOSE OBJECTS
	oRs.Close 
	Set oRs = Nothing 

 End Sub


'------------------------------------------------------------------------------
Sub DisplayCopyTo( ByVal iClassId, ByVal iTimeId )
	dim iClassSeasonId

	' get the classseasonid of the class
	iClassSeasonId = getClassSeasonId( iClassId )

	'response.write "<br />"
	response.write vbcrlf & "<fieldset>"
	response.write vbcrlf & "<legend><strong>Copy selected registrants to:</strong></legend>"

	' get the season picks
	ShowSeasonPicks iClassSeasonId	

	' get the classes for the season
	response.write "<span id=""classpicks"">"
	ShowClassPicks iClassSeasonId
	response.write "</span>"

	response.write " <input class=""button"" type=""button"" onClick=""confirm_copy();"" id=""copyattendeesbtn"" name=""copyattendeesbtn"" value=""Copy"" />"

	response.write vbcrlf & "</fieldset>"

End Sub


'------------------------------------------------------------------------------
Sub DisplayMove( ByVal iClassId )

	response.write "<br />"
	response.write vbcrlf & "<fieldset>"
	response.write vbcrlf & "<legend><strong>Transfer selected registrants to:</strong></legend>"
	'response.write "<p><input class=""button"" type=""button"" onClick=""confirm_move();"" name=""complete"" value=""Transfer selected registrants to:"" /></p>"
	'response.write "<p>"

	DisplayClassActivityChoices iClassId

	response.write " &nbsp; <input class=""button"" type=""button"" onClick=""confirm_move();"" name=""complete"" value=""Transfer"" />"

	'response.write "</p>"
	response.write vbcrlf & "</fieldset>"

End Sub 


'------------------------------------------------------------------------------
Sub DisplayClassActivityChoices( ByVal iClassId )
	Dim sSql, oRs, sMax

	sSql = "SELECT timeid, classname, activityno, max, enrollmentsize "
	sSql = sSql & " FROM egov_class C, egov_class_time T "
	sSql = sSql & " WHERE C.classid = T.classid AND C.classid = " & iClassId
	sSql = sSql & " ORDER BY activityno"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""newtimeid"" id=""newtimeid"">"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("timeid")& """>" & oRs("classname") & " &mdash; " & oRs("activityno")
		If IsNull(oRs("max")) Then
			sMax = "n/a"
		Else 
			sMax = oRs("max")
		End If 
		response.write " (max: " & sMax & ", enrld: " & oRs("enrollmentsize") & ")</option>"
		oRs.MoveNext
	Loop
	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
Sub DisplayClassEvents( ByVal iorgid )
	Dim sSql, oRs

	' GET ALL CLASS\EVENTS FOR ORG
	sSql = "SELECT classid, classname, isparent FROM egov_roster_list2 "
	sSql = sSql & "WHERE orgid = " & iorgid & " AND parentclassid IS NULL ORDER BY classname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		' LOOP THRU AND DISPLAY CLASS\EVENTS
		Do While Not oRs.EOF
			DisplayTimes oRs("classid"), oRs("classname")

			' DISPLAY CHILD CLASS\EVENTS
			If oRs("isparent") Then
				DisplayChildClassEvents iorgid, oRs("classid")
			End If

			oRs.MoveNext
		Loop 
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
Sub DisplayChildClassEvents( ByVal iorgid, ByVal iparentid )
	Dim oRs, sSql

	' GET ALL CLASS\EVENTS FOR ORG
	sSql = "SELECT classid, classname FROM egov_roster_list2 "
	sSql = sSql & "WHERE orgid = " & iorgid & " AND parentclassid = " & iparentid & " ORDER BY classname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then

		' LOOP THRU AND DISPLAY CHILD CLASS\EVENTS
		Do While Not oRs.EOF
			' DISPLAY CHILD INFORMATION
			DisplayTimes oRs("classid"), oRs("classname")
			
			oRs.MoveNext
		Loop 
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
Function fnGetPercentFull( ByVal sMax, ByVal sCurrent )

	If IsNumeric(sMax) AND IsNumeric(sCurrent) Then
 		 fnGetPercentFull = FormatNumber(clng(sCurrent) / clng(sMAX) * 100,0)  
	Else
	 	 fnGetPercentFull = "n/a"
	End If

End Function


'------------------------------------------------------------------------------
Sub DisplayTimes( ByVal iClassId, ByVal sClassName )
	Dim sSql, oRs

	sSql = "SELECT  starttime, endtime, min, max, timeid "
	sSql = sSql & "FROM egov_class_time WHERE classid = " & iClassId 
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	' INSTRUCTOR INFORMATION
	If Not oRs.EOF Then

		' DISPLAY CLASS INFORMATION
		Do While Not oRs.EOF 
			response.write "<option value=""" &iclassid & "," & oRs("timeid")& """>" & sClassName & " --- (" & oRs("starttime") & " - " & oRs("endtime") & " " & fnGetTimeDaysofWeek(iclassid) & ")</option>"
			oRs.MoveNext
		Loop
	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------
Function fnGetTimeDaysofWeek( ByVal iclassid )
	Dim sSql, oRs
	
	sReturnValue = ""

	' GET THE DAY OF THE WEEK VALUES FOR THE SPECIFIED
	sSql = "SELECT dayofweek FROM egov_class_dayofweek WHERE classid = " & iClassId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	' IF NOT EMPTY
	If Not oRs.EOF Then

		' LOOP THRU AVAILABLE DAYS OF THE WEEK
		Do While Not oRs.EOF 
			sReturnValue = sReturnValue &  WeekDayName( oRs("dayofweek"),True ) & " "
			oRs.MoveNext
		Loop
	Else
		' NO DAYS FOUND
		sReturnValue = ""
	End If

	oRs.Close
	Set oRs = Nothing

	' RETURN DAYS OF THE WEEK
	fnGetTimeDaysofWeek = Trim(sReturnValue)

End Function


'------------------------------------------------------------------------------
Sub DisplayClassTimes2( ByVal iTimeID )
	Dim sSql, oRs

	sSql = "SELECT starttime, endtime, min, max "
	sSql = sSql & "FROM egov_class_time WHERE timeid = " & itimeId 
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	' INSTRUCTOR INFORMATION
	If Not oRs.EOF Then

		' DISPLAY CLASS INFORMATION
		Do While Not oRs.EOF 
			response.write oRs("starttime") & " - " & oRs("endtime") & " " & fnGetTimeDaysofWeek(iclassid) 
			oRs.MoveNext
		Loop
	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------
Function getClassSeasonId( ByVal iClassId )
	Dim sSql, oRs

	sSql = "SELECT classseasonid FROM egov_class WHERE classid = " & iClassId 
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		getClassSeasonId = oRs("classseasonid")
	Else
		getClassSeasonId = 0
	End If

	oRs.Close
	Set oRs = Nothing

End Function


'------------------------------------------------------------------------------
Function buildTeamRosterQuery( ByVal p_classid, ByVal p_timeid )
	Dim lcl_query

	lcl_query = "SELECT "
	lcl_query = lcl_query & " isnull(u.userfname,'') as userfname, "
	lcl_query = lcl_query & " isnull(u.userlname,'') as userlname, "
	lcl_query = lcl_query & " '' as userage, "
	lcl_query = lcl_query & " isnull(u.birthdate,'') as birthdate, "
	lcl_query = lcl_query & " isnull(cl.rostergrade,'') as rostergrade, "
	lcl_query = lcl_query & " (SELECT isnull(c.teamreg_tshirt_enabled,1) FROM egov_class c WHERE c.classid = cl.classid) AS teamreg_tshirt_enabled, "
	lcl_query = lcl_query & " (SELECT isnull(c.teamreg_pants_enabled,0) FROM egov_class c WHERE c.classid = cl.classid) AS teamreg_pants_enabled, "
	lcl_query = lcl_query & " isnull(cl.rostershirtsize,'') as rostershirtsize, "
	lcl_query = lcl_query & " isnull(cl.rosterpantssize,'') as rosterpantssize, "
	lcl_query = lcl_query & " u.userhomephone, "
	lcl_query = lcl_query & " (SELECT isnull(u2.userfname,'') FROM egov_users u2 WHERE u2.userid = f.belongstouserid) as parentfirstname, "
	lcl_query = lcl_query & " (SELECT isnull(u2.userlname,'') FROM egov_users u2 WHERE u2.userid = f.belongstouserid) as parentlastname, "
	lcl_query = lcl_query & " (SELECT isnull(u2.userhomephone,'') FROM egov_users u2 WHERE u2.userid = f.belongstouserid) as parentphone, "
	lcl_query = lcl_query & " dbo.fn_BuildAddress("
	lcl_query = lcl_query &                       "isnull(u.userstreetnumber,''),"
	lcl_query = lcl_query &                       "isnull(u.userstreetprefix,''),"
	lcl_query = lcl_query &                       "isnull(u.useraddress,''),"
	lcl_query = lcl_query &                       "'',''"
	lcl_query = lcl_query &                     ") AS useraddress_complete, "
	lcl_query = lcl_query & " isnull(cl.rostercoachtype,'') as rostercoachtype, "
	lcl_query = lcl_query & " isnull(cl.rostervolunteercoachname,'') as rostervolunteercoachname, "
	lcl_query = lcl_query & " isnull(cl.rostervolunteercoachdayphone,'') as rostervolunteercoachdayphone, "
	lcl_query = lcl_query & " isnull(cl.rostervolunteercoachcellphone,'') as rostervolunteercoachcellphone, "
	lcl_query = lcl_query & " isnull(cl.rostervolunteercoachemail,'') as rostervolunteercoachemail "
	lcl_query = lcl_query & " FROM egov_class_list cl "
	lcl_query = lcl_query &      " LEFT OUTER JOIN egov_familymembers f ON cl.familymemberid = f.familymemberid "
	lcl_query = lcl_query &                  " AND cl.userid = f.belongstouserid "
	lcl_query = lcl_query &      " LEFT OUTER JOIN egov_users u ON f.userid = u.userid "
	lcl_query = lcl_query & " WHERE cl.status = 'ACTIVE' "
	lcl_query = lcl_query & " AND cl.classid = " & p_classid

	If p_timeid <> "" Then 
		lcl_query = lcl_query & " AND cl.classtimeid = " & p_timeid
	End If 

	lcl_query = lcl_query & " ORDER BY u.userlname, u.userfname "

	buildTeamRosterQuery = lcl_query

End Function 

'------------------------------------------------------------------------------
function checkTeamRosterExists(iClassID, iTimeID)
  dim lcl_return, sClassID, sTimeID

  lcl_return = false
  sClassID   = 0
  sTimeID    = 0

  if iClassID <> "" then
     sClassID = clng(iClassID)
  end if

  if iTimeID <> "" then
     sTimeID = clng(iTimeID)
  end if

  sSQL = "SELECT COUNT(u.userid) as roster_exists "
  'sSQL = sSQL & " isnull(u.userfname,'') as userfname, "
  'sSQL = sSQL & " isnull(u.userlname,'') as userlname, "
  'sSQL = sSQL & " '' as userage, "
  'sSQL = sSQL & " isnull(u.birthdate,'') as birthdate, "
  'sSQL = sSQL & " isnull(cl.rostergrade,'') as rostergrade, "
  'sSQL = sSQL & " (SELECT isnull(c.teamreg_tshirt_enabled,1) FROM egov_class c WHERE c.classid = cl.classid) AS teamreg_tshirt_enabled, "
  'sSQL = sSQL & " (SELECT isnull(c.teamreg_pants_enabled,0) FROM egov_class c WHERE c.classid = cl.classid) AS teamreg_pants_enabled, "
  'sSQL = sSQL & " isnull(cl.rostershirtsize,'') as rostershirtsize, "
  'sSQL = sSQL & " isnull(cl.rosterpantssize,'') as rosterpantssize, "
  'sSQL = sSQL & " u.userhomephone, "
  'sSQL = sSQL & " (SELECT isnull(u2.userfname,'') FROM egov_users u2 WHERE u2.userid = f.belongstouserid) as parentfirstname, "
  'sSQL = sSQL & " (SELECT isnull(u2.userlname,'') FROM egov_users u2 WHERE u2.userid = f.belongstouserid) as parentlastname, "
  'sSQL = sSQL & " (SELECT isnull(u2.userhomephone,'') FROM egov_users u2 WHERE u2.userid = f.belongstouserid) as parentphone, "
  'sSQL = sSQL & " dbo.fn_BuildAddress( "
  'sSQL = sSQL &                     " isnull(u.userstreetnumber,''), "
  'sSQL = sSQL &                     " isnull(u.userstreetprefix,''), "
  'sSQL = sSQL &                     " isnull(u.useraddress,''), "
  'sSQL = sSQL &                     " '','' "
  'sSQL = sSQL &                     ") AS useraddress_complete, "
  'sSQL = sSQL & " isnull(cl.rostercoachtype,'') as rostercoachtype, "
  'sSQL = sSQL & " isnull(cl.rostervolunteercoachname,'') as rostervolunteercoachname, "
  'sSQL = sSQL & " isnull(cl.rostervolunteercoachdayphone,'') as rostervolunteercoachdayphone, "
  'sSQL = sSQL & " isnull(cl.rostervolunteercoachcellphone,'') as rostervolunteercoachcellphone, "
  'sSQL = sSQL & " isnull(cl.rostervolunteercoachemail,'') as rostervolunteercoachemail "
  sSQL = sSQL & " FROM egov_class_list cl "
  sSQL = sSQL & "   LEFT OUTER JOIN egov_familymembers f "
  sSQL = sSQL &                   " ON cl.familymemberid = f.familymemberid "
  sSQL = sSQL &                   " AND cl.userid = f.belongstouserid "
  sSQL = sSQL & "   LEFT OUTER JOIN egov_users u "
  sSQL = sSQL &                   " ON f.userid = u.userid "
  sSQL = sSQL & " WHERE cl.status = 'ACTIVE' "
  sSQL = sSQL & " AND cl.classid = " & sClassID

  if sTimeID > 0 then
     sSQL = sSQL & " AND cl.classtimeid = " & sTimeID
  end if

 	set oCheckTeamRosterExists = Server.CreateObject("ADODB.Recordset")
	 oCheckTeamRosterExists.Open sSQL, Application("DSN"), 0, 1

  if not oCheckTeamRosterExists.eof then
     if oCheckTeamRosterExists("roster_exists") > 0 then
        lcl_return = true
     end if
  end if

  oCheckTeamRosterExists.close
  set oCheckTeamRosterExists = nothing

  checkTeamRosterExists = lcl_return

end function


'------------------------------------------------------------------------------
' void ShowSeasonPicks iClassSeasonId
'------------------------------------------------------------------------------
Sub ShowSeasonPicks( ByVal iClassSeasonId )
	Dim sSql, oRs

	sSql = "SELECT C.classseasonid, C.seasonname FROM egov_class_seasons C, egov_seasons S  "
	sSql = sSql & " WHERE C.isclosed = 0 AND C.seasonid = S.seasonid AND orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY C.seasonyear DESC, S.displayorder DESC, C.seasonname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""classseasonid"" id=""classseasonid"" onchange=""pullSeasonClasses();"">" 

		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("classseasonid") & """ "  
			If CLng(iClassSeasonId) = CLng(oRs("classseasonid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("seasonname") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------
Sub ShowClassPicks( ByVal iClassSeasonId )
	Dim sSql, oRs

	sSql = "SELECT T.timeid, C.classname, T.activityno, C.classseasonid "
	sSql = sSql & "FROM egov_class_time T, egov_class C, egov_class_status S "
	sSql = sSql & "WHERE C.classid = T.classid AND C.statusid = S.statusid "
	sSql = sSql & "AND C.orgid = " & SESSION("orgid")
	sSql = sSql & " AND C.classseasonid = " & iClassSeasonId
	sSql = sSql & " AND S.iscancelled = 0 AND T.iscanceled = 0 "
	sSql = sSql & "ORDER BY C.classname, T.activityno"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		response.write vbcrlf & " <select name=""classtimeid"" id=""classtimeid"">" 

		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("timeid") & """ "
			response.write ">" & oRs("classname") & " - " & oRs("activityno") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub


%>
