<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitinspectionedit.asp
' AUTHOR: Steve Loar
' CREATED: 07/10/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Edits permit inspections
'
' MODIFICATION HISTORY
' 1.0   07/10/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitInspectionId, iPermitId, sSql, oRs, sPermitInspectionType, sInspectionDescription, bIsReinspection
Dim sRequestReceived, sRequestedDate, sScheduledDate, sInspectedDate, sContactPhone, sContact, sInspectedtime
Dim sScheduledtime, sRequestedTime, sRequestedAmPm, iInspectionStatusId, iInspectorUserId, sSchedulingNotes
Dim sScheduledAmPm, sInspectedAmPm, bIsFinal, bPermitIsCompleted, bIsOnHold

iPermitInspectionId = CLng(request("permitinspectionid"))

sSql = "SELECT permitid, permitinspectiontype, inspectiondescription, ISNULL(inspectoruserid,0) AS inspectoruserid, isfinal, "
sSql = sSql & " inspectionstatusid, requestreceiveddate, requesteddate, requestedtime, ISNULL(requestedampm,'') AS requestedampm, "
sSql = sSql & " scheduleddate, scheduledtime, ISNULL(scheduledampm, '') AS scheduledampm, inspecteddate, inspectedtime, ISNULL(inspectedampm,'') AS inspectedampm, "
sSql = sSql & " contact, contactphone, isreinspection, schedulingnotes "
sSql = sSql & " FROM egov_permitinspections WHERE permitinspectionid = " & iPermitInspectionId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then
	iPermitId = oRs("permitid")
	sPermitInspectionType = oRs("permitinspectiontype")
	sInspectionDescription = oRs("inspectiondescription")
	bIsReinspection = oRs("isreinspection")
	sRequestReceived = oRs("requestreceiveddate")
	sRequestedDate = oRs("requesteddate")
	sScheduledDate = oRs("scheduleddate")
	sInspectedDate = oRs("inspecteddate")
	sContactPhone = oRs("contactphone")
	sContact = oRs("contact")
	sInspectedtime = oRs("inspectedtime")
	sScheduledtime = oRs("scheduledtime")
	sRequestedTime = oRs("requestedtime")
	iInspectionStatusId = oRs("inspectionstatusid")
	iInspectorUserId = oRs("inspectoruserid")
	sSchedulingNotes = oRs("schedulingnotes")
	sRequestedAmPm = oRs("requestedampm")
	sScheduledAmPm = oRs("scheduledampm")
	sInspectedAmPm = oRs("inspectedampm")
	If oRs("isfinal") Then
		bIsFinal = True 
	Else
		bIsFinal = False 
	End If 
End If 

oRs.Close
Set oRs = Nothing 

bPermitIsCompleted = GetPermitIsCompleted( iPermitId ) '	in permitcommonfunctions.asp

bIsOnHold = GetPermitIsOnHold( iPermitId ) '	in permitcommonfunctions.asp

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="JavaScript" src="../scripts/layers.js"></script>
	<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script language="Javascript">
	<!--
		
		var bHasInspectedDate = false;

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}
		function setmakeics()
		{
			$("#makeics").val("yes");
		}

		function doValidate()
		{
			var timestring;
			var timearray;
			//var bHasInspectedDate = false;

			// check the requested date
			if ($("#requesteddate").val() != '')
			{
				if (! isValidDate($("#requesteddate").val()))
				{
					alert("The requested date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#requesteddate").focus();
					return;
				}
			}

			// check the scheduled date
			if ($("#scheduleddate").val() != '' || ($("#scheduleddate").val() == '' && $("#makeics").val() != "no"))
			{
				if (! isValidDate($("#scheduleddate").val()))
				{
					alert("The scheduled date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#scheduleddate").focus();
					return;
				}
			}

			// check the inspected date
			if ($("#inspecteddate").val() != '')
			{
				if (! isValidDate($("#inspecteddate").val()))
				{
					alert("The inspected date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#inspecteddate").focus();
					return;
				}
				else
				{
					// the inspected date is valid
					bHasInspectedDate = true;
				}
			}

			// check the requested time
			if ($("#requestedtime").val() != '')
			{
				rege = /^\d{1,2}:\d{1,2}$/;
				Ok = rege.test($("#requestedtime").val());
				if ( ! Ok )
				{
					alert("The requested time must be in the format HH:MM.\nPlease correct this and try saving again.");
					$("#requestedtime").focus();
					return;
				}
				else
				{
					timestring = $("#requestedtime").val();
					timearray = timestring.split(":");
					if (parseInt(timearray[0]) > 12)
					{
						alert("The requested hour must be in the range of 0-12.\nPlease correct this and try saving again.");
						$("#requestedtime").focus();
						return;
					}
					if (parseInt(timearray[1]) > 59)
					{
						alert("The requested minute must be in the range of 0-59.\nPlease correct this and try saving again.");
						$("#requestedtime").focus();
						return;
					}
					if ( timearray[0].length < 2 )
					{
						timearray[0] = '0' + timearray[0];
					}
					if ( timearray[1].length < 2 )
					{
						timearray[1] = '0' + timearray[1];
					}
					$("#requestedtime").val(timearray[0] + ':' + timearray[1]);
				}
			}

			// check the scheduled time
			if ($("#scheduledtime").val() != '' || ($("#scheduledtime").val() == '' && $("#makeics").val() != "no"))
			{
				rege = /^\d{1,2}:\d{1,2}$/;
				Ok = rege.test($("#scheduledtime").val());
				if ( ! Ok )
				{
					alert("The scheduled time must be in the format HH:MM.\nPlease correct this and try saving again.");
					$("#scheduledtime").focus();
					return;
				}
				else
				{
					timestring = $("#scheduledtime").val();
					timearray = timestring.split(":");
					if (parseInt(timearray[0]) > 12)
					{
						alert("The scheduled hour must be in the range of 0-12.\nPlease correct this and try saving again.");
						$("#scheduledtime").focus();
						return;
					}
					if (parseInt(timearray[1]) > 59)
					{
						alert("The scheduled minute must be in the range of 0-59.\nPlease correct this and try saving again.");
						$("#scheduledtime").focus();
						return;
					}
					if ( timearray[0].length < 2 )
					{
						timearray[0] = '0' + timearray[0];
					}
					if ( timearray[1].length < 2 )
					{
						timearray[1] = '0' + timearray[1];
					}
					$("#scheduledtime").val(timearray[0] + ':' + timearray[1]);
				}
			}

			// check the inspected time
			if ($("#inspectedtime").val() != '')
			{
				rege = /^\d{1,2}:\d{1,2}$/;
				Ok = rege.test($("#inspectedtime").val());
				if ( ! Ok )
				{
					alert("The inspected time must be in the format HH:MM.\nPlease correct this and try saving again.");
					$("#inspectedtime").focus();
					return;
				}
				else
				{
					timestring = $("#inspectedtime").val();
					timearray = timestring.split(":");
					if (parseInt(timearray[0]) > 12)
					{
						alert("The inspected hour must be in the range of 0-12.\nPlease correct this and try saving again.");
						$("#inspectedtime").focus();
						return;
					}
					if (parseInt(timearray[1]) > 59)
					{
						alert("The inspected minute must be in the range of 0-59.\nPlease correct this and try saving again.");
						$("#inspectedtime").focus();
						return;
					}
					if ( timearray[0].length < 2 )
					{
						timearray[0] = '0' + timearray[0];
					}
					if ( timearray[1].length < 2 )
					{
						timearray[1] = '0' + timearray[1];
					}
					$("#inspectedtime").val(timearray[0] + ':' + timearray[1]);
				}
			}

			if (bHasInspectedDate == true)
			{
				// The page has the inspected date and it is in a valid format
				// So check that the inspected date is not in the future
				doAjax('checkinpsecteddate.asp', 'inspecteddate=' + $("#inspecteddate").val(), 'SubmitOrAlertOnDate', 'get', '0');
			}
			else
			{
				// Fire off ajax check to see if the inspected date is needed for the status as it is missing
				//alert(document.frmInspection.inspectionstatusid.options[document.frmInspection.inspectionstatusid.selectedIndex].value);
				doAjax('checkinpsecteddateneed.asp', 'inspectionstatusid=' + document.frmInspection.inspectionstatusid.options[document.frmInspection.inspectionstatusid.selectedIndex].value, 'SubmitOrAlertOnNeeded', 'get', '0');
			}
 	 }

		function SubmitOrAlertOnNeeded( sReturn )
		{
			//alert( sReturn );
			if (sReturn == 'NOT NEEDED')
			{
				// The status does not need an inspected date
				if (bHasInspectedDate == true)
				{
					alert('The status you have selected does not allow inspected dates to be entered.\nPlease remove the inspected date, or change the status, and try saving again.');
					document.frmInspection.inspecteddate.focus();
					return;
				}
				else
				{
					document.frmInspection.submit();
				}
			}
			else
			{
				if (bHasInspectedDate == true)
				{
					document.frmInspection.submit();
				}
				else
				{
					alert('An inspected date is required for the status you have selected.\nPlease input an inspected date and try saving again.');
					document.frmInspection.inspecteddate.focus();
					return;
				}
			}
		}

		function SubmitOrAlertOnDate( sReturn )
		{
			if (sReturn == 'DATE OK')
			{
				// Fire off ajax check to see if the inspected date is needed for the status
				doAjax('checkinpsecteddateneed.asp', 'inspectionstatusid=' + document.frmInspection.inspectionstatusid.options[document.frmInspection.inspectionstatusid.selectedIndex].value, 'SubmitOrAlertOnNeeded', 'get', '0');
				//document.frmInspection.submit();
			}
			else
			{
				alert('The inspected date cannot be a date in the future.\nPlease correct this and try saving again.');
				document.frmInspection.inspecteddate.focus();
				return;
			}
		}

		function ViewInspectionDetails( )
		{
			//var w = (screen.width - 680)/2;
			//var h = (screen.height - 480)/2;
			//winHandle = eval('window.open("permitinspectiondetails.asp?permitinspectionid=<%=iPermitInspectionId%>", "_details", "width=900,height=600,location=0,resizable=1,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			parent.showModal('permitinspectiondetails.asp?permitinspectionid=<%=iPermitInspectionId%>', 'Permit Inspection Details', 60, 80);
		}


		function doLoad()
		{
			setMaxLength();
		<%
		'if session("orgid") = "139" and request("success") <> "" and request("makeics") = "" then
		if 1=1 and request("success") <> "" and request("makeics") = "" then
			'Need to evaluate if the final inspection has passed and status isn't completed
			if PermitHasNoPendingInspections( iPermitId, 0) then
				'Change to green check for All Permit Inspections Passed
				%>
				parent.document.getElementById("apipimg").src = '../images/check.png';
				<%
				'Enable Button
				%>
				parent.document.getElementsByName("completepermitbtn")[0].disabled = false;
				<%
				'Remove Tooltip Class
				%>
				parent.document.getElementsByName("completepermitbtn")[0].classList.remove("tooltip");
				parent.document.getElementById("completett").style.display = "none";
				<%
			else
				'Change to red X for All Permit Inspections Passed
				%>
				parent.document.getElementById("apipimg").src = '../images/x.png';
				<%
				'Disable Button
				%>
				parent.document.getElementsByName("completepermitbtn")[0].disabled = true;
				<%
				'Add Tooltip Class
				%>
				parent.document.getElementsByName("completepermitbtn")[0].classList.add("tooltip");
				parent.document.getElementById("completett").style.display = "block";
				<%
			end if

			'Need to update this inpsection in the table %>
			parent.document.getElementById("InStatus<%=iPermitInspectionId%>").innerHTML = $("select[name=inspectionstatusid]  option:selected").text();
			parent.document.getElementById("InSchedDate<%=iPermitInspectionId%>").innerHTML = $("#scheduleddate").val();
			parent.document.getElementById("InInspectDate<%=iPermitInspectionId%>").innerHTML = $("#inspecteddate").val();
			parent.document.getElementById("InInspector<%=iPermitInspectionId%>").innerHTML = $("select[name=inspectoruserid]  option:selected").text();

			
		<%
		'elseif session("orgid") = "139" and request("success") <> "" and request("makeics") <> "" then
		elseif 1=1 and request("success") <> "" and request("makeics") <> "" then
			%>parent.RefreshPageAfterVoid( "InspectionChange|<%=iPermitInspectionId%>|<%=request("makeics")%>" );<%
		else
			If request("success") <> "" Then %>
				parent.RefreshPageAfterVoid( "InspectionChange|<%=iPermitInspectionId%>|<%=request("makeics")%>" );
<%			End If	%>
		<% 
		end if
		%>
		}

<%		If request("success") <> "" Then 
			DisplayMessagePopUp 
		End If 
%>

  $( function() {
    $( ".datepicker" ).datepicker({
      changeMonth: true,
      showOn: "both",
      buttonText: "<i class=\"fa fa-calendar\"></i>",
      changeYear: true
    });
  } );
	//-->
	</script>


</head>

<body onload="doLoad();">

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	

	<!--BEGIN: EDIT FORM-->
	<form name="frmInspection" action="permitinspectionupdate.asp" method="post">
	<input type="hidden" name="permitid" value="<%=iPermitId%>" />
	<input type="hidden" name="permitinspectionid" value="<%=iPermitInspectionId%>" />
	<input type="hidden" name="inspectionpage" value="permitinspectionedit" />
	<input type="hidden" name="makeics" id="makeics" value="no" />

	<p>
		<table cellpadding="0" border="0" cellspacing="0" id="inspectiondetails">
			<tr>
				<td class="firstcol">Inspection:</td>
				<td nowrap="nowrap" colspan="3"><span class="keyinfo"><%=sPermitInspectionType%>
<%					If bIsReinspection Then 
						response.write " &mdash; This is a reinspection"
					End If		
					If bIsFinal Then
						response.write "<br />This is the final inspection for this permit"
					End If 
%>
					</span>
				</td>
			</tr>
			<tr><td class="firstcol">Description:</td><td nowrap="nowrap" colspan="3"><%=sInspectionDescription%></td></tr>
			<tr><td class="firstcol">Inspector:</td><td class="datecol" nowrap="nowrap" colspan="3"><select name="inspectoruserid"><% ShowPermitInspectors iInspectorUserId %></td></tr>
			<tr><td class="firstcol">Status:</td><td class="datecol" nowrap="nowrap" colspan="3"><% ShowInspectionStatuses iInspectionStatusId  %></td></tr>
			<tr>
				<td class="firstcol" nowrap="nowrap">Request&nbsp;Received:</td>
				<td colspan="3" nowrap="nowrap">&nbsp;<%= sRequestReceived %></td>
			</tr>
			<tr>
				<td class="firstcol" nowrap="nowrap">Requested&nbsp;Date:</td>
				<td class="datecol" nowrap="nowrap"><input type="text" class="datepicker" name="requesteddate" id="requesteddate" value="<%=sRequestedDate%>" size="10" maxlength="10" />
				</td>
				<td class="timecol" align="right">Requested&nbsp;Time (HH:MM):</td>
				<td class="timefields"><input type="text" name="requestedtime" id="requestedtime" value="<%=sRequestedTime%>" size="5" maxlength="5" />
					<% ShowAmPmPicks "requestedampm", sRequestedAmPm	' In permitcommonfunctions.asp	 %>
				</td>
			</tr>
			<tr>
				<td class="firstcol" nowrap="nowrap">Scheduled&nbsp;Date:</td>
				<td class="datecol" nowrap="nowrap"><input type="text" class="datepicker" name="scheduleddate" id="scheduleddate" value="<%=sScheduledDate%>" size="10" maxlength="10" />
				</td>
				<td class="timecol" align="right">Scheduled&nbsp;Time (HH:MM):</td>
				<td class="timefields"><input type="text" name="scheduledtime" id="scheduledtime" value="<%=sScheduledtime%>" size="5" maxlength="5" />
					<% ShowAmPmPicks "scheduledampm", sScheduledAmPm	' In permitcommonfunctions.asp	 %>
				</td>
			</tr>
			<tr>
				<td class="firstcol" nowrap="nowrap">Inspected&nbsp;Date:</td>
				<td class="datecol" nowrap="nowrap"><input type="text" class="datepicker" name="inspecteddate" id="inspecteddate" value="<%=sInspectedDate%>" size="10" maxlength="10" />
				</td>
				<td class="timecol" align="right">Inspected&nbsp;Time (HH:MM):</td>
				<td class="timefields"><input type="text" name="inspectedtime" id="inspectedtime" value="<%=sInspectedtime%>" size="5" maxlength="5" />
					<% ShowAmPmPicks "inspectedampm", sInspectedAmPm	' In permitcommonfunctions.asp	%>
				</td>
			</tr>
			<tr>
				<td class="firstcol">Contact:</td> 
				<td class="datecol" colspan="3"><input type="text" name="contact" id="contact" value="<%=sContact%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td class="firstcol">Contact Phone:</td>
				<td class="datecol" colspan="3"><input type="text" name="contactphone" id="contactphone" value="<%=sContactPhone%>" size="25" maxlength="25" /></td>
			</tr>
			<tr>
				<td class="firstcol">Scheduling Notes:</td>
				<td class="datecol" colspan="3">&nbsp;</td>
			</tr>
			<tr>
				<td class="datecol" colspan="4"><textarea name="schedulingnotes" rows="5" cols="80" maxlength="1000"><%=sSchedulingNotes%></textarea></td>
			</tr>

		</table>
	</p>
	<p>

<%					
	tooltipclass=""
	tooltip = ""
	disabled = ""
	If bPermitIsCompleted or bIsOnHold or (bisfinal and not AllOtherInspectionsAreDone( iPermitId, iPermitinspectionid )) or not GetPermitIsIssued( iPermitId ) _
		or (bisfinal and not PermitFeesArePaid( iPermitId )) Then		' in permitcommonfunctions.asp
		tooltipclass="tooltip"
		disabled = " disabled "
		tooltip = "<span class=""tooltiptext"">You cannot save because:<br />"
		if bPermitIsCompleted then tooltip = tooltip & "The permit is complete.<br />"
		if bIsOnHoldthen then tooltip = tooltip & "The permit is on hold.<br />"
		if not GetPermitIsIssued( iPermitId ) then tooltip = tooltip & "The permit isn't issued.<br />"
		if (bisfinal and not AllOtherInspectionsAreDone( iPermitId, iPermitinspectionid )) then tooltip = tooltip & "Other inspections are incomplete.<br />"
		if (bisfinal and not PermitFeesArePaid( iPermitId )) then tooltip = tooltip & "The fees aren't paid or paid too much.<br />"
		tooltip = tooltip & "</span>"
	end if
%>
	<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" id="savebutton" onclick="doValidate();">Save Changes<%=tooltip%></button> &nbsp; &nbsp;
	<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" id="savebutton" onclick="setmakeics();doValidate();">Save Changes and Create Outlook Invitation<%=tooltip%></button> &nbsp; &nbsp;
	<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" /> &nbsp; &nbsp;
		<input type="button" class="button ui-button ui-widget ui-corner-all" value="View Inspection Details" onclick="ViewInspectionDetails();" />
	</p>
	<% if session("orgid") = "139" or session("orgid") = "181" then 
		sSQL = "SELECT * FROM egov_permitinspectionreports WHERE permitinspectionid = '" & iPermitInspectionId & "'"
		set oIR = Server.CreateObject("ADODB.RecordSet")
		oIR.Open sSQL, Application("DSN"), 3, 1
		if oIR.EOF then 
			intPermitInspectionReportID = 0
		else
			intPermitInspectionReportID = oIR("permitinspectionreportid")
			strInspectionType = oIR("inspectiontype")
			if oIR("approved") then strApproved = "checked"
			if oIR("disapproved") then strDisapproved = "checked"
			if oIR("approvedwcorr") then strApprovedWCorr = "checked"
			if oIR("coc") then strCOC = "checked"
			strRemarks = oIR("remarks")
			intInspector = oIR("permitinspectorid")
			if isnull(intInspector) then intInspector = 0
		end if

		oIR.Close
		Set oIR = Nothing
	%>
	<p>
		<div id="inspectionreport_expand" onClick="toggleDisplayShow( 'inspectionreport' );">
			<strong><span id="inspectionreportimg">&ndash;</span> <u>Inspection Report:</u></strong>
		</div>
		<div id="inspectionreport" style="border: 1px solid black;width:584px; padding:3px;">
			<input type="hidden" name="permitinspectionreportid" value="<%=intPermitInspectionReportID%>" />
			<%
				tooltipclass=""
				tooltip = ""
				disabled = ""
				if intPermitInspectionReportID = 0 then
					tooltipclass="tooltip"
					disabled = " disabled "
					tooltip = "<span class=""tooltiptext"">You must first save the Inspection Report data.</span>"
				end if
			%>
			<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="parent.showModal('inspectionreport.asp?permitinspectionreportid=<%=intPermitInspectionReportID%>','Print Inspection Report', 55, 90)" >Print<%=tooltip%></button>
			&nbsp; &nbsp; &nbsp; &nbsp;
			<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onClick="parent.showModal('emailinspectionreport.asp?permitinspectionreportid=<%=intPermitInspectionReportID%>', 'Email Permit Inspection Report', 30, 50);">Email<%=tooltip%></button>
			<br />
			Inspection Type: <input type="text" size=40 name="inspectiontype" maxlength="500" value="<%=strInspectionType%>" />
			<br />
			Inspections:<br />
			<style>
				#itlist select
				{
					margin:3px;
				}
			</style>
			<div id="itlist">
			<%
			sSQL = "SELECT inspectiontype FROM egov_permitinspectionreporttypes WHERE permitinspectionreportid = '" & intPermitInspectionReportID & "'"
			set oRs = Server.CreateObject("ADODB.RecordSet")
			oRs.Open sSQL, Application("DSN"), 3, 1
			intTypeCount = 1

			sSQL = "SELECT * FROM egov_permitinspectionreporttypeoptions"
			Set oIT = Server.CreateObject("ADODB.Recordset")
			oIT.Open sSql, Application("DSN"), 3, 1

			Do While not oRs.EOF
				ListInspectionTypes oRs("inspectiontype"), intTypeCount, oIT
				intTypeCount = intTypeCount + 1
				oRs.MoveNext
			loop
			if intTypeCount = 1 then ListInspectionTypes "", intTypeCount, oIT
			oIT.Close
			Set oIT = Nothing
			%>
			</div>
			<a href="javascript:addIT()">Add Inspection</a>
			<input type="hidden" id="itCount" value="<%=intTypeCount+1%>" />
			<script>
				function addIT() {
					var count = document.getElementById("itCount").value;
					var itm = document.getElementById("inspectionreporttype1");
					var cln = itm.cloneNode(true);
					cln.id = "inspectionreporttype" + count;
					cln.setAttribute("name","inspectionreporttype" + count);
					cln.selectedIndex = 0;
					document.getElementById("itlist").appendChild(cln);
					document.getElementById("itCount").value = parseInt(count) + 1;
				}
			</script>
			<br />
			<input type="checkbox" value="on" name="approved" <%=strApproved%> />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;APPROVED
			<br />
			<input type="checkbox" value="on" name="disapproved" <%=strDisapproved%>/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;DISAPPROVED
			<br />
			<input type="checkbox" value="on" name="approvedwcorr" <%=strApprovedWCorr%> />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;APPROVED WITH NOTED CORRECTIONS
			<br />
			<input type="checkbox" value="on" name="coc" <%=strCOC%> />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CERTIFICATE OF COMPLIANCE
			<br />
			Remarks:<br />
			<textarea style="width:95%" rows="10" name="remarks"><%=strRemarks%></textarea>
			<br />
			Inspector:<select name="permitinspectorid"><% ShowPermitInspectors intInspector %>

		</div>
	</p>
	<% end if %>
	<p>
		<div id="newinspectionnotes_expand" onClick="toggleDisplayShow( 'newinspectionnotes' );">
			<strong><span id="newinspectionnotesimg">&ndash;</span> <u>New Permit Inspection Notes:</u></strong>
		</div>
		<div id="newinspectionnotes">
		<table>
			<tr><td><strong>Internal Notes:</strong><br />
					<textarea name="internalcomment" rows="5" cols="80" maxlength="1000"></textarea>
				</td>
			</tr>
			<tr><td><strong>Public Notes:</strong><br />
					<textarea name="externalcomment" rows="5" cols="80" maxlength="1000"></textarea>
				</td>
			</tr>
		</table>
		</div> 
		<div id="inspectionnotes_expand" onClick="toggleDisplayShow( 'inspectionnotes' );">
			<strong><span id="inspectionnotesimg">&ndash;</span> <u>Prior Permit Inspection Notes:</u></strong>
		</div>
		<div id="inspectionnotes">
<%			ShowInspectionNotes iPermitInspectionId		%>
		</div>
	</p>

	</form>
	<!--END: EDIT FORM-->

	</div>
</div>
<% 
'if request.querystring("makeics") = "yes" then response.redirect "makeics.asp?permitinspectionid=" & iPermitInspectionId 
%>

<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

<%	If request("success") <> "" Then 
		SetupMessagePopUp request("success")
	End If	
%>

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub ShowInspectionNotes( iPermitInspectionId )
'--------------------------------------------------------------------------------------------------
Sub ShowInspectionNotes( iPermitInspectionId )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT entrydate, ISNULL(internalcomment,'') AS internalcomment, ISNULL(externalcomment,'') AS externalcomment, "
	sSql = sSQl & " S.inspectionstatus, U.firstname, U.lastname, ISNULL(activitycomment,'') AS activitycomment "
	sSql = sSQl & " FROM egov_permitlog L, egov_inspectionstatuses S, users U "
	sSql = sSQl & " WHERE S.inspectionstatusid = L.inspectionstatusid AND U.userid = L.adminuserid AND permitinspectionid = " & iPermitInspectionId
	sSql = sSQl & " AND L.isinspectionentry = 1 ORDER BY permitlogid DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<table id=""priorinspectionnotes"" cellpadding=""3"" cellspacing=""0"" border=""0"">"
		Do While Not oRs.EOF 
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 1 Then
				response.write " class=""altrow"" "
			End If 
			response.write "><td><strong>"
			response.write oRs("firstname") & " " & oRs("lastname") & " &ndash; " & oRs("inspectionstatus") & " &ndash; " & oRs("entrydate") & "</strong><br />"
			If oRs("activitycomment") <> "" Then 
				response.write " &nbsp; " & oRs("activitycomment") & "<br />"
			End If 
			If oRs("internalcomment") <> "" Then 
				response.write " &nbsp; <strong>Internal Note:</strong> " & oRs("internalcomment") & "<br />"
			End If 
			If oRs("externalcomment") <> "" Then 
				response.write " &nbsp; <strong>Public Note:</strong> " & oRs("externalcomment")
			End If 
			response.write "</td></tr>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</table>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowInspectionStatuses( iInspectionStatusId )
'--------------------------------------------------------------------------------------------------
Sub ShowInspectionStatuses( iInspectionStatusId )
	Dim sSql, oRs

	sSql = "SELECT inspectionstatusid, inspectionstatus FROM egov_inspectionstatuses WHERE orgid = " & session("orgid")
	sSql = sSQl & " ORDER BY inspectionstatusorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""inspectionstatusid"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option "
			If CLng(iInspectionStatusId) = CLng(oRs("inspectionstatusid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " value=""" & oRs("inspectionstatusid") & """>" & oRs("inspectionstatus") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowPermitInspectors( iInspectorUserId )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitInspectors( iInspectorUserId )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname FROM users WHERE orgid = " & session("orgid") & " AND ispermitinspector = 1 "
	sSql = sSQl & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<option value=""0"">Unassigned</option>"
		
	Do While Not oRs.EOF
		response.write vbcrlf & "<option "
		If CLng(iInspectorUserId) = CLng(oRs("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write " value=""" & oRs("userid") & """>" & oRs("firstname") & " " & oRs("lastname") & "</option>"
		oRs.MoveNext
	Loop 

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 
End Sub 

Sub ListInspectionTypes(strIT, intCount, oIT)
oIT.MoveFirst
%>
<select id="inspectionreporttype<%=intCount%>" name="inspectionreporttype<%=intCount%>">
<option value="0">Choose...</option>
<%
strGroup = ""
Do While Not oIT.EOF
	if strGroup <> oIT("optgroup") then
		if strGroup <> ""then response.write "</optgroup>"
		strGroup = oIT("optgroup")
		response.write "<optgroup label =""" & strGroup & """>"


	end if

	selected = ""
	if strIT = oIT("code") then selected = " selected"

	response.write "<option value=""" & oIT("code") & """" & selected & ">" & oIT("description") & "</option>" & vbcrlf
	oIT.MoveNext
loop
response.write "</optgroup></select>"

End Sub



%>
