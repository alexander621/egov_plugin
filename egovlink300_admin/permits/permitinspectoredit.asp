<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitinspectoredit.asp
' AUTHOR: Steve Loar
' CREATED: 08/13/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Edits permit inspections by permit inspectors
'
' MODIFICATION HISTORY
' 1.0   08/13/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitInspectionId, iPermitId, sSql, oRs, sPermitInspectionType, sInspectionDescription, bIsReinspection
Dim sRequestReceived, sRequestedDate, sScheduledDate, sInspectedDate, sContactPhone, sContact, sInspectedtime
Dim sScheduledtime, sRequestedTime, sRequestedAmPm, iInspectionStatusId, iInspectorUserId, sSchedulingNotes
Dim sScheduledAmPm, sInspectedAmPm, bIsFinal, bPermitIsCompleted, bIsOnHold, bCanSaveChanges, iPermitStatusId
Dim sAlertMsg, sAlertSetByUser, dAlertDate, sPermitLocation, sLocationType

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "edit permit inspection", sLevel	' In common.asp

iPermitInspectionId = CLng(request("permitinspectionid"))

If request("activetab") <> "" Then 
	If IsNumeric(request("activetab")) Then 
		iActiveTabId = clng(request("activetab"))
	Else
		iActiveTabId = clng(0)
	End If 
Else
	iActiveTabId = clng(0)
End If 

sPermitLocation = ""
sLocationType = ""

sSql = "SELECT permitid, permitinspectiontype, inspectiondescription, ISNULL(inspectoruserid,0) AS inspectoruserid, isfinal, "
sSql = sSql & " inspectionstatusid, requestreceiveddate, requesteddate, requestedtime, ISNULL(requestedampm,'') AS requestedampm, "
sSql = sSql & " scheduleddate, scheduledtime, ISNULL(scheduledampm, '') AS scheduledampm, inspecteddate, inspectedtime, ISNULL(inspectedampm,'') AS inspectedampm, "
sSql = sSql & " contact, contactphone, isreinspection, schedulingnotes "
sSql = sSql & " FROM egov_permitinspections WHERE orgid = " & session("orgid") & " AND permitinspectionid = " & iPermitInspectionId

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
	If IsNull(oRs("inspecteddate")) Then 
		sInspectedDate = FormatDateTime(Date(),2)
	Else 
		sInspectedDate = oRs("inspecteddate")
	End If 
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

iPermitStatusId = GetPermitStatusId( iPermitId )	' in permitcommonfunctions.asp

bCanSaveChanges = StatusAllowsSaveChanges( iPermitStatusId ) 	' in permitcommonfunctions.asp

GetPermitAlertDetails iPermitId, sAlertMsg, sAlertSetByUser, dAlertDate ' in permitcommonfunctions.asp

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../yui/build/tabview/assets/skins/sam/tabview.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<!--
	<script type="text/javascript" src="../yui/build/yahoo-dom-event/yahoo-dom-event.js"></script>
	<script type="text/javascript" src="../yui/build/element/element-beta.js"></script>
	<script type="text/javascript" src="../yui/build/tabview/tabview.js"></script>
	-->
	<script type="text/javascript" src="../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../yui/element-min.js"></script>  
	<script type="text/javascript" src="../yui/tabview-min.js"></script>

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="JavaScript" src="../scripts/layers.js"></script>
	<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script language="Javascript">
	<!--
  		$( function() {

    			$( ".datepicker" ).datepicker({
      				changeMonth: true,
      				showOn: "both",
      				buttonText: "<i class=\"fa fa-calendar\"></i>",
      				changeYear: true
    			}); 
		});

		var tabView;
		var winHandle;
		var bHasInspectedDate = false;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
			//tabView.set('activeIndex', 0); 
			tabView.set('activeIndex', <%=iActiveTabId%>);
		})();

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function doValidate()
		{
			var timestring;
			var timearray;
			//var bHasInspectedDate = false;

			// Set the active tab
			$("#activetab").val(tabView.get("activeIndex"));

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
			if ($("#scheduleddate").val() != '')
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
			if (document.getElementById("scheduledtime").value != '')
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
				alert('The inspected date is either invalid or has been set to a date in the future.\nPlease correct this and try saving again.');
				document.frmInspection.inspecteddate.focus();
				return;
			}
		}

		function ViewInspectionDetails( )
		{
			var w = (screen.width - 680)/2;
			var h = (screen.height - 480)/2;
			//winHandle = eval('window.open("permitinspectiondetails.asp?permitinspectionid=<%=iPermitInspectionId%>", "_details", "width=900,height=600,location=0,resizable=1,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			parent.showModal('permitinspectiondetails.asp?permitinspectionid=<%=iPermitInspectionId%>', 'Permit Inspection Details', 60, 80);
		}


		function doLoad()
		{
			setMaxLength();
		}

		function ViewAttachment( iAttachmentId )
		{
			location.href = "permitattachmentview.asp?permitattachmentid=" + iAttachmentId;
		}

		function AddAttachments( )
		{
			var w = (screen.width - 640)/2;
			var h = (screen.height - 480)/2;
			//winHandle = eval('window.open("permitattachment.asp?permitid=<%=iPermitId%>", "_contact", "width=800,height=350,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('permitattachment.asp?permitid=<%=iPermitId%>', 'Add An Attachment', 40, 40);
		}

		function ViewDetails()
		{
			var w = (screen.width - 680)/2;
			var h = (screen.height - 480)/2;
			//winHandle = eval('window.open("viewpermitdetails.asp?permitid=<%=iPermitId%>", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
			showModal('viewpermitdetails.asp?permitid=<%=iPermitId%>', 'Permit Details', 50, 80);
		}

		function GoToList()
		{
			location.href = 'permitinspectorlist.asp';
		}

		function RefreshPageAfterVoid( sResults )
		{
			//alert(sResults);
			setTimeout(function() {location.href = "permitinspectoredit.asp?permitinspectionid=<%=iPermitInspectionId%>&activetab=" + tabView.get("activeIndex");}, 200);
		}

<%		If request("success") <> "" Then 
			'DisplayMessagePopUp %>
  		$( function() {
			$("#successmessage").show();
			$("#successmessage").fadeOut(2000);
		});
			
		<%End If %>

	//-->
	</script>

</head>
<body class="yui-skin-sam" onload="doLoad();">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	<div class="gutters">
	
	<!--BEGIN: PAGE TITLE-->
	<p>
		<font size="+1"><strong>Permit Inspection</strong></font><br /><br />
	</p>
	<!--END: PAGE TITLE-->

	<!--BEGIN: EDIT FORM-->
	<form name="frmInspection" action="permitinspectionupdate.asp" method="post">
	<input type="hidden" name="permitid" value="<%=iPermitId%>" />
	<input type="hidden" name="permitinspectionid" value="<%=iPermitInspectionId%>" />
	<input type="hidden" name="activetab" id="activetab" value="<%=iActiveTabId%>" />
	<input type="hidden" name="inspectionpage" value="permitinspectoredit" />

	<p>
		<table cellpadding="0" border="0" cellspacing="0" style="width:50%" id="inspectiondetails">
			<tr><td class="firstcol">Permit #:</td><td nowrap="nowrap" colspan="3"><span class="keyinfo"><%=GetPermitNumber( iPermitId )%></span></td></tr>
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
<%				sLocationType = GetPermitLocationType( iPermitId )
				If sLocationType = "address" Then	%>
					<tr><td class="firstcol">Address:</td><td nowrap="nowrap" colspan="3"><%=GetPermitJobSite( iPermitId )%></td></tr>
<%				End If		

				If sLocationType = "location" Then	%>
					<tr><td valign="top" class="firstcol">Location:</td><td nowrap="nowrap" colspan="3"><%=Replace(GetPermitPermitLocation( iPermitId ),Chr(10),"<br />")%></td></tr>
<%				End If		%>

			<tr><td class="firstcol">Permit Type:</td><td nowrap="nowrap" colspan="3"><%=GetPermitTypeDesc( iPermitId, True ) %></td></tr>
			<tr><td class="firstcol">Description of Work:</td><td nowrap="nowrap" colspan="3"><%=GetDescriptionOfWork( iPermitId )%></td></tr>
<%				If sAlertMsg <> "" Then %>
					<tr>
						<td class="firstcol" valign="top">Alert:</td>
						<td colspan="3"><% response.write "<span id=""permitalertmsg"">" & sAlertMsg & "</span><br />Set by " & sAlertSetByUser & " on " & FormatDateTime(dAlertDate,2)  %>
						</td>
					</tr>
<%				End If		%>
			<tr><td class="firstcol">Inspector:</td><td class="datecol" nowrap="nowrap" colspan="3"><select name="inspectoruserid"><% ShowPermitInspectors iInspectorUserId %></td></tr>
			<tr><td class="firstcol">Inspection Status:</td><td class="datecol" nowrap="nowrap" colspan="3"><% ShowInspectionStatuses iInspectionStatusId  %></td></tr>
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
				<td colspan="4"><textarea name="schedulingnotes" rows="5" cols="80" maxlength="1000"><%=sSchedulingNotes%></textarea></td>
			</tr>
		</table>
	</p>
	<p>
		<input type="button" class="button ui-button ui-widget ui-corner-all" value="<< Back to Inspection List" onclick="GoToList();" /> &nbsp; &nbsp;
<%		If InspectionCanSaveChanges( iPermitId, iPermitInspectionId ) And Not bPermitIsCompleted And Not bIsOnHold Then		' in permitcommonfunctions.asp		%>
			<input type="button" class="button ui-button ui-widget ui-corner-all" value="Save Changes" id="savebutton" onclick="doValidate();" /> &nbsp; &nbsp;
<%		End If		%>
		<input type="button" class="button ui-button ui-widget ui-corner-all" value="View Inspection Details" onclick="ViewInspectionDetails();" /> &nbsp; &nbsp;
		<input type="button" class="button ui-button ui-widget ui-corner-all" value="View Permit Details" onclick="ViewDetails();" />
	</p>

	<div id="demo" class="yui-navset">
			<ul class="yui-nav">
				<li><a href="#tab1"><em>Notes</em></a></li>
				<li><a href="#tab2"><em>Attachments</em></a></li>
				<% if session("orgid") = "139" or session("orgid") = "181" then %>
				<li><a href="#tab3"><em>Inspection Report</em></a></li>
				<% end if %>
			</ul>            
			<div class="yui-content">
				<div id="tab1"> <!-- Notes -->
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
				</div>
				<div id="tab2"> <!-- Attachments -->
					<p class="tabpage">
<%					If bCanSaveChanges Then		%>
						&nbsp; <input type="button" class="button ui-button ui-widget ui-corner-all" value="Add An Attachment" onclick="AddAttachments( );" /> 
<%					End If %>
					</p>
					<p>
						<table cellpadding="2" cellspacing="0" border="0" class="feetable" id="attachmentlist">
							<tr><th>File Name</th><th>Description</th><th>Date Added</th><th>Added By</th></tr>
<%							iMaxAttachments = ShowAttachmentList( iPermitId )		%>		
						</table>
						<input type="hidden" id="maxattachments" name="maxattachments" value="<%=iMaxAttachments%>" />
					</p>
				</div>
				<% if session("orgid") = "139" or session("orgid") = "181" then %>
				<div id="tab3"> <!-- Inspection Report -->
					<p class="tabpage">
						<%
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
								display:block;
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
					</p>
				</div>
				<% end if %>
			</div>
		</div>

	</form>
	<!--END: EDIT FORM-->

	</div>
	</div>
</div>

<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  
<!--#Include file="modal.asp"-->  

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
' void ShowInspectionNotes( iPermitInspectionId )
'--------------------------------------------------------------------------------------------------
Sub ShowInspectionNotes( ByVal iPermitInspectionId )
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
' void ShowInspectionStatuses( iInspectionStatusId )
'--------------------------------------------------------------------------------------------------
Sub ShowInspectionStatuses( ByVal iInspectionStatusId )
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
' void ShowPermitInspectors( iInspectorUserId )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitInspectors( ByVal iInspectorUserId )
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


'--------------------------------------------------------------------------------------------------
' integer ShowAttachmentList( iPermitId )
'--------------------------------------------------------------------------------------------------
Function ShowAttachmentList( ByVal iPermitId )
	Dim sSql, oRs, iRecCount

	iRecCount = 0

	sSql = "SELECT permitattachmentid, attachmentname, ISNULL(description,'') AS description, attachmentpath, "
	sSql = sSql & " ISNULL(adminuserid,0) AS adminuserid, dateadded, fileextension "
	sSql = sSql & " FROM egov_permitattachments WHERE permitid = " & iPermitId
	sSql = sSql & " ORDER BY 1 DESC"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			iRecCount = iRecCount + 1
			response.write vbcrlf & "<tr"
			If iRecCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write ">"

'			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
'			response.write "<td align=""center"" title=""Click to View"" onclick=""ViewAttachment(" & oRs("permitattachmentid") & ");"">" & oRs("dateadded") & "</td>"
'			response.write "<td align=""center"" title=""Click to View"" onclick=""ViewAttachment(" & oRs("permitattachmentid") & ");"">" & GetAdminName( oRs("adminuserid") ) & "</td>"
'			response.write "<td align=""center"" title=""Click to View"" onclick=""ViewAttachment(" & oRs("permitattachmentid") & ");"">" & oRs("attachmentname") & "</td>"
'			response.write "<td align=""center"" title=""Click to View"" onclick=""ViewAttachment(" & oRs("permitattachmentid") & ");"">" & oRs("description") & "</td>"

			If oRs("attachmentpath") = "..\permitattachments" Then 
				sLink = "<a class=""permitattachments"" href='" & oRs("attachmentpath") & "/" & oRs("permitattachmentid") & "." & oRs("fileextension") & "' target=""_blank"">"
			Else
				sLink = "<a class=""permitattachments"" href='" & oRs("attachmentpath") & "/" & oRs("permitattachmentid") & "_" & oRs("attachmentname") & "' target=""_blank"">"
			End If
			response.write "<td align=""center"" title=""Click to View"">" & sLink & oRs("attachmentname") & "</a></td>"
			response.write "<td align=""center"">" & oRs("description") & "</td>"
			response.write "<td align=""center"">" & DateValue(oRs("dateadded")) & "</td>"
			response.write "<td align=""center"">" & GetAdminName( oRs("adminuserid") ) & "</td>"
			
			response.write "</tr>"
			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

	ShowAttachmentList = iRecCount

End Function 


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
