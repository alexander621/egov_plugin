<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitrevieweredit.asp
' AUTHOR: Steve Loar
' CREATED: 08/05/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Edits permit reviews from the reviewer list
'
' MODIFICATION HISTORY
' 1.0   08/05/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitReviewId, iPermitId, iPermitStatusId, sLegalDescription, sListedOwner, iPermitAddressId
Dim sPermitStatus, sPermitNo, sPermitReviewType, sReviewDescription, iReviewerId, iReviewStatusId
Dim sRequired, bPermitIsCompleted, bIsOnHold, bBlockReviewStatusChange, iActiveTabId, bCanSaveChanges
Dim sAlertMsg, sAlertSetByUser, dAlertDate, sPermitLocation, sLocationType

sLevel = "../" ' Override of value from common.asp

iPermitReviewId = CLng(request("permitreviewid"))

If request("activetab") <> "" Then 
	If IsNumeric(request("activetab")) Then 
		iActiveTabId = clng(request("activetab"))
	Else
		iActiveTabId = clng(0)
	End If 
Else
	iActiveTabId = clng(0)
End If 

iPermitId = GetPermitIdByPermitReviewId( iPermitReviewId ) '	in permitcommonfunctions.asp

bPermitIsCompleted = GetPermitIsCompleted( iPermitId ) '	in permitcommonfunctions.asp

bIsOnHold = GetPermitIsOnHold( iPermitId ) '	in permitcommonfunctions.asp

bBlockReviewStatusChange = GetPermitStatusBlockReview( iPermitId )	' in permitcommonfunctions.asp

iPermitStatusId = GetPermitStatusId( iPermitId )	' in permitcommonfunctions.asp

bCanSaveChanges = StatusAllowsSaveChanges( iPermitStatusId ) 	' in permitcommonfunctions.asp

sPermitNo = ""
sProposedUse = ""
sExistingUse = ""
sDescriptionOfWork = ""
sPermitStatus = ""
sRequired = ""
sPermitReviewType = ""
sReviewDescription = ""
iReviewerId = 0
iReviewStatusId = 0
sPermitLocation = ""
sLocationType = ""

GetReviewDetails iPermitReviewId, sPermitReviewType, sReviewDescription, iReviewerId, iReviewStatusId, sRequired 

GetPermitAlertDetails iPermitId, sAlertMsg, sAlertSetByUser, dAlertDate ' in permitcommonfunctions.asp

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../yui/build/tabview/assets/skins/sam/tabview.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script type="text/javascript" src="../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../yui/element-min.js"></script>  
	<script type="text/javascript" src="../yui/tabview-min.js"></script>

	<script type="text/javascript" src="../scripts/layers.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script language="Javascript">
	<!--
		var tabView;
		var winHandle;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
			//tabView.set('activeIndex', 0); 
			tabView.set('activeIndex', <%=iActiveTabId%>);
		})();

		function doValidate()
		{
			$("#activetab").val(tabView.get("activeIndex"));
			// Post the page
			document.frmReview.submit();
		}

		function doLoad()
		{
			setMaxLength();
		}

		function ViewDetails()
		{
			var w = (screen.width - 680)/2;
			var h = (screen.height - 480)/2;
			//winHandle = eval('window.open("viewpermitdetails.asp?permitid=<%=iPermitId%>", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
			showModal('viewpermitdetails.asp?permitid=<%=iPermitId%>', 'Permit Details', 50, 80);

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

		function GoToList()
		{
			location.href = 'permitreviewerlist.asp';
		}

		function RefreshPageAfterVoid( sResults )
		{
			//alert(sResults);
			setTimeout(function() {location.href = "permitrevieweredit.asp?permitreviewid=<%=iPermitReviewId%>&activetab=" + tabView.get("activeIndex");}, 200);
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
			<font size="+1"><strong>Permit Review</strong></font><br /><br />
		</p>
		<!--END: PAGE TITLE-->

		<!--BEGIN: EDIT FORM-->
		<form name="frmReview" action="permitreviewupdate.asp" method="post">
		<input type="hidden" name="permitid" value="<%=iPermitId%>" />
		<input type="hidden" name="permitreviewid" value="<%=iPermitReviewId%>" />
		<input type="hidden" name="reviewpage" value="permitrevieweredit" />
		<input type="hidden" name="activetab" id="activetab" value="<%=iActiveTabId%>" />
		<p>
			<table cellpadding="2" border="0" cellspacing="0" id="reviewdetails">
				<tr><td>Permit #:</td><td><span class="keyinfo"><%=GetPermitNumber( iPermitId )%></span></td></tr>
				<tr><td>Review:</td><td nowrap="nowrap"><span class="keyinfo"><%=sPermitReviewType%></span> <%=sRequired%></td></tr>
				<tr><td>Description:</td><td nowrap="nowrap"><%=sReviewDescription%></td></tr>
				<tr><td>Reviewer:</td><td nowrap="nowrap"><% ShowPermitReviewers iReviewerId %></td></tr>
				<tr><td>Review Status:</td><td nowrap="nowrap"><% ShowReviewStatuses iReviewStatusId, bBlockReviewStatusChange %></td></tr>
				
<%				sLocationType = GetPermitLocationType( iPermitId )
				If sLocationType = "address" Then	%>
					<tr><td>Address:</td><td nowrap="nowrap"><%=GetPermitJobSite( iPermitId )%></td></tr>
<%				End If		

				If sLocationType = "location" Then	%>
					<tr><td>Location:</td><td nowrap="nowrap"><%=Replace(GetPermitPermitLocation( iPermitId ),Chr(10),"<br />")%></td></tr>
<%				End If		%>

				<tr><td>Permit Type:</td><td nowrap="nowrap"><%=GetPermitTypeDesc( iPermitId, True ) %></td></tr>
				<tr><td>Description of Work:</td><td nowrap="nowrap"><%=GetDescriptionOfWork( iPermitId )%></td></tr>
<%				If sAlertMsg <> "" Then %>
					<tr>
						<td valign="top">Alert:</td>
						<td><% response.write "<span id=""permitalertmsg"">" & sAlertMsg & "</span><br />Set by " & sAlertSetByUser & " on " & FormatDateTime(dAlertDate,2)  %>
						</td>
					</tr>
<%				End If		%>
			</table>
		</p>
		<p>
			<input type="button" class="button ui-button ui-widget ui-corner-all" value="<< Back to Review List" onclick="GoToList();" /> &nbsp; &nbsp;
	<%		If Not bPermitIsCompleted And Not bIsOnHold Then	%>
				<input type="button" class="button ui-button ui-widget ui-corner-all" value="Save Changes" id="savebutton" onclick="doValidate();" /> &nbsp; &nbsp;
	<%		End If		%>
			<input type="button" class="button ui-button ui-widget ui-corner-all" value="View Permit Details" onclick="ViewDetails();" />
		</p>

		<div id="demo" class="yui-navset">
			<ul class="yui-nav">
				<li><a href="#tab1"><em>Notes</em></a></li>
				<li><a href="#tab2"><em>Attachments</em></a></li>
			</ul>            
			<div class="yui-content">
				<div id="tab1"> <!-- Notes -->
					<p>
						<div id="newreviewnotes_expand" onClick="toggleDisplayShow( 'newreviewnotes' );">
							<strong><span id="newreviewnotesimg">&ndash;</span> <u>New Permit Review Notes:</u></strong>
						</div>
						<div id="newreviewnotes">
						<table>
							<tr><td><strong>Internal Notes:</strong><br />
									<textarea id="internalcomment" name="internalcomment" rows="7" cols="100" maxlength="2500"></textarea>
								</td>
							</tr>
							<tr><td><strong>Public Notes:</strong><br />
									<textarea id="externalcomment" name="externalcomment" rows="7" cols="100" maxlength="2500"></textarea>
								</td>
							</tr>
						</table>
						</div> 
						<div id="reviewnotes_expand" onClick="toggleDisplayShow( 'reviewnotes' );">
							<strong><span id="reviewnotesimg">&ndash;</span> <u>Prior Permit Review Notes:</u></strong>
						</div>
						<div id="reviewnotes">
				<%			ShowReviewNotes iPermitReviewId		%>
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
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function GetPermitDetails( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitDetails( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT permitnumberprefix, permitnumberyear, ISNULL(permitnumber,0) AS permitnumber, "
	sSql = sSql & " P.permitstatusid, S.permitstatus, ISNULL(P.descriptionofwork,'') AS descriptionofwork, "
	sSql = sSql & " ISNULL(proposeduse, '') AS proposeduse, ISNULL(existinguse, '') AS existinguse "
	sSql = sSql & " FROM egov_permits P, egov_permitstatuses S "
	sSql = sSql & " WHERE P.permitstatusid = S.permitstatusid "
	sSql = sSql & " AND permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("permitnumber")) > CLng(0) Then 
			sPermitNo = oRs("permitnumberyear") & oRs("permitnumberprefix") & oRs("permitnumber")
		Else
			sPermitNo = "None"
		End If 
		sProposedUse = oRs("proposeduse")
		sExistingUse = oRs("existinguse")
		sDescriptionOfWork= oRs("descriptionofwork")
		sPermitStatus = oRs("permitstatus")
	Else
		sPermitNo = ""
		sProposedUse = ""
		sExistingUse = ""
		sDescriptionOfWork= ""
		sPermitStatus = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void GetReviewDetails iPermitReviewId, sPermitReviewType, sReviewDescription, iReviewerId, iReviewStatusId, sRequired 
'--------------------------------------------------------------------------------------------------
Sub GetReviewDetails( ByVal iPermitReviewId, ByRef sPermitReviewType, ByRef sReviewDescription, ByRef iReviewerId, ByRef iReviewStatusId, ByRef sRequired )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(permitreviewtype,'') AS permitreviewtype, ISNULL(reviewdescription,'') AS reviewdescription, "
	sSql = sSql & " ISNULL(revieweruserid,0) AS revieweruserid, reviewstatusid, isrequired "
	sSql = sSql & " FROM egov_permitreviews WHERE permitreviewid = " & iPermitReviewId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sPermitReviewType = oRs("permitreviewtype")
		sReviewDescription = oRs("reviewdescription")
		iReviewerId = oRs("revieweruserid")
		iReviewStatusId = oRs("reviewstatusid")
		If oRs("isrequired") Then
			sRequired = " &ndash; This Review is Required"
		End If 
	Else
		sPermitReviewType = ""
		sReviewDescription = ""
		iReviewerId = 0
		iReviewStatusId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPermitReviewers iReviewerId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitReviewers( ByVal iReviewerId )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname,isdeleted FROM users WHERE orgid = " & session("orgid") & " AND ispermitreviewer = 1 "
	sSql = sSQl & " ORDER BY isdeleted,lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""revieweruserid"">"
	response.write vbcrlf & "<option value=""0"">Unassigned</option>"
		
	Do While Not oRs.EOF
		response.write vbcrlf & "<option "
		If CLng(iReviewerId) = CLng(oRs("userid")) Then
			response.write " selected=""selected"" "
		End If 

		strReviewerName = oRs("firstname") & " " & oRs("lastname")
		if oRs("isdeleted") then strReviewerName = "[" & strReviewerName & "]"

		response.write " value=""" & oRs("userid") & """>" & strReviewerName & "</option>"
		oRs.MoveNext
	Loop 
		
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowReviewStatuses iReviewStatusId, bBlockReviewStatusChange 
'--------------------------------------------------------------------------------------------------
Sub ShowReviewStatuses( ByVal iReviewStatusId, ByVal bBlockReviewStatusChange )
	Dim sSql, oRs

	sSql = "SELECT reviewstatusid, reviewstatus FROM egov_reviewstatuses WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY reviewstatusorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""reviewstatusid"""
		If bBlockReviewStatusChange Then
			response.write " disabled=""disabled"" "
		End If 
		response.write ">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option "
			If CLng(iReviewStatusId) = CLng(oRs("reviewstatusid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " value=""" & oRs("reviewstatusid") & """>" & oRs("reviewstatus") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowReviewNotes iPermitReviewId 
'--------------------------------------------------------------------------------------------------
Sub ShowReviewNotes( ByVal iPermitReviewId )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT entrydate, ISNULL(internalcomment,'') AS internalcomment, ISNULL(externalcomment,'') AS externalcomment, "
	sSql = sSql & " S.reviewstatus, U.firstname, U.lastname, ISNULL(activitycomment,'') AS activitycomment "
	sSql = sSql & " FROM egov_permitlog L, egov_reviewstatuses S, users U "
	sSql = sSql & " WHERE S.reviewstatusid = L.reviewstatusid AND U.userid = L.adminuserid AND permitreviewid = " & iPermitReviewId
	sSql = sSql & " AND L.isreviewentry = 1 ORDER BY permitlogid DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<table id=""priorreviewnotes"" cellpadding=""3"" cellspacing=""0"" border=""0"">"
		Do While Not oRs.EOF 
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 1 Then
				response.write " class=""altrow"" "
			End If 
			response.write "><td><strong>"
			response.write oRs("firstname") & " " & oRs("lastname") & " &ndash; " & oRs("reviewstatus") & " &ndash; " & oRs("entrydate") & "</strong><br />"
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



%>
