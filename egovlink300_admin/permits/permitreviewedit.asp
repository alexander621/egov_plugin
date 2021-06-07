<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitreviewedit.asp
' AUTHOR: Steve Loar
' CREATED: 06/22/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Edits permit reviews
'
' MODIFICATION HISTORY
' 1.0   06/22/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitReviewId, iPermitId, iPermitStatusId, sLegalDescription, sListedOwner, iPermitAddressId
Dim sPermitStatus, sPermitNo, sPermitReviewType, sReviewDescription, iReviewerId, iReviewStatusId
Dim sRequired, bPermitIsCompleted, bIsOnHold, bBlockReviewStatusChange

iPermitReviewId = CLng(request("permitreviewid"))

iPermitId = GetPermitIdByPermitReviewId( iPermitReviewId ) '	in permitcommonfunctions.asp

bPermitIsCompleted = GetPermitIsCompleted( iPermitId ) '	in permitcommonfunctions.asp

bIsOnHold = GetPermitIsOnHold( iPermitId ) '	in permitcommonfunctions.asp

bBlockReviewStatusChange = GetPermitStatusBlockReview( iPermitId )	' in permitcommonfunctions.asp

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

GetReviewDetails iPermitReviewId, sPermitReviewType, sReviewDescription, iReviewerId, iReviewStatusId, sRequired 

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css?v=1" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script type="text/javascript" src="../scripts/layers.js"></script>
	<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="Javascript">
	<!--

		function doClose()
		{
			//window.close();
			//window.opener.focus();
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function doValidate()
		{
			parent.RefreshPageAfterVoid( "Review Change" );

			// Post the page
			document.frmReview.submit();
		}

		function doLoad()
		{
			setMaxLength();

<%		If request("success") <> "" Then %>
			parent.RefreshPageAfterVoid( "Review Change" );
<%		End If	%>
		}

<%		If request("success") <> "" Then 
			DisplayMessagePopUp
		End If 
%>

	//-->
	</script>

</head>

<body onload="doLoad();">

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	

	<!--BEGIN: EDIT FORM-->
	<form name="frmReview" action="permitreviewupdate.asp" method="post">
	<input type="hidden" name="permitid" value="<%=iPermitId%>" />
	<input type="hidden" name="permitreviewid" value="<%=iPermitReviewId%>" />
	<input type="hidden" name="reviewpage" value="permitreviewedit" />
	<p>
		<table cellpadding="3" border="0" cellspacing="0" id="reviewdetails">
			<tr><td>Review:</td><td nowrap="nowrap"><span class="keyinfo"><%=sPermitReviewType%></span> <%=sRequired%></td></tr>
			<tr><td>Description:</td><td nowrap="nowrap"><%=sReviewDescription%></td></tr>
			<tr><td>Reviewer:</td><td nowrap="nowrap"><% ShowPermitReviewers iReviewerId %></td></tr>
			<tr><td>Status:</td><td nowrap="nowrap"><% ShowReviewStatuses iReviewStatusId, bBlockReviewStatusChange %></td></tr>
		</table>
	</p>
	<p>
<%					
	tooltipclass=""
	tooltip = ""
	disabled = ""
	If bPermitIsCompleted or bIsOnHold Then
		tooltipclass="tooltip"
		disabled = " disabled "
		tooltip = "<span class=""tooltiptext"">You cannot save changes because:<br />The permit is complete or on hold.</span>"
	end if
	%>
		<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>"id="savebutton" onclick="doValidate();">Save Changes<%=tooltip%></button> &nbsp; &nbsp;
		<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" />
	</p>
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
	<p>
<%					
	tooltipclass=""
	tooltip = ""
	disabled = ""
	If bPermitIsCompleted or bIsOnHold Then
		tooltipclass="tooltip"
		disabled = " disabled "
		tooltip = "<span class=""tooltiptext"">You cannot save changes because:<br />The permit is complete or on hold.</span>"
	end if
	%>
		<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>"id="savebutton" onclick="doValidate();">Save Changes<%=tooltip%></button> &nbsp; &nbsp;
		<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" />
	</p>
		<div id="reviewnotes_expand" onClick="toggleDisplayShow( 'reviewnotes' );">
			<strong><span id="reviewnotesimg">&ndash;</span> <u>Prior Permit Review Notes:</u></strong>
		</div>
		<div id="reviewnotes">
<%			ShowReviewNotes iPermitReviewId		%>
		</div>
	</p>

	</form>
	<!--END: EDIT FORM-->

	</div>
</div>

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
' Function GetPermitDetails( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitDetails( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT permitnumberprefix, permitnumberyear, ISNULL(permitnumber,0) AS permitnumber, "
	sSql = sSql & " P.permitstatusid, S.permitstatus, ISNULL(P.descriptionofwork,'') AS descriptionofwork, "
	sSql = sSql & " ISNULL(proposeduse, '') AS proposeduse, ISNULL(existinguse, '') AS existinguse "
	sSql = sSql & " FROM egov_permits P, egov_permitstatuses S "
	sSql = sSql & " WHERE P.permitstatusid = S.permitstatusid "
	sSql = sSQl & " AND permitid = " & iPermitId 

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
' Sub GetReviewDetails( iPermitReviewId, sPermitReviewType, sReviewDescription, iReviewerId, iReviewStatusId, sRequired )
'--------------------------------------------------------------------------------------------------
Sub GetReviewDetails( ByVal iPermitReviewId, ByRef sPermitReviewType, ByRef sReviewDescription, ByRef iReviewerId, ByRef iReviewStatusId, ByRef sRequired )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(permitreviewtype,'') AS permitreviewtype, ISNULL(reviewdescription,'') AS reviewdescription, ISNULL(revieweruserid,0) AS revieweruserid, reviewstatusid, isrequired "
	sSql = sSQl & " FROM egov_permitreviews WHERE permitreviewid = " & iPermitReviewId

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
' Sub ShowPermitReviewers( iReviewerId )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitReviewers( iReviewerId )
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
' Sub ShowReviewStatuses( iReviewStatusId, bBlockReviewStatusChange )
'--------------------------------------------------------------------------------------------------
Sub ShowReviewStatuses( iReviewStatusId, bBlockReviewStatusChange )
	Dim sSql, oRs

	sSql = "SELECT reviewstatusid, reviewstatus FROM egov_reviewstatuses WHERE orgid = " & session("orgid")
	sSql = sSQl & " ORDER BY reviewstatusorder"

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
		If bBlockReviewStatusChange Then response.write vbcrlf & " <span class=""red"">Permit Must Be Released</span>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowReviewNotes( iPermitReviewId )
'--------------------------------------------------------------------------------------------------
Sub ShowReviewNotes( iPermitReviewId )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT entrydate, ISNULL(internalcomment,'') AS internalcomment, ISNULL(externalcomment,'') AS externalcomment, "
	sSql = sSQl & " S.reviewstatus, U.firstname, U.lastname, ISNULL(activitycomment,'') AS activitycomment "
	sSql = sSQl & " FROM egov_permitlog L, egov_reviewstatuses S, users U "
	sSql = sSQl & " WHERE S.reviewstatusid = L.reviewstatusid AND U.userid = L.adminuserid AND permitreviewid = " & iPermitReviewId
	sSql = sSQl & " AND L.isreviewentry = 1 ORDER BY permitlogid DESC"

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
				response.write replace(oRs("activitycomment"),vbcrlf,"<br />") & "<br />"
			End If 
			If oRs("internalcomment") <> "" Then 
				response.write "<strong>Internal Note:</strong><br />" & replace(oRs("internalcomment"),vbcrlf,"<br />") & "<br /><br />"
			End If 
			If oRs("externalcomment") <> "" Then 
				response.write "<strong>Public Note:</strong><br />" & replace(oRs("externalcomment"),vbcrlf,"<br />")
			End If 
			response.write "</td></tr>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</table>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 



%>
