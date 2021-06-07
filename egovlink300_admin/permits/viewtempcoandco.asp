<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewtempcoandco.asp
' AUTHOR: Steve Loar
' CREATED: 11/20/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the temporary CO and the Certificate of Occupancy
'
' MODIFICATION HISTORY
' 1.0   11/20/2008	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sCOType, sTitle, sShowApprovedAs, sShowConstTypeOn, sShowOccTypeon, sShowOccupantsOn
Dim sFlagType, sPermitLocation, sOccupancyType, sApprovedAs, sConstType, sOccupants, sListedOwner
Dim sTopText, sBottomText, sCodeRef, sApproval, sFooter, sSubfooter, iPermitStatusId, bIssued

iPermitId = CLng(request("permitid"))

If request("cotype") <> "" Then
	sCOType = request("cotype")
	sTitle = "Temporary Certificate of Occupancy"
	sFlagType = "tco"
Else
	sCOType = ""
	sTitle = "Certificate of Occupancy"
	sFlagType = "co"
End If 

bIssued = False 

If GetPermitDate( iPermitId, sCOType & "coissueddate" ) = "" Then
	' If the issued date is blank then issue the Temp CO or CO
	IssueCO iPermitId, sCOType & "coissueddate"

	' Push out the expiration date
	PushOutPermitExpirationDate iPermitId   ' in permitcommonfunctions.asp

	' Get the status of the permit for the notes
	iPermitStatusId = GetPermitStatusId( iPermitId )		' in permitcommonfunctions.asp

	'MakeAPermitLogEntry( iPermitid, sActivity, sActivityComment, sInternalComment, sExternalComment, iPermitStatusId, iIsInspectionEntry, iIsReviewEntry, iIsActivityEntry, iPermitReviewId, iPermitInspectionId, iReviewStatusId, iInspectionStatusId )
	MakeAPermitLogEntry iPermitId, "'" & sTitle & " Issued'", "'" & sTitle & " Issued'", "NULL", "NULL", iPermitStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"

	bIssued = True 
End If 

sPermitLocation = GetPermitStreetAddress( iPermitId )

GetCOFlags iPermitId, sFlagType, sShowApprovedAs, sShowConstTypeOn, sShowOccTypeon, sShowOccupantsOn


%>

<html>
<head>
	<title>E-Gov <%=sTitle%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitprint.css" media="print" />

	<script language="Javascript">
	<!--

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		window.onload = function()
		{
<%		If bIssued Then		%>
		  //window.opener.location.reload(true);		//StatusReturn( 'UPDATED' );
		  window.opener.StatusReturn( 'UPDATED' );
<%		End If				%>
		}

	//-->
	</script>

</head>

<body id="permitbody">
 
<div id="idControls" class="noprint">
	<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
	<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" /> 
</div>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

<%		ShowHeader iPermitId, sCOType	%>

		<hr />

		<table class="codisplay" cellpadding="0" cellspacing="0" border="0">
			<tr>
				<td class="firstcocell"><span class="tempcoandcolabel">Permit Number: </span><%	response.write GetPermitNumber( iPermitId ) %></td>
				<td><span class="tempcoandcolabel">Date Issued: </span><%=GetPermitDate( iPermitId, sCOType & "coissueddate" )%></td>
			</tr>
<%			If sShowApprovedAs Or sShowOccTypeon Then	%>
				<tr>
					<td class="firstcocell">
<%					If sShowApprovedAs Then	
						sApprovedAs = GetPermitDetailItemAsString( iPermitId, "approvedas" )
						If sApprovedAs = "" Then
							sApprovedAs = "&nbsp;"
						End If 
%>
						<span class="tempcoandcolabel">Approved As: </span><%=sApprovedAs%>
<%					Else 		%>
						&nbsp;
<%					End If		%>	
					</td>
					<td>
<%	
					If sShowOccTypeon Then 
						sOccupancyType = GetPermitOccupancyTypeGroup( iPermitId )
						If sOccupancyType = "" Then 
							sOccupancyType = "&nbsp;"
						End If		%>
						<span class="tempcoandcolabel">Occupancy Use: </span><%=sOccupancyType%>
<%					Else 		%>
						&nbsp;
<%					End If		%>					
					</td>
				</tr>
<%			End If		%>

<%			If sShowConstTypeOn Or sShowOccupantsOn Then	%>
				<tr>
					<td class="firstcocell">
<%					If sShowConstTypeOn Then	
						sConstType = GetPermitConstructionType( iPermitId )
						If sConstType = "" Then
							sConstType = "&nbsp;"
						End If 
%>
						<span class="tempcoandcolabel">Type of Construction: </span><%=sConstType%>
<%					Else 		%>
						&nbsp;
<%					End If		%>	
					</td>
					<td>
<%					If sShowOccupantsOn Then	
						sOccupants = GetPermitDetailItemAsNumber( iPermitId, "occupants", "integer" )
						If sOccupants = "" Then
							sOccupants = "&nbsp;"
						End If 
%>
						<span class="tempcoandcolabel">Occupants: </span><%=sOccupants%>
<%					Else 		%>
						&nbsp;
<%					End If		%>	
					</td>
				</tr>
<%			End If		%>

			<tr><td colspan="2"><hr /></td></tr>

			<tr>
				<td class="firstcocell"><span class="tempcoandcolabel">Project Address: </span><strong><%=sPermitLocation%></strong></td>
				<td><span class="tempcoandcolabel">Owner: </span>
<%					sListedOwner = GetPermitListedOwner( iPermitId )
					If sListedOwner = "" Then
						sListedOwner = "&nbsp;"
					End If 
					response.write "<strong>" & sListedOwner &  "</strong>"
%>
				</td>
			</tr>
		</table>
<%
		sTopText = GetPermitDocumentValue( iPermitId, sCOType & "cotoptext" )
		If sTopText <> "" Then		%>
			<p class="cobodydetails">
				<%=sTopText%>
			</p>
<%		End If		%>

		<p class="cobodydetails">
			Stipulations, Conditions, Variances:<br /><br />
<%			response.write GetPermitDetailItemAsString( iPermitId, sCOType & "conotes" )		%>
		</p>

<%
		sBottomText = GetPermitDocumentValue( iPermitId, sCOType & "cobottomtext" )
		If sBottomText <> "" Then		%>
			<p class="cobodydetails">
				<%=sBottomText%>
			</p>
<%		End If		%>

<%
		sCodeRef = GetPermitDocumentValue( iPermitId, sCOType & "cocoderef" )
		If sCodeRef <> "" Then		%>
			<p class="codetails">
				<%=sCodeRef%>
			</p>
<%		End If		%>

<%
		sApproval = GetPermitDocumentValue( iPermitId, sCOType & "coapproval" )
		If sApproval <> "" Then		%>
			<p class="codetails">
				<%=sApproval%>
			</p>
<%		End If		%>

<%
		sFooter = GetPermitDocumentValue( iPermitId, sCOType & "cofooter" )
		If sFooter <> "" Then		%>
			<p class="cofooter">
				<%=sFooter%>
			</p>
<%		End If		%>

<%
		sSubooter = GetPermitDocumentValue( iPermitId, sCOType & "cosubfooter" )
		If sSubooter <> "" Then		%>
			<p class="cosubfooter">
				<%=sSubooter%>
			</p>
<%		End If		%>

	</div>
</div>
<!--END: PAGE CONTENT-->


</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub ShowtHeader( iPermitId, sCOType )
'--------------------------------------------------------------------------------------------------
Sub ShowHeader( iPermitId, sCOType )
	Dim sLogo

	response.write vbcrlf & "<table id=""coheader"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
	response.write vbcrlf & "<tr>"

	sPermitLogo = GetPermitDocumentValue( iPermitId, sCOType & "cologo" )
	If sPermitLogo <> "" Then
		response.write "<td><img src=""" & sPermitLogo & """ alt=""logo"" border=""0"" /></td>"
	Else
		response.write "<td>&nbsp;</td>"
	End If 
	response.write "<td>&nbsp;</td>"
	' Permit Number and any right titles
	response.write "<td valign=""top"" id=""coaddress"">"
	response.write GetPermitDocumentValue( iPermitId, sCOType & "coaddress" )
	response.write "</td>"
	response.write "</tr>"

	response.write vbcrlf & "<tr><td>&nbsp;</td>"
	' Center title including title and sub title
	response.write "<td align=""center"" valign=""bottom"">"
	response.write "<span id=""cotitle"">" & GetPermitDocumentValue( iPermitId, sCOType & "cotitle" ) & "</span><br />"
	response.write "<span id=""cosubtitle"">" & GetPermitDocumentValue( iPermitId, sCOType & "cosubtitle" ) & "</span>"
	response.write "</td><td>&nbsp;</td>"
	response.write vbcrlf & "</tr>"

	response.write vbcrlf & "</table>"

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub GetCOFlags( iPermitId, sFlagType, sShowApprovedAs, sShowConstTypeOn, sShowOccTypeon, sShowOccupantsOn )
'--------------------------------------------------------------------------------------------------
Sub GetCOFlags( ByVal iPermitId, ByVal sFlagType, ByRef sShowApprovedAs, ByRef sShowConstTypeOn, ByRef sShowOccTypeon, ByRef sShowOccupantsOn )
	Dim sSql, oRs

	sSql = "SELECT showapprovedason" & sFlagType & " AS showapprovedason, showconsttypeon" & sFlagType & " AS showconsttypeon, "
	sSql = sSql & " showocctypeon" & sFlagType & " AS showocctypeon, showoccupantson" & sFlagType & " AS showoccupantson "
	sSql = sSql & " FROM egov_permitpermittypes WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("showapprovedason") Then 
			sShowApprovedAs = True 
		Else
			sShowApprovedAs = False 
		End If 
		If oRs("showconsttypeon") Then 
			sShowConstTypeOn = True 
		Else
			sShowConstTypeOn = False 
		End If 
		If oRs("showocctypeon") Then 
			sShowOccTypeon = True 
		Else
			sShowOccTypeon = False 
		End If 
		If oRs("showoccupantson") Then 
			sShowOccupantsOn = True 
		Else
			sShowOccupantsOn = False 
		End If 
	Else
		sShowApprovedAs = False
		sShowConstTypeOn = False
		sShowOccTypeon = False
		sShowOccupantsOn = False
	End If 
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub IssueCO( iPermitId, sDateField )
'--------------------------------------------------------------------------------------------------
Sub IssueCO( iPermitId, sDateField )
	Dim sSql

	sSql = "UPDATE egov_permits SET " & sDateField & " = dbo.GetLocalDate(" & Session("OrgID") & ",getdate()) "
	sSql = sSql & " WHERE permitid = " & iPermitId

	RunSQL sSql

End Sub 


%>
