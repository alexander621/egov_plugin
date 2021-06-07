<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="permitscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitsearch.asp
' AUTHOR: Steve Loar
' CREATED: 04/22/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Public accessable search for permits by address
'
' MODIFICATION HISTORY
' 1.0   04/22/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, sNumber, sStreetName, sSearch

If iorgid = 7 Then
	sTitle = sOrgName
Else
	sTitle = "E-Gov Services " & sOrgName
End If

sSearch = ""

If request("searchnumber") <> "" Then
	sNumber = request("searchnumber")
	sSearch = sSearch & "AND A.residentstreetnumber = '" & dbready_string( sNumber, 10 ) & "' "
Else
	sNumber = ""
End If 

If request("searchstreet") <> "" Then
	sStreetName = request("searchstreet")
	sSearch = sSearch & " AND (A.residentstreetname LIKE '%" & dbready_string( sStreetName, 50 ) & "%' "
	sSearch = sSearch & " OR A.residentstreetname + ' ' + A.streetsuffix = '" & dbready_string( sStreetName, 50 ) & "' "
	sSearch = sSearch & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix = '" & dbready_string( sStreetName, 50 ) & "' "
	sSearch = sSearch & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix + ' ' + A.streetdirection = '" & dbready_string( sStreetName, 50 ) & "' )"
Else
	sStreetName = ""
End If 


%>

<html>
<head>

	<title><%=sTitle%></title>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permitsstyles.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script language="JavaScript" src="../scripts/jquery-1.4.2.min.js"></script>

	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="Javascript">
	<!--

		function searchStreet()
		{
			if ($("#searchnumber").value == "" && $("searchstreet").value == "")
			{
				inlineMsg($("searchstreet").id,'<strong>Missing Values: </strong>Please enter a number or street name or both, then try your search again.',10,'searchstreet');
				return;
			}
			document.frmSearch.submit();
		}

		$(document).ready(function(){
			$(':input:visible:enabled:first').focus(); 
		});

	//-->
	</script>

</head>

<!--#Include file="../include_top.asp"-->

<!--BEGIN PAGE CONTENT-->

<%	RegisteredUserDisplay( "../" ) %>

<p>
	<font class="pagetitle"><%=GetOrgFeatureName( "building permits" )%></font>
	<br />
</p>


<!--BEGIN: Page Top Display-->
<% 
	If OrgHasDisplay( iorgid, "permit searchpagetop" ) Then
		response.write vbcrlf & "<div id=""permitsearchpagetop"">" & GetOrgDisplay( iOrgId, "permit searchpagetop" ) & "</div>"
	End If 
%>
<!--END: Page Top Display-->

<form name="frmSearch" method="post" action="permitsearch.asp">

	<p>
	<table cellpadding="2" cellspacing="0" border="0" id="permitlocationsearch">
		<tr>
			<td>
				<input type="text" id="searchnumber" name="searchnumber" value="<%=sNumber%>" size="10" maxlength="10" onkeypress="if(event.keyCode=='13'){searchStreet();return false;}" /><br />Number
			</td>
			<td>
				<input type="text" id="searchstreet" name="searchstreet" value="<%=sStreetName%>" size="50" maxlength="50" onkeypress="if(event.keyCode=='13'){searchStreet();return false;}" /><br />Street Name
			</td>
			<td valign="top" nowrap="nowrap">
				<input type="button" class="button" value="Search" onclick="searchStreet();" />
			</td>
		</tr>
	</table>
	</p>

</form>


<div id="permitsearchresults">

<%	ShowSearchResults sSearch		%>

</div>

<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ShowSearchResults sSearch
'--------------------------------------------------------------------------------------------------
Sub ShowSearchResults( ByVal sSearch )
	Dim sSql, oRs, iRowCount

	' initially there will be nothing to search on, so skip the search 
	If sSearch <> "" Then
		'response.write sSearch & "<br /><br />"

		iRowCount = 0

		sSql = "SELECT  P.permitid, P.permitnumberprefix, P.permitnumberyear, P.permitnumber, P.isonhold, P.isvoided, P.isexpired, ISNULL(A.residentstreetprefix,'') AS residentstreetprefix, "
		sSql = sSql & " permitnumberdisplay = CASE WHEN P.permitnumber IS NULL THEN '' ELSE P.permitnumberyear+P.permitnumberprefix+CAST(P.permitnumber AS varchar) END, "
		sSql = sSql & " P.applieddate, P.releaseddate, P.approveddate, P.issueddate, P.completeddate, P.expirationdate, T.permittype, T.permittypedesc, S.permitstatus, A.residentstreetnumber, ISNULL(A.residentunit,'') AS residentunit, "
		sSql = sSql & " A.residentstreetname, A.listedowner, ISNULL(A.streetsuffix,'') AS streetsuffix, ISNULL(A.streetdirection,'') AS streetdirection, ISNULL(A.residentcity,'') AS residentcity, S.statusdatedisplayed, S.permitstatusorder, "
		sSql = sSql & " ISNULL(completeddate,ISNULL(issueddate,ISNULL(approveddate,ISNULL(releaseddate,applieddate)))) AS sortdate, permitsort = CASE WHEN P.permitnumber IS NULL THEN '' ELSE 'z' END, "
		sSql = sSql & " ISNULL(A.latitude,0.00) AS latitude, ISNULL(A.longitude,0.00) AS longitude "
		sSql = sSql & " FROM egov_permits P, egov_permitpermittypes T, egov_permitstatuses S, egov_permitaddress A " 
		sSql = sSql & " WHERE P.orgid = " & iOrgid & " AND P.isbuildingpermit = 1 AND T.permitid = P.permitid AND S.isinitialstatus = 0 "
		sSql = sSql & " AND P.permitstatusid = S.permitstatusid AND A.permitid = P.permitid AND P.isvoided = 0 " & sSearch 
		sSql = sSql & " ORDER BY P.permitnumberyear DESC, P.permitnumber DESC"
		'response.write sSql & "<br /><br />"
		session("ShowSearchResultsSql") = sSql

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1
		session("ShowSearchResultsSql") = ""

		If Not oRs.EOF Then 
			response.write vbcrlf & "<table id=""permitresults"" cellpadding=""1"" cellspacing=""0"" border=""0"">"
			response.write vbcrlf & "<tr><th>Permit #</th><th>Permit<br />Type</th><th>Address/<br />Owner</th><th>Applicant</th><th>Status</th><th>Status<br />Date</th></tr>"

			Do While Not oRs.EOF
				iRowCount = iRowCount + 1
				response.write vbcrlf & "<tr id=""" & iRowCount & """"
				If iRowCount Mod 2 = 0 Then
					response.write " class=""altrow"" "
				End If 
				response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"

				response.write "<td title=""click to view"" onClick=""location.href='permitdetails.asp?p=" & oRs("permitid") & "';"" nowrap=""nowrap"">"
				If oRs("permitnumberdisplay") = "" Then
					response.write "&nbsp;" & iRowCount
				Else
					response.write "&nbsp;" & GetPermitNumber( oRs("permitid") )
				End If 
				response.write "</td>"

				response.write "<td align=""left"" title=""click to view"" onClick=""location.href='permitdetails.asp?p=" & oRs("permitid") & "';"">" & oRs("permittype") & " &ndash; " & oRs("permittypedesc") & "</td>"

				response.write "<td title=""click to view"" onClick=""location.href='permitdetails.asp?p=" & oRs("permitid") & "';"">"
				response.write "&nbsp;" & oRs("residentstreetnumber")
				If oRs("residentstreetprefix") <> "" Then
					response.write " " & oRs("residentstreetprefix")
				End If 
				response.write " " & oRs("residentstreetname")
				If oRs("streetsuffix") <> "" Then
					response.write " " & oRs("streetsuffix")
				End If 
				If oRs("streetdirection") <> "" Then
					response.write " " & oRs("streetdirection")
				End If 
				response.write " " & oRs("residentunit")
				response.write "<br />" & " &nbsp;&nbsp;" & oRs("listedowner")
				response.write "</td>"

				response.write "<td title=""click to view"" onClick=""location.href='permitdetails.asp?p=" & oRs("permitid") & "';"">"
				response.write GetPermitApplicantName( oRs("permitid") )
				response.write "</td>"

				If oRs("isonhold") Or oRs("isvoided") Or oRs("isexpired") Then 
					response.write "<td align=""center"" title=""click to view"" onClick=""location.href='permitdetails.asp?p=" & oRs("permitid") & "';"">"
					If oRs("isonhold") Then 
						response.write "On Hold"
					Else
						response.write "Expired"
					End If 
					response.write "</td>"
					response.write "<td align=""center"" title=""click to view"" onClick=""location.href='permitdetails.asp?p=" & oRs("permitid") & "';"">"
					If oRs("isexpired") And Not IsNull(oRs("expirationdate")) Then 
						response.write DateValue(oRs("expirationdate"))
					Else 
						response.write GetLastLogDate( oRs("permitid") )   ' in permitcommonfunctions.asp
					End If 
					response.write "</td>"
				Else 

					response.write "<td align=""center"" title=""click to view"" onClick=""location.href='permitdetails.asp?p=" & oRs("permitid") & "';"">" & oRs("permitstatus") & "</td>"

					response.write "<td align=""center"" title=""click to view"" onClick=""location.href='permitdetails.asp?p=" & oRs("permitid") & "';"">"
					Select Case oRs("statusdatedisplayed") 
						Case "applieddate"
							response.write DateValue(oRs("applieddate"))
						Case "releaseddate"
							response.write DateValue(oRs("releaseddate"))
						Case "approveddate"
							response.write DateValue(oRs("approveddate"))
						Case "issueddate"
							response.write DateValue(oRs("issueddate"))
						Case "completeddate"
							response.write DateValue(oRs("completeddate"))
					End Select 
					response.write "</td>"
				End If 

				response.write "</tr>"
				oRs.MoveNext 
			Loop 

			response.write vbcrlf & "</table>"
		Else
			response.write vbcrlf & "<p>No permits could be found that match your search criteria.</p>"
		End If

		oRs.Close
		Set oRs = Nothing 

	End If 

End Sub





%>
