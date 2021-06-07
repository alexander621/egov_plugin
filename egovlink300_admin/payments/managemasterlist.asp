<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: managemasterlist.asp
' AUTHOR: Steve Loar
' CREATED: 08/1/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows editing of the Rye Master List for commuter parking renewals
'
' MODIFICATION HISTORY
' 1.0   08/1/2011   Steve Loar - INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sApplicantFirstNameSearch, sApplicantLastNameSearch, sPermitHolderTypeSearch, sLoadMsg

sLevel = "../"  'Override of value from common.asp

'Check the page availability and user access rights in one call
PageDisplayCheck "managemasterlist", sLevel	 'In common.asp

If request("permitholdertypesearch") <> "" Then 
	sPermitHolderTypeSearch = request("permitholdertypesearch")
Else
	sPermitHolderTypeSearch = ""
End If 

If request("applicantfirstnamesearch") <> "" Then 
	sApplicantFirstNameSearch = request("applicantfirstnamesearch")
Else
	sApplicantFirstNameSearch = ""
End If 

If request("applicantlastnamesearch") <> "" Then 
	sApplicantLastNameSearch = request("applicantlastnamesearch")
Else
	sApplicantLastNameSearch = ""
End If 

If request("s") <> "" Then
	If request("s") = "u" Then
		sLoadMsg = "displayScreenMsg('Your Changes Were Successfully Saved');"
	End If 
End If 


%>
<html>
<head>
 <title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="payment_styles.css" />

	<script type="text/javascript" src="https://code.jquery.com/jquery-1.5.min.js"></script>

	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="javascript">
	<!--

		function displayScreenMsg( iMsg ) 
		{
			if(iMsg!="") 
			{
				$("#screenMsg").html( "*** " + iMsg + " ***" );
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html( "" );
		}

		function validateSearch()
		{
			var OkToSearch = true;
			
			if ( $("#applicantlastnamesearch").val() == "" )
			{
				$("#applicantlastnamesearch").focus();
				inlineMsg("applicantlastnamesearch",'<strong>Missing Value: </strong>Please enter an Applicant Last Name.',8,"applicantlastnamesearch");
				OkToSearch = false;
			}

			if ( $("#applicantfirstnamesearch").val() == "" )
			{
				$("#applicantfirstnamesearch").focus();
				inlineMsg("applicantfirstnamesearch",'<strong>Missing Value: </strong>Please enter an Applicant First Name.',8,"applicantfirstnamesearch");
				OkToSearch = false;
			}

			if ( OkToSearch )
			{
				document.frmSearch.submit();
			}
		}

		function validate()
		{
			var OkToSave = true;

			if ( $("#applicantzip").val() == "" )
			{
				$("#applicantzip").focus();
				inlineMsg("applicantzip",'<strong>Missing Value: </strong>Please enter an Applicant Zip.',8,"applicantzip");
				OkToSave = false;
			}

			if ( $("#applicantstate").val() == "" )
			{
				$("#applicantstate").focus();
				inlineMsg("applicantstate",'<strong>Missing Value: </strong>Please enter an Applicant State.',8,"applicantstate");
				OkToSave = false;
			}

			if ( $("#applicantcity").val() == "" )
			{
				$("#applicantcity").focus();
				inlineMsg("applicantcity",'<strong>Missing Value: </strong>Please enter an Applicant City.',8,"applicantcity");
				OkToSave = false;
			}

			if ( $("#applicantaddress").val() == "" )
			{
				$("#applicantaddress").focus();
				inlineMsg("applicantaddress",'<strong>Missing Value: </strong>Please enter an Applicant Address.',8,"applicantaddress");
				OkToSave = false;
			}
			
			if ( $("#applicantlastname").val() == "" )
			{
				$("#applicantlastname").focus();
				inlineMsg("applicantlastname",'<strong>Missing Value: </strong>Please enter an Applicant Last Name.',8,"applicantlastname");
				OkToSave = false;
			}

			if ( $("#applicantfirstname").val() == "" )
			{
				$("#applicantfirstname").focus();
				inlineMsg("applicantfirstname",'<strong>Missing Value: </strong>Please enter an Applicant First Name.',8,"applicantfirstname");
				OkToSave = false;
			}

			if ( OkToSave )
			{
				document.frmApplicant.submit();
			}
		}

		function SetUpPage()
		{
			<%=sLoadMsg%>
			$("#permitholdertypesearch").focus();
		}

	//-->
	</script>

</head>

<body onload="SetUpPage();">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->

<div id="content">
	<div id="centercontent">

		<!--BEGIN: PAGE TITLE-->
		<p>
			<font size="+1"><strong>Commuter Renewal/Waitlist Master List</strong></font><br />
		</p>
		<!--END: PAGE TITLE-->

		<div id="screenmsgbox">
			<span id="screenMsg"></span>
			<input type="button" class="button" value="Export Non-Renewers" onClick="location.href='nonrenewalexport.asp'" />
		</div>

		<form name="frmSearch" method="post" action="managemasterlist.asp">
			<fieldset id="scanentryfieldset">
				<strong>Search for an applicant</strong>

				<table id="searchpicks" cellspacing="0" cellpadding="3" border="0">
					<tr>
						<td class="searchlabel" align="right"><strong>Permit Holder Type:</strong></td>
						<td>
<%
							ShowPermitTypes sPermitHolderTypeSearch, "permitholdertypesearch"
%>
						</td>
					</tr>
					<tr>
						<td class="searchlabel" align="right"><strong>Applicant First Name:</strong></td>
						<td>
							<input type="text" name="applicantfirstnamesearch" id="applicantfirstnamesearch" size="50" maxlength="50" placeholder="Applicant First Name" value="<%= sApplicantFirstNameSearch %>" />
						</td>
					</tr>
					<tr>
						<td class="searchlabel" align="right"><strong>Applicant Last Name:</strong></td>
						<td>
							<input type="text" name="applicantlastnamesearch" id="applicantlastnamesearch" size="50" maxlength="50" placeholder="Applicant Last Name" value="<%= sApplicantLastNameSearch %>" />
						</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td>
							<input type="button" class="button" value="Search" onclick="validateSearch();" />
						</td>
					</tr>
				</table>
			</fieldset>
		</form>

<%
		If sPermitHolderTypeSearch <> "" Then 
			' Show data here
			ShowApplicantData sPermitHolderTypeSearch, sApplicantFirstNameSearch, sApplicantLastNameSearch
		End If 
%>

	</div>
</div>



<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ShowPermitTypePicks sPermitHolderTypeSearch, sSelectName
'--------------------------------------------------------------------------------------------------
Sub ShowPermitTypes( ByVal sPermitHolderTypeSearch, ByVal sSelectName )
	Dim sSql, oRs, sSelected

	sSql = "SELECT permitholdertype "
	sSql = sSql & "FROM egov_ryepermitrenewals "
	sSql = sSql & "WHERE orgid = " & session("orgid")
	sSql = sSql & " GROUP BY permitholdertype ORDER BY permitholdertype"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select id=""" & sSelectName & """ name=""" & sSelectName & """>"
		Do While Not oRs.EOF
			If oRs("permitholdertype") = sPermitHolderTypeSearch Then
				sSelected = " selected=""selected"" "
			Else
				sSelected = ""
			End If 

			if instr(",Current Railroad Permit Holder,Current Highland/Cedar Permit Holder,","," & oRs("permitholdertype") & ",") > 0 then
				sSelected = sSelected & " style=""font-weight:bold"""
			end if
			response.write vbcrlf & "<option value=""" & oRs("permitholdertype") & """" & sSelected & ">" & oRs("permitholdertype") & "</option>"
			oRs.MoveNext 
		Loop
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Sub 
Sub ShowPermitTypePicks( ByVal sPermitHolderTypeSearch, ByVal sSelectName )
	Dim sSql, oRs, sSelected

	sSql = "SELECT permitholdertype "
	sSql = sSql & "FROM egov_ryepermitholdertypes "
	sSql = sSql & "WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY permitholdertype"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select id=""" & sSelectName & """ name=""" & sSelectName & """>"
		Do While Not oRs.EOF
			If oRs("permitholdertype") = sPermitHolderTypeSearch Then
				sSelected = " selected=""selected"" "
			Else
				sSelected = ""
			End If 
			response.write vbcrlf & "<option value=""" & oRs("permitholdertype") & """" & sSelected & ">" & oRs("permitholdertype") & "</option>"
			oRs.MoveNext 
		Loop
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowApplicantData sPermitHolderTypeSearch, sApplicantFirstNameSearch, sApplicantLastNameSearch
'--------------------------------------------------------------------------------------------------
Sub ShowApplicantData( ByVal sPermitHolderTypeSearch, ByVal sApplicantFirstNameSearch, ByVal sApplicantLastNameSearch )
	Dim sSql, oRs

	sSql = "SELECT renewalid, permitholdertype, applicantfirstname, applicantlastname, ISNULL(applicantaddress,'') AS applicantaddress, "
	sSql = sSql & "ISNULL(applicantcity,'') AS applicantcity, ISNULL(applicantstate,'') AS applicantstate, "
	sSql = sSql & "ISNULL(applicantzip,'') AS applicantzip, ISNULL(applicantphone,'') AS applicantphone, "
	sSql = sSql & "ISNULL(vehiclelicense,'') AS vehiclelicense, ISNULL(rr,'') AS rr, ISNULL(hc,'') AS hc "
	sSql = sSql & "FROM egov_ryepermitrenewals "
	sSql = sSql & "WHERE permitholdertype = '" & dbsafe(sPermitHolderTypeSearch) & "' " 
	sSql = sSql & "AND LOWER(applicantfirstname) = '" & LCase(dbsafe(sApplicantFirstNameSearch)) & "' " 
	sSql = sSql & "AND LOWER(applicantlastname) = '" & LCase(dbsafe(sApplicantLastNameSearch)) & "' " 
	sSql = sSql & "AND orgid = " & session("orgid") & " AND year = " & Year(Now()) 
	sSql = sSql & " ORDER BY renewalid"

	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<form name=""frmApplicant"" method=""post"" action=""managemasterlistupdate.asp"">"

		response.write vbcrfl& "<input type=""hidden"" id=""renewalid"" name=""renewalid"" value=""" & oRs("renewalid") & """ />"

		response.write vbcrlf & "<table id=""applicantdata"" cellspacing=""0"" cellpadding=""3"" border=""0"">"

		response.write vbcrlf & "<tr>"
		response.write "<td class=""datalabel"" align=""right"">Permit Holder Type:</td>"
		response.write "<td>" 
		ShowPermitTypePicks oRs("permitholdertype"), "permitholdertype"
		response.write "</td>"
		response.write "</tr>"

		response.write vbcrlf & "<tr>"
		response.write "<td class=""datalabel"" align=""right""><span class=""requiredindicator"">* </span>Applicant First Name:</td>"
		response.write "<td><input type=""text"" id=""applicantfirstname"" name=""applicantfirstname"" value=""" & oRs("applicantfirstname") & """ size=""50"" maxlength=""50"" /></td>"
		response.write "</tr>"

		response.write vbcrlf & "<tr>"
		response.write "<td class=""datalabel"" align=""right""><span class=""requiredindicator"">* </span>Applicant Last Name:</td>"
		response.write "<td><input type=""text"" id=""applicantlastname"" name=""applicantlastname"" value=""" & oRs("applicantlastname") & """ size=""50"" maxlength=""50"" /></td>"
		response.write "</tr>"

		response.write vbcrlf & "<tr>"
		response.write "<td class=""datalabel"" align=""right""><span class=""requiredindicator"">* </span>Applicant Address:</td>"
		response.write "<td><input type=""text"" id=""applicantaddress"" name=""applicantaddress"" value=""" & oRs("applicantaddress") & """ size=""100"" maxlength=""100"" /></td>"
		response.write "</tr>"

		response.write vbcrlf & "<tr>"
		response.write "<td class=""datalabel"" align=""right""><span class=""requiredindicator"">* </span>Applicant City:</td>"
		response.write "<td><input type=""text"" id=""applicantcity"" name=""applicantcity"" value=""" & oRs("applicantcity") & """ size=""50"" maxlength=""50"" /></td>"
		response.write "</tr>"

		response.write vbcrlf & "<tr>"
		response.write "<td class=""datalabel"" align=""right""><span class=""requiredindicator"">* </span>Applicant State:</td>"
		response.write "<td><input type=""text"" id=""applicantstate"" name=""applicantstate""  value=""" & oRs("applicantstate") & """ size=""2"" maxlength=""2"" /></td>"
		response.write "</tr>"

		response.write vbcrlf & "<tr>"
		response.write "<td class=""datalabel"" align=""right""><span class=""requiredindicator"">* </span>Applicant Zip:</td>"
		response.write "<td><input type=""text"" id=""applicantzip"" name=""applicantzip"" value=""" & oRs("applicantzip") & """ size=""10"" maxlength=""10"" /></td>"
		response.write "</tr>"

		response.write vbcrlf & "<tr>"
		response.write "<td class=""datalabel"" align=""right"">Applicant Phone:</td>"
		response.write "<td><input type=""text"" id=""applicantphone"" name=""applicantphone"" value=""" & oRs("applicantphone") & """ size=""25"" maxlength=""25"" /></td>"
		response.write "</tr>"

		response.write vbcrlf & "<tr>"
		response.write "<td class=""datalabel"" align=""right"">Vehicle License:</td>"
		response.write "<td><input type=""text"" id=""vehiclelicense"" name=""vehiclelicense"" value=""" & oRs("vehiclelicense") & """ size=""25"" maxlength=""25"" /></td>"
		response.write "</tr>"

'		response.write vbcrlf & "<tr>"
'		response.write "<td class=""datalabel"" align=""right"">RR:</td>"
'		response.write "<td><input type=""text"" id=""rr"" name=""rr"" value=""" & oRs("rr") & """ size=""5"" maxlength=""5"" /></td>"
'		response.write "</tr>"
'
'		response.write vbcrlf & "<tr>"
'		response.write "<td class=""datalabel"" align=""right"">H/C:</td>"
'		response.write "<td><input type=""text"" id=""hc"" name=""hc"" value=""" & oRs("hc") & """ size=""5"" maxlength=""5"" /></td>"
'		response.write "</tr>"

		response.write vbcrlf & "<tr>"
		response.write "<td class=""datalabel"">&nbsp;</td>"
		response.write "<td><span class=""requiredindicator"">* Indicates the field is required</span></td>"
		response.write "</tr>"

		response.write vbcrlf & "<tr>"
		response.write "<td class=""datalabel"">&nbsp;</td>"
		response.write "<td><input type=""button"" class=""button"" value=""Save Changes"" onclick=""validate();"" /></td>"
		response.write "</tr>"

		response.write vbcrlf & "</table>"

		response.write vbcrlf & "</form>"
	Else
		response.write vbcrlf & "<p>No information was found. Please check your search criteria and try again.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Sub



%>


