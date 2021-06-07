<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: constructiontypeedit.asp
' AUTHOR: Steve Loar
' CREATED: 12/13/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   12/13/07	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, iOccupancyTypeId, sUseGroupCode, sOccupancyType, iMin, iMax

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "construction type rates" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

If isFeatureOffline("construction type rates") = "Y" Then 
    response.redirect "../admin/outage_feature_offline.asp"
End If 

iOccupancyTypeId = CLng(request("occupancytypeid") )

If CLng(iOccupancyTypeId) > CLng(0) Then
	sTitle = "Edit"
	GetOccupancyType iOccupancyTypeId
Else
	sTitle = "New"
End If 

iMin = CLng(40000)
iMax = CLng(0)

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="Javascript">
	<!--

		function Another()
		{
			location.href="constructiontypeedit.asp?occupancytypeid=0";
		}

		function Validate()
		{
			var rege;
			var Ok; 

			if (document.frmRates.usegroupcode.value == '')
			{
				alert("Please provide a use group code.");
				document.frmRates.usegroupcode.focus();
				return;
			}

			if (document.frmRates.occupancytype.value == '')
			{
				alert("Please provide an occupancy type.");
				document.frmRates.occupancytype.focus();
				return;
			}

			for (var p = parseInt(document.frmRates.min.value); p <= parseInt(document.frmRates.max.value); p++)
			{
				// see if the rate exists
				if (document.getElementById("rate" + p))
				{
					// check it for being blank
					if (document.getElementById("rate" + p).value == '')
					{
						alert("All rates need a numberic value in the format '###.####', or 'NP'.");
						document.getElementById("rate" + p).focus();
						return;
					}
					else
					{
						document.getElementById("rate" + p).value = document.getElementById("rate" + p).value.toUpperCase();
						if (document.getElementById("rate" + p).value != 'NP')
						{
							// validate the rate's format
							rege = /^\d{0,3}\.{0,1}\d{0,4}$/;
							Ok = rege.test(document.getElementById("rate" + p).value);
							if (! Ok)
							{
								alert("All rates need a numberic value in the format '###.####', or 'NP'.");
								document.getElementById("rate" + p).focus();
								return;
							}
						}
					}
				}
			}
			//alert('OK');
			// All is OK so submit
			document.frmRates.submit();
		}

		function Delete() 
		{
			if (confirm("Do you wish to delete this?"))
			{
				location.href="constructiontypedelete.asp?oid=<%=iOccupancyTypeId%>";
			}
		}

<%		If request("success") <> "" Then 
			DisplayMessagePopUp 
		End If 
%>

	//-->
	</script>

</head>

<body>

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong><%=sTitle%> Construction Type Rates</strong></font><br /><br />
				<a href="constructiontypelist.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
			</p>
			<!--END: PAGE TITLE-->

		<div id="functionlinks">
<%		If CLng(iOccupancyTypeId) = CLng(0) Then %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Create" />
<%		Else %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Update" /> &nbsp; &nbsp;
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" /> &nbsp; &nbsp;
<%			If request("success") <> "" Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Another();" value="Create Another" />
<%			End If		%>
			<br />

<%		End If %>
		</div>

		<form name="frmRates" action="constructiontypeupdate.asp" method="post">
		<input type="hidden" name="occupancytypeid" value="<%=iOccupancyTypeId%>" />
		<div class="shadow">
			<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
				<tr>
					<td>
						<table cellpadding="5" cellspacing="0" border="0">
							<tr>
								<td align="right">Use Group:</td><td><input type="text" name="usegroupcode" value="<%=sUseGroupCode%>" size="25" maxlength="25" /></td>
							</tr>
							<tr>
								<td align="right">Occupancy Type:</td><td><input type="text" name="occupancytype" value="<%=sOccupancyType%>" size="100" maxlength="150" /></td>
							</tr>
							<tr>
								<td colspan="2">&nbsp;</td>
							</tr>
							<tr>
								<td class="ratetitle">&nbsp;</td><td class="ratetitle"><strong>Rates</strong></td>
							</tr>
							<tr><td colspan="2">
							<table border="0" cellspacing="0" cellpadding="0" id="ratesinput">
<%
							ShowConstructionTypeRow( session("orgid") ) 

							If CLng(iOccupancyTypeId) > CLng(0) Then
								GetRates iOccupancyTypeId
							Else
								GetInitialRates
							End If 
%>
							</table>
							</td></tr>
							<tr><td>NP=Not Permitted</td></tr>
						</table>
					</td>
				</tr>
			</table>
		</div>

		<input type="hidden" name="min" value="<%=iMin%>" />
		<input type="hidden" name="max" value="<%=iMax%>" />
		</form>
		<!--END: EDIT FORM-->

		</div>
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
' Sub ShowConstructionTypeRow( iOrgid ) 
'--------------------------------------------------------------------------------------------------
Sub ShowConstructionTypeRow( iOrgid ) 
	Dim sSql, oRs

	sSql = "SELECT constructiontype FROM egov_constructiontypes WHERE orgid = " & iOrgid & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<tr>"
		Do While Not oRs.EOF 
			response.write "<td align=""center"">" & oRs("constructiontype") & "</td>"
			oRs.MoveNext
		Loop 
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub GetOccupancyType( iOccupancyTypeId )
'--------------------------------------------------------------------------------------------------
Sub GetOccupancyType( iOccupancyTypeId )
	Dim sSql, oRs

	sSql = "SELECT usegroupcode, occupancytype FROM egov_occupancytypes WHERE occupancytypeid = " & iOccupancyTypeId
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sUseGroupCode = Replace(oRs("usegroupcode"),"""","&quot;")
		sOccupancyType = Replace(oRs("occupancytype"),"""","&quot;")
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub GetRates( iOccupancyTypeId )
'--------------------------------------------------------------------------------------------------
Sub GetRates( iOccupancyTypeId )
	Dim sSql, oRs

	sSql = "SELECT F.constructiontypeid, T.constructiontype, F.constructiontyperate, F.isnotpermitted, T.displayorder "
	sSql = sSql & " FROM egov_constructionfactors F, egov_constructiontypes T "
	sSql = sSql & " WHERE F.occupancytypeid = " & iOccupancyTypeId
	sSql = sSql & " AND F.constructiontypeid = T.constructiontypeid AND T.orgid = "& session("orgid" )
	sSql = sSql & " ORDER BY T.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<tr>"
		Do While Not oRs.EOF
			If iMin > CLng(oRs("constructiontypeid")) Then
				iMin = CLng(oRs("constructiontypeid"))
			End If 
			If iMax < CLng(oRs("constructiontypeid")) Then
				iMax = CLng(oRs("constructiontypeid"))
			End If 
	'		response.write "<td align=""right"">" & oRs("constructiontype") & ":</td>"
			response.write "<td align=""center""><input type=""text"" name=""rate" & oRs("constructiontypeid") & """ id=""rate" & oRs("constructiontypeid") & """  value="""
			If Not oRs("isnotpermitted") Then 
				response.write oRs("constructiontyperate") 
			Else
				response.write "NP"
			End If 
			response.write """ size=""8"" maxlength=""8"" /></td>"
			oRs.MoveNext
		Loop  
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub GetInitialRates( )
'--------------------------------------------------------------------------------------------------
Sub GetInitialRates( )
	Dim sSql, oRs

	sSql = "SELECT constructiontypeid, constructiontype, displayorder "
	sSql = sSql & " FROM egov_constructiontypes "
	sSql = sSql & " WHERE orgid = "& session("orgid" )
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<tr>"
		Do While Not oRs.EOF
			If iMin > CLng(oRs("constructiontypeid")) Then
				iMin = CLng(oRs("constructiontypeid"))
			End If 
			If iMax < CLng(oRs("constructiontypeid")) Then
				iMax = CLng(oRs("constructiontypeid"))
			End If 
	'		response.write "<td align=""right"">" & oRs("constructiontype") & ":</td><td>"
			response.write "<td align=""center""><input type=""text"" name=""rate" & oRs("constructiontypeid") & """ id=""rate" & oRs("constructiontypeid") & """ value="""" size=""8"" maxlength=""8"" /></td>"
			oRs.MoveNext
		Loop  
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 



%>
