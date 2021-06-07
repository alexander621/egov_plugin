<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!--#Include file="facility_functions.asp"-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CLIENT_TEMPLATE_PAGE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/17/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/06/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFacilityId, sFacilityName, oRs, sSql, iRowCount, x

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "edit facilities", sLevel	' In common.asp

If request("facilityid") = "" Then
	response.redirect( "facility_management.asp" )
Else 
	iFacilityId = CLng(request("facilityid"))
End If

sFacilityName = GetFacilityName( iFacilityId )

%>

<html>
<head>
	<title>E-Gov Facility Waivers</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css">
	<link rel="stylesheet" type="text/css" href="facility.css">

	<script language="Javascript">
	  <!--

		function ConfirmDelete( sWaiverName, iWaiverid, iFacilityId ) 
		{
			var msg = "Deleting this waiver removes it from all facilities. \n\n Do you wish to remove " + sWaiverName + "?";
			if (confirm(msg))
			{
				location.href='waiver_delete.asp?iWaiverid='+ iWaiverid + '&iFacilityId=' + iFacilityId;
			}
		}

		function ChangeCheck( field, iWaiverid, iFacilityId )
		{
			if (field.checked == true)
			{
	//			alert("checked");
				location.href='waiver_include.asp?iWaiverId='+ iWaiverid + '&iFacilityId=' + iFacilityId;
			}
			else
			{
	//			alert("unchecked");
				location.href='waiver_remove.asp?iWaiverId='+ iWaiverid + '&iFacilityId=' + iFacilityId;
			}
		}

		function EditWaiver( iWaiverId, iFacilityId )
		{
			location.href='waiver_edit.asp?iWaiverId=' + iWaiverId + '&iFacilityId=' + iFacilityId
		}

	  //-->
	 </script>
</head>
<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
<div id="centercontent">	
	<p>
	<h3>Recreation: Facility Waivers - <%=sFacilityName%></h3>
	<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
	</p>

	<div id="functionlinks">
		<a href="waiver_edit.asp?iWaiverId=0&iFacilityId=<%=iFacilityId%>" id="new_waiver"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;Add Waivers</a>&nbsp;&nbsp;
	</div>

	<div class="shadow">
	<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
		<tr>
			<th>Include</th><th>Waiver</th><th>Description</th><th>&nbsp;</th>
		</tr>

<%
		sSql = "SELECT waiverid, orgid, name, description FROM egov_waivers "
		sSql = sSql & "WHERE orgid = " & Session("orgid") & " ORDER BY name"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1
		
		If Not oRs.EOF Then
			iRowCount = 0
			Do While Not oRs.EOF
				' print out the lines here
				iRowCount = iRowCount + 1
				If iRowCOunt Mod 2 = 0 Then
					response.write "<tr class=""alt_row"">"
				Else
					response.write "<tr>"
				End If
				
%>
				<td>
					<form name="waiverform<%=iRowCount%>" method="post" action="availability_save.asp">
					<input type="hidden" name="waiverid" value="<%=oRs("waiverid")%>" />
					<input type="hidden" name="iFacilityId" value="<%=iFacilityId%>" />
					<input type="checkbox" name="include" value="included" <%=CheckWaiverDisplay(iFacilityId, oRs("waiverid"))%> onclick="ChangeCheck(this,<%=oRs("waiverid")%>,<%=iFacilityId%>);" />
				</td>
				<td>
					<!-- <a href="HTTPS://SECURE.ECLINK.COM/EGOVLINK/DISPLAY_WAIVER.ASP?MASK=<%=oRs("waiverid")%>" target="waiverpop"><%=oRs("name")%></a>-->
					<a href="display_waiver.aspx?MASK=<%=oRs("waiverid")%>" target="waiverpop"><%=oRs("name")%></a>
				</td>
				<td><%=oRs("description")%></td>
				<td class="action">
					<a href="javascript:EditWaiver(<%=oRs("waiverid")%>,<%=iFacilityId%>);">Edit</a>&nbsp;&nbsp;
					<a href="javascript:ConfirmDelete('<%=oRs("name")%>',<%=oRs("waiverid")%>,<%=iFacilityId%>);">Delete</a>
					</form>	
				</td>

				</tr>
<%
				oRs.MoveNext
			Loop 
		End If 
		oRs.close
		Set oRs = Nothing 
%>
	</table>
	</div>
</div>
</div>
<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  

</body>


</html>




