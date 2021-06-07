<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: LOCATION_MGMT.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 03/21/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/17/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1   04/26/06   TERRY FOSTER - MADE FUNCTIONAL
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sName, sAddress1, sAddress2, sCity, sState, sZip
Dim ilocationID 

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "locations" ) Then
	If Not UserHasPermission( Session("UserId"), "rentallocations" ) Then
		response.redirect sLevel & "permissiondenied.asp"
	End If 
End If 


' GET location ID
If request("locationid") = "" Or Not IsNumeric(request("locationid")) Or request("locationid") = 0 Then
	' CREATE NEW location
	ilocationID = 0
	sTitle = "Add New Location"
	sLinkText = "Create Location"
Else
	' EDIT EXISTING location
	ilocationID = request("locationid")
	sTitle = "Edit Location"
	sLinkText = "Save Changes"
	
	' GET location INFORMATION
	GetlocationInfo ilocationID
End If

If request("msg") <> "" Then
	If request("msg") = "i" Then
		sLoadMsg = "displayScreenMsg('This Location Was Successfully Created');"
	End If
	If request("msg") = "s" Then
		sLoadMsg = "displayScreenMsg('Your Changes Were Successfully Saved');"
	End If 
End If 


%>


<html>
<head>
 	<title>E-Gov Administration Console</title>

	 <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

 	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="classes.css" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="Javascript">
	<!--

		function deleteConfirm() 
		{
			if(confirm('Do you wish to delete this location?')) 
			{
				window.location="location_delete.asp?iLocationid=<%=ilocationID%>";
			}
		}

		function displayScreenMsg( iMsg ) 
		{
			if(iMsg!="") 
			{
				$("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("screenMsg").innerHTML = "";
		}

		function SetUpPage()
		{
			<%=sLoadMsg%>
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
	<font size="+1"><strong>Recreation: <%=sTitle%></strong></font><br />
</p>
<!--END: PAGE TITLE-->

<p>
	<span id="screenMsg"></span>
</p>


<!--BEGIN: FUNCTION LINKS-->
<p>
	<input type="button" class="button" value="<< Back" onclick="location.href='location_mgmt.asp';" /> &nbsp; 
	<input type="button" class="button" value="Delete" onclick="deleteConfirm();" /> &nbsp; 
	<input type="button" class="button" value="<%=sLinkText%>" id="savebutton" onclick="javascript:document.frmlocation.submit();" />
</p>
<!--END: FUNCTION LINKS-->


<!--BEGIN: EDIT FORM-->
<form name="frmlocation" action="location_save.asp" method="post">
<input type="hidden" name="ilocationid" value="<%=ilocationID%>" >

<div class="shadow">
	<table cellpadding="5" cellspacing="0" border="0" class="locationlist">
		<tr>
			<th>Location Information</th>
		</tr>
		<tr>
			<td>
				<table>
					<tr>
						<td>Name:</td><td><input type="text" name="sName" maxlength="50" size="50" value="<%=sName%>" /></td>
					</tr>
					<tr>
						<td>Address 1:</td><td><input type="text" name="sAddress1" size="100" maxlength="100" value="<%=sAddress1%>" /></td>
					</tr>
					<tr>
						<td>Address 2:</td><td><input type="text" name="sAddress2" size="100" maxlength="100" value="<%=sAddress2%>" /></td>
					</tr>
					<tr>
						<td>City:</td><td><input type="text" name="sCity" size="40" maxlength="40" value="<%=sCity%>" /></td>
					</tr>
					<tr>
						<td>State:</td><td><input type="text" name="sState" size="2" maxlength="2" value="<%=sState%>" /></td>
					</tr>
					<tr>
						<td>Zip:</td><td><input type="text" name="sZip" size="10" maxlength="10" value="<%=sZip%>" /></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</div>

</form>
<!--END: EDIT FORM-->

	</div>
</div>
<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  

</body>


</html>



<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void GETlocationINFO(IlocationID)
'--------------------------------------------------------------------------------------------------
Sub GetlocationInfo( ByVal ilocationID )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_class_location WHERE locationid = " & ilocationID 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If NOT oRs.EOF Then
		sName = oRs("name")
		sAddress1 = oRs("address1")
		sAddress2 = oRs("address2")
		sCity = oRs("city")
		sState = oRs("state")
		sZip = oRs("zip")
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub 


%>


