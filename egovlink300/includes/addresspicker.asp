<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#Include file="../include_top_functions.asp"-->
<!-- #include file="common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: addresspicker.asp
' AUTHOR: Steve Loar
' CREATED: 08/28/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This displays valid addresses for selection
'
' MODIFICATION HISTORY
' 1.0   08/28/07	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />

		<script language="Javascript">
		<!--

			//window.onunload = function (e) {saybye('<%=request("saving")%>'); register(e)};

			function saybye( sString )
			{
				if (sString == 'yes')
				{
					//alert('You need to click \'Create User\' again to complete this registration.');
				}
			}

			function doSelect()
			{
				// check that an address was picked
				if (document.frmAddress.stnumber.selectedIndex < 0)
				{
					alert('Please select an address from the list first.');
					return false;
				}
				window.opener.document.register.residentstreetnumber.value = document.frmAddress.stnumber.value;
				window.opener.document.register.egov_users_useraddress.value = '';
				<% if request("sCheckType") = "FinalCheck" then %>
					window.opener.finalCheckValidate();
				<% end if %>
				window.close();
			}

			function doKeep()
			{
				window.opener.document.register.egov_users_useraddress.value = document.frmAddress.oldstnumber.value + ' ' + document.frmAddress.stname.value;
				window.opener.document.register.residentstreetnumber.value = '';
				window.opener.document.register.skip_address.selectedIndex = 0;
				<% if request("sCheckType") = "FinalCheck" then %>
					window.opener.finalCheckValidate();
				<% end if %>
				window.close();
			}


		//-->
		</script>

	</head>
	<body>
	<div id="content">
	<div id="addresspickcontent">

		<font size="+1"><b>Invalid Address Selection</b></font><br /><br />
		<p class="addresspicker">
			The address you entered does not match any in our system. You can select a valid address from the list or, 
			if you are certain the address you entered is correct, select to use the address you supplied and continue registration.
		</p>
		<form name="frmAddress" action="addresspicker.asp" method="post">
			<strong>The address you entered</strong><br />
			<input type="text" name="oldstnumber" value="<%=request("stnumber")%>" disabled="disabled" size="8" maxlength="10" /> &nbsp; 
			<input type="text" name="stname" value="<%=request("stname")%>" disabled="disabled" size="50" maxlength="50" />
			<div id="addresspicklist">
				<strong>Valid Address Choices </strong><br />
<%				ShowAddressPicks( request("stname") )	%>
			</div>
			<input type="button" class="button" name="validpick" value="Use the valid address selected" onclick="doSelect();" /> &nbsp;OR&nbsp;
			<input type="button" class="button" name="invalidpick" value="Use the address I entered" onclick="doKeep();" />
		</form>
	</div>
	</div>
	</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' Sub ShowAddressPicks( sStreetName )
'--------------------------------------------------------------------------------------------------
Sub ShowAddressPicks( sStreetName )
	Dim sSql, oAddress, sOption

	sSql = "SELECT DISTINCT residentstreetnumber, residentstreetname, CAST(residentstreetnumber AS INT) AS ordernumb, "
	sSql = sSQl & " ISNULL(residentstreetprefix,'') AS residentstreetprefix, ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & " FROM egov_residentaddresses WHERE orgid = " & iOrgid
	sSql = sSql & " AND (residentstreetname = '" & dbsafe(sStreetName) & "' "
	sSql = sSql & " OR residentstreetname + ' ' + streetsuffix = '" & dbsafe(sStreetName) & "' "
	sSQL = sSQL & " OR residentstreetname + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
	sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
	sSql = sSql & " OR residentstreetprefix + ' ' + residentstreetname = '" & dbsafe(sStreetName) & "' "
	sSql = sSql & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = '" & dbsafe(sStreetName) & "' "
	sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
	sSql = sSql & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' )"
	sSql = sSQl & " ORDER BY 2, 5, 6, 4, 3, 1"

	Set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSQL, Application("DSN"), 0, 1
	
	If not oAddress.EOF Then
		response.write vbcrlf & "<select id=""stnumber"" name=""stnumber"" size=""10"">"
		Do While NOT oAddress.EOF 
			sOption = oAddress("residentstreetnumber")
			If oAddress("residentstreetprefix") <> "" Then
				sOption = sOption & " " & oAddress("residentstreetprefix")
			End If 
			sOption = sOption & " " & oAddress("residentstreetname")
			If oAddress("streetsuffix") <> "" Then
				sOption = sOption & " "  & oAddress("streetsuffix")
			End If
			If oAddress("streetdirection") <> "" Then
				sOption = sOption & " "  & oAddress("streetdirection")
			End If

			response.write vbcrlf & "<option value=""" & oAddress("residentstreetnumber") & """ >"  
			response.write sOption & "</option>"
			oAddress.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oAddress.close
	Set oAddress = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowAddressPicksOld( sStreetName )
'--------------------------------------------------------------------------------------------------
Sub ShowAddressPicksOld( sStreetName )
	Dim sSql, oAddress

	sSql = "Select residentstreetnumber, residentstreetname From egov_residentaddresses Where orgid = " & iOrgid
	sSql = sSql & " and residentstreetname = '" & dbsafe(sStreetName) & "' ORDER BY residentstreetname, CAST(residentstreetnumber as int)"

	Set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSQL, Application("DSN"), 0, 1
	
	If not oAddress.EOF Then
		response.write vbcrlf & "<select id=""stnumber"" name=""stnumber"" size=""10"">"
		Do While NOT oAddress.EOF 
			response.write vbcrlf & "<option value=""" & oAddress("residentstreetnumber") & """ >"  
			response.write oAddress("residentstreetnumber") & "  " & oAddress("residentstreetname") & "</option>"
			oAddress.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oAddress.close
	Set oAddress = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
	If Not VarType( strDB ) = vbString Then 
		DBsafe = strDB
	Else 
		DBsafe = Replace( strDB, "'", "''" )
	End If 
End Function



%>