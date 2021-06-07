<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!-- #include file="../includes/common.asp" //-->

<%
%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="Javascript">
		<!--


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
				parent.document.<%=request("parentform")%>.residentstreetnumber.value = document.frmAddress.stnumber.value;
				parent.document.<%=request("parentform")%>.useraddress.value = '';
				<% if request("sCheckType") = "FinalCheck" then %>
					parent.checkDuplicateCitizens( 'FinalUserCheckFailed' );
				<% elseif request("sCheckType") = "FinalCheckValidate" then %>
					parent.finalCheckValidate();
				<% end if %>
				parent.hideModal(window.frameElement.getAttribute("data-close"));
			}

			function doKeep()
			{
				parent.document.<%=request("parentform")%>.useraddress.value = document.frmAddress.oldstnumber.value + ' ' + document.frmAddress.stname.value;
				parent.document.<%=request("parentform")%>.residentstreetnumber.value = '';
				parent.document.<%=request("parentform")%>.address.selectedIndex = 0;
				<% if request("sCheckType") = "FinalCheck" then %>
					parent.checkDuplicateCitizens( 'FinalUserCheckFailed' );
				<% elseif request("sCheckType") = "FinalCheckValidate" then %>
					parent.finalCheckValidate();
				<% end if %>
				parent.hideModal(window.frameElement.getAttribute("data-close"));
			}


		//-->
		</script>

	</head>
	<body>
	<div id="content">
	<div id="addresspickcontent">
		<p class="addresspicker">
			The address you entered does not match any in the system. You can select a valid address from the list or, 
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
			<input type="button" class="button ui-button ui-widget ui-corner-all" name="validpick" value="Use the valid address selected" onclick="doSelect();" /> &nbsp;OR&nbsp;
			<input type="button" class="button ui-button ui-widget ui-corner-all" name="invalidpick" value="Use the address I entered" onclick="doKeep();" />
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
	sSql = sSql & " FROM egov_residentaddresses WHERE orgid = " & SESSION("orgid")
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
	oAddress.Open sSQL, Application("DSN"), 3, 1
	
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


%>
