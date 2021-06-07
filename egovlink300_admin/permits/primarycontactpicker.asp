<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: primarycontactpicker.asp
' AUTHOR: Steve Loar
' CREATED: 03/27/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Selects primary contacts from the registered users
'
' MODIFICATION HISTORY
' 1.0   03/27/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iPrimaryContactUserId, bCanSelect

bCanSelect = False 
iPermitId = CLng(request("permitid"))
iPrimaryContactUserId = GetPrimaryContactUserId( iPermitId )

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="Javascript">
		<!--

		function doNewContact()
		{
			location.href="permitapplicantedit.asp?userid=0&updatetitle=1&detailid=primarycontactdetails";
		}

		function doSelect()
		{
			if (document.frmContact.userid.options[document.frmContact.userid.selectedIndex].value > 0)
			{
				window.opener.document.getElementById("primarycontactdetails").innerHTML = document.frmContact.userid.options[document.frmContact.userid.selectedIndex].text;
			}
			else
			{
				window.opener.document.getElementById("primarycontactdetails").innerHTML = 'None Selected ';
			}
			window.opener.document.getElementById("isprimarycontactuserid").value = document.frmContact.userid.options[document.frmContact.userid.selectedIndex].value;
			doClose();
		}

		function doClose()
		{
			window.close();
			window.opener.focus();
		}
		
		function SearchCitizens( iSearchStart )
		{
			var optiontext;
			var optionchanged;
			//alert(document.BuyerForm.searchname.value);
			var searchtext = document.frmContact.searchname.value;
			var searchchanged = searchtext.toLowerCase();

			iSearchStart = parseInt(iSearchStart) + 1;
			
			for (x=iSearchStart; x < document.frmContact.userid.length ; x++)
			{
				optiontext = document.frmContact.userid.options[x].text;
				optionchanged = optiontext.toLowerCase();
				if (optionchanged.indexOf(searchchanged) != -1)
				{
					document.frmContact.userid.selectedIndex = x;
					document.frmContact.results.value = 'Possible Match Found.';
					document.getElementById('searchresults').innerHTML = 'Possible Match Found.';
					document.frmContact.searchstart.value = x;
					return;
				}
			}
			document.frmContact.results.value = 'No Match Found.';
			document.getElementById('searchresults').innerHTML = 'No Match Found.';
			document.frmContact.searchstart.value = -1;
		}

		function ClearSearch()
		{
			document.frmContact.searchstart.value = -1;
		}

		function UserPick()
		{
			document.frmContact.searchname.value = '';
			document.frmContact.results.value = '';
			document.getElementById('searchresults').innerHTML = '';
			document.frmContact.searchstart.value = -1;
		}

		//-->
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent">
				<font size="+1"><strong>Primary Contact Selection</strong></font><br /><br />
				<form name="frmContact" action="primarycontactpicker.asp" method="post">
					<p>
					<% ShowPrimaryContactPicks iPrimaryContactUserId, bCanSelect %>
					<% If bCanSelect Then %>
						<br /><br />
						<input type="text" name="searchname" size="25" maxlength="50" onkeypress="if(event.keyCode=='13'){SearchCitizens(document.frmContact.searchstart.value);return false;}" />
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Search" onclick="SearchCitizens(document.frmContact.searchstart.value);" /> &nbsp;
						<input type="hidden" name="searchstart" value="-1" />
						<input type="hidden" name="results" value="" />
						<span id="searchresults"></span>
						</p><p>
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Select" onclick="doSelect();" /> &nbsp; &nbsp; 
					<% Else %>
						</p><p>
					<% End If %> 
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="Cancel" onclick="doClose();" /> &nbsp; &nbsp; 
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="New Contact" onclick="doNewContact();" />
					</p>
				</form>
			</div>
		</div>
	</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

Function GetPrimaryContactUserId( iPermitId )
End Function 


Sub ShowPrimaryContactPicks( iPrimaryContactUserId, bCanSelect )
	Dim oCmd, oRs

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserWithAddressList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, Session("OrgID"))
	    Set oRs = .Execute
	End With

	If Not oRs.EOF Then 
		bCanSelect = True 
		response.write vbcrlf & "<select name=""userid"" onchange=""UserPick();"">"
		response.write vbcrlf & "<option value=""0"">Select a registered user from the list...</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("userid") & """"
			If CLng(iUserId) = CLng(oRs("userid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("userlname") & ", " & oRs("userfname")
			If Not IsNull( oRs("useraddress")) And  oRs("useraddress") <> "" Then 
				response.write " &ndash; " & oRs("useraddress")
			End If 
			response.write "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	Else
		response.write "No users have been registered."
	End If 
		
	oRs.Close
	Set oRs = Nothing
	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowUserDropDown( iUserId )
'--------------------------------------------------------------------------------------------------
Sub ShowUserDropDown( iUserId )
	Dim oCmd, oResident

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserWithAddressList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, Session("OrgID"))
	    Set oResident = .Execute
	End With

	Do While Not oResident.eof 
		response.write vbcrlf & "<option value=""" & oResident("userid") & """"
		If CLng(iUserId) = CLng(oResident("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oResident("userlname") & ", " & oResident("userfname") & " &ndash; " & oResident("useraddress") & "</option>"
		oResident.movenext
	Loop 
		
	oResident.close
	Set oResident = Nothing
	Set oCmd = Nothing
End Sub  


%>
