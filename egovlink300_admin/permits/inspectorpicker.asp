<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: inspectorpicker.asp
' AUTHOR: Steve Loar
' CREATED: 08/14/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Selects permit inspectors
'
' MODIFICATION HISTORY
' 1.0   08/14/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitInspectionId, iInspectorUserId, iRowId

iPermitInspectionId = CLng(request("permitinspectionid"))
iInspectorUserId = CLng(request("inspectoruserid"))
iRowId = CLng(request("rowid"))

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

		<script language="Javascript">
		<!--

		function doSelect()
		{
			if (document.frmContact.inspectoruserid.options[document.frmContact.inspectoruserid.selectedIndex].value > 0)
			{
				parent.document.getElementById("Inspector<%=iRowId%>").innerHTML = document.frmContact.inspectoruserid.options[document.frmContact.inspectoruserid.selectedIndex].text;
				doAjax('changeinspector.asp', 'permitinspectionid=<%=iPermitInspectionId%>&inspectoruserid=' + document.frmContact.inspectoruserid.options[document.frmContact.inspectoruserid.selectedIndex].value, '', 'get', '0');
			}
			doClose();
		}

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function SearchCitizens( iSearchStart )
		{
			var optiontext;
			var optionchanged;
			//alert(document.BuyerForm.searchname.value);
			var searchtext = document.frmContact.searchname.value;
			var searchchanged = searchtext.toLowerCase();

			iSearchStart = parseInt(iSearchStart) + 1;
			
			for (x=iSearchStart; x < document.frmContact.inspectoruserid.length ; x++)
			{
				optiontext = document.frmContact.inspectoruserid.options[x].text;
				optionchanged = optiontext.toLowerCase();
				if (optionchanged.indexOf(searchchanged) != -1)
				{
					document.frmContact.inspectoruserid.selectedIndex = x;
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
				<form name="frmContact" action="inspectorpicker.asp" method="post">
					<p>
					<% ShowPermitInspectors iInspectorUserId  %>
						<br /><br />
						<input type="text" name="searchname" size="25" maxlength="50" onchange="ClearSearch();" />
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Search" onclick="SearchCitizens(document.frmContact.searchstart.value);" /> &nbsp;
						<input type="hidden" name="searchstart" value="-1" />
						<input type="hidden" name="results" value="" />
						<span id="searchresults"></span>
						</p>
						<p>
							<input type="button" class="button ui-button ui-widget ui-corner-all" value="Select" onclick="doSelect();" /> &nbsp; &nbsp; 
							<input type="button" class="button ui-button ui-widget ui-corner-all" value="Cancel" onclick="doClose();" />
						</p>
					</p>
				</form>
			</div>
		</div>
	</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ShowPermitInspectors iInspectorUserId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitInspectors( ByVal iInspectorUserId )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname FROM users WHERE orgid = " & session("orgid") & " AND ispermitinspector = 1 "
	sSql = sSQl & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""inspectoruserid"" onchange=""UserPick();"">"
		response.write vbcrlf & "<option value=""0"">Select an inspector from the list...</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option "
			If CLng(iInspectorUserId) = CLng(oRs("userid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " value=""" & oRs("userid") & """>" & oRs("firstname") & " " & oRs("lastname") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
