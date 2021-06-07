<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: listedownerpicker.asp
' AUTHOR: Steve Loar
' CREATED: 12/10/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Selects existing listed owners for new addresses
'
' MODIFICATION HISTORY
' 1.0   12/10/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="global.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits/permits.css" />

		<script language="Javascript">
		<!--

			function doClose()
			{
				if (window.opener)
				{
					window.close();
					window.opener.focus();
				}
				else
				{
					parent.hideModal(window.frameElement.getAttribute("data-close"));
				}
			}

			function SearchOwners( iSearchStart )
			{
				var optiontext;
				var optionchanged;
				var searchtext = document.frmListedOwner.searchname.value;
				var searchchanged = searchtext.toLowerCase();

				iSearchStart = parseInt(iSearchStart) + 1;
				
				for (x=iSearchStart; x < document.frmListedOwner.listedowner.length ; x++)
				{
					optiontext = document.frmListedOwner.listedowner.options[x].text;
					optionchanged = optiontext.toLowerCase();
					if (optionchanged.indexOf(searchchanged) != -1)
					{
						document.frmListedOwner.listedowner.selectedIndex = x;
						document.frmListedOwner.results.value = 'Possible Match Found.';
						document.getElementById('searchresults').innerHTML = 'Possible Match Found.';
						document.frmListedOwner.searchstart.value = x;
						return;
					}
				}
				document.frmListedOwner.results.value = 'No Match Found.';
				document.getElementById('searchresults').innerHTML = 'No Match Found.';
				document.frmListedOwner.searchstart.value = -1;
			}

			function doSelect()
			{
				if (document.frmListedOwner.listedowner.options[document.frmListedOwner.listedowner.selectedIndex].value > 0)
				{
					if (window.opener)
					{
						window.opener.document.getElementById("owner").value = document.frmListedOwner.listedowner.options[document.frmListedOwner.listedowner.selectedIndex].text;
					}
					else
					{
						parent.document.getElementById("owner").value = document.frmListedOwner.listedowner.options[document.frmListedOwner.listedowner.selectedIndex].text;
					}
				}
				doClose();
			}

			function UserPick()
			{
				document.frmListedOwner.searchname.value = '';
				document.frmListedOwner.results.value = '';
				document.getElementById('searchresults').innerHTML = '';
				document.frmListedOwner.searchstart.value = -1;
			}

		//-->
		if (window.top!=window.self)
		{
			document.getElementById('title').style.display = 'none';
		}
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent">
				<font id="title" size="+1"><strong>Listed Owner Selection</strong></font><br /><br />
		<script>
		if (window.top!=window.self)
		{
			document.getElementById('title').style.display = 'none';
		}
		</script>
				<form name="frmListedOwner" action="listedownerpicker.asp" method="post">
					<p>
					<% ShowListedOwnerPicks %>
					<br /><br /><!--  onchange="ClearSearch();"  -->
					<input type="text" name="searchname" size="25" maxlength="50" onkeypress="if(event.keyCode=='13'){SearchOwners(document.frmListedOwner.searchstart.value);return false;}" />
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="Search" onclick="SearchOwners(document.frmListedOwner.searchstart.value);" /> &nbsp;
					<input type="hidden" name="searchstart" value="-1" />
					<input type="hidden" name="results" value="" />
					<span id="searchresults"></span>
					</p>
					<p>
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Select" onclick="doSelect();" /> &nbsp; &nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Cancel" onclick="doClose();" /> &nbsp; &nbsp; 
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

'--------------------------------------------------------------------------------------------------
' Sub ShowListedOwnerPicks( )
'--------------------------------------------------------------------------------------------------
Sub ShowListedOwnerPicks( )
	Dim sSql, oRs, iOwnerCount

	iOwnerCount = CLng(0)

	sSql = "SELECT DISTINCT listedowner FROM egov_residentaddresses "
	sSql = sSQl & " WHERE orgid = " & session("orgid") & " AND listedowner IS NOT NULL ORDER BY listedowner"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""listedowner"" onchange=""UserPick();"">"
		response.write vbcrlf & "<option value=""0"">Select an existing listed owner from the list...</option>"
		Do While Not oRs.EOF
			If oRs("listedowner") <> "" Then 
				iOwnerCount = iOwnerCount + CLng(1)
				response.write vbcrlf & "<option value=""" & iOwnerCount & """>"
				response.write oRs("listedowner")
				response.write "</option>"
			End If 
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	Else
		response.write "No listed owners are in the system."
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 

%>
