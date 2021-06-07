<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: contactuserpicker.asp
' AUTHOR: Steve Loar
' CREATED: 03/12/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Selects permit contacts
'
' MODIFICATION HISTORY
' 1.0   03/12/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitContactTypeid

iPermitContactTypeid = CLng(request("permitcontacttypeid"))

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

		<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

		<script language="Javascript">
		<!--

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function doCloseAfterAdd( sReturn )
		{
			doClose();
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
					$('searchresults').innerHTML = 'Possible Match Found.';
					document.frmContact.searchstart.value = x;
					return;
				}
			}
			document.frmContact.results.value = 'No Match Found.';
			$('searchresults').innerHTML = 'No Match Found.';
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
			$('searchresults').innerHTML = '';
			document.frmContact.searchstart.value = -1;
		}

		function doSelect()
		{
			if (document.frmContact.userid.options[document.frmContact.userid.selectedIndex].value > 0)
			{
					parent.document.getElementById("maxusers").value = parseInt(parent.document.getElementById("maxusers").value) + 1;
					var tbl = parent.document.getElementById("contractoruserlist");
					var lastRow = tbl.rows.length;
					var newRow = parseInt(parent.document.getElementById("maxusers").value);
					var row = tbl.insertRow(lastRow);
					// if it is an odd number put a class on it??
					if (newRow % 2 == 0)
					{
						row.className = 'altrow';
					}

					// Add the Remove Row checkbox
					var cell = row.insertCell(0);
					cell.align = 'center';
					var e = parent.document.createElement('input');
					e.type = 'checkbox';
					e.name = 'removeuser' + newRow;
					e.id = 'removeuser' + newRow;
					e.value = document.frmContact.userid.options[document.frmContact.userid.selectedIndex].value;
					cell.appendChild(e);

					// hidden userid
					e = parent.document.createElement('input');
					e.type = 'hidden';
					e.name = 'userid' + newRow;
					e.value = document.frmContact.userid.options[document.frmContact.userid.selectedIndex].value;
					cell.appendChild(e);

					// Add the User name 
					cell = row.insertCell(1);
					var cellText = parent.document.createTextNode(document.frmContact.userid.options[document.frmContact.userid.selectedIndex].text);
					cell.appendChild(cellText);

					// Add can add others
					cell = row.insertCell(2);
					cell.align = 'center';
					e = parent.document.createElement('input');
					e.type = 'checkbox';
					e.name = 'canaddothers' + newRow;
					e.id = 'canaddothers' + newRow;
					e.value = document.frmContact.userid.options[document.frmContact.userid.selectedIndex].value;
					cell.appendChild(e);

					//Add the Primary Contact Radio button
					cell = row.insertCell(3);
					cell.align = 'center';
					//e = parent.document.createElement('input');
					//e.type = 'radio';
					//e.name = 'isprimarycontact';
					//e.value = document.frmContact.userid.options[document.frmContact.userid.selectedIndex].value;
					cell.innerHTML = '&nbsp;';

					// Fire off an Ajax call to add the user to the contractor
					doAjax('addpermitcontactuser.asp', 'permitcontacttypeid=<%=iPermitContactTypeid%>&userid=' + document.frmContact.userid.options[document.frmContact.userid.selectedIndex].value, 'doCloseAfterAdd', 'get', '0');
			}
			else
			{
				doClose();
			}
		}

		//-->
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent">
				<form name="frmContact" action="contactuserpicker.asp" method="post">
					<p>
<%						ShowUserDropDown iPermitContactTypeid		%>						
						<br /><br /><!--  onchange="ClearSearch();"  -->
						<input type="text" name="searchname" size="25" maxlength="50" onkeypress="if(event.keyCode=='13'){SearchCitizens(document.frmContact.searchstart.value);return false;}" />
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Search" onclick="SearchCitizens(document.frmContact.searchstart.value);" /> &nbsp;
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

'------------------------------------------------------------------------------
' Sub ShowUserDropDown( iPermitContactTypeid )
'------------------------------------------------------------------------------
Sub ShowUserDropDown( iPermitContactTypeid )
	Dim sSql, oRs

	sSql = "SELECT userid, userfname, userlname, ISNULL(userhomephone,'') AS userphone FROM egov_users "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND userregistered = 1 "
	sSql = sSql & " AND headofhousehold = 1 AND userfname IS NOT NULL AND userlname IS NOT NULL "
	sSql = sSql & " AND userfname != '' and userlname != '' AND userid NOT IN ( SELECT userid "
	sSql = sSql & " FROM egov_permitcontacttypes_to_users WHERE permitcontacttypeid = " & iPermitContactTypeid
	sSql = sSql & " ) ORDER BY userlname, userfname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""userid"" id=""userid"" onchange=""UserPick();"">"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("userid") & """>" & oRs("userlname") & ", " & oRs("userfname")
			If oRs("userphone") <> "" Then 
				response.write " &mdash; " & FormatPhoneNumber( oRs("userphone") )
			End If 
			response.write "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 
		
	oRs.Close
	Set oRs = Nothing

End Sub 

%>
