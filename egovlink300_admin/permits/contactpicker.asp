<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: contactpicker.asp
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
Dim iPermitId, sType, sTitle, iContactTypeId, sContactType, bCanSelect

iPermitId = CLng(request("permitid"))
sType = request("stype")
sContactType = sType
bCanSelect = False 

Select Case sType
	Case "isprimarycontractor"
		sTitle = "Primary Contractor"
		sContactType = "iscontractor"
	Case "isarchitect"
		sTitle = "Architect/Engineer"
	Case Else 
		sTitle = "Other Contractor"
End Select 

If sType = "iscontractor" Then
	iContactTypeId = 0
Else 
	If sType <> "" Then 
		iContactTypeId = GetCurrentContact( iPermitId, sType )
	Else
		iContactTypeId = 0
	End If 
End If 

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

  		<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
		<script language="Javascript">
		<!--
		var lastSearch = "";

		function GetSelected(ID)
		{
			$.get( "contractorlist.asp?icontacttypeid=" + ID, function( data ) {
				$("#permitcontacttypeid").html(data);
			});
		}
		function LiveSearchCitizens( ID )
		{
			if (lastSearch != ID)
			{
				$.get( "contractorlist.asp?query=" + ID, function( data ) {
					$("#permitcontacttypeid").html(data);
				});
				lastSearch = ID;
			}
			else
			{
				$('#permitcontacttypeid option:selected').next().attr('selected', 'selected');
			}
		}

		function doSelect()
		{
			//alert(document.frmContact.permitcontacttypeid.options[document.frmContact.permitcontacttypeid.selectedIndex].text);
			//alert($("#permitcontacttypeid").val());
<%			If sType = "iscontractor" Then %>
				if ($("#permitcontacttypeid").val() > 0)
				{
					parent.document.getElementById("maxcontractors").value = parseInt(parent.document.getElementById("maxcontractors").value) + 1;
					var tbl = parent.document.getElementById("contractorlist");
					var lastRow = tbl.rows.length;
					var newRow = parseInt(parent.document.getElementById("maxcontractors").value);
					var row = tbl.insertRow(lastRow);

					// Add the Remove Row checkbox
					var cellZero = row.insertCell(0);
					cellZero.align = 'left';
					//cellZero.className = 'contactpick';
					var e = parent.document.createElement('input');
					e.type = 'checkbox';
					e.name = 'removepermitcontactid' + newRow;
					e.id = 'removepermitcontactid' + newRow;
					cellZero.appendChild(e);

					//var space = document.createTextNode( '\u00A0' );
					//cellZero.appendChild(space);
					//cellZero.appendChild(document.createTextNode( ' ' ));
					//cellZero.appendChild(space);

					// Add the Contact name text
					//var cellOne = row.insertCell(1);
					var e1 = parent.document.createElement('input');
					e1.type = 'hidden';
					e1.name = 'contractor' + newRow;
					e1.id = 'contractor' + newRow;
					e1.value = $("#permitcontacttypeid").val();
					cellZero.appendChild(e1);

					var e2 = parent.document.createElement('input');
					e2.type = 'hidden';
					e2.name = 'permitcontactid' + newRow;
					e2.id = 'permitcontactid' + newRow;
					e2.value = 0;
					cellZero.appendChild(e2);

					var cellText = parent.document.createTextNode(document.frmContact.permitcontacttypeid.options[document.frmContact.permitcontacttypeid.selectedIndex].text);
					cellZero.appendChild(cellText);
					parent.document.getElementById("addcontractor").focus();

					//parent.NiceTitles.autoCreated.anchors.addElements(parent.document.getElementsByTagName("a"), "title");
				}
<%			Else  %>
				if ($("#permitcontacttypeid").val() > 0)
				{
					parent.document.getElementById("<%=sType%>display").innerHTML = $( "#permitcontacttypeid option:selected" ).text()
				}
				else
				{
					parent.document.getElementById("<%=sType%>display").innerHTML = 'None Selected ';
				}
				parent.document.getElementById("<%=sType%>permitcontacttypeid").value = $("#permitcontacttypeid").val();
<%			End If	%>
			doClose();
		}

		function doClose()
		{
			//window.close();
			//parent.focus();
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
			
			for (x=iSearchStart; x < document.frmContact.permitcontacttypeid.length ; x++)
			{
				optiontext = document.frmContact.permitcontacttypeid.options[x].text;
				optionchanged = optiontext.toLowerCase();
				if (optionchanged.indexOf(searchchanged) != -1)
				{
					document.frmContact.permitcontacttypeid.selectedIndex = x;
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

		function doNewContact()
		{
			location.href="permitcontactedit.asp?permitcontactid=0&permitid=<%=iPermitId%>&type=<%=sType%>&updatetitle=1";
		}

		//-->
		</script>

	</head>
	<body onload="document.getElementById('searchname').focus()">
		<div id="content">
			<div id="centercontent">
				<script>parent.document.getElementById('modaltitle'+window.frameElement.getAttribute("data-close")).innerHTML = '<%=sTitle%>';</script>
				<form name="frmContact" action="contactpicker.asp" method="post">
					<p>
					<%
						'if session("orgid") = "139" then
						if 1=1 then
							response.write vbcrlf & "<select id=""permitcontacttypeid"" name=""permitcontacttypeid"" onchange=""UserPick();"">"
							response.write "</select>"
							'Call to get selected contractor
							%>
							<br />
							<input id="searchname" type="text" name="searchname" size="25" maxlength="50" onkeypress="if(event.keyCode=='13'){LiveSearchCitizens(document.frmContact.searchname.value);return false;}" />
							<input type="button" class="button ui-button ui-widget ui-corner-all" value="Search" onclick="LiveSearchCitizens(document.frmContact.searchname.value);" /> &nbsp;
							<script> GetSelected(<%=iContactTypeId%>);  </script>
							<%
						else
							ShowContactPicks iContactTypeId, sContactType, bCanSelect
							tooltipclass=""
							tooltip = ""
							disabled = ""
							If not bCanSelect Then
								tooltipclass="tooltip"
								disabled = " disabled "
								tooltip = "<span class=""tooltiptext"">No contacts are available.</span>"
							end if
							%>
							<br /><br /><!--  onchange="ClearSearch();"  -->
							<input id="searchname" type="text" name="searchname" size="25" maxlength="50" onkeypress="if(event.keyCode=='13'){SearchCitizens(document.frmContact.searchstart.value);return false;}" />
							<input type="button" class="button ui-button ui-widget ui-corner-all" value="Search" onclick="SearchCitizens(document.frmContact.searchstart.value);" /> &nbsp;
							<input type="hidden" name="searchstart" value="-1" />
							<input type="hidden" name="results" value="" />
							<span id="searchresults"></span>
							<%
						end if
						%>
					</p>
					<p>
					<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="doSelect();">Select<%=tooltip%></button> &nbsp; &nbsp; 
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="Cancel" onclick="doClose();" /> &nbsp; &nbsp; 
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="New Contractor" onclick="doNewContact();" />
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
' integer GetCurrentContact( iPermitId, sContactType )
'--------------------------------------------------------------------------------------------------
Function GetCurrentContact( ByVal iPermitId, ByVal sContactType )
	Dim sSql, oRs

	sSql = "SELECT permitcontacttypeid FROM egov_permitcontacts "
	sSql = sSQl & " WHERE " & dbsafe(sContactType) & " = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId

	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetCurrentContact = oRs("permitcontacttypeid")
	Else 
		GetCurrentContact = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' void ShowContactPicks iContactTypeId, sContactType, bCanSelect
'--------------------------------------------------------------------------------------------------
Sub ShowContactPicks( ByVal iContactTypeId, ByVal sContactType, ByRef bCanSelect )
	Dim sSql, oRs, bName, sContractorType

	sSql = "SELECT permitcontacttypeid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSQl & "ISNULL(company,'') AS company, ISNULL(company,'') + ISNULL(lastname,'') + ISNULL(firstname,'') AS sortname, "
	sSql = sSql & "ISNULL(contractortypeid,0) AS contractortypeid "
	sSql = sSQl & "FROM egov_permitcontacttypes WHERE isorganization = 0 AND orgid = " & session("orgid") & " ORDER BY 5"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		bCanSelect = True 
		response.write vbcrlf & "<select id=""permitcontacttypeid"" name=""permitcontacttypeid"" onchange=""UserPick();"">"
		response.write vbcrlf & "<option value=""0"">Select a contractor from the list...</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("permitcontacttypeid") & """"
			If CLng(iContactTypeId) = CLng(oRs("permitcontacttypeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">"
			If oRs("company") <> "" Then
				response.write oRs("company")
				If oRs("firstname") <> "" Then 
					response.write " &mdash; "
				End If 
			End If 
			If oRs("firstname") <> "" Then 
				response.write oRs("lastname") & ", " & oRs("firstname")
			End If 

			' Contractor type
			sContractorType = GetContractorType( oRs("contractortypeid") ) 
			If sContractorType <> "" Then 
				response.write " &ndash; (<strong>" & sContractorType & "</strong>)"
			End If 


'			If oRs("firstname") <> "" Then
'				response.write oRs("firstname") & " " & oRs("lastname")
'				bName = True 
'			Else
'				bName = False 
'			End If 
'			If oRs("company") <> "" Then
'				If bName Then 
'					response.write " ("
'				End If 
'				response.write oRs("company")
'				If bName Then
'					response.write ")"
'				End If 
'			End If 
			response.write "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	Else
		response.write "No contacts have been created."
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
