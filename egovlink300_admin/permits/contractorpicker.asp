<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: contractorpicker.asp
' AUTHOR: Steve Loar
' CREATED: 08/14/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Selects permit contractors
'
' MODIFICATION HISTORY
' 1.0   08/14/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sFieldId

sFieldId = request("fieldid")

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="Javascript">
		<!--

		function doSelect()
		{
			if (document.frmContact.permitcontacttypeid.options[document.frmContact.permitcontacttypeid.selectedIndex].value > 0)
			{
				parent.document.getElementById("<%=sFieldId%>").value = document.frmContact.permitcontacttypeid.options[document.frmContact.permitcontacttypeid.selectedIndex].text;
			}
			doClose();
		}

		function doClose()
		{
			//window.close();
			//window.opener.focus();
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

		//-->
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent">
				<form name="frmContact" action="contactpicker.asp" method="post">
					<p>
					<% ShowContactPicks %>
						<br /><br /><!--  onchange="ClearSearch();"  -->
						<input type="text" name="searchname" size="25" maxlength="50" onkeypress="if(event.keyCode=='13'){SearchCitizens(document.frmContact.searchstart.value);return false;}" />
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
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function GetCurrentContact( iPermitId, sContactType )
'--------------------------------------------------------------------------------------------------
Function GetCurrentContact( ByVal iPermitId, ByVal sContactType )
	Dim sSql, oRs

	sSql = "SELECT permitcontacttypeid FROM egov_permitcontacts "
	sSql = sSQl & " WHERE " & dbsafe(sContactType) & " = 1 AND permitid = " & iPermitId

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
' Sub ShowContactPicks( )
'--------------------------------------------------------------------------------------------------
Sub ShowContactPicks( )
	Dim sSql, oRs, bName, sContractorName

	sContractorName = ""

	sSql = "SELECT permitcontacttypeid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSQl & " ISNULL(company,'') AS company, ISNULL(lastname,'') + ISNULL(firstname,'') + ISNULL(company,'') AS sortname "
	sSql = sSQl & " FROM egov_permitcontacttypes WHERE orgid = " & session("orgid") & " ORDER BY 5"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		bCanSelect = True 
		response.write vbcrlf & "<select name=""permitcontacttypeid"" onchange=""UserPick();"">"
		response.write vbcrlf & "<option value=""0"">Select a contractor from the list...</option>"
		Do While Not oRs.EOF
			sContractorName = ""
			If oRs("firstname") <> "" Then
				sContractorName = oRs("firstname") & " " & oRs("lastname")
				bName = True 
			Else
				bName = False 
			End If 
			If oRs("company") <> "" Then
				If bName Then 
					sContractorName = sContractorName &  " ("
				End If 
				sContractorName = sContractorName & oRs("company")
				If bName Then
					sContractorName = sContractorName & ")"
				End If 
			End If 
			response.write vbcrlf & "<option value=""" & oRs("permitcontacttypeid") & """>"
			response.write sContractorName
			response.write "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	Else
		response.write "No contractors have been created."
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 



%>
