<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitfixturetypeedit.asp
' AUTHOR: Steve Loar
' CREATED: 12/19/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   12/19/07	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, iPermitFixtureTypeid, sPermitFixture, sUnitQty, sUnitFeeAmount, sUseStepTable, iMaxRows

sLevel = "../" ' Override of value from common.asp
sUseStepTable = ""
iMaxRows = 1

PageDisplayCheck "permit fixture types", sLevel	' In common.asp

iPermitFixtureTypeid = CLng(request("permitfixturetypeid") )

If CLng(iPermitFixtureTypeid) > CLng(0) Then
	sTitle = "Edit"
	GetPermitFixture iPermitFixtureTypeid
Else
	sTitle = "New"
	sUnitQty = "0"
	sUnitFeeAmount = "0.00"
End If 

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="JavaScript" src="../scripts/formatnumber.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>
	<script language="JavaScript" src="../scripts/removecommas.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="Javascript">
	<!--

		function Another()
		{
			location.href="permitfixturetypeedit.asp?permitfixturetypeid=0";
		}

		function NewStepRow()
		{
			document.frmFixtures.maxrows.value = parseInt(document.frmFixtures.maxrows.value) + 1;
			var tbl = document.getElementById("feesteptable");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmFixtures.maxrows.value);
			var row = tbl.insertRow(lastRow);

			var cellOne = row.insertCell(0);
			cellOne.align = 'center';
			var inputOne = document.createElement('input');
			inputOne.type = 'text';
			inputOne.id = 'atleastqty' + newRow;
			inputOne.name = 'atleastqty' + newRow;
			if (newRow > 1)
			{
				inputOne.value = document.getElementById("notmorethanqty" + (newRow - 1)).value;
			}
			else
			{
				inputOne.value = '0';
			}
			inputOne.size = '9';
			inputOne.maxLength = '9';
			cellOne.appendChild(inputOne);

			var cellTwo = row.insertCell(1);
			cellTwo.align = 'center';
			var inputTwo = document.createElement('input');
			inputTwo.type = 'text';
			inputTwo.id = 'notmorethanqty' + newRow;
			inputTwo.name = 'notmorethanqty' + newRow;
			inputTwo.value = '999999999';
			inputTwo.size = '9';
			inputTwo.maxLength = '9';
			cellTwo.appendChild(inputTwo);
			
			var cellFive = row.insertCell(2);
			cellFive.align = 'center';
			var inputFive = document.createElement('input');
			inputFive.type = 'text';
			inputFive.id = 'unitamount' + newRow;
			inputFive.name = 'unitamount' + newRow;
			inputFive.value = '0.00';
			inputFive.size = '9';
			inputFive.maxLength = '9';
			cellFive.appendChild(inputFive);

			var cellFour = row.insertCell(3);
			cellFour.align = 'center';
			var inputFour = document.createElement('input');
			inputFour.type = 'text';
			inputFour.id = 'unitqty' + newRow;
			inputFour.name = 'unitqty' + newRow;
			inputFour.value = '1';
			inputFour.size = '9';
			inputFour.maxLength = '9';
			cellFour.appendChild(inputFour);

			var cellThree = row.insertCell(4);
			cellThree.align = 'center';
			var inputThree = document.createElement('input');
			inputThree.type = 'text';
			inputThree.id = 'baseamount' + newRow;
			inputThree.name = 'baseamount' + newRow;
			inputThree.value = '0.00';
			inputThree.size = '9';
			inputThree.maxLength = '9';
			cellThree.appendChild(inputThree);

		}

		function Validate()
		{
			var rege;
			var Ok; 

			// Check for a fixture name
			if (document.frmFixtures.permitfixture.value == '')
			{
				alert("Please provide a fixture type, then try saving again.");
				document.frmFixtures.permitfixture.focus();
				return;
			}
			
			var iStepRows = 0;
			// Check the step table values entered
			for (var t = 1; t <= parseInt(document.frmFixtures.maxrows.value); t++)
			{
				if (document.getElementById("atleastqty" + t).value != '')
				{
					iStepRows += 1;
					// Remove any extra spaces
					document.getElementById("atleastqty" + t).value = removeSpaces(document.getElementById("atleastqty" + t).value);
					//Remove commas that would cause problems in validation
					document.getElementById("atleastqty" + t).value = removeCommas(document.getElementById("atleastqty" + t).value);
		
					// Validate the at least quantity format
					rege = /^\d+$/;
					Ok = rege.test(document.getElementById("atleastqty" + t).value);
					if ( ! Ok )
					{
						alert("The 'At Least Quantity' should be blank or a whole number value.\nPlease correct this and try saving again.");
						document.getElementById("atleastqty" + t).focus();
						return;
					}

					// Validate the not more than quantity format
					if (document.getElementById("notmorethanqty" + t).value != '')
					{
						// Remove any extra spaces
						document.getElementById("notmorethanqty" + t).value = removeSpaces(document.getElementById("notmorethanqty" + t).value);
						//Remove commas that would cause problems in validation
						document.getElementById("notmorethanqty" + t).value = removeCommas(document.getElementById("notmorethanqty" + t).value);
		
						rege = /^\d+$/;
						Ok = rege.test(document.getElementById("notmorethanqty" + t).value);
						if ( ! Ok )
						{
							alert("The 'Not More Than Quantity' must be a whole number value.\nPlease correct this and try saving again.");
							document.getElementById("notmorethanqty" + t).focus();
							return;
						}
					}
					else
					{
						alert("The 'Not More Than Quantity' cannot be blank and must be a whole number value.\nPlease correct this and try saving again.");
						document.getElementById("notmorethanqty" + t).focus();
						return;
					}

					// Validate the Base Fee Amount format
					if (document.getElementById("baseamount" + t).value != '')
					{
						// Remove any extra spaces
						document.getElementById("baseamount" + t).value = removeSpaces(document.getElementById("baseamount" + t).value);
						//Remove commas that would cause problems in validation
						document.getElementById("baseamount" + t).value = removeCommas(document.getElementById("baseamount" + t).value);
		
						rege = /^\d*\.?\d{0,2}$/;
						Ok = rege.test(document.getElementById("baseamount" + t).value);
						if ( ! Ok )
						{
							alert("The 'Base Fee Amount' must be in currency format.\nPlease correct this and try saving again.");
							document.getElementById("baseamount" + t).focus();
							return;
						}
						else
						{
							document.getElementById("baseamount" + t).value = format_number(Number(document.getElementById("baseamount" + t).value),2);
						}
					}
					else
					{
						alert("The 'Base Fee Amount' cannot be blank and must be in currency format.\nPlease correct this and try saving again.");
						document.getElementById("baseamount" + t).focus();
						return;
					}

					// Validate the unit quantity format
					if (document.getElementById("unitqty" + t).value != '')
					{
						// Remove any extra spaces
						document.getElementById("unitqty" + t).value = removeSpaces(document.getElementById("unitqty" + t).value);
						//Remove commas that would cause problems in validation
						document.getElementById("unitqty" + t).value = removeCommas(document.getElementById("unitqty" + t).value);
		
						rege = /^\d+$/;
						Ok = rege.test(document.getElementById("unitqty" + t).value);
						if ( ! Ok )
						{
							alert("The 'Unit Quantity' must be a whole number value.\nPlease correct this and try saving again.");
							document.getElementById("unitqty" + t).focus();
							return;
						}
					}
					else
					{
						alert("The 'Unit Quantity' cannot be blank and must be a whole number value.\nPlease correct this and try saving again.");
						document.getElementById("unitqty" + t).focus();
						return;
					}

					// Validate the Unit Amount format
					if (document.getElementById("unitamount" + t).value != '')
					{
						// Remove any extra spaces
						document.getElementById("unitamount" + t).value = removeSpaces(document.getElementById("unitamount" + t).value);
						//Remove commas that would cause problems in validation
						document.getElementById("unitamount" + t).value = removeCommas(document.getElementById("unitamount" + t).value);
		
						rege = /^\d*\.?\d{0,2}$/;
						Ok = rege.test(document.getElementById("unitamount" + t).value);
						if ( ! Ok )
						{
							alert("The 'Unit Amount' must be in currency format.\nPlease correct this and try saving again.");
							document.getElementById("unitamount" + t).focus();
							return;
						}
						else
						{
							document.getElementById("unitamount" + t).value = format_number(Number(document.getElementById("unitamount" + t).value),2);
						}
					}
					else
					{
						alert("The 'Unit Amount' cannot be blank and must be in currency format.\nPlease correct this and try saving again.");
						document.getElementById("unitamount" + t).focus();
						return;
					}

				}
			}
			if (iStepRows == 0)
			{
				alert("You have not input any data into the fee table.\nPlease add some data to the table and try saving again.");
				return;
			}

			//alert("All was OK");
			// All is OK so submit
			document.frmFixtures.submit();
		}

		function Delete() 
		{
			if (confirm("Do you wish to delete this permit fixture type?"))
			{
				location.href="permitfixturetypedelete.asp?permitfixturetypeid=<%=iPermitFixtureTypeid%>";
			}
		}

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}
		function commonIFrameUpdateFunction()
		{
			UpdateParentFixtures()
		}
		function UpdateParentFixtures()
		{

			//Get New Values
			var request = new XMLHttpRequest();
			request.open('GET', 'popselectbox.asp?type=fixtures', false);  // `false` makes the request synchronous
			request.send();

			if (request.status === 200) {
  				newDDVals = request.responseText;

				//Update the unselected values
				var unSelVals = parent.document.getElementById('newpermitfixturetypeid');
				unSelVals.innerHTML = newDDVals;
				unSelVals.value = unSelVals.options[0].value;

				//Update the already selected values
				var selVals = parent.document.getElementById('permitfixturetypeid').options;
				for (var i = 0; i < selVals.length; i++) {
					var unSelOptions = unSelVals.options;
					for (var j = 0; j < unSelOptions.length; j++) {

						//Find Matching option in unselected values
						if (selVals[i].value == unSelOptions[j].value)
						{
							//Update Name in Selected Values
							selVals[i].innerHTML = unSelOptions[j].innerHTML;
	
							//Purge from unselected values
							unSelVals.removeChild(unSelOptions[j]);
						}
					}
				}

			}

		}

<%		If request("success") <> "" Then 
			DisplayMessagePopUp 
		End If 
%>

	//-->
	</script>

</head>

<body>

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong><%=sTitle%> Permit Fixture Type</strong></font><br /><br />
				<a href="permitfixturetypelist.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
			</p>
			<!--END: PAGE TITLE-->

		<div id="functionlinks">
<%		If CLng(iPermitFixtureTypeid) = CLng(0) Then %>
			<input type="button" id="savebutton" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Create" /><br />
<%		Else %>
			<input type="button" id="savebutton" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Save Changes" /> &nbsp; &nbsp;
			<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" />&nbsp; &nbsp; 
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" /> &nbsp; &nbsp;
<%			If request("success") <> "" Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Another();" value="Create Another" />
<%			End If		%>
			<br />
<%		End If %>
		</div>

		<form name="frmFixtures" action="permitfixturetypeupdate.asp" method="post">
		<input type="hidden" name="permitfixturetypeid" value="<%=iPermitFixtureTypeid%>" />
		
		<p>
			Fixture Type: &nbsp;&nbsp; <input type="text" name="permitfixture" value="<%=sPermitFixture%>" size="100" maxlength="150" />
		</p>
		<p>
			<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add Row" id="addref" onClick="NewStepRow()" />
		</p>
		<div class="shadow">
			<table cellpadding="5" cellspacing="0" border="0" class="tableadmin" id="feesteptable">
				<tr><th>Qty Is<br />At Least</th><th>And Is<br />Less Than</th><th>Price Per<br />Unit</th><th>Unit Is<br />This Qty</th><th>Then Add<br />This Amount</th></tr>
<%				
				iMaxRows = ShowFixtureStepTable( iPermitFixtureTypeid ) 
%>
			</table>
		</div>
		* To remove a row, blank out the &quot;Qty Is At Least&quot; value. The row will be removed when your changes are saved.
		<input type="hidden" name="maxrows" id="maxrows" value="<%=iMaxRows%>" />

		</form>
		<!--END: EDIT FORM-->
		<div id="functionlinks">
<%		If CLng(iPermitFixtureTypeid) = CLng(0) Then %>
			<input type="button" id="savebutton" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Create" /><br />
<%		Else %>
			<input type="button" id="savebutton" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Save Changes" /> &nbsp; &nbsp;
			<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" />&nbsp; &nbsp; 
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" /> &nbsp; &nbsp;
<%			If request("success") <> "" Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Another();" value="Create Another" />
<%			End If		%>
			<br />
<%		End If %>
		</div>

		</div>
		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

<%	If request("success") <> "" Then 
		SetupMessagePopUp request("success")
	End If	
%>

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub GetPermitFixture( iPermitFixtureTypeid )
'--------------------------------------------------------------------------------------------------
Sub GetPermitFixture( iPermitFixtureTypeid )
	Dim sSql, oRs

	'sSql = "SELECT permitfixture, unitqty, unitfeeamount, usesteptable FROM egov_permitfixturetypes WHERE permitfixturetypeid = " & iPermitFixtureTypeid
	sSql = "SELECT permitfixture FROM egov_permitfixturetypes WHERE permitfixturetypeid = " & iPermitFixtureTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sPermitFixture = Replace(oRs("permitfixture"),"""","&quot;")
'		sUnitQty = oRs("unitqty")
'		sUnitFeeAmount = FormatNumber(oRs("unitfeeamount"),2,,,0)
'		If oRs("usesteptable") Then 
'			sUseStepTable = " checked=""checked"" "
'		Else
'			sUseStepTable = ""
'		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function ShowFixtureStepTable( iPermitFixtureTypeid )
'--------------------------------------------------------------------------------------------------
Function ShowFixtureStepTable( iPermitFixtureTypeid )
	Dim sSql, oRs, iRowCount 

	sSql = "SELECT atleastqty, notmorethanqty, baseamount, unitqty, unitamount FROM egov_permitfixturetypestepfees WHERE permitfixturetypeid = " & iPermitFixtureTypeid & " ORDER BY atleastqty"
	iRowCount = CLng(0)
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF 
			iRowCount = iRowCount + CLng(1)
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write ">"
			response.write "<td align=""center""><input type=""text"" id=""atleastqty" & iRowCount &""" name=""atleastqty" & iRowCount &""" value=""" & oRs("atleastqty") & """ size=""9"" maxlength=""9"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""notmorethanqty" & iRowCount &""" name=""notmorethanqty" & iRowCount &""" value=""" & oRs("notmorethanqty") & """ size=""9"" maxlength=""9"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""unitamount" & iRowCount &""" name=""unitamount" & iRowCount &""" value=""" & FormatNumber(oRs("unitamount"),2,,,0) & """ size=""9"" maxlength=""9"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""unitqty" & iRowCount &""" name=""unitqty" & iRowCount &""" value=""" & oRs("unitqty") & """ size=""9"" maxlength=""9"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""baseamount" & iRowCount &""" name=""baseamount" & iRowCount &""" value=""" & FormatNumber(oRs("baseamount"),2,,,0) & """ size=""9"" maxlength=""9"" /></td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop 
	Else
		iRowCount = iRowCount + CLng(1)
		response.write "<tr>"
		response.write "<td align=""center""><input type=""text"" id=""atleastqty1"" name=""atleastqty1"" value=""0"" size=""9"" maxlength=""9"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""notmorethanqty1"" name=""notmorethanqty1"" value=""999999999"" size=""9"" maxlength=""9"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""unitamount1"" name=""unitamount1"" value=""0.00"" size=""9"" maxlength=""9"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""unitqty1"" name=""unitqty1"" value=""1"" size=""9"" maxlength=""9"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""baseamount1"" name=""baseamount1"" value=""0.00"" size=""9"" maxlength=""9"" /></td>"
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowFixtureStepTable = iRowCount
End Function 
%>
