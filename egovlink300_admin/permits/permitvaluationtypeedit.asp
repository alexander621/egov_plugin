<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitvaluationtypeedit.asp
' AUTHOR: Steve Loar
' CREATED: 04/14/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/14/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, iPermitValuationTypeid, sPermitValuation, sUnitQty, sUnitFeeAmount, sUseStepTable, iMaxRows

sLevel = "../" ' Override of value from common.asp
sUseStepTable = ""
iMaxRows = 1

PageDisplayCheck "permit valuation types", sLevel	' In common.asp

iPermitValuationTypeid = CLng(request("permitvaluationtypeid") )

If CLng(iPermitValuationTypeid) > CLng(0) Then
	sTitle = "Edit"
	GetPermitValuation iPermitValuationTypeid
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
			location.href="permitvaluationtypeedit.asp?permitvaluationtypeid=0";
		}

		function NewStepRow()
		{
			document.frmValuation.maxrows.value = parseInt(document.frmValuation.maxrows.value) + 1;
			var tbl = document.getElementById("feesteptable");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmValuation.maxrows.value);
			var row = tbl.insertRow(lastRow);

			var cellOne = row.insertCell(0);
			cellOne.align = 'center';
			var inputOne = document.createElement('input');
			inputOne.type = 'text';
			inputOne.id = 'atleastvalue' + newRow;
			inputOne.name = 'atleastvalue' + newRow;
			if (newRow > 1)
			{
				inputOne.value = document.getElementById("notmorethanvalue" + (newRow - 1)).value;
			}
			else
			{
				inputOne.value = '0';
			}
			inputOne.size = '12';
			inputOne.maxLength = '12';
			cellOne.appendChild(inputOne);

			var cellTwo = row.insertCell(1);
			cellTwo.align = 'center';
			var inputTwo = document.createElement('input');
			inputTwo.type = 'text';
			inputTwo.id = 'notmorethanvalue' + newRow;
			inputTwo.name = 'notmorethanvalue' + newRow;
			inputTwo.value = '999999999.99';
			inputTwo.size = '12';
			inputTwo.maxLength = '12';
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

			// Check for a valuation name
			if (document.frmValuation.permitvaluation.value == '')
			{
				alert("Please provide a valuation type, then try saving again.");
				document.frmValuation.permitvaluation.focus();
				return;
			}
			
			var iStepRows = 0;
			// Check the step table values entered
			for (var t = 1; t <= parseInt(document.frmValuation.maxrows.value); t++)
			{
				if (document.getElementById("atleastvalue" + t).value != '')
				{
					iStepRows += 1;
					// Remove any extra spaces
					document.getElementById("atleastvalue" + t).value = removeSpaces(document.getElementById("atleastvalue" + t).value);
					//Remove commas that would cause problems in validation
					document.getElementById("atleastvalue" + t).value = removeCommas(document.getElementById("atleastvalue" + t).value);
		
					// Validate the at least value format
					//rege = /^\d+$/;
					rege = /^\d*\.?\d{0,2}$/;
					Ok = rege.test(document.getElementById("atleastvalue" + t).value);
					if ( ! Ok )
					{
						alert("The 'At Least Value' should be blank or in currency format.\nPlease correct this and try saving again.");
						document.getElementById("atleastvalue" + t).focus();
						return;
					}
					else
					{
						document.getElementById("atleastvalue" + t).value = format_number(Number(document.getElementById("atleastvalue" + t).value),2);
					}

					// Validate the not more than quantity format
					if (document.getElementById("notmorethanvalue" + t).value != '')
					{
						// Remove any extra spaces
						document.getElementById("notmorethanvalue" + t).value = removeSpaces(document.getElementById("notmorethanvalue" + t).value);
						//Remove commas that would cause problems in validation
						document.getElementById("notmorethanvalue" + t).value = removeCommas(document.getElementById("notmorethanvalue" + t).value);
		
						//rege = /^\d+$/;
						rege = /^\d*\.?\d{0,2}$/;
						Ok = rege.test(document.getElementById("notmorethanvalue" + t).value);
						if ( ! Ok )
						{
							alert("The 'Not More Than Value' must be in currency format.\nPlease correct this and try saving again.");
							document.getElementById("notmorethanvalue" + t).focus();
							return;
						}
						else
						{
							document.getElementById("notmorethanvalue" + t).value = format_number(Number(document.getElementById("notmorethanvalue" + t).value),2);
						}
					}
					else
					{
						alert("The 'Not More Than Quantity' cannot be blank and must be in currency format.\nPlease correct this and try saving again.");
						document.getElementById("notmorethanvalue" + t).focus();
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
			document.frmValuation.submit();
		}

		function Delete() 
		{
			if (confirm("Do you wish to delete this permit valuation type?"))
			{
				location.href="permitvaluationtypedelete.asp?permitvaluationtypeid=<%=iPermitValuationTypeid%>";
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

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong><%=sTitle%> Permit Valuation Type</strong></font><br /><br />
				<a href="permitvaluationtypelist.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
			</p>
			<!--END: PAGE TITLE-->

<%		If CLng(iPermitValuationTypeid) = CLng(0) Then %>
			<input type="button" id="savebutton" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Create" /><br />
<%		Else %>
			<input type="button" id="savebutton" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Save Changes" /> &nbsp; &nbsp;
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" /> &nbsp; &nbsp;
<%			If request("success") <> "" Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Another();" value="Create Another" />
<%			End If		%>
			<br />
<%		End If %>

		<form name="frmValuation" action="permitvaluationtypeupdate.asp" method="post">
		<input type="hidden" name="permitvaluationtypeid" value="<%=iPermitValuationTypeid%>" />
		
		<p>
			Valuation Type: &nbsp;&nbsp; <input type="text" name="permitvaluation" value="<%=sPermitValuation%>" size="100" maxlength="150" />
		</p>
		<p>
			<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add Row" id="addref" onClick="NewStepRow()" />
		</p>
		<div class="shadow">
			<table cellpadding="5" cellspacing="0" border="0" class="tableadmin" id="feesteptable">
				<tr><th>Value Is<br />At Least</th><th>And Is<br />Less Than</th><th>Price Per<br />Unit</th><th>Unit Is<br />This Qty</th><th>Then Add<br />This Amount</th></tr>
<%				
				iMaxRows = ShowValuationStepTable( iPermitValuationTypeid ) 
%>
			</table>
		</div>
		* To remove a row, blank out the &quot;Value Is At Least&quot; field for that row. The row will be removed when your changes are saved.
		<input type="hidden" name="maxrows" id="maxrows" value="<%=iMaxRows%>" />

		</form>
		<!--END: EDIT FORM-->

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
' Sub GetPermitValuation( iPermitValuationTypeid )
'--------------------------------------------------------------------------------------------------
Sub GetPermitValuation( iPermitValuationTypeid )
	Dim sSql, oRs

	sSql = "SELECT permitvaluation FROM egov_permitvaluationtypes WHERE permitvaluationtypeid = " & iPermitValuationTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sPermitValuation = Replace(oRs("permitvaluation"),"""","&quot;")
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function ShowValuationStepTable( iPermitValuationTypeid )
'--------------------------------------------------------------------------------------------------
Function ShowValuationStepTable( iPermitValuationTypeid )
	Dim sSql, oRs, iRowCount 

	sSql = "SELECT atleastvalue, notmorethanvalue, baseamount, unitqty, unitamount FROM egov_permitvaluationtypestepfees WHERE permitvaluationtypeid = " & iPermitValuationTypeid & " ORDER BY atleastvalue"
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
			response.write "<td align=""center""><input type=""text"" id=""atleastvalue" & iRowCount &""" name=""atleastvalue" & iRowCount &""" value=""" & FormatNumber(oRs("atleastvalue"),2,,,0) & """ size=""12"" maxlength=""12"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""notmorethanvalue" & iRowCount &""" name=""notmorethanvalue" & iRowCount &""" value=""" & FormatNumber(oRs("notmorethanvalue"),2,,,0) & """ size=""12"" maxlength=""12"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""unitamount" & iRowCount &""" name=""unitamount" & iRowCount &""" value=""" & FormatNumber(oRs("unitamount"),2,,,0) & """ size=""9"" maxlength=""9"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""unitqty" & iRowCount &""" name=""unitqty" & iRowCount &""" value=""" & oRs("unitqty") & """ size=""9"" maxlength=""9"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""baseamount" & iRowCount &""" name=""baseamount" & iRowCount &""" value=""" & FormatNumber(oRs("baseamount"),2,,,0) & """ size=""9"" maxlength=""9"" /></td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop 
	Else
		iRowCount = iRowCount + CLng(1)
		response.write "<tr>"
		response.write "<td align=""center""><input type=""text"" id=""atleastvalue1"" name=""atleastvalue1"" value=""0.00"" size=""12"" maxlength=""12"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""notmorethanvalue1"" name=""notmorethanvalue1"" value=""999999999.99"" size=""12"" maxlength=""12"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""unitamount1"" name=""unitamount1"" value=""0.00"" size=""9"" maxlength=""9"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""unitqty1"" name=""unitqty1"" value=""1"" size=""9"" maxlength=""9"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""baseamount1"" name=""baseamount1"" value=""0.00"" size=""9"" maxlength=""9"" /></td>"
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowValuationStepTable = iRowCount
End Function 
%>
