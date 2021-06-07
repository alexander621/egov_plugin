<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitresidentialunittypeedit.asp
' AUTHOR: Steve Loar
' CREATED: 10/30/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   11/30/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, iResidentialUnitTypeid, sResidentialUnitType, sUnitFeeAmount, sUseStepTable, iMaxRows

sLevel = "../" ' Override of value from common.asp
sUseStepTable = ""
iMaxRows = 1

PageDisplayCheck "residential unit types", sLevel	' In common.asp

iResidentialUnitTypeId = CLng(request("residentialunittypeid") )

If CLng(iResidentialUnitTypeId) > CLng(0) Then
	sTitle = "Edit"
	GetResidentialUnitDetails iResidentialUnitTypeid
Else
	sTitle = "New"
	sUnitFeeAmount = "0.00"
	sResidentialUnitType = ""
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
			location.href="permitresidentialunittypeedit.asp?permitresidentialunittypeid=0";
		}

		function NewStepRow()
		{
			document.frmResidentialUnit.maxrows.value = parseInt(document.frmResidentialUnit.maxrows.value) + 1;
			var tbl = $("feesteptable");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmResidentialUnit.maxrows.value);
			var row = tbl.insertRow(lastRow);

			var cellOne = row.insertCell(0);
			cellOne.align = 'center';
			var inputOne = document.createElement('input');
			inputOne.type = 'text';
			inputOne.id = 'atleastqty' + newRow;
			inputOne.name = 'atleastqty' + newRow;
			if (newRow > 1)
			{
				inputOne.value = $("notmorethanqty" + (newRow - 1)).value;
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

			var cellThree = row.insertCell(3);
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

			// Check for a residential unit type name
			if ($F("residentialunittype") == '')
			{
				alert("Please provide a residential unit type name, then try saving again.");
				$("residentialunittype").focus();
				return;
			}
			
			var iStepRows = 0;
			// Check the step table values entered
			for (var t = 1; t <= parseInt(document.frmResidentialUnit.maxrows.value); t++)
			{
				if ($("atleastqty" + t).value != '')
				{
					iStepRows += 1;
					// Remove any extra spaces
					$("atleastqty" + t).value = removeSpaces($("atleastqty" + t).value);
					//Remove commas that would cause problems in validation
					$("atleastqty" + t).value = removeCommas($("atleastqty" + t).value);
		
					// Validate the at least quantity format
					rege = /^\d+$/;
					Ok = rege.test($("atleastqty" + t).value);
					if ( ! Ok )
					{
						alert("The 'At Least Quantity' should be blank or a whole number value.\nPlease correct this and try saving again.");
						$("atleastqty" + t).focus();
						return;
					}

					// Validate the not more than quantity format
					if ($("notmorethanqty" + t).value != '')
					{
						// Remove any extra spaces
						$("notmorethanqty" + t).value = removeSpaces($("notmorethanqty" + t).value);
						//Remove commas that would cause problems in validation
						$("notmorethanqty" + t).value = removeCommas($("notmorethanqty" + t).value);
		
						rege = /^\d+$/;
						Ok = rege.test($("notmorethanqty" + t).value);
						if ( ! Ok )
						{
							alert("The 'Not More Than Quantity' must be a whole number value.\nPlease correct this and try saving again.");
							$("notmorethanqty" + t).focus();
							return;
						}
					}
					else
					{
						alert("The 'Not More Than Quantity' cannot be blank and must be a whole number value.\nPlease correct this and try saving again.");
						$("notmorethanqty" + t).focus();
						return;
					}

					// Validate the Base Fee Amount format
					if ($("baseamount" + t).value != '')
					{
						// Remove any extra spaces
						$("baseamount" + t).value = removeSpaces($("baseamount" + t).value);
						//Remove commas that would cause problems in validation
						$("baseamount" + t).value = removeCommas($("baseamount" + t).value);
		
						rege = /^\d*\.?\d{0,2}$/;
						Ok = rege.test($("baseamount" + t).value);
						if ( ! Ok )
						{
							alert("The 'Base Fee Amount' must be in currency format.\nPlease correct this and try saving again.");
							$("baseamount" + t).focus();
							return;
						}
						else
						{
							$("baseamount" + t).value = format_number(Number($("baseamount" + t).value),2);
						}
					}
					else
					{
						alert("The 'Base Fee Amount' cannot be blank and must be in currency format.\nPlease correct this and try saving again.");
						$("baseamount" + t).focus();
						return;
					}

					// Validate the Unit Amount format
					if ($("unitamount" + t).value != '')
					{
						// Remove any extra spaces
						$("unitamount" + t).value = removeSpaces($("unitamount" + t).value);
						//Remove commas that would cause problems in validation
						$("unitamount" + t).value = removeCommas($("unitamount" + t).value);
		
						rege = /^\d*\.?\d{0,2}$/;
						Ok = rege.test($("unitamount" + t).value);
						if ( ! Ok )
						{
							alert("The 'Unit Amount' must be in currency format.\nPlease correct this and try saving again.");
							$("unitamount" + t).focus();
							return;
						}
						else
						{
							$("unitamount" + t).value = format_number(Number($("unitamount" + t).value),2);
						}
					}
					else
					{
						alert("The 'Unit Amount' cannot be blank and must be in currency format.\nPlease correct this and try saving again.");
						$("unitamount" + t).focus();
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
			document.frmResidentialUnit.submit();
		}

		function Delete() 
		{
			if (confirm("Do you wish to delete this residential unit type?"))
			{
				location.href="permitresidentialunittypedelete.asp?permitresidentialunittypeid=<%=iResidentialUnitTypeid%>";
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
				<font size="+1"><strong><%=sTitle%> Residential Unit Type</strong></font><br /><br />
				<a href="permitresidentialunittypelist.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
			</p>
			<!--END: PAGE TITLE-->

<%		If CLng(iResidentialUnitTypeid) = CLng(0) Then %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Create" /><br />
<%		Else %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Update" /> &nbsp; &nbsp;
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" /> &nbsp; &nbsp;
<%			If request("success") <> "" Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Another();" value="Create Another" />
<%			End If		%>
			<br />
<%		End If %>

		<form name="frmResidentialUnit" action="permitresidentialunittypeupdate.asp" method="post">
		<input type="hidden" name="residentialunittypeid" value="<%=iResidentialUnitTypeid%>" />
		
		<p>
			Residential Unit Type Name: &nbsp;&nbsp; <input type="text" id="residentialunittype" name="residentialunittype" value="<%=sResidentialUnitType%>" size="100" maxlength="150" />
		</p>
		<p>
			<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add Row" id="addref" onClick="NewStepRow()" />
		</p>
		<div class="shadow">
			<table cellpadding="5" cellspacing="0" border="0" class="tableadmin" id="feesteptable">
				<tr><th>Has<br />At Least</th><th>And Is<br />Less Than</th><th>Price Per<br />Unit</th><th>Then Add<br />This Amount</th></tr>
<%				
				iMaxRows = ShowStepTable( iResidentialUnitTypeid ) 
%>
			</table>
		</div>
		* To remove a row, blank out the &quot;Has At Least&quot; value. The row will be removed when your changes are saved.
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
' Sub GetResidentialUnitDetails( iResidentialUnitTypeid )
'--------------------------------------------------------------------------------------------------
Sub GetResidentialUnitDetails( iResidentialUnitTypeid )
	Dim sSql, oRs

	sSql = "SELECT residentialunittype FROM egov_permitresidentialunittypes WHERE residentialunittypeid = " & iResidentialUnitTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sResidentialUnitType = Replace(oRs("residentialunittype"),"""","&quot;")
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function ShowStepTable( iResidentialUnitTypeid )
'--------------------------------------------------------------------------------------------------
Function ShowStepTable( iResidentialUnitTypeid )
	Dim sSql, oRs, iRowCount 

	sSql = "SELECT atleastqty, notmorethanqty, baseamount, unitamount FROM egov_permitresidentialunittypestepfees WHERE residentialunittypeid = " & iResidentialUnitTypeid & " ORDER BY atleastqty"
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
		response.write "<td align=""center""><input type=""text"" id=""baseamount1"" name=""baseamount1"" value=""0.00"" size=""9"" maxlength=""9"" /></td>"
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowStepTable = iRowCount
End Function 
%>
