<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitvaluationfeetypeedit.asp
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
Dim sTitle, iPermitFeeTypeid, sPermitFee, sPermitFeePrefix, sMinAmount, sIsForBuildingPermits, iAccountId
Dim sCategoryId, iPermitValuationTypeid, bOnSewerFeeReport, bOnBBSFeeReport, iFeeReportingTypeId

sLevel = "../" ' Override of value from common.asp

'PageDisplayCheck "permit fee types", sLevel	' In common.asp
PageDisplayCheck "permit types", sLevel	' In common.asp


iPermitFeeTypeid = CLng(request("permitfeetypeid") )

iPermitValuationTypeid = GetFeeTypeValuation( iPermitFeeTypeid )

If CLng(iPermitFeeTypeid) > CLng(0) Then
	sTitle = "Edit"
	GetPermitFeeType iPermitFeeTypeid
Else
	sTitle = "New"
	sPermitFee = ""
	sPermitFeePrefix = ""
	sMinAmount = "0.00"
	sIsForBuildingPermits = "" 
	iAccountId = 0
	sCategoryId = 0
	iFeeReportingTypeId = 0
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
  	<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  	<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script language="Javascript">
	<!--

		function CopyFee( iPermitFeeTypeId )
		{
			if (confirm("Make a copy of this fee?"))
			{
				location.href="permitfeetypecopy.asp?permitfeetypeid=" + iPermitFeeTypeId + "&redirectpage=permitvaluationfeetypeedit";
			}
		}

		function Another()
		{
			location.href="permitvaluationfeetypeedit.asp?permitfeetypeid=0";
		}

		function NewStepRow()
		{
			document.frmFeeTypes.maxrows.value = parseInt(document.frmFeeTypes.maxrows.value) + 1;
			var tbl = document.getElementById("feesteptable");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmFeeTypes.maxrows.value);
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

			// Check for a fee type name
			if (document.frmFeeTypes.permitfee.value == '')
			{
				alert("Please provide a Fee Type Name, then try saving again.");
				document.frmFeeTypes.permitfee.focus();
				return;
			}

			// Validate the Minimum Amount format
			if (document.getElementById("minimumamount").value != '')
			{
				// Remove any extra spaces
				document.getElementById("minimumamount").value = removeSpaces(document.getElementById("minimumamount").value);
				//Remove commas that would cause problems in validation
				document.getElementById("minimumamount").value = removeCommas(document.getElementById("minimumamount").value);

				rege = /^\d*\.?\d{0,2}$/;
				Ok = rege.test(document.getElementById("minimumamount").value);
				if ( ! Ok )
				{
					alert("The 'Minimum Amount' must be in currency format.\nPlease correct this and try saving again.");
					document.getElementById("minimumamount").focus();
					return;
				}
				else
				{
					document.getElementById("minimumamount").value = format_number(Number(document.getElementById("minimumamount").value),2);
				}
			}
			else
			{
				document.getElementById("minimumamount").value = '0.00';
			}

			var iStepRows = 0;
			// Check the step table values entered
			for (var t = 1; t <= parseInt(document.frmFeeTypes.maxrows.value); t++)
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

			//alert('OK');
			document.frmFeeTypes.submit();
		}

		function appendValuation()
		{
			var elOld = document.getElementById('newpermitvaluationtypeid');
			var elSel = document.getElementById('permitvaluationtypeid');
			if (elOld.length > 0)
			{
				var sNewText = elOld.options[elOld.selectedIndex].text;  
				var sNewValue = elOld.options[elOld.selectedIndex].value; 
				var elOptNew = document.createElement('option');
				elOptNew.text = sNewText;
				elOptNew.value = sNewValue;
				try 
				{
					elSel.add(elOptNew, null); // standards compliant; doesn't work in IE6
				}
				catch(ex) 
				{
					elSel.add(elOptNew); // IE6 only
				}
				if (elSel.length > 0 && elSel.length < 21)
				{
					elSel.size = elSel.size + 1;
				}
				elOld.remove(elOld.selectedIndex);
			}
		}

		function removeValuation()
		{
			var elOld = document.getElementById('newpermitvaluationtypeid');
			var elSel = document.getElementById('permitvaluationtypeid');
			if (elSel.length > 0)
			{
				var i;
				for (i = elSel.length - 1; i>=0; i--) 
				{
					if (elSel.options[i].selected) 
					{
						var elOptNew = document.createElement('option');
						elOptNew.text = elSel.options[i].text;
						elOptNew.value = elSel.options[i].value;
						try 
						{
							elOld.add(elOptNew, null); // standards compliant; doesn't work in IE6
						}
						catch(ex) 
						{
							elOld.add(elOptNew); // IE6 only
						}
						elSel.remove(i);
						elSel.size = elSel.size - 1;
						break;
					}
				}
			}
		}

		function removeAllValuations()
		{
			var elOld = document.getElementById('newpermitvaluationtypeid');
			var elSel = document.getElementById('permitvaluationtypeid');
			var elOptNew;
			if (elSel.length > 0)
			{
				var i;
				for (i = elSel.length - 1; i>=0; i--) 
				{
					elOptNew = document.createElement('option');
					elOptNew.text = elSel.options[i].text;
					elOptNew.value = elSel.options[i].value;
					try 
					{
						elOld.add(elOptNew, null); // standards compliant; doesn't work in IE6
					}
					catch(ex) 
					{
						elOld.add(elOptNew); // IE6 only
					}
					elSel.remove(i);
				}
				elSel.size = 0;
			}
		}

		function Delete() 
		{
			if (confirm("Do you wish to delete this fee type?"))
			{
				location.href="permitfeetypedelete.asp?permitfeetypeid=<%=iPermitFeeTypeid%>";
			}
		}
		function EditPermitFeeCategories()
		{
			showModal('permitfeecategories.asp', 'Permit Fee Categories', 65, 55);
		}
		function EditGLAccounts()
		{
			showModal('../purchases/gl_account_mgmt.asp', 'GL Accounts', 85, 95);
		}

<%		If request("success") <> "" Then 
			DisplayMessagePopUp 
		End If 
%>

	//-->
	</script>
	<script language="JavaScript" src="permitfeedd.js"></script>

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
				<font size="+1"><strong><%=sTitle%> Valuation Fee</strong></font><br /><br />
				<a href="permitfeetypelist.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
			</p>
			<!--END: PAGE TITLE-->

<%		If CLng(iPermitFeeTypeid) = CLng(0) Then %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" id="savebutton" value="Create" /><br />
<%		Else %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" id="savebutton" value="Save Changes" /> &nbsp; &nbsp;
			<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" /> &nbsp; &nbsp;
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" /> &nbsp; &nbsp;
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="CopyFee(<%=iPermitFeeTypeid%>);" value="Copy Fee" /> &nbsp; &nbsp;
<%			If request("success") <> "" Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Another();" value="Create Another" />
<%			End If		%>
			<br />
<%		End If %>

		<form name="frmFeeTypes" action="permitvaluationfeetypeupdate.asp" method="post">
		<input type="hidden" name="permitfeetypeid" value="<%=iPermitFeeTypeid%>" />
		<input type="hidden" name="isvaluationtypefee" value="1" />
		<input type="hidden" name="isbuildingpermitfee" value="1" />
		<p>
			Fee Type Name: &nbsp;&nbsp; <input type="text" id="permitfee" name="permitfee" value="<%=sPermitFee%>" size="100" maxlength="150" />
		</p>
		<p>
			Fee Category: &nbsp;&nbsp; <input type="text" name="permitfeeprefix" value="<%=sPermitFeePrefix%>" size="15" maxlength="15" />
			 &nbsp;&nbsp; 
			Minimum Amount: &nbsp;&nbsp; <input type="input" id="minimumamount" name="minimumamount" value="<%=sMinAmount%>" size="10" maxlength="10" />
		</p>
		<p>
			Invoice Category: &nbsp;&nbsp; <select name="permitfeecategorytypeid">
<%
			ShowCategories sCategoryId
%>
			</select>
			<input type="button" value="Edit Categories" onClick="EditPermitFeeCategories();" class="button ui-button ui-widget ui-corner-all" />
		</p>
<%		If OrgHasFeature( "gl accounts" ) Then %>
		<p>
			GL Account: &nbsp;&nbsp; <select name="accountid" class="glaccountDD">
<%
			ShowAccounts iAccountId
%>
			</select>
				<input type="button" value="Edit GL Accounts" onClick="EditGLAccounts();" class="button ui-button ui-widget ui-corner-all" />
		</p>
<%		End If %>
		
		<p>
			Reporting Type: &nbsp;&nbsp; <%	ShowFeeReportingTypes iFeeReportingTypeId %>
		</p>

		<p>
			<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add Row" id="addref" onClick="NewStepRow()" />
		</p>
		<div class="shadow">
			<table cellpadding="5" cellspacing="0" border="0" class="tableadmin" id="feesteptable">
				<tr><th>Value Is<br />At Least</th><th>And Is<br />Less Than</th><th>Price Per<br />Unit</th><th>Unit Is<br />This Qty</th><th>Then Add<br />This Amount</th></tr>
<%				
				iMaxRows = ShowValuationStepTable( iPermitFeeTypeid ) 
%>
			</table>
		</div>
		* To remove a row, blank out the &quot;Value Is At Least&quot; field for that row. The row will be removed when your changes are saved.
		<input type="hidden" name="maxrows" id="maxrows" value="<%=iMaxRows%>" />

		</form>
		<!--END: EDIT FORM-->
<%		If CLng(iPermitFeeTypeid) = CLng(0) Then %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" id="savebutton" value="Create" /><br />
<%		Else %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" id="savebutton" value="Save Changes" /> &nbsp; &nbsp;
			<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" /> &nbsp; &nbsp;
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" /> &nbsp; &nbsp;
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="CopyFee(<%=iPermitFeeTypeid%>);" value="Copy Fee" /> &nbsp; &nbsp;
<%			If request("success") <> "" Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Another();" value="Create Another" />
<%			End If		%>
			<br />
<%		End If %>

		</div>
		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  
	<!--#Include file="modal.asp"-->  

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
' Sub GetPermitFeeType( iPermitFeeTypeid )
'--------------------------------------------------------------------------------------------------
Sub GetPermitFeeType( iPermitFeeTypeid )
	Dim sSql, oRs

	sSql = "SELECT permitfee, ISNULL(permitfeeprefix,'') AS permitfeeprefix, ISNULL(minimumamount, 0.00) AS minimumamount, "
	sSql = sSql & " isbuildingpermitfee, ISNULL(accountid,0) AS accountid, permitfeecategorytypeid, ISNULL(feereportingtypeid,0) AS feereportingtypeid "
	sSql = sSql & " FROM egov_permitfeetypes WHERE permitfeetypeid = " & iPermitFeeTypeid 
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sPermitFee = Replace(oRs("permitfee"),"""","&quot;")
		sPermitFeePrefix = Replace(oRs("permitfeeprefix"),"""","&quot;")
		sMinAmount = FormatNumber(oRs("minimumamount"),2,,,0)
		iAccountId = oRs("accountid")
		sCategoryId = oRs("permitfeecategorytypeid")
		iFeeReportingTypeId = CLng(oRs("feereportingtypeid"))
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowCategories( sCategoryId )
'--------------------------------------------------------------------------------------------------
Sub ShowCategories( sCategoryId )
	Dim sSql, oRs

	sSql = "SELECT permitfeecategorytypeid, permitfeecategory, iscommercial FROM egov_permitfeecategorytypes WHERE orgid = " & session("orgid" )
	sSql = sSql & " ORDER BY displayorder, permitfeecategory "
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("permitfeecategorytypeid") & """ "
		If CLng(sCategoryId) = CLng(oRs("permitfeecategorytypeid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write " >" & oRs("permitfeecategory")
		If oRs("iscommercial") Then 
			response.write " (Commercial)"
		Else
			response.write " (Residential)"
		End If  
		response.write "</option>"
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing 
	
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowIncludedValuations( iPermitFeeTypeid )
'--------------------------------------------------------------------------------------------------
Sub ShowIncludedValuations( iPermitFeeTypeid )
	Dim sSql, oRs, iSize

	iSize = GetValuationCount( iPermitFeeTypeid )

	sSql = "SELECT F.permitvaluation, T.permitvaluationtypeid FROM egov_permitfeetypes_to_permitvaluationtypes T, egov_permitvaluationtypes F "
	sSql = sSql & " WHERE T.permitvaluationtypeid = F.permitvaluationtypeid AND T.permitfeetypeid = " & iPermitFeeTypeid
	sSql = sSql & " ORDER BY permitvaluation"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1
	
	response.write vbcrlf & "<select id=""permitvaluationtypeid"" name=""permitvaluationtypeid"" size=""" & iSize & """>"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("permitvaluationtypeid") & """>" & oRs("permitvaluation") & "</option>"
		oRs.MoveNext
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetValuationCount( iPermitFeeTypeid )
'--------------------------------------------------------------------------------------------------
Function GetValuationCount( iPermitFeeTypeid )
	Dim sSql, oRs
	
	sSql = "SELECT COUNT(permitvaluationtypeid) AS hits FROM egov_permitfeetypes_to_permitvaluationtypes "
	sSql = sSql & " WHERE permitfeetypeid = " & iPermitFeeTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			If CLng(oRs("hits")) > CLng(20) Then
				GetValuationCount = 20
			Else 
				GetValuationCount = oRs("hits")
			End If 
		Else
			GetValuationCount = 0
		End If 
	Else
		GetValuationCount = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowUnusedValuations( iPermitFeeTypeid )
'--------------------------------------------------------------------------------------------------
Sub ShowUnusedValuations( iPermitFeeTypeid )
	Dim sSql, oRs

	sSql = "SELECT permitvaluation, permitvaluationtypeid FROM egov_permitvaluationtypes "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " AND permitvaluationtypeid NOT IN (SELECT permitvaluationtypeid "
	sSql = sSql & " FROM egov_permitfeetypes_to_permitvaluationtypes WHERE permitfeetypeid = " & iPermitFeeTypeid
	sSql = sSql & " ) ORDER BY permitvaluation"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""newpermitvaluationtypeid"" name=""newpermitvaluationtypeid"">"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("permitvaluationtypeid") & """>" & oRs("permitvaluation") & "</option>"
		oRs.MoveNext
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowAccounts( iAccountId )
'--------------------------------------------------------------------------------------------------
Sub ShowAccounts( iAccountId )
	Dim sSql, oRs

	sSql = "SELECT accountid, accountname, ISNULL(accountnumber,'') AS accountnumber FROM egov_accounts WHERE orgid = " & session("orgid" )
	sSql = sSql & " AND accountstatus = 'A' ORDER BY accountname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(iAccountid) = CLng(0) Then 
			response.write vbcrlf & "<option value=""0"" selected=""selected"" >Select an Account</option>"
		End If 
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("accountid") & """ "
			If CLng(iAccountId) = CLng(oRs("accountid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " >" & oRs("accountname")
			If oRs("accountnumber") <> "" Then
				response.write " (" & oRs("accountnumber") & ")"
			End If 
			response.write "</option>"
			oRs.MoveNext 
		Loop
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowValuations( iPermitValuationTypeid )
'--------------------------------------------------------------------------------------------------
Sub ShowValuations( iPermitValuationTypeid )
	Dim sSql, oRs

	sSql = "SELECT permitvaluation, permitvaluationtypeid FROM egov_permitvaluationtypes "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY permitvaluation"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""permitvaluationtypeid"" name=""permitvaluationtypeid"">"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("permitvaluationtypeid") & """"
		If CLng(iPermitValuationTypeid) = CLng(oRs("permitvaluationtypeid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("permitvaluation") & "</option>"
		oRs.MoveNext
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetFeeTypeValuation( iPermitFeeTypeid )
'--------------------------------------------------------------------------------------------------
Function GetFeeTypeValuation( iPermitFeeTypeid )
	Dim sSql, oRs

	sSql = "SELECT permitvaluationtypeid FROM egov_permitfeetypes_to_permitvaluationtypes "
	sSql = sSql & " WHERE permitfeetypeid = " & iPermitFeeTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		GetFeeTypeValuation = CLng(oRs("permitvaluationtypeid"))
	Else
		GetFeeTypeValuation = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Function ShowValuationStepTable( iPermitValuationTypeid )
'--------------------------------------------------------------------------------------------------
Function ShowValuationStepTable( iPermitValuationTypeid )
	Dim sSql, oRs, iRowCount 

	sSql = "SELECT atleastvalue, notmorethanvalue, baseamount, unitqty, unitamount FROM egov_permitvaluationtypestepfees "
	sSql = sSql & " WHERE permitfeetypeid = " & iPermitFeeTypeid & " ORDER BY atleastvalue"
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
