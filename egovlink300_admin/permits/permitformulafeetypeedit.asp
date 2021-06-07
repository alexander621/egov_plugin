<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitformulafeetypeedit.asp
' AUTHOR: Steve Loar
' CREATED: 01/10/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This allows the creation and edit of formula type fees types for permits.
'
' MODIFICATION HISTORY
' 1.0   01/10/2008	Steve Loar - INITIAL VERSION
' 1.1	09/29/2008	Steve Loar - Sewer fee report flag added for Papillion
' 1.3	10/29/2008	Steve Loar - Fee Reporting types added 
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, iPermitFeeTypeid, sPermitFee, sUnitQty, sUnitAmount, sCategoryId, sIsReinspectionFee
Dim sIsForBuildingPermits, iAccountId, sMinAmount, sPermitFeePrefix, sBaseAmount, sIsUpfrontFee
Dim iFeeMethodId, bOnSewerFeeReport, sUpFrontAmount, bShowUpFrontAmount, bOnBBSFeeReport
Dim iFeeReportingTypeId

sLevel = "../" ' Override of value from common.asp

'PageDisplayCheck "permit fee types", sLevel	' In common.asp
PageDisplayCheck "permit types", sLevel	' In common.asp


iPermitFeeTypeid = CLng(request("permitfeetypeid") )

If CLng(iPermitFeeTypeid) > CLng(0) Then
	sTitle = "Edit"
	GetPermitFeeType iPermitFeeTypeid
Else
	sTitle = "New"
	sUnitQty = "1"
	sUnitAmount = "0.0000"
	sMinAmount = "0.00"
	sBaseAmount = "0.00"
	sCategoryId = 0
	iAccountId = 0
	iFeeMethodId = 0
	sIsForBuildingPermits = ""
	sPermitFeePrefix = ""
	sUpFrontAmount = "0.00"
	iFeeReportingTypeId = 0
End If 

bShowUpFrontAmount = OrgHasFeature( "up front fees" )

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
				location.href="permitfeetypecopy.asp?permitfeetypeid=" + iPermitFeeTypeId + "&redirectpage=permitformulafeetypeedit";
			}
		}

		function Another()
		{
			location.href="permitformulafeetypeedit.asp?permitfeetypeid=0";
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

			// Validate the Base Fee Amount format
			if (document.getElementById("baseamount").value != '')
			{
				// Remove any extra spaces
				document.getElementById("baseamount").value = removeSpaces(document.getElementById("baseamount").value);
				//Remove commas that would cause problems in validation
				document.getElementById("baseamount").value = removeCommas(document.getElementById("baseamount").value);

				rege = /^(-?)\d*\.?\d{0,2}$/; // Allows for negative
				//rege = /^\d*\.?\d{0,2}$/;
				Ok = rege.test(document.getElementById("baseamount").value);
				if ( ! Ok )
				{
					alert("The 'Base Fee Amount' must be in currency format.\nPlease correct this and try saving again.");
					document.getElementById("baseamount").focus();
					return;
				}
				else
				{
					document.getElementById("baseamount").value = format_number(Number(document.getElementById("baseamount").value),2);
				}
			}
			else
			{
				document.getElementById("baseamount").value = '0.00';
			}

			// Validate the Unit Qty format
			if (document.getElementById("unitqty").value != '')
			{
				// Remove any extra spaces
				document.getElementById("unitqty").value = removeSpaces(document.getElementById("unitqty").value);
				//Remove commas that would cause problems in validation
				document.getElementById("unitqty").value = removeCommas(document.getElementById("unitqty").value);

				rege = /^\d+$/;
				Ok = rege.test(document.getElementById("unitqty").value);
				if ( ! Ok )
				{
					alert("The 'Unit Amount' must be a whole number value.\nPlease correct this and try saving again.");
					document.getElementById("unitqty").focus();
					return;
				}
				//else
				//{
				//	document.getElementById("unitqty").value = format_number(Number(document.getElementById("unitqty").value),0);
				//}
			}
			else
			{
				document.getElementById("unitqty").value = '0';
			}

			// Validate the Unit Amount format
			if (document.getElementById("unitamount").value != '')
			{
				// Remove any extra spaces
				document.getElementById("unitamount").value = removeSpaces(document.getElementById("unitamount").value);
				//Remove commas that would cause problems in validation
				document.getElementById("unitamount").value = removeCommas(document.getElementById("unitamount").value);

				rege = /^\d*\.?\d{0,4}$/;
				Ok = rege.test(document.getElementById("unitamount").value);
				if ( ! Ok )
				{
					alert("The 'Unit Amount' must be in currency format.\nPlease correct this and try saving again.");
					document.getElementById("unitamount").focus();
					return;
				}
				else
				{
					document.getElementById("unitamount").value = format_number(Number(document.getElementById("unitamount").value),4);
				}
			}
			else
			{
				document.getElementById("unitamount").value = '0.0000';
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

			// Validate the Up Front Amount
			if ($("#upfrontamount").length > 0)
			{
				if ($("#upfrontamount").val() != '')
				{
					// Remove any extra spaces
					$("#upfrontamount").val(removeSpaces($("#upfrontamount").val()));
					//Remove commas that would cause problems in validation
					$("#upfrontamount").val(removeCommas($("#upfrontamount").val()));

					rege = /^\d*\.?\d{0,2}$/;
					Ok = rege.test($("#upfrontamount").val());
					if ( ! Ok )
					{
						alert("The 'Up Front Amount' must be in currency format.\nPlease correct this and try saving again.");
						$("#upfrontamount").focus();
						return;
					}
					else
					{
						$("#upfrontamount").val(format_number(Number($("#upfrontamount").val()),2));
					}
				}
				else
				{
					$("#upfrontamount").val('0.00');
				}
			}

			//Select all the included multipliers so they are passed to the saving page
			var elSel = document.getElementById('feemultipliertypeid');
			if (elSel.length > 1)
			{
				elSel.multiple = true;
				var i;
				for (i = elSel.length - 1; i>=0; i--) 
				{
					elSel.options[i].selected = true;
				}
			}

			//alert('OK');
			document.frmFeeTypes.submit();
		}

		
		function appendMultiplier()
		{
			var elOld = document.getElementById('newfeemultipliertypeid');
			var elSel = document.getElementById('feemultipliertypeid');
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
				if (elSel.length > 0 && elSel.length < 4)
				{
					elSel.size = elSel.size + 1;
				}
				elOld.remove(elOld.selectedIndex);
			}
		}

		function removeMultiplier()
		{
			var elOld = document.getElementById('newfeemultipliertypeid');
			var elSel = document.getElementById('feemultipliertypeid');
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


		function Delete() 
		{
			if (confirm("Do you wish to delete this permit fee type?"))
			{
				location.href="permitfeetypedelete.asp?permitfeetypeid=<%=iPermitFeeTypeid%>";
			}
		}

<%		If request("success") <> "" Then 
			DisplayMessagePopUp 
		End If 
%>
		function EditPermitFeeCategories()
		{
			showModal('permitfeecategories.asp', 'Permit Fee Categories', 65, 55);
		}
		function EditGLAccounts()
		{
			showModal('../purchases/gl_account_mgmt.asp', 'GL Accounts', 85, 95);
		}
		function EditMultipliers()
		{
			showModal('feemultiplierlist.asp?id=<%=iPermitFeeTypeid%>', 'Fee Multiplier Rates', 85, 95);
		}

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
				<font size="+1"><strong><%=sTitle%> Formula Fee</strong></font><br /><br />
				<a href="permitfeetypelist.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
			</p>
			<!--END: PAGE TITLE-->

		<div id="functionlinks">
<%		If CLng(iPermitFeeTypeid) = CLng(0) Then	%>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" id="savebutton" value="Create" /><br />
<%		Else	%>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" id="savebutton" value="Save Changes" /> &nbsp; &nbsp;
			<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" />

     			<div class="dropdown right">
  				<button class="ui-button ui-widget ui-corner-all dd-green"><i class="fa fa-bars" aria-hidden="true"></i> Tools</button>
  				<div class="dropdown-content">
					<a href="javascript:Delete();">Delete</a>
					<a href="javascript:CopyFee(<%=iPermitFeeTypeid%>;)">Copy Fee</a>
<%			If request("success") <> "" Then %>
					<a href="javascript:Another();">Create Another</a>
<%			End If		%>
				</div>
			</div>
			<br />
<%		End If %>
		</div>

		<form name="frmFeeTypes" action="permitformulafeetypeupdate.asp" method="post">
		<input type="hidden" name="permitfeetypeid" value="<%=iPermitFeeTypeid%>" />
		<input type="hidden" name="isbuildingpermitfee" value="1" />
		
		<p>
			Fee Type Name: &nbsp;&nbsp; <input type="text" name="permitfee" value="<%=sPermitFee%>" size="100" maxlength="150" />
		</p>
		<p>
			Fee Category: &nbsp;&nbsp; <input type="text" name="permitfeeprefix" value="<%=sPermitFeePrefix%>" size="15" maxlength="15" />
			 &nbsp;&nbsp; 
			
			Invoice Category: &nbsp;&nbsp; <select name="permitfeecategorytypeid">
<%
			ShowCategories sCategoryId
%>
			</select>
			<input type="button" value="Edit Categories" onClick="EditPermitFeeCategories();" class="button ui-button ui-widget ui-corner-all" />
			 &nbsp;&nbsp; 
		</p>
		
<%		If OrgHasFeature( "gl accounts" ) Then %>
			<p>
				GL Account: &nbsp;&nbsp; <select name="accountid" class="glaccountDD">
<%
				ShowAccounts iAccountId
%>
				</select>
				<input type="button" value="Edit GL Accounts" onClick="EditGLAccounts();" class="button ui-button ui-widget ui-corner-all" />
			<p>
<%		End If %>

		<p>
			Reporting Type: &nbsp;&nbsp; <%	ShowFeeReportingTypes iFeeReportingTypeId %>
		</p>

		<p>
		<div class="shadow">
			<table cellpadding="3" cellspacing="0" border="0" class="tableadmin" id="feesteptable">
				<tr><th>Calculation Method</th><th>Base Fee<br />Amount</th><th>Unit<br />Qty</th><th>Unit<br />Amount</th><th>Multipliers</th><th>Minimum<br />Amount</th>
<%					If bShowUpFrontAmount Then	%>
						<th>Up Front<br />Amount</th>
<%					End If						%>
				</tr>
				<tr>
					<td align="center"><select name="permitfeemethodid"><% ShowFeeMethods iFeeMethodId %></select></td>
					<td align="center"><input type="input" id="baseamount" name="baseamount" value="<%=sBaseAmount%>" size="9" maxlength="9" /></td>
					<td align="center"><input type="input" id="unitqty" name="unitqty" value="<%=sUnitQty%>" size="9" maxlength="9" /></td>
					<td align="center"><input type="input" id="unitamount" name="unitamount" value="<%=sUnitAmount%>" size="9" maxlength="9" /></td>
					<td align="center"><% ShowIncludedMultipliers iPermitFeeTypeid %><br />
										<input type="button" name="removeone" value="Remove" class="button ui-button ui-widget ui-corner-all" onclick="removeMultiplier();" />
					</td>
					<td align="center"><input type="input" id="minimumamount" name="minimumamount" value="<%=sMinAmount%>" size="9" maxlength="9" /></td>
<%					If bShowUpFrontAmount Then	%>
						<td align="center"><input type="input" id="upfrontamount" name="upfrontamount" value="<%=sUpFrontAmount%>" size="9" maxlength="9" /></td>
<%					End If						%>
				</tr>
			</table>
		</div>
		</p>
		<p>
			Multipliers to include: <% ShowUnusedMultipliers iPermitFeeTypeid %> &nbsp; &nbsp;
			<input type="button" name="add" value="Add" class="button ui-button ui-widget ui-corner-all" onclick="appendMultiplier();" />
			<input type="button" value="Edit Multipliers" onClick="EditMultipliers();" class="button ui-button ui-widget ui-corner-all" />
		</p>
		<div id="functionlinks">
<%		If CLng(iPermitFeeTypeid) = CLng(0) Then	%>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" id="savebutton" value="Create" /><br />
<%		Else	%>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" id="savebutton" value="Save Changes" /> &nbsp; &nbsp;
			<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" />
<%		End If %>
		</div>
		</form>
		<!--END: EDIT FORM-->

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
	sSql = sSql & " ISNULL(baseamount, 0.00) AS baseamount, ISNULL(unitqty, 0) AS unitqty, ISNULL(unitamount, 0.0000) AS unitamount, "
	sSql = sSql & " ISNULL(accountid,0) AS accountid, permitfeecategorytypeid, permitfeemethodid, ISNULL(upfrontamount,0.00) AS upfrontamount, "
	sSql = sSql & " ISNULL(feereportingtypeid,0) AS feereportingtypeid "
	sSql = sSql & " FROM egov_permitfeetypes WHERE permitfeetypeid = " & iPermitFeeTypeid 
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sPermitFee = Replace(oRs("permitfee"),"""","&quot;")
		sPermitFeePrefix = Replace(oRs("permitfeeprefix"),"""","&quot;")
		sMinAmount = FormatNumber(oRs("minimumamount"),2,,,0)
		sBaseAmount = FormatNumber(oRs("baseamount"),2,,,0)
		sUnitQty = FormatNumber(oRs("unitqty"),0,,,0)
		sUnitAmount = FormatNumber(oRs("unitamount"),4,,,0)
		iAccountId = oRs("accountid")
		sCategoryId = oRs("permitfeecategorytypeid")
		iFeeMethodId = oRs("permitfeemethodid")
		sUpFrontAmount = FormatNumber(oRs("upfrontamount"),2,,,0)
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
' Sub ShowFeeMethods( iFeeMethodId )
'--------------------------------------------------------------------------------------------------
Sub ShowFeeMethods( iFeeMethodId )
	Dim sSql, oRs

	sSql = "SELECT permitfeemethodid, permitfeemethod FROM egov_permitfeemethods WHERE orgid = " & session("orgid" )
	sSql = sSql & " AND isfixture = 0 AND isvaluation = 0 AND isconstructiontypemethod = 0 AND ispercentage = 0 "
	sSql = sSql & " ORDER BY permitfeemethod "
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("permitfeemethodid") & """ "
		If CLng(iFeeMethodId) = CLng(oRs("permitfeemethodid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write " >" & oRs("permitfeemethod")
		response.write "</option>"
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing 
	
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowUnusedMultipliers( iPermitFeeTypeid )
'--------------------------------------------------------------------------------------------------
Sub ShowUnusedMultipliers( iPermitFeeTypeid )
	Dim sSql, oRs

	sSql = "SELECT feemultiplier, feemultipliertypeid FROM egov_feemultipliertypes "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " AND feemultipliertypeid NOT IN (SELECT feemultipliertypeid "
	sSql = sSql & " FROM egov_permitfeetypes_to_feemultipliertypes WHERE permitfeetypeid = " & iPermitFeeTypeid
	sSql = sSql & " ) ORDER BY feemultiplier"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""newfeemultipliertypeid"" name=""newfeemultipliertypeid"" class=""feemultiplierDD"">"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("feemultipliertypeid") & """>" & oRs("feemultiplier") & "</option>"
		oRs.MoveNext
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowIncludedMultipliers( iPermitFeeTypeid )
'--------------------------------------------------------------------------------------------------
Sub ShowIncludedMultipliers( iPermitFeeTypeid )
	Dim sSql, oRs, iSize

	iSize = GetMultiplierCount( iPermitFeeTypeid )

	sSql = "SELECT F.feemultiplier, T.feemultipliertypeid FROM egov_permitfeetypes_to_feemultipliertypes T, egov_feemultipliertypes F "
	sSql = sSql & " WHERE T.feemultipliertypeid = F.feemultipliertypeid AND T.permitfeetypeid = " & iPermitFeeTypeid
	sSql = sSql & " ORDER BY feemultiplier"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""feemultipliertypeid"" name=""feemultipliertypeid"" size=""" & iSize & """>"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("feemultipliertypeid") & """>" & oRs("feemultiplier") & "</option>"
		oRs.MoveNext
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetMultiplierCount( iPermitFeeTypeid )
'--------------------------------------------------------------------------------------------------
Function GetMultiplierCount( iPermitFeeTypeid )
	Dim sSql, oRs
	
	sSql = "SELECT COUNT(feemultipliertypeid) AS hits FROM egov_permitfeetypes_to_feemultipliertypes "
	sSql = sSql & " WHERE permitfeetypeid = " & iPermitFeeTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			If CLng(oRs("hits")) > CLng(3) Then
				GetMultiplierCount = 3
			Else 
				GetMultiplierCount = oRs("hits")
			End If 
		Else
			GetMultiplierCount = 0
		End If 
	Else
		GetMultiplierCount = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


%>
