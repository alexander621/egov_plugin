<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitpercentagetypefeeedit.asp
' AUTHOR: Steve Loar
' CREATED: 09/08/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This allows the creation and edit of percentage type fees types for permits.
'
' MODIFICATION HISTORY
' 1.0   09/08/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, iPermitFeeTypeid, sPermitFee, sPercentage, sCategoryId, iFeeMethodId
Dim iAccountId, sMinAmount, sPermitFeePrefix, bOnSewerFeeReport, bOnBBSFeeReport
Dim iFeeReportingTypeId

sLevel = "../" ' Override of value from common.asp

'PageDisplayCheck "permit fee types", sLevel	' In common.asp
PageDisplayCheck "permit types", sLevel	' In common.asp


iPermitFeeTypeid = CLng(request("permitfeetypeid") )

iFeeMethodId = GetPermitFeeMethodIdByType( "ispercentage" )

If CLng(iPermitFeeTypeid) > CLng(0) Then
	sTitle = "Edit"
	GetPermitFeeType iPermitFeeTypeid
Else
	sTitle = "New"
	sPercentage = "0.0000"
	sMinAmount = "0.00"
	sCategoryId = 0
	iAccountId = 0
	sPermitFeePrefix = ""
	bOnSewerFeeReport = False 
	bOnBBSFeeReport = False 
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
				location.href="permitfeetypecopy.asp?permitfeetypeid=" + iPermitFeeTypeId + "&redirectpage=permitpercentagetypefeeedit";
			}
		}
		
		function Another()
		{
			location.href="permitpercentagetypefeeedit.asp?permitfeetypeid=0";
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

			// Validate the Percentage format
			if (document.getElementById("percentage").value != '')
			{
				// Remove any extra spaces
				document.getElementById("percentage").value = removeSpaces(document.getElementById("percentage").value);
				//Remove commas that would cause problems in validation
				document.getElementById("percentage").value = removeCommas(document.getElementById("percentage").value);

				rege = /^\d{0,1}\.\d{0,4}$/;
				Ok = rege.test(document.getElementById("percentage").value);
				if ( ! Ok )
				{
					alert("The 'Percentage' must be a number only in the format of #.####.\nPlease correct this and try saving again.");
					document.getElementById("percentage").focus();
					return;
				}
				else
				{
					if (Number(document.getElementById("percentage").value) > 1)
					{
						alert("The 'Percentage' must be a number not greater than 1.\nPlease correct this and try saving again.");
						document.getElementById("percentage").focus();
						return;
					}

					if (Number(document.getElementById("percentage").value) == 0)
					{
						alert("The 'Percentage' must be a number greater than 0.\nPlease correct this and try saving again.");
						document.getElementById("percentage").focus();
						return;
					}
					
					document.getElementById("percentage").value = format_number(Number(document.getElementById("percentage").value),4);
				}
			}
			else
			{
				//document.getElementById("percentage").value = '0.0000';
				alert("Please provide a Percentage, then try saving again.");
				document.frmFeeTypes.percentage.focus();
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
			//alert(document.getElementById("percentage").value);
			//alert('OK');
			document.frmFeeTypes.submit();
		}

		function Delete() 
		{
			if (confirm("Do you wish to delete this permit fee type?"))
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
				<font size="+1"><strong><%=sTitle%> Percentage Fee</strong></font><br /><br />
				<a href="permitfeetypelist.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
			</p>
			<!--END: PAGE TITLE-->

		<div id="functionlinks">
<%		If CLng(iPermitFeeTypeid) = CLng(0) Then	%>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" id="savebutton" value="Create" /><br />
<%		Else	%>
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

		<form name="frmFeeTypes" action="permitpercentagetypeupdate.asp" method="post">
		<input type="hidden" name="permitfeetypeid" value="<%=iPermitFeeTypeid%>" />
		<input type="hidden" name="isbuildingpermitfee" value="1" />
		<input type="hidden" name="permitfeemethodid" value="<%=iFeeMethodId%>" />
		
		<p>
			Fee Type Name: &nbsp;&nbsp; <input type="text" name="permitfee" value="<%=sPermitFee%>" size="100" maxlength="150" />
		</p>
		<p>
			Fee Category: &nbsp;&nbsp; <input type="text" name="permitfeeprefix" value="<%=sPermitFeePrefix%>" size="15" maxlength="15" />
		</p>
		<p>
<%		If OrgHasFeature( "gl accounts" ) Then %>

			GL Account: &nbsp;&nbsp; <select name="accountid" class="glaccountDD">
<%
			ShowAccounts iAccountId
%>
			</select>
			<input type="button" value="Edit GL Accounts" onClick="EditGLAccounts();" class="button ui-button ui-widget ui-corner-all" />
		
<%		End If %>
		</p>
		
		<p>
			Reporting Type: &nbsp;&nbsp; <%	ShowFeeReportingTypes iFeeReportingTypeId %>
		</p>

		<p>
		<div class="shadow">
			<table cellpadding="3" cellspacing="0" border="0" class="tableadmin" id="feesteptable">
				<tr><th>Invoice<br />Category</th><th>Percentage</th><th>Minimum<br />Amount</th></tr>
				<tr>
					<td align="center"><select name="permitfeecategorytypeid"><% ShowCategories sCategoryId %></select><input type="button" value="Edit Categories" onClick="EditPermitFeeCategories();" class="button ui-button ui-widget ui-corner-all" /></td>
					<td align="center"><input type="input" id="percentage" name="percentage" value="<%=sPercentage%>" size="6" maxlength="6" /></td>
					<td align="center"><input type="input" id="minimumamount" name="minimumamount" value="<%=sMinAmount%>" size="9" maxlength="9" /></td>
				</tr>
			</table>
		</div>
		</p>
		</form>
		<!--END: EDIT FORM-->
		<div id="functionlinks">
<%		If CLng(iPermitFeeTypeid) = CLng(0) Then	%>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" id="savebutton" value="Create" /><br />
<%		Else	%>
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
	sSql = sSql & " ISNULL(percentage, 0.0000) AS percentage, ISNULL(accountid,0) AS accountid, permitfeecategorytypeid, "
	sSql = sSql & " ISNULL(feereportingtypeid,0) AS feereportingtypeid "
	sSql = sSql & " FROM egov_permitfeetypes WHERE permitfeetypeid = " & iPermitFeeTypeid 
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sPermitFee = Replace(oRs("permitfee"),"""","&quot;")
		sPermitFeePrefix = Replace(oRs("permitfeeprefix"),"""","&quot;")
		sMinAmount = FormatNumber(oRs("minimumamount"),2,,,0)
		sPercentage = FormatNumber(oRs("percentage"),4,,,0)
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

	sSql = "SELECT permitfeecategorytypeid, permitfeecategory, iscommercial "
	sSql = sSql & " FROM egov_permitfeecategorytypes WHERE orgid = " & session("orgid" )
	sSql = sSql & " ORDER BY displayorder, permitfeecategory "
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	response.write vbcrlf & "<option value=""-1"" "
	If CLng(sCategoryId) = CLng(-1) Then
			response.write " selected=""selected"" "
		End If 
		response.write " >All Fees</option>"

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


%>
