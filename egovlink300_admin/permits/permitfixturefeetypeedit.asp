<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitfixturefeetypeedit.asp
' AUTHOR: Steve Loar
' CREATED: 01/07/08
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/07/08	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, iPermitFeeTypeid, sPermitFee, sPermitFeePrefix, sMinAmount, sIsForBuildingPermits, iAccountId
Dim sCategoryId, bOnSewerFeeReport, sUpFrontAmount, bShowUpFrontAmount, bOnBBSFeeReport
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
	sPermitFee = ""
	sPermitFeePrefix = ""
	sMinAmount = "0.00"
	sIsForBuildingPermits = "" 
	iAccountId = 0
	sCategoryId = 0
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
				location.href="permitfeetypecopy.asp?permitfeetypeid=" + iPermitFeeTypeId + "&redirectpage=permitfixturefeetypeedit";
			}
		}

		function Another()
		{
			location.href="permitfixturefeetypeedit.asp?permitfeetypeid=0";
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

			//Select all the included fixtures so they are passed to the saving page
			var elSel = document.getElementById('permitfixturetypeid');
			elSel.multiple = true;
			var i;
			for (i = elSel.length - 1; i>=0; i--) 
			{
				elSel.options[i].selected = true;
			}

			//alert('OK');
			document.frmFeeTypes.submit();
		}

		function appendFixture()
		{
			var elOld = document.getElementById('newpermitfixturetypeid');
			var elSel = document.getElementById('permitfixturetypeid');
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

		function removeFixture()
		{
			var elOld = document.getElementById('newpermitfixturetypeid');
			var elSel = document.getElementById('permitfixturetypeid');
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

		function removeAllFixtures()
		{
			var elOld = document.getElementById('newpermitfixturetypeid');
			var elSel = document.getElementById('permitfixturetypeid');
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

		function init()
		{
<%			'if request("success") <> "" then 
			'	response.write "showDialog('Success','" & request("success") & "','success', .7);"
			'end if 
%>
			Sortable.create('permitfixturetypeid');
		}
		function EditPermitFeeCategories()
		{
			showModal('permitfeecategories.asp', 'Permit Fee Categories', 65, 55);
		}
		function EditGLAccounts()
		{
			showModal('../purchases/gl_account_mgmt.asp', 'GL Accounts', 85, 95);
		}
		function EditPermitFixtureTypes()
		{
			showModal('permitfixturetypelist.asp', 'Edit Permit Fixture Types', 75, 65);
		}

<%		If request("success") <> "" Then 
			DisplayMessagePopUp 
		End If 
%>

	//-->
	</script>
	<script language="JavaScript" src="permitfeedd.js"></script>

</head>

<body onload="init()">

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong><%=sTitle%> Fixture Fee</strong></font><br /><br />
				<a href="permitfeetypelist.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
			</p>
			<!--END: PAGE TITLE-->

		<div id="functionlinks">
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

		<form name="frmFeeTypes" action="permitfixturefeetypeupdate.asp" method="post">
		<input type="hidden" name="permitfeetypeid" value="<%=iPermitFeeTypeid%>" />
		<input type="hidden" name="isfixturetypefee" value="1" />
		<input type="hidden" name="isbuildingpermitfee" value="1" />
		<p>
			Fee Type Name: &nbsp;&nbsp; <input type="text" id="permitfee" name="permitfee" value="<%=sPermitFee%>" size="100" maxlength="150" />
		</p>
		<p>
			Fee Category: &nbsp;&nbsp; <input type="text" name="permitfeeprefix" value="<%=sPermitFeePrefix%>" size="15" maxlength="15" />
			 &nbsp;&nbsp; 
			Minimum Amount: &nbsp;&nbsp; <input type="input" id="minimumamount" name="minimumamount" value="<%=sMinAmount%>" size="9" maxlength="9" />
			 &nbsp;&nbsp; 
			Up Front Amount: &nbsp;&nbsp; <input type="input" id="upfrontamount" name="upfrontamount" value="<%=sUpFrontAmount%>" size="9" maxlength="9" />
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
		
		<fieldset>
		<legend><strong>Fixtures</strong></legend>
			<p>
				Fixtures to Include: <% ShowUnusedFixtures iPermitFeeTypeid %> &nbsp; &nbsp;
				<input type="button" name="add" value="Add" class="button ui-button ui-widget ui-corner-all" onclick="appendFixture();" /> &nbsp; &nbsp;
				<input type="button" name="add" value="Edit Permit Fixture Types" class="button ui-button ui-widget ui-corner-all" onclick="EditPermitFixtureTypes();" />
			</p>
			<p>
				<div id="includedfixturetop">
					Included Fixtures:  &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
					<span id="includedfixturebuttons">
					<input type="button" name="removeone" class="button ui-button ui-widget ui-corner-all" value="Remove Selected" onclick="removeFixture();" /> &nbsp; &nbsp;
					<input type="button" name="removeall" class="button ui-button ui-widget ui-corner-all" value="Remove All" onclick="removeAllFixtures();" />		
					</span>
				</div>
<%
				ShowIncludedFixtures iPermitFeeTypeid
%>
				
			</p>

		</fieldset>

		</form>
		<!--END: EDIT FORM-->
		<div id="functionlinks">
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
	sSql = sSql & " isbuildingpermitfee, ISNULL(accountid,0) AS accountid, permitfeecategorytypeid, "
	sSql = sSql & " ISNULL(upfrontamount,0.00) AS upfrontamount, ISNULL(feereportingtypeid,0) AS feereportingtypeid "
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
' Sub ShowIncludedFixtures( iPermitFeeTypeid )
'--------------------------------------------------------------------------------------------------
Sub ShowIncludedFixtures( iPermitFeeTypeid )
	Dim sSql, oRs, iSize

	iSize = GetFixtureCount( iPermitFeeTypeid )

	sSql = "SELECT F.permitfixture, T.permitfixturetypeid FROM egov_permitfeetypes_to_permitfixturetypes T, egov_permitfixturetypes F "
	sSql = sSql & " WHERE T.permitfixturetypeid = F.permitfixturetypeid AND T.permitfeetypeid = " & iPermitFeeTypeid
	sSql = sSql & " ORDER BY F.displayorder, permitfixture"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1
	
	response.write vbcrlf & "<select id=""permitfixturetypeid"" name=""permitfixturetypeid"" size=""" & iSize & """>"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("permitfixturetypeid") & """>" & oRs("permitfixture") & "</option>"
		oRs.MoveNext
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetFixtureCount( iPermitFeeTypeid )
'--------------------------------------------------------------------------------------------------
Function GetFixtureCount( iPermitFeeTypeid )
	Dim sSql, oRs
	
	sSql = "SELECT COUNT(permitfixturetypeid) AS hits FROM egov_permitfeetypes_to_permitfixturetypes "
	sSql = sSql & " WHERE permitfeetypeid = " & iPermitFeeTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			If CLng(oRs("hits")) > CLng(20) Then
				GetFixtureCount = 20
			Else 
				GetFixtureCount = oRs("hits")
			End If 
		Else
			GetFixtureCount = 0
		End If 
	Else
		GetFixtureCount = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowUnusedFixtures( iPermitFeeTypeid )
'--------------------------------------------------------------------------------------------------
Sub ShowUnusedFixtures( iPermitFeeTypeid )
	Dim sSql, oRs

	sSql = "SELECT F.permitfixture, F.permitfixturetypeid FROM egov_permitfixturetypes F "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " AND permitfixturetypeid NOT IN (SELECT permitfixturetypeid "
	sSql = sSql & " FROM egov_permitfeetypes_to_permitfixturetypes WHERE permitfeetypeid = " & iPermitFeeTypeid
	sSql = sSql & " ) ORDER BY F.displayorder, F.permitfixture"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""newpermitfixturetypeid"" name=""newpermitfixturetypeid"">"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("permitfixturetypeid") & """>" & oRs("permitfixture") & "</option>"
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



%>
