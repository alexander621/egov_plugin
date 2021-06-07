<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: payment_accounts.asp
' AUTHOR: Steve Loar
' CREATED: 4/17/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the assignment of gl accounts to payment methods and refund accounts
'
' MODIFICATION HISTORY
' 1.0   4/17/2007   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMaxRow, sLoadMsg

iMaxRow = 0
sLoadMsg = ""
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "payment accounts" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

If request("s") <> "" Then
	If request("s") = "u" Then
		sLoadMsg = "eGovLink.Accounts.displayScreenMsg('Changes Saved');"
	End If 
End If 

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="../rentals/rentalsstyles.css" />
	<link rel="stylesheet" type="text/css" href="account_styles.css" />

	<script type="text/javascript" src="../scripts/jquery-1.7.2.min.js"></script>

	<script language="JavaScript" src="../scripts/formatnumber.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>
	<script language="JavaScript" src="../scripts/removecommas.js"></script>
	<script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="Javascript">
	<!--

		// create the egov NameSpace
		var eGovLink = eGovLink || {}; 

		// create the sub-NameSpace with the methods inside   
		eGovLink.Accounts = (function() 
		{

			var Validate = function()
			{
				
//				var bValid = true;
				var itemCount = Number($("#maxrows").val());

				for (x=1; x <= itemCount; x++ )
				{
					if ( $("#defaultamount" + x).length != 0 )
					{
						// Remove any extra spaces
						$("#defaultamount" + x).val(removeSpaces($("#defaultamount" + x).val()));
						//Remove commas that would cause problems in validation
						$("#defaultamount" + x).val(removeCommas($("#defaultamount" + x).val()));

						// Validate the format of the amount
						if ($("#defaultamount" + x).val() != "")
						{
							var rege = /^\d*\.?\d{0,2}$/
							var Ok = rege.exec($("#defaultamount" + x).val());
							if ( Ok )
							{
								$("#defaultamount" + x).val(format_number(Number($("#defaultamount" + x).val()),2));
							}
							else 
							{
								inlineMsg("defaultamount" + x,'<strong>Invalid Value: </strong> This must be a number in currency format.',5,"defaultamount" + x);
								$("#defaultamount" + x).focus();
								return false;
							}
						}
					}
				}
				
				document.formPayments.submit();

			}

			var displayScreenMsg = function( iMsg ) 
			{
				if( iMsg != "" ) 
				{
					$("#screenMsg").html( "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;" );
					window.setTimeout("eGovLink.Accounts.clearScreenMsg()", (10 * 1000));
				}
			}

			var clearScreenMsg = function() 
			{
				$("#screenMsg").html( "" );
			}
		
			// This makes the functions publicly accessible
			return {
				Validate: Validate,
				displayScreenMsg: displayScreenMsg,
				clearScreenMsg: clearScreenMsg
			};

		}());


<%		if sLoadMsg <> "" then %>
			$(document).ready(function() {
				<%=sLoadMsg%>
			});
<%		end if  %>


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
			<font size="+1"><strong>Payment and Refund Accounts</strong></font>
		</p><br />
		<!--END: PAGE TITLE-->

		<table id="screenMsgtable">
			<tr><td>
				<span id="screenMsg"></span>
			</td></tr>
		</table>

		<!--BEGIN: FUNCTION LINKS-->
		<div id="functionlinks">
			<input type="button" class="button" value="Save Changes" onclick="eGovLink.Accounts.Validate();" />
		</div><br /><br />
		<!--END: FUNCTION LINKS-->


		<!--BEGIN: EDIT FORM-->
		<form name="formPayments" action="payment_accounts_update.asp" method="post">

				<table cellpadding="7" cellspacing="0" border="0" class="tableadmin" id="paymentAccountList">
					<tr><th>Payment Method</th><th>GL Account</th><th>Public Method</th><th>Admin Method</th>
<%					bHasRefundDebit = OrgHasRefundDebitAcct() 
					If bHasRefundDebit Then
						response.write "<th>Default Amount</th>"
					End If
%>					
					</tr>
<%
					iMaxRow = ShowPaymentAccounts( bHasRefundDebit )
%>
				</table>

			<input type="hidden" id="maxrows" name="maxrows" value="<%=iMaxRow%>" />
		</form>
		<!--END: EDIT FORM-->

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>



<%

'--------------------------------------------------------------------------------------------------
' integer ShowPaymentAccounts( bHasRefundDebit )
'--------------------------------------------------------------------------------------------------
Function ShowPaymentAccounts( ByVal bHasRefundDebit )
	Dim sSql, oPayments, iRows

	iRows = 0

	sSql = "SELECT P.paymenttypeid, P.paymenttypename, P.ispublicmethod, P.isadminmethod, "
	sSql = sSql & " ISNULL(O.accountid,0) AS accountid, P.hasdefaultamount, ISNULL(O.defaultamount,0.00) AS defaultamount "
	sSql = sSql & " FROM egov_paymenttypes P, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE P.paymenttypeid = O.paymenttypeid AND O.orgid = " & Session("OrgId")
	sSql = sSql & " ORDER BY P.displayorder"

	Set oPayments = Server.CreateObject("ADODB.Recordset")
	oPayments.Open sSQL, Application("DSN"), 3, 1

	Do While Not oPayments.EOF

		iRows = iRows + 1
		response.write vbcrlf & "<tr"
		If iRows Mod 2 = 0 Then response.write " class=""altrow"" "
		response.write "><td align=""center"">" & oPayments("paymenttypename") & "<input type=""hidden"" name=""paymenttypeid" & iRows & """ value=""" & oPayments("paymenttypeid") & """ /></td>"
		
		response.write "<td align=""center"">"
		ShowAccountPicks oPayments("accountid"), iRows  ' In common.asp
		response.write "</td>"

		response.write "<td align=""center"">"
		If oPayments("ispublicmethod") Then
			response.write "Yes"
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"

		response.write "<td align=""center"">"
		If oPayments("isadminmethod") Then
			response.write "Yes"
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"

		If bHasRefundDebit Then 
			response.write "<td align=""center"">"
			If oPayments("hasdefaultamount") Then
				response.write "<input type=""text"" id=""defaultamount" & iRows & """ name=""defaultamount" & iRows & """ value=""" & FormatNumber(CDbl(oPayments("defaultamount")),2,,,0) & """ size=""6"" maxlength=""6"" />"
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"
		End If 

		response.write "</tr>"

		oPayments.MoveNext

	Loop
	
	oPayments.Close
	Set oPayments = Nothing 

	ShowPaymentAccounts = iRows

End Function 


%>