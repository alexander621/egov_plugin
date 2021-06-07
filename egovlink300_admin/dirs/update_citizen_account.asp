<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: update_citizen_account.asp
' AUTHOR: Steve Loar
' CREATED: 1/10/2007 - Copied from update_citizen.asp
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  citizen account deposits and withdrawls.
'
' MODIFICATION HISTORY
' 1.0   1/10/2007	Steve Loar - Initial code 
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sCitizenName, sAmount, sEntryType, sNotes, iUserID, cCurrentBalance, iJournalEntryTypeID, sType
Dim iJournalId, iPaymentTypeId, sShowType, iFamilyid, x, iLedgerId, iAccountId, sPaymentEntryType
Dim sPlusMinus, sRt, sJet

sLevel = "../" ' Override of value from common.asp

' Check the admin person's right to this page
If Not UserHasPermission( Session("UserId"), "edit citizens" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

iUserID = request("uid")
cCurrentBalance = GetCitizenCurrentBalance( iUserId )
sEntryType = request("entrytype")
If sEntryType = "credit" Then
	sShowType = "Deposit Funds"
	sShowButton = "Complete the Deposit"
Else ' Debit
	'sShowType = Proper( sEntryType )
	sShowType = "Withdraw Funds"
	sShowButton = "Complete the Withdrawal"
End If 
'iFamilyid = GetFamilyId( iUserId ) 

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then   ' This should be the page saving to itself
	'iOrgID = Session("OrgID") - already set in common.asp
	sAmount = request("amount")
	sNotes = dbsafe(request("notes"))
	
	If sEntryType = "credit" Then
		sType = "deposit"  ' 5
		sPaymentEntryType = "debit"
		sPlusMinus = "+"
	Else	' credit
		sType = "withdrawl"  ' 6
		sPaymentEntryType = "credit"
		sPlusMinus = "-"
	End If 
	
	iJournalEntryTypeID = GetJournalEntryTypeID( sType )

	iPaymentTypeId = GetPaymentTypeId1( "Citizen Account" )

	iItemTypeId = GetItemTypeId( "citizen account" )

	If sEntryType = "credit" Then
		' Handle deposits
		sRt = "c"
		sJet = "d"
		' This is the same as the paymentid in other payment scripts
		iJournalId = InsertJournalEntry( session("orgid"), iUserID, Session("UserID"), sAmount, iJournalEntryTypeID, sNotes )

		' Get the max payment id then loop thru and do inserts for those that have amounts
		iMaxPaymentTypes = GetmaxPaymentTypeId( Session("Orgid") )
		x = 1
		Do While x <= iMaxPaymentTypes
			If request("amount" & x) <> "" Then 
				If HasChecks( x ) Then
					' Check
					sCheck = "'" & dbsafe(request("checkno")) & "'"
					InsertPaymentRecord_Checks iJournalId, x, request("amount" & x), "APPROVED", sCheck
					' get the accountid of the payment type 
					iAccountId = GetPaymentAccountId( Session("Orgid"), x )  ' In ../includes/common.asp
					' Make the ledger entry for the payment 
					iLedgerId = MakeLedgerEntry( Session("Orgid"), iAccountId, iJournalId, CDbl(request("amount" & x)), 2, sPaymentEntryType, sPlusMinus, "NULL", 1, x, cCurrentBalance, "NULL" )
					'           MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iPriceTypeid )
				Else
					If HasCitizensAccounts( x ) Then
						' Transfer
						InsertPaymentRecord_Transfer iJournalId, x, request("amount" & x), "APPROVED", request("accountid")
						' Credit the account that was the source of the funds; includes a ledger entry
						AdjustCitizenAccount request("accountid"), iJournalId, session("orgid"), "debit", request("amount" & x), iItemTypeId
					Else
						' Charge, Cash and Other
						InsertPaymentRecord iJournalId, x, CDbl(request("amount" & x)), "APPROVED"
						' get the accountid of the payment type 
						iAccountId = GetPaymentAccountId( Session("Orgid"), x )  ' In ../includes/common.asp
						' Make the ledger entry for the payment 
						iLedgerId = MakeLedgerEntry( Session("Orgid"), iAccountId, iJournalId, CDbl(request("amount" & x)), 2, sPaymentEntryType, sPlusMinus, "NULL", 1, x, cCurrentBalance, "NULL" )
						'           MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iPriceTypeid )
					End If
				End If 
			End If 
			x = x + 1
		Loop 
	Else
		' handle Withdrawals
		sRt = "d"
		sJet = "w"
		'iPaymentTypeId = GetRefundPaymentTypeId( )  ' 6 is the refund voucher. In ../includes/common.asp
		iPaymentTypeId = request("paymenttypeid") ' the selected issue to pick
		
		If CLng(iPaymentTypeId) <> CLng(4) Then 
			' This is the same as the paymentid in other payment scripts
			iJournalId = InsertJournalEntry( session("orgid"), iUserID, Session("UserID"), sAmount, iJournalEntryTypeID, sNotes )

			' get the account id of the payment type. These are not other citizens
			iAccountId = GetPaymentAccountId( Session("Orgid"), iPaymentTypeId)  '  In ../includes/common.asp

			' Make the ledger entry for the account that gets the refund 
			iLedgerId = MakeLedgerEntry( Session("Orgid"), iAccountId, iJournalId, CDbl(sAmount), 2, sPaymentEntryType, sPlusMinus, "NULL", 1, iPaymentTypeId, cCurrentBalance, "NULL" )
			'           MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iPriceTypeid )
			' Because this is a withdrawl, no verisign payment records are made
		Else
			' 4 is citizen account transfer
			' transfers to another account have to be handled as a deposit to the receiving account, so this will never run??
			sJet = "d"
			' get the deposit journal entry type
			iJournalEntryTypeID = GetJournalEntryTypeID( "deposit" )

			' get the userid of the person getting the money
			iAccountId = request("accountid") ' who is getting the money

			' do the journal entry as a deposit to the person getting the money
			iJournalId = InsertJournalEntry( session("orgid"), iAccountId, Session("UserID"), sAmount, iJournalEntryTypeID, sNotes )

			
			' do the verisign payment record as money from the user
			InsertPaymentRecord_Transfer iJournalId, 4, sAmount, "APPROVED", iUserID

			' Credit the account that is getting the funds; includes a ledger entry
			AdjustCitizenAccount iAccountId, iJournalId, session("orgid"), "credit", sAmount, iItemTypeId
		End If 
		
	End If 

	' Adjust their account; includes a ledger entry
	AdjustCitizenAccount iUserID, iJournalId, session("orgid"), sEntryType, sAmount, iItemTypeId
	
	' see if the org has the undo feature and set the session variable'
	If OrgHasFeature("undo on citizen receipt") Then
		' In ../includes/common.asp'
		SetUnDoBtnDisplay iJournalId, True
	Else
		'More stuff here'
	End If 

	' want to route to the receipt here instead of the history page'
	'response.redirect "citizen_account_history.asp?u=" & iUserID'
	response.redirect "../purchases/viewjournal.asp?uid=" & iUserID & "&pid=" & iJournalId & "&rt=" & sRt & "&it=ci&jet=" & sJet & "&src=cah"

End If 

sCitizenName = GetCitizenName( iUserId )
cCurrentBalance = GetCitizenCurrentBalance( iUserId )

%>

<html lang="en">
<head>
	<meta charset="UTF-8">

	<title><%=langBSCommittees%></title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="reservationliststyles.css" />

	<script src="../scripts/jquery-1.6.1.min.js"></script>

	<script src="../scripts/ajaxLib.js"></script>
	<script src="../scripts/formatnumber.js"></script>
	<script src="../scripts/removespaces.js"></script>
	<script src="../scripts/removecommas.js"></script>

	<script>
	<!--
		function validate() 
		{
			// handle withdrawals
			if ($("#entrytype").val() == "debit")
			{
				//alert(document.accountForm.amount.value);
				// Check that amount is entered
				if ($("#amount").val() === "")
				{
					alert('Please enter an Amount for this entry.');
					$("#amount").focus();
					return;
				}

				// Check that amount is currency
				//var rege = /^(?:\d+|\d{1,3}(?:,\d{3})*)(?:\.\d{1,2}){0,1}$/;
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec($("#amount").val());
				if ( ! Ok )
				{
					alert("The amount must be a number without any formatting.\nExample: 1234.56\n\nPlease try again.");
					$("#amount").focus();
					return;
				}
				
				// check the amount to withdraw is not more than the current balance.
				var cAmount = Number( $("#currentbalance").val() );
				var wAmount = Number( $("#amount").val() );

				if (wAmount === 0)
				{
					alert("The amount must be a number greater than 0.\nPlease try again.");
					$("#amount").focus();
					return;
				}
				
				if (wAmount > cAmount)
				{
					alert("The amount must be a number less than or equal to the current balance.\nPlease try again.");
					$("#amount").focus();
					return;
				}

				// Submit the credit
				document.accountForm.submit();
				//alert("Ok to submit.");
			}
			else  //Credit - deposits
			{
				// Check the debit fields - We cheat in the javascript code to handle the input correctly
				// Do we have any charge amount
				if ($("#amount1").val() != "")
				{
					var rege1 = /^\d*\.?\d{0,2}$/
					var Ok1 = rege1.exec($("#amount1").val());
					if ( ! Ok1 )
					{
						alert("The Charge Card amount must be a number without any formatting.\nExample: 1234.56\n\nPlease try again.");
						$("#amount1").focus();
						return;
					}
					//alert(document.accountForm.amount1.value);
				}
				// Do we have any check amount
				if ($("#amount2").val() != "")
				{
					var rege2 = /^\d*\.?\d{0,2}$/
					var Ok2 = rege2.exec($("#amount2").val());
					if ( ! Ok2 )
					{
						alert("The Check amount must be a number without any formatting.\nExample: 1234.56\n\nPlease try again.");
						$("#amount2").focus();
						return;
					}
					//alert(document.accountForm.amount2.value);
					if ($("#checkno").val() == "")
					{
						if ( ! confirm('You have entered a check amount without a check number. \nDo you wish to continue?'))
						{
							$("#checkno").focus();
							return;
						}
					}
				}
				// Do we have any cash amount
				if ($("#amount3").val() != "")
				{
					var rege3 = /^\d*\.?\d{0,2}$/
					var Ok3 = rege3.exec($("#amount3").val());
					if ( ! Ok3 )
					{
						alert("The Cash amount must be a number without any formatting.\nExample: 1234.56\n\nPlease try again.");
						$("#amount3").focus();
						return;
					}
					//alert(document.accountForm.amount3.value);
				}
				// Do we have any Other amount
				if ($("#amount8").val() != "")
				{
					var rege8 = /^\d*\.?\d{0,2}$/
					var Ok8 = rege8.exec($("#amount8").val());
					if ( ! Ok8 )
					{
						alert("The Other amount must be a number without any formatting.\nExample: 1234.56\n\nPlease try again.");
						$("#amount8").focus();
						return;
					}
					//alert(document.accountForm.amount8.value);
				}

				// Do we have any Account transfer amount
				if ($("#amount4").length > 0)
				{
					if ($("#amount4").val() != "")
					{
						var rege4 = /^\d*\.?\d{0,2}$/
						var Ok4 = rege4.exec($("#amount4").val());
						if ( ! Ok4 )
						{
							alert("The Citizen Account amount must be a number without any formatting.\nExample: 1234.56\n\nPlease try again.");
							$("#amount4").focus();
							return;
						}

						//  check that the selected account has enough money to cover the transfered amount via Ajax  
						//doAjax('../includes/checkaccountamount.asp', 'amt=' + $("#amount4").val() + '&uid=' + document.accountForm.accountid.options[document.accountForm.accountid.selectedIndex].value, 'AccountCheck', 'get', '0');
						doAjax('../includes/checkaccountamount.asp', 'amt=' + $("#amount4").val() + '&uid=' + $("#accountid").val(), 'AccountCheck', 'get', '0');
					}
					else
					{
						//alert('Submitting');
						document.accountForm.submit();
					}
				}
				else
				{
					//alert('Submitting');
					document.accountForm.submit();
				}
			}


			// Submit form - Moved to AccountCheck()
			//document.accountForm.submit();
		}

	function AccountCheck( check )
	{
		if (check == 'OK')
		{
			//alert(check);
			//alert(document.accountForm.amount4.value);

			document.accountForm.submit(); 
			//alert("Ok to submit.");
		}
		else
		{
			alert( 'The account you selected does not have sufficient funds to cover the transfer.\n\nPlease enter a different method or reduce the amount.');
			$("#amount4").focus();
		}
	}

	function checkAmount()
	{
		//Credit Amounts
		if ($("#amount").val() != "")
		{
			var rege = /^\d*\.?\d{0,2}$/
			var Ok = rege.exec($("#amount").val());
			if ( Ok )
			{
				$("#amount").val( Number($("#amount").val()) );
				$("#amount").val( format_number($("#amount").val(),2) );
			}
			else
			{
				alert('Values should be currency or blank.\nPlease correct this amount.');
				$("#amount").focus();
			}
		}
	}

	function addTotal()
	{
		var total = 0.00;
		//Charge
		if ($("#amount1").val() != "")
		{
			var rege1 = /^\d*\.?\d{0,2}$/
			var Ok1 = rege1.exec($("#amount1").val());
			if ( Ok1 )
			{
				total += Number($("#amount1").val());
				$("#amount1").val( Number($("#amount1").val()) );
				$("#amount1").val( format_number($("#amount1").val(),2) );
			}
			else
			{
				alert('Values should be currency or blank.\nPlease correct this amount.');
				$("#amount1").focus();
			}
		}
		// Check
		if ($("#amount2").val() != "")
		{
			var rege2 = /^\d*\.?\d{0,2}$/
			var Ok2 = rege2.exec($("#amount2").val());
			if ( Ok2 )
			{
				total += Number($("#amount2").val());
				$("#amount2").val( Number($("#amount2").val()) );
				$("#amount2").val( format_number($("#amount2").val(),2) );
			}
			else
			{
				alert('Values should be currency or blank.\nPlease correct this amount.');
				$("#amount2").focus();
			}
		}
		// Cash
		if ($("#amount3").val() != "")
		{
			var rege3 = /^\d*\.?\d{0,2}$/
			var Ok3 = rege3.exec($("#amount3").val());
			if ( Ok3 )
			{
				total += Number($("#amount3").val());
				$("#amount3").val( Number($("#amount3").val()) );
				$("#amount3").val( format_number($("#amount3").val(),2) );
			}
			else
			{
				alert('Values should be currency or blank.\nPlease correct this amount.');
				$("#amount3").focus();
			}
		}
		// Other
		if ($("#amount8").val() != "")
		{
			var rege8 = /^\d*\.?\d{0,2}$/
			var Ok8 = rege8.exec($("#amount8").val());
			if ( Ok8 )
			{
				total += Number($("#amount8").val());
				$("#amount8").val( Number($("#amount8").val()) );
				$("#amount8").val( format_number($("#amount8").val(),2) );
			}
			else
			{
				alert('Values should be currency or blank.\nPlease correct this amount.');
				$("#amount8").focus();	
			}
		}
		// Account transfer 
		if ($("#amount4").length > 0)
		{
			if ($("#amount4").val() != "")
			{
				var rege4 = /^\d*\.?\d{0,2}$/
				var Ok4 = rege4.exec($("#amount4").val());
				if ( Ok4 )
				{
					total += Number($("#amount4").val());
					$("#amount4").val( Number($("#amount4").val()) );
					$("#amount4").val( format_number($("#amount4").val(),2) );
				}
				else
				{
					alert('Values should be currency or blank.\nPlease correct this amount.');
					$("#amount4").focus();
				}
			}
		}

		$("#total").html( format_number(total,2) );
		$("#amount").val( total );
		//alert(document.accountForm.amount.value);
	}

	function toggleAccountDisplay()
	{
		if ($("#paymenttypeid").val() === "4")
		{
			$("#citizenaccountpicks").show();
		}
		else
		{
			$("#citizenaccountpicks").hide();
		}
	}

	//-->
	</script>

</head>

<body>

  <% ShowHeader sLevel %>
  <!--#Include file="../menu/menu.asp"--> 

<div id="content">
	<div id="centercontent">
	
	<font size="+1"><b><%=sCitizenName%> Account: <%=sShowType%></b></font><br /><br />

	<input type="button" class="button" value="<< Back" onclick="location.href='citizen_account_history.asp?u=<%=iUserID%>'" /> 
	<br /><br />

	<h3>
		Current Balance = <%=FormatCurrency(cCurrentBalance,2)%>
	</h3>

	<form method="post" name="accountForm" action="update_citizen_account.asp">
		<input type="hidden" id="uid" name="uid" value="<%=iUserID%>" />
		<input type="hidden" id="entrytype" name="entrytype" value="<%=sEntryType%>" />
		<input type="hidden" id="currentbalance" name="currentbalance" value="<%=cCurrentBalance%>" />

		<table border="0" class="tableadmin" cellpadding="4" cellspacing="0" width="80%">
			<tr><th align="left">Property</th><th align="left">Value</th></tr>

<%				' if withdrawal of funds (debit)
				If sEntryType = "debit" Then %>
				<tr>
					<td class="label" align="right" nowrap="nowrap">
						<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color="red">*</font></span> 
							Amount:
						</span>
					</td>
					<td>
						<span class="cot-text-emphasized" title="This field is required"> 
							<input type="text" value="" id="amount" name="amount" size="20" maxlength="20" onblur="checkAmount()" />
						</span>
					</td>
				</tr>
				<tr>
					<td class="label" align="right" nowrap="nowrap">Issue To:</td>
					<td><%ShowWithdrawTypes( session("orgid") ) %></td>

				</tr>

<%			Else ' Deposit funds (credit)  %>
				<input type="hidden" value="" id="amount" name="amount" id="amount" />
<%				ShowAmounts iUserID
%>
				<tr>
					<td class="label" align="right" nowrap="nowrap">Total:</td><td><span id="total">0.00</span></td>
				</tr>
<%			End If %>

			<tr>
				<td class="label" align="right" valign="top" nowrap="nowrap">
					<span class="cot-text-emphasized"><span class="cot-text-emphasized"></span> 
						Notes:
					</span>
				</td>
				<td>
					<span class="cot-text-emphasized"> 
						<textarea name="notes" id="citizenaccountupdatenotes" class="blockednotes"></textarea>
					</span>
				</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td>
					<input type="button" value="<%=sShowButton%>" class="button" onClick="validate()" />
				</td>
			</tr>
		</table>
	</form>

	</div>
</div>

<!--#Include file="../admin_footer.asp"-->  
  

</body>
</html>

<!--#Include file="inc_dbfunction.asp"-->

<%
'--------------------------------------------------------------------------------------------------
' Sub ShowWithdrawTypes iOrgId 
'--------------------------------------------------------------------------------------------------
Sub ShowWithdrawTypes( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT P.paymenttypeid, P.paymenttypename, P.requirescitizenaccount FROM egov_paymenttypes P, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE P.isforcitizenaccounts = 1 AND P.isCitizenAccountRefundMethod = 1 AND P.paymenttypeid = O.paymenttypeid "
	sSql = sSql & " AND O.orgid = " & iOrgId
	sSql = sSql & " ORDER BY P.displayorder"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		
		response.write vbcrlf & "<select id=""paymenttypeid"" name=""paymenttypeid"" onchange=""toggleAccountDisplay()"">"
		Do While Not oRs.EOF 
			If (Not oRs("requirescitizenaccount")) Or (oRs("requirescitizenaccount") And HasTransferableAccounts( iUserId, "withdrawal") ) Then
				response.write vbcrlf & "<option value=""" & oRs("paymenttypeid") & """"
				If CLng(oRs("paymenttypeid")) = CLng(6) Then
					response.write " selected=""selected"""
				End If 
				response.write ">" & oRs("paymenttypename") & "</option>"
			End If 
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"
	End If 
	
	oRs.CLose
	Set oRs = Nothing 

	If HasTransferableAccounts( iUserId, "withdrawal" ) Then 
		response.write "&nbsp; <span id=""citizenaccountpicks"">To:" & showWithdrawlCitizenAccounts( iUserID ) & "</span>"
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function InsertJournalEntry( iOrgID, iUserID, sAdminUserId, sAmount, iJournalEntryTypeID, sNotes )
'--------------------------------------------------------------------------------------------------
Function InsertJournalEntry( ByVal iOrgID, ByVal iUserID, ByVal sAdminUserId, ByVal sAmount, ByVal iJournalEntryTypeID, ByVal sNotes )
	Dim sSql, oInsert, iJournalId, iAdminLocationId

	iJournalId = 0

	' this is where the admin person is working today
	If Session("LocationId") <> "" Then
		iAdminLocationId = Session("LocationId")
	Else
		iAdminLocationId = 0 
	End If 

	' Note: paymentlocation id is set to 1 so that these always show as happening in the office not the public website
	sSql = "Insert into egov_class_payment (paymentdate, orgid, userid, adminuserid, adminlocationid, paymentlocationid, paymenttotal, journalentrytypeid, notes ) Values (dbo.GetLocalDate(" & iOrgID & ",getdate()), "
	sSql = sSql & iOrgID & ", " & iUserId & ", " & sAdminUserId & ", " & iAdminLocationId & ", 1, " & sAmount & ", " & iJournalEntryTypeID & ", '" & sNotes & "' )"
	sSql = "SET NOCOUNT ON;" & sSql & ";SELECT @@IDENTITY AS ROWID;"
'	response.write sSQL & "<br /><br />"
'	response.End 

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSql, Application("DSN"), 3, 3

	iJournalId = oInsert("ROWID")

	oInsert.close
	Set oInsert = Nothing

	InsertJournalEntry = iJournalId

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetPaymentTypeId( sPaymentType )
'--------------------------------------------------------------------------------------------------
Function GetPaymentTypeId1( ByVal sPaymentType )
	Dim sSql, oEntry, sTypeId

	sSql = "Select paymenttypeid from egov_paymenttypes Where paymenttypename = '" & sPaymentType & "'"

	Set oEntry = Server.CreateObject("ADODB.Recordset")
	oEntry.Open sSql, Application("DSN"), 0, 1

	If Not oEntry.EOF Then 
		sTypeId = oEntry("paymenttypeid") 
	Else 
		sTypeId = 0
	End If 

	oEntry.close
	Set oEntry = Nothing

	GetPaymentTypeId1 = sTypeId

End Function 


'--------------------------------------------------------------------------------------------------
' Sub AdjustCitizenAccount( iUserID, iJournalId, iOrgID, sEntryType, sAmount, iItemTypeId )
'--------------------------------------------------------------------------------------------------
Sub AdjustCitizenAccount( ByVal iUserID, ByVal iJournalId, ByVal iOrgID, ByVal sEntryType, ByVal sAmount, ByVal iItemTypeId )
	Dim oCmd, sNewBalance, iLedgerId, cPriorBalance

	cPriorBalance = GetCitizenCurrentBalance( iUserId )

	' Create the Ledger record
	iLedgerId = InsertLedgerEntry( iOrgID, iUserID, iJournalId, sAmount, iItemTypeId, sEntryType, cPriorBalance )

	If sEntryType = "credit" Then
		sNewBalance = CDbl(cPriorBalance) + CDbl(sAmount)
	Else  ' credit
		sNewBalance = CDbl(cPriorBalance) - CDbl(sAmount)
	End If 

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		' Update the account balance
		.CommandText = "UPDATE egov_users SET accountbalance = " & sNewBalance & " WHERE userid = " & iUserID
		.Execute
	End With

	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function InsertLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, cPriorBalance )
'--------------------------------------------------------------------------------------------------
Function InsertLedgerEntry( ByVal iOrgID, ByVal iAccountId, ByVal iJournalId, ByVal cAmount, ByVal iItemTypeId, ByVal sEntryType, ByVal cPriorBalance )
	Dim sSql, oInsert, iLedgerId, sPlusMinus

	iLedgerId = 0

	If sEntryType = "credit" Then
		sPlusMinus = "'+'"
	Else  ' credit
		sPlusMinus = "'-'"
	End If 

	sSql = "Insert Into egov_accounts_ledger (paymentid,orgid,entrytype,accountid,amount,itemtypeid, priorbalance, plusminus, paymenttypeid) Values ("
	sSql = sSql & iJournalId & ", " & iOrgID & ", '" & sEntryType & "', " & iAccountId & ", " & cAmount & ", " & iItemTypeId & "," & cPriorBalance & ", " & sPlusMinus & ", 4 )"
	sSql = "SET NOCOUNT ON;" & sSql & ";SELECT @@IDENTITY AS ROWID;"
'	response.write sSQL & "<br /><br />"
'	response.End 

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSql, Application("DSN"), 3, 3

	iLedgerId = oInsert("ROWID")

	oInsert.close
	Set oInsert = Nothing

	InsertLedgerEntry = iLedgerId

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowAmounts()
'--------------------------------------------------------------------------------------------------
Sub ShowAmounts( ByVal iUserID )
	Dim sSql, oRs

	sSql = "SELECT paymenttypeid, paymenttypename, requirescheckno, requirescitizenaccount FROM egov_paymenttypes " 
	sSql = sSql & " WHERE isrefundmethod = 0 AND isrefunddebit = 0  AND isadminmethod = 1 "
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF 
		If (Not oRs("requirescitizenaccount")) Or (oRs("requirescitizenaccount") And HasTransferableAccounts( iUserId, "deposit" ) ) Then 
			response.write vbcrlf & "<tr>"
			response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">"
			response.write oRs("paymenttypename") & ": "
			response.write "</td><td>"
			response.write "<input type=""text"" value="""" id=""amount" & oRs("paymenttypeid") & """ name=""amount" & oRs("paymenttypeid") & """ size=""20"" maxlength=""20"" onblur=""addTotal()"" />"
			If oRs("requirescheckno") Then
				response.write " &nbsp;  Check #:<input type=""text"" value="""" id=""checkno"" name=""checkno"" size=""8"" maxlength=""8"" />"
			End If 
			If oRs("requirescitizenaccount") Then
				response.write "&nbsp; From:" & ShowTransferableAccounts( iUserID )
			End If 
			response.write "</td></tr>"
		End If 

		oRs.MoveNext

	Loop
	
	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function  ShowTransferableAccounts( iUserID )
'--------------------------------------------------------------------------------------------------
Function  ShowTransferableAccounts( ByVal iUserID )
	Dim sSql, oAccounts, sText

	sSql = "Select userfname, userlname, userid, accountbalance from egov_users where accountbalance > 0.00 and userid <> " & iUserID & " and familyid = " & GetFamilyId( iUserId ) & " Order by userlname, userfname"

	Set oAccounts = Server.CreateObject("ADODB.Recordset")
	oAccounts.Open sSql, Application("DSN"), 0, 1

	If Not oAccounts.EOF Then 
		sText = vbcrlf & "<select id=""accountid"" name=""accountid"">"
		Do While Not oAccounts.EOF
			sText = sText & vbcrlf & "<option value=""" & oAccounts("userid") & """>" & oAccounts("userfname") & " " & oAccounts("userlname") & " (" & FormatNumber(oAccounts("accountbalance"),2) & ") " & "</option>"
			oAccounts.MoveNext
		Loop 
		sText = sText & vbcrlf & "</select>"
	Else
		sText = "None"
	End If 

	oAccounts.close
	Set oAccounts = Nothing 

	ShowTransferableAccounts = sText
End Function 


'--------------------------------------------------------------------------------------------------
' string showWithdrawlCitizenAccounts( iUserId )
'--------------------------------------------------------------------------------------------------
Function showWithdrawlCitizenAccounts( ByVal iUserId )
	Dim sSql, oRs, sText

	sSql = "SELECT userfname, userlname, userid, ISNULL(accountbalance,0) AS accountbalance FROM egov_users "
	sSql = sSql & "WHERE orgid = " & session( "OrgId" ) & " AND userid <> " & iUserID & " AND familyid = " & GetFamilyId( iUserId ) & " ORDER BY userlname, userfname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		sText = vbcrlf & "<select id=""accountid"" name=""accountid"">"
		Do While Not oRs.EOF
			sText = sText & vbcrlf & "<option value=""" & oRs("userid") & """>" & oRs("userfname") & " " & oRs("userlname") & " (" & FormatNumber(oRs("accountbalance"),2) & ") " & "</option>"
			oRs.MoveNext
		Loop 
		sText = sText & vbcrlf & "</select>"
	Else
		sText = "None"
	End If 

	oRs.Close
	Set oRs = Nothing 

	showWithdrawlCitizenAccounts = sText

End Function 



'--------------------------------------------------------------------------------------------------
' Function HasTransferableAccounts( iUserId, sTransferType )
'--------------------------------------------------------------------------------------------------
Function HasTransferableAccounts( ByVal iUserId, ByVal sTransferType )
	Dim sSql, oAccounts, sWhere

	If LCase(sTransferType) = "deposit" Then
		sWhere = " accountbalance > 0.00 AND "
	Else
		sWhere = ""
	End If 

	sSql = "SELECT Count(userid) AS hits FROM egov_users WHERE " & sWhere & " userid <> " & iUserID & " AND familyid = " & GetFamilyId( iUserId ) 

	Set oAccounts = Server.CreateObject("ADODB.Recordset")
	oAccounts.Open sSql, Application("DSN"), 0, 1

	If clng(oAccounts("hits")) > clng(0) Then 
		HasTransferableAccounts = True 
	Else
		HasTransferableAccounts = False 
	End If 

	oAccounts.close
	Set oAccounts = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Sub InsertPaymentRecord( iPaymentId, iPaymentTypeId, sAmount, sStatus )
'--------------------------------------------------------------------------------------------------
Sub InsertPaymentRecord( ByVal iPaymentId, ByVal iPaymentTypeId, ByVal sAmount, ByVal sStatus )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "Insert Into egov_verisign_payment_information (paymentid, paymenttypeid, amount, paymentstatus) Values (" & iPaymentid & ", " & iPaymentTypeId & ", " & sAmount & ", '" & sStatus & "' )"
		.Execute
	End With
	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub InsertPaymentRecord_Checks( iPaymentId, iPaymentTypeId, sAmount, sStatus, sCheckNo )
'--------------------------------------------------------------------------------------------------
Sub InsertPaymentRecord_Checks( ByVal iPaymentId, ByVal iPaymentTypeId, ByVal sAmount, ByVal sStatus, ByVal sCheckNo )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "Insert Into egov_verisign_payment_information (paymentid, paymenttypeid, amount, paymentstatus, checkno) Values (" & iPaymentid & ", " & iPaymentTypeId & ", " & sAmount & ", '" & sStatus & "', " & sCheckNo & " )"
		.Execute
	End With
	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub InsertPaymentRecord_Transfer( iPaymentId, iPaymentTypeId, sAmount, sStatus, iAccountId )
'--------------------------------------------------------------------------------------------------
Sub InsertPaymentRecord_Transfer( ByVal iPaymentId, ByVal iPaymentTypeId, ByVal sAmount, ByVal sStatus, ByVal iAccountId )
	Dim oCmd

'	response.write "Insert Into egov_verisign_payment_information (paymentid, paymenttypeid, amount, paymentstatus, citizenuserid) Values (" & iPaymentid & ", " & iPaymentTypeId & ", " & sAmount & ", '" & sStatus & "', " & iAccountId & " )"
'	response.End
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "Insert Into egov_verisign_payment_information (paymentid, paymenttypeid, amount, paymentstatus, citizenuserid) Values (" & iPaymentid & ", " & iPaymentTypeId & ", " & sAmount & ", '" & sStatus & "', " & iAccountId & " )"
		.Execute
	End With
	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetJournalEntryTypeID( sType )
'--------------------------------------------------------------------------------------------------
Function GetJournalEntryTypeID( ByVal sType )
	Dim sSql, oEntry, sTypeId

	sSql = "Select journalentrytypeid from egov_journal_entry_types Where journalentrytype = '" & sType & "'"

	Set oEntry = Server.CreateObject("ADODB.Recordset")
	oEntry.Open sSql, Application("DSN"), 0, 1

	If Not oEntry.EOF Then 
		sTypeId = oEntry("journalentrytypeid") 
	Else 
		sTypeId = 0
	End If 

	oEntry.close
	Set oEntry = Nothing

	GetJournalEntryTypeID = sTypeId
End Function


'--------------------------------------------------------------------------------------------------
' Function MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, cPriorBalance, iPriceTypeid )
' This is the same as the one used by classes and events
'--------------------------------------------------------------------------------------------------
Function MakeLedgerEntry( ByVal iOrgID, ByVal iAccountId, ByVal iJournalId, ByVal cAmount, ByVal iItemTypeId, ByVal sEntryType, ByVal sPlusMinus, ByVal iItemId, ByVal iIsPaymentAccount, ByVal iPaymentTypeId, ByVal cPriorBalance, ByVal iPriceTypeid )
	Dim sSql, oInsert, iLedgerId

	iLedgerId = 0

	sSql = "Insert Into egov_accounts_ledger ( paymentid,orgid,entrytype,accountid,amount,itemtypeid,plusminus, "
	sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, pricetypeid ) Values ( "
	sSql = sSql & iJournalId & ", " & iOrgID & ", '" & sEntryType & "', " & iAccountId & ", " & cAmount & ", " & iItemTypeId & ", '" & sPlusMinus & "', " 
	sSql = sSql & iItemId & ", " & iIsPaymentAccount & ", " & iPaymentTypeId & ", " & cPriorBalance & ", " & iPriceTypeid & " )"
	sSql = "SET NOCOUNT ON;" & sSql & ";SELECT @@IDENTITY AS ROWID;"
	response.write sSQL & "<br /><br />"
'	response.End 

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSql, Application("DSN"), 3, 3

	iLedgerId = oInsert("ROWID")

	oInsert.close
	Set oInsert = Nothing

	MakeLedgerEntry = iLedgerId

End Function 


%>
