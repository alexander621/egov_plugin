<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../classes/class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewjournal.asp
' AUTHOR: Steve Loar
' CREATED: 02/07/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the items for a given purchaseid.
'
' MODIFICATION HISTORY
' 1.0   02/07/07	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPaymentId, iUid, sReceiptType, sJEType, iDisplayType, sBackTo, bShowUnDoBtn
Dim iUserId, nTotal, dPaymentDate, sSql, nRowTotal, bMultiWeeks

response.Expires = 60
response.Expiresabsolute = Now() - 1
response.AddHeader "pragma","no-store"
response.AddHeader "cache-control","private"
response.CacheControl = "no-store" 'HTTP prevent back button

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "registration" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

iPaymentId = CLng(request("pid"))
sReceiptType = request("rt")  ' Debit or Credit
sJEType = request("jet") ' Deposit, Withdrawl, Purchase, Refund
sId = request("it") ' citizen account, classes, pool, facility, gifts
iUid = CLng(request("uid"))

If request("src") = "cah" Then
	sBackTo = "../dirs/citizen_account_history.asp?u=" & iUid
Else
	sBackTo = "javascript:history.back();"
End If 

If sId = "ci" Then  ' Citizen account activities
	If sReceiptType = "c" And sJEType = "d" Then
		' Deposit to account
		iDisplayType = 1
	ElseIf sReceiptType = "d" And sJEType = "d" Then
		' Transfer to another account
		iDisplayType = 2
	ElseIf sReceiptType = "d" And sJEType = "w" Then
		' Withdrawl from account
		iDisplayType = 3
	End If 
Else ' Purchases and refunds
	If sReceiptType = "d" And sJEType = "p" Then
		' Purchase using account funds
		iDisplayType = 4
	ElseIf sReceiptType = "c" And sJEType = "r" Then
		' Refund into account funds
		iDisplayType = 5
	End If 
End If 

bShowUnDoBtn = False  
If IsUnDoBtnDisplayed( iPaymentId ) Then 
	bShowUnDoBtn = True  
	SetUnDoBtnDisplay iPaymentId, False
Else
	If UserHasPermission( Session("UserId"), "undo citizen account transaction" ) Then
		bShowUnDoBtn = True
	End If 
End If 


%>
<html lang="en">
<head>
	<meta charset="UTF-8">

	<title>E-Gov Receipt</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="account_styles.css" />
	<link rel="stylesheet" href="receiptprint.css" media="print" />

	<style>
		#content1 {
			border: 1px solid blue;
			}
		#centercontent1 {
			border: 1px solid red;
			}
		#topright1 {
			background-color: yellow;
			}
		div#receiptdatenumberblock 
		{
			display: block;
			font-weight: bold;
		}
		
	</style>
	
	<script src="https://code.jquery.com/jquery-1.5.min.js"></script>
	
	<script>
	
		var unDoTransaction = function( paymentId ) {
			var okToProceed = confirm("This will permanently remove this tranaction from the system. Do you wish to continue?");
			
			if ( okToProceed ) {
				//alert("Firing off undo script on " + paymentId );
				//return false;
				
				var request = jQuery.ajax({  
					url: "./undo_transaction.asp",  
					type: "POST",  
					dataType: "text",
					data: { 
						paymentId : paymentId
				 	 },  
					contentType: 'application/x-www-form-urlencoded; charset=UTF-8'
				}); 

				request.done( function( data ) { 
					
					if (data === 'Success') {
						alert("This transaction has successfully been removed. \nPrint this receipt if you need to. Once you leave this page, all related information will be gone.");
						$("#undobutton").hide();
					}
					else {
						alert("Failed: This transaction was not successfully removed.");
					}
				});
				
				request.fail( function(jqXHR, textStatus) { 
					alert( "Failed: " + textStatus );
				});
		
			}
		};
	
	</script>
	

</head>
<body>
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="idControls" class="noprint">
	<input type="button" class="button" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
	<input type="button" class="button" value="<< Back" onclick="location.href='<% = sBackTo %>'" />
	
<%	
	If bShowUnDoBtn Then 
		response.write "&nbsp;&nbsp;<input type=""button"" class=""button"" id=""undobutton"" value=""Undo This Transaction"" onclick=""unDoTransaction(" & iPaymentId & ");"" />"
	End If 
%>
</div>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

<%
	ShowReceiptHeader iDisplayType

	response.write "<hr />"
	
	If GetPaymentDetails( iPaymentId, iUserId, nTotal, dPaymentDate ) Then 
		response.write vbcrlf & "<div id=""receiptdatenumberblock"">"
		response.write vbcrlf & "Date: " & FormatDateTime(CDate(dPaymentDate),2) 
		response.write " <span id=""receiptnumberdisplay"">Receipt #: " & iPaymentId & "</span>"
		response.write vbcrlf & "</div>"
		response.write vbcrlf & "<hr />"

		response.write vbcrlf & "<div id=""topright"">"
'		response.write vbcrlf & "<p>Total: " & FormatCurrency(nTotal) & "</p>"
		ShowAccountChange iPaymentId, iUid
		response.write vbcrlf & "</div>"

		response.write ShowUserInfo( iUid )
		response.write "<hr />"

		If iDisplayType = 1 Or iDisplayType = 4 Then 
			' show payment sources for deposits and purchases
			ShowPaymentTypes iPaymentId 
			response.write "<hr />"
		End If 

		response.write "<div id=""transactiontype"">Transaction: "
		response.write getPaymentTypeName( iPaymentId )
		response.write "</div><hr />"

		ShowJournalDetails iPaymentId, iUid, iDisplayType
	Else 
		response.write "<P>No Details could be found.</p>"
	End If 

	response.write "<hr />"
	ShowReceiptFooter iDisplayType


%>

	</div>
</div>
<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' string getPaymentTypeName( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function getPaymentTypeName( ByVal iPaymentId )
	Dim sSql, oRs, sPaymentTypeName

	sSql = "SELECT ISNULL(P.paymenttypename,'') AS paymenttypename FROM egov_accounts_ledger A, egov_paymenttypes P "
	sSql = sSql & "WHERE A.entrytype = 'credit' AND a.ispaymentaccount = 1 "
	sSql = sSql & "AND A.paymenttypeid = P.paymenttypeid AND A.paymentid = " & iPaymentId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		getPaymentTypeName = oRs("paymenttypename")
	Else
		getPaymentTypeName = "Deposit to Account"
	End If
	
	oRs.close
	Set oRs = Nothing

End Function  


'--------------------------------------------------------------------------------------------------
' ShowPaymentTypes iPaymentId 
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentTypes( ByVal iPaymentId )
	Dim sSql, oRs, cTotal

	cTotal = 0.00
	sSql = "SELECT paymenttypeid, paymenttypename, requirescheckno, requirescitizenaccount FROM egov_paymenttypes " 
	sSql = sSql & " WHERE isrefundmethod = 0 AND isrefunddebit = 0 ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write "<table id=""journalreceiptpayments"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<tr>"
			response.write "<td class=""label"" align=""right"" nowrap=""nowrap"" width=""40%"">"
			response.write oRs("paymenttypename") 
			response.write ": &nbsp;</td><td>"
			cAmount = GetAmount( iPaymentId, oRs("paymenttypeid") )
			cTotal = cTotal + cAmount
			response.write FormatCurrency(cAmount, 2)
			If oRs("requirescheckno") Then
				response.write " &nbsp;  Check # " 
				If CDbl(cAmount) > CDbl(0.00) Then 
					response.write GetCheckNo( iPaymentId, oRs("paymenttypeid") )
				End If 
			End If 
			If oRs("requirescitizenaccount") Then
				response.write " &nbsp; From:" 
				If CDbl(cAmount) > CDbl(0.00) Then 
					response.write GetAccountName( iPaymentId, oRs("paymenttypeid") )
				End If 
			End If 
			response.write "</td></tr>"
			oRs.MoveNext
		Loop
		response.write "<tr><td class=""label totalpayment"" align=""right"" nowrap=""nowrap"" width=""40%"">Total: &nbsp;</td><td class=""journaltotalpayment"">" & FormatCurrency(cTotal,2) & "</td><tr>"
		response.write "</table>"
	End If 
	
	oRs.close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' double GetAmount( iPaymentId, iPaymentTypeId )
'--------------------------------------------------------------------------------------------------
Function GetAmount( ByVal iPaymentId, ByVal iPaymentTypeId )
	Dim sSql, oRs, cAmount

	sSql = "SELECT amount FROM egov_verisign_payment_information WHERE paymentid = " & iPaymentId & " AND paymenttypeid = " & iPaymentTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		cAmount = CDbl(oRs("amount"))
	Else
		cAmount = 0.00
	End If 

	oRs.close
	Set oRs = Nothing 

	GetAmount = cAmount

End Function 


'--------------------------------------------------------------------------------------------------
' string GetCheckNo( iPaymentId, iPaymentTypeId )
'--------------------------------------------------------------------------------------------------
Function GetCheckNo( ByVal iPaymentId, ByVal iPaymentTypeId )
	Dim sSql, oRs

	sSql = "SELECT checkno FROM egov_verisign_payment_information WHERE paymentid = " & iPaymentId
	sSql = sSql & " AND paymenttypeid = " & iPaymentTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetCheckNo = oRs("checkno")
	Else
		GetCheckNo = ""
	End If 

	oRs.close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetAccountName( iPaymentId, iPaymentTypeId )
'--------------------------------------------------------------------------------------------------
Function GetAccountName( ByVal iPaymentId, ByVal iPaymentTypeId )
	Dim sSql, oRs

	sSql = "SELECT userfname, userlname FROM egov_verisign_payment_information, egov_users "
	sSql = sSql & "WHERE paymentid = " & iPaymentId & " AND paymenttypeid = " & iPaymentTypeId
	sSql = sSql & " AND citizenuserid = userid "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetAccountName = oRs("userfname") & " " & oRs("userlname")
	Else
		GetAccountName = ""
	End If 

	oRs.close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetTransferedTo( iPaymentId, iUid )
'--------------------------------------------------------------------------------------------------
Function GetTransferedTo( ByVal iPaymentId, ByVal iUid )
	Dim sSql, oRs

	sSql = "SELECT userfname, userlname FROM egov_accounts_ledger, egov_users "
	sSql = sSql & " WHERE accountid = userid AND paymentid = " & iPaymentId & " AND accountid <> " & iUid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetTransferedTo = oRs("userfname") & " " & oRs("userlname")
	Else
		GetTransferedTo = ""
	End If 

	oRs.close
	Set oRs = Nothing 
	
End Function 


'--------------------------------------------------------------------------------------------------
' ShowJournalDetails iPaymentId, iUid, iDisplayType 
'--------------------------------------------------------------------------------------------------
Sub ShowJournalDetails( ByVal iPaymentId, ByVal iUid, ByVal iDisplayType )
	Dim sSql, oRs
	
	' Get the activities that they were part of - Should give 1 row
	sSql = "SELECT P.paymentid, P.paymentdate, L.entrytype, L.amount, U.firstname + ' ' + U.lastname as adminname, P.notes "
	sSql = sSql & "FROM egov_class_payment P, egov_accounts_ledger L, egov_item_types I, users U "
	sSql = sSql & "WHERE P.paymentid = L.paymentid AND L.itemtypeid = I.itemtypeid AND "
	sSql = sSql & "P.adminuserid = U.userid AND L.accountid = " & iUid & " AND P.paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRS.EOF Then 
		Response.Write vbcrlf & "<table border=""0"" cellspacing=""0"" cellpadding=""2"" id=""citizenaccountreceipt"">"
		Response.Write vbcrlf & vbtab & "<tr><th align=""left"">Date</th><th align=""left"">Admin</th>"
		If iDisplayType = 2 Then 
			response.write "<th align=""left"">Transfered To</th>"
		End If 
		response.write "<th align=""left"">Notes</th><th align=""right"">Deposit</th><th align=""right"">Withdrawl</th></tr>"

		' LOOP AND DISPLAY THE RECORDS
		Do While Not oRs.EOF 
			response.Write vbcrlf & "<tr>"
			response.write "<td nowrap>" & FormatDateTime(CDate(oRs("paymentdate")),2) & "</td>"
			response.write "<td nowrap>" & oRs("adminname") & "</td>"
			If iDisplayType = 2 Then 
				' show who got the transfer
				response.write "<td>" & GetTransferedTo( iPaymentId, iUid ) & "</td>"
			End If 
			response.write "<td>" & oRs("notes") & "</td>"
			response.write "<td nowrap align=""right"">"
			If oRs("entrytype") = "credit" Then
				response.write FormatCurrency(oRs("amount"),2) 
			Else
				response.write "&nbsp;"
			End If
			response.write "</td>"
			response.write "<td nowrap align=""right"">" 
			If oRs("entrytype") = "debit" Then
				response.write FormatCurrency(oRs("amount"),2) 
			Else 
				response.write "&nbsp;"
			End If 
			response.write "</td>"
			response.Write vbcrlf & "</tr>"
			oRs.MoveNext
		Loop
		response.write "</table>"
	End If 

	oRs.close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' ShowAccountChange iPaymentId, iUid 
'--------------------------------------------------------------------------------------------------
Sub ShowAccountChange( ByVal iPaymentId, ByVal iUid )
	Dim sSql, oRs, cAmount, cPriorBalance, cCurrentBalance

	' Get the activities that they were part of - Should give 1 row
	sSql = "SELECT entrytype, amount, priorbalance, plusminus "
	sSql = sSql & "FROM egov_accounts_ledger "
	sSql = sSql & "WHERE accountid = " & iUid & " AND paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRS.EOF Then 
		response.write vbcrlf & "<h3>Account Information</h3>"
		response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" id=""accountchange"">"
		cPriorBalance = CDbl(oRs("priorbalance"))
		response.write vbcrlf & "<tr><td nowrap=""nowrap"">Prior Balance.............................</td><td align=""right"">" & FormatCurrency(cPriorBalance,2) & "</td></tr>"
		cAmount = CDbl(oRs("amount"))
		If oRs("entrytype") = "credit" then
			cCurrentBalance = cPriorBalance + cAmount
			cPrefix = oRs("plusminus")
		Else
			cCurrentBalance = cPriorBalance - cAmount
			cPrefix = oRs("plusminus")
		End If 
		
		response.write vbcrlf & "<tr><td nowrap=""nowrap"">Change.....................................</td><td id=""changecell"" align=""right"">" & cPrefix & FormatCurrency(cAmount,2) & "</td></tr>"
		response.write vbcrlf & "<tr><td nowrap=""nowrap"">Current Balance.........................</td><td align=""right"">" & FormatCurrency(cCurrentBalance,2) & "</td></tr>"

		response.write vbcrlf & "</table>"
	End If 

	oRs.close
	Set oRs = Nothing 
	
End Sub 


'--------------------------------------------------------------------------------------------------
' string ShowUserInfo( iUserId )
'--------------------------------------------------------------------------------------------------
Function ShowUserInfo( ByVal iUserId )
	Dim oCmd, sResidentDesc, sUserType, sSql, oUser
	ShowUserInfo = ""

	sUserType = GetUserResidentType(iUserid)
	' If they are not one of these (R, N), we have to figure which they are
	If sUserType <> "R" And sUserType <> "N" Then
		' This leaves E and B - See if they are a resident, also
		sUserType = GetResidentTypeByAddress(iUserid, Session("OrgID"))
	End If 

	sResidentDesc = GetResidentTypeDesc(sUserType)
	
	sSql = "SELECT userfname, userlname, useraddress, useraddress2, userunit, usercity, userstate, "
	sSql = sSql & "userzip, usercountry, useremail, userhomephone, userworkphone, userfax, "
	sSql = sSql & "userbusinessname, userpassword, userregistered, residenttype, registrationblocked, "
	sSql = sSql & "blockeddate, blockedadminid, blockedexternalnote, blockedinternalnote, isdeleted "
	sSql = sSql & "FROM egov_users WHERE userid = " & iUserId

	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open sSql, Application("DSN"), 3, 1

	If Not oUser.EOF Then 

		ShowUserInfo = "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""receiptuserinfo"">"
		ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Name:</td><td nowrap=""nowrap"">" & oUser("userfname") & " " & oUser("userlname")
		If oUser("isdeleted") Then 
			ShowUserInfo = ShowUserInfo & " (deleted)"
		End If 
		ShowUserInfo = ShowUserInfo & "&nbsp;&nbsp;&nbsp;<strong>" & sResidentDesc & "</strong></td></tr>"
		ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Address:</td><td nowrap=""nowrap"">" & oUser("useraddress") 
		If oUser("userunit") <> "" Then 
			ShowUserInfo = ShowUserInfo & "&nbsp;&nbsp;" & oUser("userunit") 
		End If
		If oUser("useraddress2") = "" Then 
			ShowUserInfo = ShowUserInfo & "<br />" & oUser("useraddress2") 
		End If 
		ShowUserInfo = ShowUserInfo & "<br />" & oUser("usercity") & ", " & oUser("userstate") & " " & oUser("userzip") & "</td></tr>"
		ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Email:</td><td>" & GetFamilyEmail( iUserId ) & "</td></tr>"
		ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Phone:</td><td>" & FormatPhone(oUser("userhomephone")) & "</td></tr>"
		'ShowUserInfo = ShowUserInfo & "<tr><td width=""85"" align=""right"" valign=""top"">Business:</td><td>" & oUser("userbusinessname") & "</td></tr>"
		ShowUserInfo = ShowUserInfo & "</table>"
	End If 

	oUser.Close
	Set oUser = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' boolean GetPaymentDetails( iPaymentId, ByRef iUserId, ByRef nTotal, ByRef dPaymentDate )
'--------------------------------------------------------------------------------------------------
Function GetPaymentDetails( ByVal iPaymentId, ByRef iUserId, ByRef nTotal, ByRef dPaymentDate )
	Dim sSql, oRs

	sSql = "SELECT userid, paymenttotal, paymentdate FROM egov_class_payment WHERE paymentid = " & iPaymentId & " AND orgid = " & Session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		iUserId = oRs("userid")
		nTotal = oRs("paymenttotal")
		dPaymentDate = oRs("paymentdate")
		GetPaymentDetails = True 
	Else 
		GetPaymentDetails = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function  


'--------------------------------------------------------------------------------------------------
' ShowReceiptHeader iDisplayType 
'--------------------------------------------------------------------------------------------------
Sub ShowReceiptHeader( ByVal iDisplayType )

	Select Case iDisplayType
		Case 1   ' Deposit to account
			If OrgHasDisplay( Session("OrgID"), "receipt header" ) Then
				response.write "<p class=""receiptheader"">" & GetOrgDisplay( Session("OrgID"), "receipt header" ) 
				response.write "<br /><br />Account Deposit Receipt"
				response.write "</p>"
			Else  
				response.write "<h3>" & Session("sOrgName") & " Deposit Receipt</h3><br /><br />"
			End If 
		Case 2  ' Transfer to another account
			If OrgHasDisplay( Session("OrgID"), "refund header" ) Then
				response.write "<p class=""receiptheader"">" & GetOrgDisplay( Session("OrgID"), "refund header" ) 
				response.write "<br /><br />Transfer Receipt"
				response.write "</p>"
			Else  
				response.write "<h3>" & Session("sOrgName") & " Transfer Receipt</h3><br /><br />"
			End If 
		Case 3  ' Withdrawl
			If OrgHasDisplay( Session("OrgID"), "refund header" ) Then
				response.write "<p class=""receiptheader"">" & GetOrgDisplay( Session("OrgID"), "refund header" ) 
				response.write "<br /><br />Account Withdrawl Receipt"
				response.write "</p>"
			Else  
				response.write "<h3>" & Session("sOrgName") & " Withdrawl Receipt</h3><br /><br />"
			End If 
		Case 4   ' Purchase
			If OrgHasDisplay( Session("OrgID"), "receipt header" ) Then
				response.write "<p class=""receiptheader"">" & GetOrgDisplay( Session("OrgID"), "receipt header" ) 
				response.write "<br /><br />Purchase Receipt"
				response.write "</p>"
			Else  
				response.write "<h3>" & Session("sOrgName") & " Purchase Receipt</h3><br /><br />"
			End If 
		Case 5   ' Refund voucher??
			If OrgHasDisplay( Session("OrgID"), "refund header" ) Then
				response.write "<p class=""receiptheader"">" & GetOrgDisplay( Session("OrgID"), "refund header" ) 
				response.write "<br /><br />Purchase Refund Voucher"
				response.write "</p>"
			Else  
				response.write "<h3>" & Session("sOrgName") & " Purchase Refund Voucher</h3><br /><br />"
			End If 
	End Select 

End Sub 


'--------------------------------------------------------------------------------------------------
' ShowReceiptFooter iDisplayType 
'--------------------------------------------------------------------------------------------------
Sub ShowReceiptFooter( ByVal iDisplayType )

	If clng(iDisplayType) = clng(1) Or clng(iDisplayType) = clng(4) Then  ' Deposit and Purchase
		If OrgHasDisplay( Session("OrgID"), "receipt footer" ) Then
			response.write "<p>" & GetOrgDisplay( Session("OrgID"), "receipt footer" ) & "</p>"
		End If 
	ElseIf iDisplayType <> clng(2) Then   ' Withdrawl and Refund
		If OrgHasDisplay( Session("OrgID"), "refund footer" ) Then
			response.write "<p>" & GetOrgDisplay( Session("OrgID"), "refund footer" ) & "</p>"
		End If 
	End If 

End Sub 


%>

