<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: undo_transaction.asp
' AUTHOR: Steve Loar
' CREATED: 03/12/2014
' COPYRIGHT: Copyright 2014 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This removes a citizen account transaction from the system.
'
' MODIFICATION HISTORY
' 1.0	03/12/2014	Steve Loar - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sSql, oRs, iPaymentId, sEntryType

iPaymentId = CLng(request("paymentId"))

' log who is doing the deletion, orgid, citizen userid for the purchase and when this is happening'
sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'Starting citizen account transaction undo.' )"
RunSQLStatement sSql

' Pull any account ledger rows that might be from citizen account balances applied to the purchase'
sSql = "SELECT ledgerid, accountid, entrytype, ISNULL(amount,0) AS amount FROM egov_accounts_ledger "
sSql = sSql & "WHERE paymenttypeid = 4 AND orgid = " & session("orgid") & " AND paymentid = " & iPaymentId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

Do While Not oRs.EOF
	if oRs("entrytype") = "debit" Then
		' These are the withdrawals and transfers from account, so putting money back'
		sEntryType = "credit"
	Else 
		' These are deposits and transfers to account, so taking money away'
		sEntryType = "debit"
	End If 
	
	' This is in common.asp'
	AdjustCitizenAccountBalance oRs("accountid"), sEntryType, oRs("amount") 
	
	sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'citizen account balance adjusted. userid: " & oRs("accountid") & " entrytype: " & sEntryType & " amount: " & oRs("amount") & "' )"
	RunSQLStatement sSql
	
	oRs.MoveNext
Loop

oRs.Close 
Set oRs = Nothing


' Remove the rows from accounts ledger table'
sSql = "DELETE FROM egov_accounts_ledger WHERE orgid = " & session("orgid") & " AND paymentid = " & iPaymentId
RunSQLStatement sSql

sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'accounts ledger cleared.' )"
RunSQLStatement sSql

' Remove the class payment row and the transaction is now gone'
sSql = "DELETE FROM egov_class_payment WHERE orgid = " & session("orgid") & " AND paymentid = " & iPaymentId
RunSQLStatement sSql

sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'egov_class_payment cleared.' )"
RunSQLStatement sSql

sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'Completed citizen account transaction undo.' )"
RunSQLStatement sSql

response.write "Success"


%>