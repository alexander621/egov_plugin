<!-- #include file="../includes/common.asp" //-->
<!--#Include file="class_global_functions.asp"-->  
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: undo_purchase.asp
' AUTHOR: Steve Loar
' CREATED: 03/10/2014
' COPYRIGHT: Copyright 2014 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This removes a purchase transaction from the system.
'
' MODIFICATION HISTORY
' 1.0	03/10/2014	Steve Loar - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sSql, oRs, iPaymentId, buyWait, updateClassCount

iPaymentId = CLng(request("paymentId"))

' log who is doing the deletion, orgid, citizen userid for the purchase and when this is happening'
sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'Starting class purchase undo.' )"
RunSQLStatement sSql

' Remove rows from the verisign payment table
sSql = "DELETE FROM egov_verisign_payment_information WHERE paymentid = " & iPaymentId
RunSQLStatement sSql

sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'verisign payments cleared.' )"
RunSQLStatement sSql

' Remove rows from the egov_journal_item_status table
sSql = "DELETE FROM egov_journal_item_status WHERE paymentid = " & iPaymentId
RunSQLStatement sSql

sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'egov_journal_item_status cleared.' )"
RunSQLStatement sSql

' Pull the rows from the class list that have that payment id'
sSql = "SELECT classlistid, classtimeid, ISNULL(quantity,0) AS quantity, status "
sSql = sSql & "FROM egov_class_list WHERE paymentid = " & iPaymentId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

Do While Not oRs.EOF
	Select Case UCASE(oRs("status"))
		Case "ACTIVE"
			buyWait = "B"
			updateClassCount = true 
		Case "DROPIN"
			buyWait = "B"
			updateClassCount = true
		Case "DROPPED"
			updateClassCount = false 
		Case "WAITLIST"
			buyWait = "W"
			updateClassCount = true
		Case "WAITLIST REMOVED"
			updateClassCount = false
	End Select
	
	If updateClassCount Then 
		' update the enrollment. This is in class_global_functions.asp
		UpdateClassTime oRs("classtimeid"), -CLng(oRs("quantity")), buyWait
	End If
	
	' delete the row from the class list'
	sSql = "DELETE FROM egov_class_list WHERE classlistid = " & oRs("classlistid")
	RunSQLStatement sSql
	
	sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'egov_class_list cleared. classtimeid: " & oRs("classtimeid") & " qty: " & oRs("quantity") & " status: " & oRs("status") & "' )"
	RunSQLStatement sSql
	
	oRs.MoveNext
Loop

oRs.Close 
Set oRs = Nothing



' Pull any account ledger rows that might be from citizen account balances applied to the purchase'
sSql = "SELECT ledgerid, accountid, ISNULL(amount,0) AS amount FROM egov_accounts_ledger "
sSql = sSql & "WHERE paymenttypeid = 4 AND orgid = " & session("orgid") & " AND paymentid = " & iPaymentId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

Do While Not oRs.EOF
	' This is in common.asp'
	AdjustCitizenAccountBalance oRs("accountid"), "credit", oRs("amount") 
	
	sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'citizen account balance adjusted. userid: " & oRs("accountid") & " amount: " & oRs("amount") & "' )"
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

sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'Completed class purchase undo.' )"
RunSQLStatement sSql

response.write "Success"


%>