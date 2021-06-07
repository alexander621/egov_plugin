<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: undo_payment.asp
' AUTHOR: Steve Loar
' CREATED: 03/14/2014
' COPYRIGHT: Copyright 2014 eclink, inc.
'			 All Rights Reserved.
'
' Description: This removes a rental transaction from the system.
'
' MODIFICATION HISTORY
' 1.0	03/14/2014	Steve Loar - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sSql, oRs, iPaymentId, reservationId, tableName, idField, idFieldValue, amount, currentAmount, newAmount, totalPaid
Dim sEntryType, totalRefunded

iPaymentId = CLng(request("paymentId"))
totalPaid = 0.00
totalRefunded = 0.00
totalRefundFees = 0.00

' log who is doing the deletion, orgid, citizen userid for the purchase and when this is happening'
sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'Starting rental transaction undo.' )"
RunSQLStatement sSql


' Pull any account ledger rows that is for citizen account balances'
sSql = "SELECT ledgerid, accountid, entrytype, ISNULL(amount,0) AS amount FROM egov_accounts_ledger "
sSql = sSql & "WHERE paymenttypeid = 4 AND orgid = " & session("orgid") & " AND paymentid = " & iPaymentId
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

Do While Not oRs.EOF
	if oRs("entrytype") = "debit" Then
		' These are the payments from account, so putting money back into the account on the undo'
		sEntryType = "credit"
	Else 
		' These are refunds to account, so taking money away on the undo'
		sEntryType = "debit"
	End If 
	
	' This is in common.asp. It puts the money back in the citizen account when credit is passed'
	AdjustCitizenAccountBalance oRs("accountid"), sEntryType, oRs("amount") 
	
	sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'Citizen account balance adjusted. userid: " & oRs("accountid") & " entrytype: " & sEntryType & " amount: " & oRs("amount") & "' )"
	RunSQLStatement sSql
	
	oRs.MoveNext
Loop

oRs.Close 
Set oRs = Nothing


' Pull out account ledger items that are rental fees and take that away from the paid amount of the fees they were against'
sSql = "SELECT reservationid, entrytype, reservationfeetype, reservationfeetypeid, ISNULL(amount,0) AS amount FROM egov_accounts_ledger "
sSql = sSql & "WHERE reservationfeetype IS NOT NULL AND orgid = " & session("orgid") & " AND paymentid = " & iPaymentId
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

If Not oRs.EOF Then
	reservationId = oRs("reservationid")
	Do While Not oRs.EOF
		Select Case oRs("reservationfeetype")
	  		Case "reservationfeeid"
	  			idField = "reservationfeeid"
	  			tableName = "egov_rentalreservationfees"
	  		Case "reservationdatefeeid"
	  			idField = "reservationdatefeeid"
	  			tableName = "egov_rentalreservationdatefees"
	  		Case "reservationdateitemid"
	  			idField = "reservationdateitemid"
	  			tableName = "egov_rentalreservationdateitems"
	  	End Select
	  	
		idFieldValue = oRs("reservationfeetypeid")
		amount = oRs("amount")
		
		' If it is a credit(undo of payments) you subtract it. Add it on debits(undo of refund)'
		If oRs("entrytype") = "credit" Then 
			currentAmount = GetCurrentPaidAmount(idFieldValue, idField, tableName)		'In rentalscommonfunctions.asp
			'response.write sSql & "<br /><br />"
			newAmount = CDbl(currentAmount) - CDbl(amount )
			setCurrentPaidAmount idFieldValue, idField, tableName, newAmount, iPaymentId
			sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", '" & tableName & " id(" & idFieldValue & ") paidamount set to " & newPaidAmount & "' )"
			RunSQLStatement sSql
		Else 
			currentAmount = GetCurrentRefundAmount(idFieldValue, idField, tableName)	'In rentalscommonfunctions.asp
			'response.write "currentAmount = " & currentAmount & "<br />"
			' need to handle straight refunds and refund fees'
			newAmount = CDbl(currentAmount) - CDbl(amount )
			'response.write "newAmount = " & newAmount & "<br />"
			setCurrentRefundAmount idFieldValue, idField, tableName, newAmount, iPaymentId
			sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", '" & tableName & " id(" & idFieldValue & ") refundamount set to " & newAmount & "' )"
			RunSQLStatement sSql
		End If 

		oRs.MoveNext
	Loop
	
	'Calculate a new paid total 
	totalPaid = CalculateReservationTotal(reservationId, "paidamount" )
	'response.write "totalPaid = " & totalPaid & "<br />"
	totalRefunded = CalculateReservationTotal(reservationId, "refundamount" )
	'response.write "totalRefunded = " & totalRefunded & "<br />"

	' Update the reservation totals'
	sSql = "UPDATE egov_rentalreservations SET totalpaid = " & totalPaid
	sSql = sSql & ", totalrefunded = " & totalRefunded
	sSql = sSql & " WHERE reservationid = " & reservationId
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql

	sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'Reservation total paid set to: " & totalPaid & "' )"
	RunSQLStatement sSql

End If 

oRs.Close 
Set oRs = Nothing

' Remove the rows from accounts ledger table'
sSql = "DELETE FROM egov_accounts_ledger WHERE orgid = " & session("orgid") & " AND paymentid = " & iPaymentId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'accounts ledger cleared.' )"
RunSQLStatement sSql

' Remove rows from the verisign payment table
sSql = "DELETE FROM egov_verisign_payment_information WHERE paymentid = " & iPaymentId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'verisign payments cleared.' )"
RunSQLStatement sSql

' Remove the class payment row and the transaction is now gone'
sSql = "DELETE FROM egov_class_payment WHERE orgid = " & session("orgid") & " AND paymentid = " & iPaymentId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'egov_class_payment cleared.' )"
RunSQLStatement sSql

sSql = "INSERT INTO egov_deleted_transaction_logs (orgid, userid, paymentid, notes) VALUES ( " & session("orgid") & ", " & Session("UserID") & ", " & iPaymentId & ", 'Completed rental payment undo.' )"
RunSQLStatement sSql

response.write "Success"



%>
