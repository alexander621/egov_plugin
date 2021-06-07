<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME:  CHANGE_STATUS.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 04/27/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/27/06   JOHN STULLENBERGER - INITIAL VERSION
' 2.0	04/27/07   Steve Loar  -  Overhauled for Menlo Park Project
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPaymentId, iCitizenId, iAdminUserId, sAmount, iJournalEntryTypeID, sNotes, iMaxPaymentTypes
Dim sCheck, iCitizenAccountId, sPlusMinus, cPriorBalance, iAccountId, iPaymentLocationId, iAdminLocationId
Dim cAmount, iLedgerId, iQuantity, iClasslistId

iPaymentLocationId = request("PaymentLocationId") 

' this is where the admin person is working today
If Session("LocationId") <> "" Then
	iAdminLocationId = Session("LocationId")
Else
	iAdminLocationId = 0 
End If 

iClasslistId = CLng(request("iclasslistid"))

' check that this is not a duplicate payment by checking the status of the student to be WAITLIST
If getEnrollmentStatus( iClasslistId ) <> "WAITLIST" Then
	' Take them to the roster so they can figure it out.'
	response.redirect("view_roster.asp?classid=" & request("classid") & "&timeid=" & request("timeid") )
End If 


iCitizenId = request("iUserId") ' Purchasing citizen (Head of Household)
iAdminUserId = Session("UserID")
response.write "paymenttotal = [" & request("paymenttotal") & "]<br />"
sAmount = CDbl(request("paymenttotal")) ' Payment total
iJournalEntryTypeID = GetJournalEntryTypeID( "purchase" )
sNotes = dbsafe(request("notes"))
iQuantity = request("quantity")


' UPDATE THE PAYMENT TABLE
'iPaymentId = UpdatePaymentInfo( request("classpaymentid"), request("paymenttypeid"), request("paymentlocationid"), request("paymentamount") )

' Insert the egov_class_payment row
iPaymentId = MakeJournalEntry( iPaymentLocationId, iAdminLocationId, iCitizenId, iAdminUserId, sAmount, iJournalEntryTypeID, sNotes, request("oldpaymentid") )

If sAmount > 0 Then ' This is for the Payment Ledger Data
	' Loop through each payment and make a ledger entry and a payment Info row
	' Get the max payment id then loop thru and do inserts for those that have amounts
	iMaxPaymentTypes = GetmaxPaymentTypeId( Session("Orgid") )  ' In common.asp
	x = 1
	Do While x <= iMaxPaymentTypes
		If request("amount" & x) <> "" Then 
			If HasChecks( x ) Then
				' Check
				sCheck = "'" & request("checkno") & "'"
				iCitizenAccountId = "NULL"
				sPlusMinus = "+"
				cPriorBalance = "NULL"
				iAccountId = GetPaymentAccountId( Session("Orgid"), x )
			Else
				sCheck = "NULL"
				If HasCitizensAccounts( x ) Then
					iCitizenAccountId = request("accountid")
					iAccountId = iCitizenAccountId
					sPlusMinus = "-"
					cPriorBalance = GetCitizenCurrentBalance( iCitizenAccountId )
					' Debit the account that was the source of the funds
					AdjustCitizenAccountBalance iCitizenAccountId, "debit", request("amount" & x) 
				Else
					iCitizenAccountId = "NULL"
					sPlusMinus = "+"
					cPriorBalance = "NULL"
					iAccountId = GetPaymentAccountId( Session("Orgid"), x )
					' Charge and Cash
				End If
			End If 
		
			' Make the ledger entry for the payment
			'MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iPriceTypeId )
			iLedgerId = MakeLedgerEntry( Session("Orgid"), iAccountId, iPaymentId, CDbl(request("amount" & x)), "NULL", "debit", sPlusMinus, "NULL", 1, x, cPriorBalance, "NULL" )

			' Make the entry in the egov_verisign_payment_information table
			InsertPaymentInformation iPaymentId, iLedgerId, x, CDbl(request("amount" & x)), "APPROVED", sCheck, iCitizenAccountId
		End If 
		x = x + 1
	Loop 
End If 

iItemTypeId = GetItemTypeId( "recreation activity" )

' Loop through the price fields and process any checked
For Each iPriceTypeId In request("pricetypeid")
	' Pull any accountids - These may not exist, so pull seperately
	iAccountId = GetAccountId( iPriceTypeId, request("classid") )  ' In class_global_functions.asp
	price = request("price" & iPriceTypeId )
	if price = "" then price = "0"
	cAmount = CDbl(price) * CDbl(iQuantity)

	' create the ledger rows for the class accounts
	'MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iPriceTypeId )
	iLedgerId = MakeLedgerEntry( Session("Orgid"), iAccountId, iPaymentId, cAmount, iItemTypeId, "credit", "+", iClasslistId, 0, "NULL", "NULL", iPriceTypeId )
Next  


' UPDATE THE CLASS LIST
UpdateStatus iClasslistId, "ACTIVE", request("paymenttotal"), iPaymentId

' Add to egov_journal_item_status
CreateJournalItemStatus iPaymentId, iItemTypeId, iClasslistId, "ACTIVE", "B"


' UPDATE ENROLLEMENT AND WAITLIST INFORMATION
UpdateEnrollment request("timeid"), iQuantity 

' RETURN TO ROSTER VIEW
' response.redirect("view_roster.asp?classid=" & request("classid") & "&timeid=" & request("timeid") )

' Take them to the receipt'
response.redirect( "view_receipt.asp?iPaymentid=" & iPaymentId )

%>

<!--#Include file="class_global_functions.asp"-->  

<!-- #include file="../includes/common.asp" //-->

<%

'--------------------------------------------------------------------------------------------------
' Function getEnrollmentStatus( iClasslistId )
'--------------------------------------------------------------------------------------------------
Function getEnrollmentStatus( ByVal iClasslistId )
	Dim sSql, oRs, sStatus
	
	sStatus = "NOSTATUS"
	
	sSql = "SELECT ISNULL(status,'NOSTATUS') AS status FROM egov_class_list WHERE classlistid = " & iClasslistId
	'response.write sSql & "<br><br>"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		sStatus = oRs("status")
	End If
	
	oRs.Close
	Set oRs = Nothing
	
	getEnrollmentStatus = sStatus
	
End Function


'--------------------------------------------------------------------------------------------------
' Function MakeJournalEntry( iPaymentLocationId, iAdminLocationId, iCitizenId, iAdminUserId, sAmount, iJournalEntryTypeID, sNotes, iRelatedPaymentId )
'--------------------------------------------------------------------------------------------------
Function MakeJournalEntry( ByVal iPaymentLocationId, ByVal iAdminLocationId, ByVal iCitizenId, ByVal iAdminUserId, ByVal sAmount, ByVal iJournalEntryTypeID, ByVal sNotes, ByVal iRelatedPaymentId )
	Dim sSql

	MakeClassPayment = 0

	sSql = "Insert into egov_class_payment (paymentdate, paymentlocationid, orgid, adminlocationid, "
	sSql = sSql & " userid, adminuserid, paymenttotal, journalentrytypeid, notes, relatedpaymentid) Values (dbo.GetLocalDate(" & Session("orgid") & ",GetDate()), " 
	sSql = sSql & iPaymentLocationId & ", " & Session("orgid") & ", " & iAdminLocationId & ", "
	sSql = sSql & iCitizenId & ", " & iAdminUserId & ", " & sAmount & ", " & iJournalEntryTypeID & ", '" & sNotes & "', " & iRelatedPaymentId & " )"
	'response.write sSQL & "<br /><br />"
	
	MakeJournalEntry = RunInsertStatement( sSql )

End Function 


'--------------------------------------------------------------------------------------------------
' SUB UPDATESTATUS(ICLASSLISTID,SSTATUS,SAMOUNT, iPaymentId)
'--------------------------------------------------------------------------------------------------
Sub UpdateStatus( ByVal iClasslistId, ByVal sStatus, ByVal sAmount, ByVal iPaymentId )
	Dim sSql

	sSql = "Update egov_class_list Set status = '" & sStatus & "', amount = " & sAmount & ", paymentid = " & iPaymentId & " Where classlistid = " & iClasslistId
	
	RunSQLStatement sSql

End Sub


'--------------------------------------------------------------------------------------------------
' SUB UPDATEPAYMENTINFO( IPAYMENTID, IPAYMENTTYPE, IPAYMENTLOCATION, CURPAYMENTTOTAL )
'--------------------------------------------------------------------------------------------------
Sub UpdatePaymentInfo( ByVal iPaymentId, ByVal iPaymentType, ByVal iPaymentLocation, ByVal curPaymentTotal )
	Dim sSql

	sSql = "Update egov_class_payment set paymenttypeid = " & iPaymentType & ", paymentlocationid = " & iPaymentLocation & ",paymenttotal = " & curPaymentTotal & " WHERE paymentid = " & iPaymentId
	
	RunSQLStatement sSql

End Sub 


'--------------------------------------------------------------------------------------------------
' SUB UPDATEENROLLMENT(ICLASSTIMEID, iQuantity)
'--------------------------------------------------------------------------------------------------
Sub UpdateEnrollment( ByVal iclasstimeid, ByVal iQuantity )
	Dim sSql

	sSQL = "UPDATE EGOV_CLASS_TIME SET enrollmentsize = enrollmentsize + " & iQuantity & ", waitlistsize = waitlistsize - " & iQuantity & " WHERE TIMEID = " & iclasstimeid 
	
	RunSQLStatement sSql

End Sub


'--------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'  Make buffer Database 'safe'
'  Useful in building SQL Strings
'    strSQL="SELECT *....WHERE Value='" & DBSafe(strValue) & "';"
'--------------------------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )

	If Not VarType( strDB ) = vbString Then 
		DBsafe = strDB
	Else 
		DBsafe = Replace( strDB, "'", "''" )
	End If 
	
End Function



%>
