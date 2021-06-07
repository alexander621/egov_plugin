<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: DROP_REGISTRANT.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 04/19/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   	04/19/06 	JOHN STULLENBERGER - INITIAL VERSION
' 2.0		06/05/2007	Steve Loar - Overhauled for Menlo Park Project
' 2.1 	11/06/2013	Steve Loar - Added drop reason
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sStatus, iCitizenId, iAdminUserId, iJournalEntryTypeID, sNotes, iAdminLocationId, sTotalRefund, iPaymentId
Dim iLedgerId, cAmount, iAccountId, iClassListId, iPaymentLocationId, sTotalPaid, sRefundFee, oRs, iPriceTypeId
Dim iDropReasonId, iOldPaymentId

If request("iclasslistid") <> "" Then 
	iClassListId = CLng(request("iclasslistid"))
Else
	iClassListId = CLng(0)
End If 

sStatus = GetRegistrantStatus( iClassListId )

sStatus = UCase(sStatus)

If sStatus = "ACTIVE" or sStatus = "DROPIN" Then  ' This should be ACTIVE or DROPIN status only for dropping, otherwise just take them to the class list page

	' this is where the admin person is working today
	If Session("LocationId") <> "" Then
		iAdminLocationId = Session("LocationId")
	Else
		iAdminLocationId = 0 
	End If 

	iCitizenId = request("iUserId") ' Purchasing citizen (Head of Household)
	iAdminUserId = Session("UserID")
	'response.write "amount = [" & request("totalprice") & "]<br />"
	sTotalPaid = CDbl(request("totalpaid"))
	sRefundFee = CDbl(request("refundfee"))
	sTotalRefund = CDbl(request("totalrefund")) ' Refund total
	iJournalEntryTypeID = GetJournalEntryTypeID( "refund" )
	sNotes = dbsafe(request("notes"))
	iItemTypeId = GetItemTypeId( "recreation activity" ) ' This is the kind of thing they bought - in class_global_functions.asp
	iClassListId = CLng(request("iclasslistid"))
	iPaymentLocationId = request("PaymentLocationId")
	iDropReasonId = clng(request("dropreasonid"))
	iOldPaymentId = CLng(request("oldpaymentid"))

	' Create a Journal entry for the refund
	iPaymentId = MakeJournalEntry( iPaymentLocationId, iAdminLocationId, iCitizenId, iAdminUserId, sTotalPaid, iJournalEntryTypeID, sNotes, iOldPaymentId, iDropReasonId )
	'response.write "iPaymentId = " & iPaymentId & "<br /><br />"

	sSql = "SELECT * FROM egov_accounts_ledger WHERE ispaymentaccount = 0 AND itemid = " & iClassListId & " AND paymentid = " & iOldPaymentId
	'response.write sSql & "<br /><br />"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	' Create Ledger entries for debiting (-) the class accounts
	Do While Not oRs.EOF
		If IsNull(oRs("accountid")) Then 
			iAccountId = 0
		Else 
			iAccountId = oRs("accountid")
		End If 
		
		iItemTypeId = oRs("itemtypeid")
		iClassListId = oRs("itemid")
		iPriceTypeId = oRs("pricetypeid")
		'cAmount = CDbl(oRs("amount"))
		cAmount = CDbl(request("amount" & iPriceTypeId ))

		iLedgerId = MakeLedgerEntry( Session("OrgId"), iAccountId, iPaymentId, cAmount, iItemTypeId, "debit", "-", iClassListId, 0, "NULL", "NULL", iPriceTypeId )
		'response.write "class accounts iLedgerId = " & iLedgerId & "<br /><br />"
		oRs.MoveNext
	Loop 
	oRs.Close
	Set oRs = Nothing 

	If CLng(request("accountid")) = CLng(0) Then
		' if Refund voucher was selected create a ledger row for that
		iPaymentTypeId = GetRefundPaymentTypeId( )  ' In common.asp
		iAccountId = GetPaymentAccountId( Session("OrgId"), iPaymentTypeId )  ' In common.asp
		
		If request("isccrefund") <> "" Then
			iIsCCRefund = request("isccrefund")
		Else
			iIsCCRefund = 0
		End If 

		iLedgerId = MakeRefundLedgerEntry( Session("OrgId"), iAccountId, iPaymentId, sTotalRefund, iItemTypeId, "credit", "-", iClassListId, 1, iPaymentTypeId, "NULL", iIsCCRefund )
		'response.write "refund voucher iLedgerId = " & iLedgerId & "<br /><br />"
	Else 
		' Else Create Ledger entry for crediting (-) user account 
		iAccountId = CLng(request("accountid"))
		cPriorBalance = GetCitizenCurrentBalance( iAccountId )

		' Credit the Citizen account that is to get the refund
		AdjustCitizenAccountBalance iAccountId, "credit", sTotalRefund
		'response.write "AdjustCitizenAccountBalance<br /><br />"

		iLedgerId = MakeLedgerEntry( Session("OrgId"), iAccountId, iPaymentId, sTotalRefund, iItemTypeId, "credit", "-", iClassListId, 0, 4, cPriorBalance, "NULL" )
		'response.write "citizen account iLedgerId = " & iLedgerId & "<br /><br />"
	End If 

	' Handle the refund fee amount
	iAccountId = GetRefundDebitAccountId( )  ' In common.asp
	'response.write "iAccountId = " & iAccountId & "<br /><br />"

	iLedgerId = MakeLedgerEntry( Session("OrgId"), iAccountId, iPaymentId, CDbl(request("refundfee")), iItemTypeId, "credit", "+", iClassListId, 1, request("refundfeeid"), "NULL", "NULL" )
	'response.write "refund fee iLedgerId = " & iLedgerId & "<br /><br />"

	' UPDATE THEIR STATUS
	DropRegistrant request("iclasslistid"), sTotalRefund, iPaymentId
	'response.write "DropRegistrant<br /><br />"

	' Add to egov_journal_item_status
	CreateJournalItemStatus iPaymentId, iItemTypeId, iClassListId, "DROPPED", "D"
	'response.write "CreateJournalItemStatus<br /><br />"

	' UPDATE THE ENROLLMENT COUNT
	UpdateEnrollment request("timeid"), sStatus, request("quantity")
	'response.write "UpdateEnrollment<br /><br />"

	' If there are child classes and events fix their counts and update the status also
	UpdateChildEnrollment request("classid"),  request("quantity"), request("oldpaymentid")
	'response.write "UpdateChildEnrollment<br /><br />"

End If 

'response.write "All Done<br /><br />"
' RETURN TO ROSTER VIEW
response.redirect("view_roster.asp?classid=" & request("classid") & "&timeid=" & request("timeid") )

%>

<!--#Include file="class_global_functions.asp"-->  
<!-- #include file="../includes/common.asp" //-->

<%

'--------------------------------------------------------------------------------------------------
' DropRegistrant iclasslistid, curRefundAmount, iPaymentId 
'--------------------------------------------------------------------------------------------------
Sub DropRegistrant( ByVal iclasslistid, ByVal curRefundAmount, ByVal iPaymentId )
	Dim sSql
	
	' SET DEFAULT REFUND AMOUNT TO ZERO
	If curRefundAmount = "" Then
		curRefundAmount = 0
	End If

	sSql = "UPDATE egov_class_list SET status = 'DROPPED', refundamount=" & curRefundAmount & ", paymentid = " & iPaymentId
	sSql = sSql & " WHERE classlistid = " & iclasslistid 
	'response.write sSql & "<br /><br />"
	
	RunSQLStatement sSql

End Sub


'--------------------------------------------------------------------------------------------------
' UpdateEnrollment iclasstimeid, sStatus, iQty 
'--------------------------------------------------------------------------------------------------
Sub UpdateEnrollment( ByVal iclasstimeid, ByVal sStatus, ByVal iQty )
	Dim sSql, sListType

	If sStatus = "ACTIVE" Then 
		sListType = "enrollmentsize"
	Else
		sListType = "waitlistsize"
	End If 

	sSql = "UPDATE EGOV_CLASS_TIME SET " & sListType & " = " & sListType & " - " & clng(iQty) & " WHERE TIMEID = " & iclasstimeid
	'response.write sSql & "<br /><br />"

	RunSQLStatement sSql

End Sub


'--------------------------------------------------------------------------------------------------
' string GetRegistrantStatus( iclasslistid )
'--------------------------------------------------------------------------------------------------
Function GetRegistrantStatus( ByVal iclasslistid )
	Dim sSql, oRs

	GetRegistrantStatus = "UNKNOWN"
	sSql = "SELECT status FROM egov_class_list WHERE classlistid = "  & CLng( iclasslistid )
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetRegistrantStatus = Trim(UCase(oRs("status")))
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' integer MakeJournalEntry( iPaymentLocationId, iAdminLocationId, iCitizenId, iAdminUserId, sAmount, iJournalEntryTypeID, sNotes, iRelatedPaymentId )
'--------------------------------------------------------------------------------------------------
Function MakeJournalEntry( ByVal iPaymentLocationId, ByVal iAdminLocationId, ByVal iCitizenId, ByVal iAdminUserId, ByVal sAmount, ByVal iJournalEntryTypeID, ByVal sNotes, ByVal iRelatedPaymentId, ByVal iDropReasonId )
	Dim sSql

	MakeClassPayment = 0

	sSql = "INSERT INTO egov_class_payment ( paymentdate, paymentlocationid, orgid, adminlocationid, "
	sSql = sSql & "userid, adminuserid, paymenttotal, journalentrytypeid, notes, relatedpaymentid, "
	sSql = sSql & "dropreasonid ) VALUES ( dbo.GetLocalDate(" & Session("orgid") & ",GETDATE()), " 
	sSql = sSql & iPaymentLocationId & ", " & Session("orgid") & ", " & iAdminLocationId & ", "
	sSql = sSql & iCitizenId & ", " & iAdminUserId & ", " & sAmount & ", " & iJournalEntryTypeID & ", '" & sNotes & "', " 
	sSql = sSql & iRelatedPaymentId & ", " & iDropReasonId & " )"
	'response.write sSql & "<br /><br />"
	
	session( "JournalEntrySQL" ) = sSql
	MakeJournalEntry  = RunIdentityInsertStatement( sSql )
	session( "JournalEntrySQL" ) = ""

End Function 


'--------------------------------------------------------------------------------------------------
' void UpdateChildEnrollment( iClassid,  iQuantity, iPaymentid )
'--------------------------------------------------------------------------------------------------
Sub UpdateChildEnrollment( ByVal iClassid,  ByVal iQuantity, ByVal iPaymentid )
	Dim sSql, oRs

	sSql = "Select C.classid, T.timeid,  CL.classlistid From egov_class C, egov_class_time T, egov_class_list CL "
	sSql = sSql & " Where C.classid = T.classid and CL.classid = C.classid and CL.classtimeid = T.timeid "
	sSql = sSql & " and C.parentclassid = " & iClassid & " and CL.paymentid = " & iPaymentid
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		' Change the status of the enrollee to DROPPED
		DropRegistrant oRs("classlistid"), 0.00, iPaymentId

		' Update the enrollment of the child time
		UpdateEnrollment oRs("timeid"), "ACTIVE", iQuantity
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


%>
