<!-- #include file="../includes/common.asp" //-->
<!--#Include file="class_global_functions.asp"-->  
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: dropinprocessing.asp
' AUTHOR: Steve Loar
' CREATED: 07/14/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This processes drop-in registrations from ID cards.
'
' MODIFICATION HISTORY
' 1.0	07/14/2011	Steve Loar - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sSql, iCitizenUserId, dblAccountBalance, iActivityCount, iPaymentLocationId, iAdminUserId, iAdminLocationId, x
Dim dblPurchaseTotal, sPurchaseNotes, iJournalEntryTypeID, iPaymentId, iPaymentTypeId, iLedgerId
Dim iFamilymemberId, iClassId, iTimeId, iClassListId, dAmount, iItemTypeId, iAccountId
Dim iPriceTypeId

iCitizenUserId = CLng(request("userid"))

iFamilymemberId = GetCitizenFamilyId( iCitizenUserId )	' In class_global_functions

dblAccountBalance = GetCitizenCurrentBalance( iCitizenUserId ) ' in common.asp

iActivityCount = clng(request("activitycount"))

sPurchaseNotes = DBsafe(Trim(request("purchasenotes")))

iPaymentLocationId = GetPaymentLocationId( "Walk In" )	' In class_global_functions

iItemTypeId = GetItemTypeId( "recreation activity" )   	' In class_global_functions

' This is where the admin person is working today
If session("LocationId") <> "" Then 
	iAdminLocationId = session("LocationId")
Else 
	iAdminLocationId = 0 
End If 

iAdminUserId = Session("UserID")

dblPurchaseTotal = CDbl(0.00)
For x = 1 To iActivityCount
	If request("activity" & x) = "on" Then
		dblPurchaseTotal = dblPurchaseTotal + CDbl(request("amount" & x))
	End If 
Next 

iJournalEntryTypeID = GetJournalEntryTypeID( "purchase" ) 	' In class_global_functions

iPaymentTypeId = GetPaymentTypeId( "Citizen Accounts" ) 	' In class_global_functions

'response.write "iCitizenUserId = " & iCitizenUserId & "<br />"
'response.write "dblPurchaseTotal = " & dblPurchaseTotal & "<br />"

'Insert the egov_class_payment row (the Journal entry)
iPaymentId = MakeJournalEntry( iPaymentLocationId, iAdminLocationId, iCitizenUserId, iAdminUserId, dblPurchaseTotal, iJournalEntryTypeID, sPurchaseNotes )

'response.write "iPaymentId = " & iPaymentId & "<br />"

'Debit the account that was the source of the funds - in common.asp
AdjustCitizenAccountBalance iCitizenUserId, "debit", dblPurchaseTotal


'Make the ledger entry for the payment - In class_global_functions.asp 
' MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iPriceTypeid )
iLedgerId = MakeLedgerEntry( Session("Orgid"), iCitizenUserId, iPaymentId, dblPurchaseTotal, "NULL", "debit", "-", "NULL", 1, iPaymentTypeId, dblAccountBalance, "NULL" )

'Make the entry in the egov_verisign_payment_information table - This is in ../includes/common.asp
InsertPaymentInformation iPaymentId, iLedgerId, iPaymentTypeId, dblPurchaseTotal, "APPROVED", "NULL", iCitizenUserId

' Loop through the selected activities and add them
For x = 1 To iActivityCount
	If request("activity" & x) = "on" Then

		iClassId = CLng(request("classid" & x))
		iTimeId = CLng(request("timeid" & x))
		dAmount = CDbl(request("amount" & x))
		iAccountId = CLng(request("accountid" & x))
		iPriceTypeId = CLng(request("pricetypeid" & x))

		'Add to the Class List
		iClassListId = AddToClassList( iCitizenUserId, iClassId, iTimeId, iFamilymemberId, dAmount, iPaymentId, iCitizenUserId )

		'Add to egov_journal_item_status  - In class_global_functions.asp 
		CreateJournalItemStatus iPaymentId, iItemTypeId, iClassListId, "DROPIN", "B"

		' Put in a ledger row for each activity selected  - In class_global_functions.asp 
		iLedgerId = MakeLedgerEntry( Session("Orgid"), iAccountId, iPaymentId, dAmount, iItemTypeId, "credit", "+", iClassListId, 0, "NULL", "NULL", iPriceTypeId )
	End If 
Next 

'response.write "<br />Completed.<br />"

' see if the org has the undo feature and set the session variable'
If OrgHasFeature("undo on receipt") Then
	' In ../includes/common.asp'
	SetUnDoBtnDisplay iPaymentId, True
End If 

'take them to the receipt viewing page
response.redirect "view_receipt.asp?iPaymentId=" & iPaymentId & "&return=1"


'------------------------------------------------------------------------------
' integer AddToClassList( iUserId, iClassId, iTimeId, iFamilymemberId, fAmount, iPaymentId, iAttendeeUserId )
'------------------------------------------------------------------------------
Function AddToClassList( ByVal iUserId, ByVal iClassId, ByVal iTimeId, ByVal iFamilymemberId, ByVal fAmount, ByVal iPaymentId, ByVal iAttendeeUserId )
	Dim sSql, oInsert

	AddToClassList = 0

	sSql = "INSERT INTO egov_class_list ( userid, classid, status, "
	sSql = sSql & "quantity, classtimeid, familymemberid, amount, "
	sSql = sSql & "paymentid, attendeeuserid, isdropin, dropindate "
	sSql = sSql & ") VALUES ( " 
	sSql = sSql & iUserId & ", "
	sSql = sSql & iClassId & ", "
	sSql = sSql & "'DROPIN', "
	sSql = sSql & "1, "
	sSql = sSql & iTimeId & ", "
	sSql = sSql & iFamilymemberId & ", "
	sSql = sSql & fAmount & ", "
	sSql = sSql & iPaymentId & ", "
	sSql = sSql & iAttendeeUserId & ", "
	sSql = sSql & "1, "
	sSql = sSql & "'" & DateValue(Now()) & "' "
	sSql = sSql & " )"

	'response.write sSql & "<br /><br />" & vbcrlf

	AddToClassList = RunInsertStatement( sSql )

End Function 

%>
