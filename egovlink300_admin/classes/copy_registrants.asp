<!-- #include file="../includes/common.asp" //-->

<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME:  copy_registrants.asp
' AUTHOR: Steve Loar
' CREATED: 11/04/2013
' COPYRIGHT: Copyright 2013 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This copies registrants from one class to another class as 'Waitlist' status.
'
' MODIFICATION HISTORY
' 1.0   11/04/2013   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iClassListId, iCurrentTimeId, iNewTimeId, iQuantity, iNewClassId, iItemTypeId, iAdminLocationId
Dim iJournalEntryTypeId, iAdminUserId, iPaymentLocationId, iPaymentId, iNewClassListId
Dim iUserId, iFamilyMemberId, iAttendeeUserId, iLedgerId

iCurrentTimeId = CLng(Request("timeid"))
'response.write "iCurrentTimeId: " & iCurrentTimeId & "<br />"

iNewTimeId = CLng(Request("classtimeid"))
'response.write "iNewTimeId: " & iNewTimeId & "<br />"

' get the new classid 
iNewClassId = getNewClassId( iNewTimeId ) 
'response.write "iNewClassId: " & iNewClassId & "<br />"

' get the payment location
iPaymentLocationId = 1  	' This is "Walk In", which no one is doing, but you gotta have something.

' get the itemtypeid 
iItemTypeId = getItemTypeId("recreation activity") 
'response.write "iItemTypeId: " & iItemTypeId & "<br />"

iAdminUserId = Session("UserID")
'response.write "iAdminUserId: " & iAdminUserId & "<br />"

'this is where the admin person is working today
If session("LocationId") <> "" Then 
	iAdminLocationId = session("LocationId")
Else 
	iAdminLocationId = 0 
End If 
'response.write "iAdminLocationId: " & iAdminLocationId & "<br />"

iJournalEntryTypeId = GetJournalEntryTypeID( "purchase" )	' in class_global_functions
'response.write "iJournalEntryTypeId: " & iJournalEntryTypeId & "<br />"

' Loop through the checked classlistids - listcheck
For Each iClassListId In request("classlistid")
	'response.write "iClassListId: " & iClassListId & "<br />"
	' Get the citizenid, attendeeuserid, familymenberid, quantity for this classlistid
	GetClassListInfo iClassListId, iUserId, iFamilyMemberId, iAttendeeUserId, iQuantity		' need to make this
	'response.write "iUserId: " & iUserId & "<br />"
	'response.write "iFamilyMemberId: " & iFamilyMemberId & "<br />"
	'response.write "iAttendeeUserId: " & iAttendeeUserId & "<br />"
	'response.write "iQuantity: " & iQuantity & "<br />"

	' insert the egov_class_payment row
	iPaymentId = MakeJournalEntry( iPaymentLocationId, iAdminLocationId, iUserId, iAdminUserId, "0.00", iJournalEntryTypeID, "Copied to Waitlist." )  ' in class_global_functions
	'response.write "iPaymentId: " & iPaymentId & "<br />"

	' insert the egov_class_list row
	iNewClassListId = AddToClassList( iUserId, iNewClassId, "WAITLIST", iQuantity, iNewTimeId, iFamilyMemberId, "0.00", iPaymentId, iAttendeeUserId )
	'response.write "iNewClassListId: " & iNewClassListId & "<br />"

	'Add to egov_journal_item_status
	CreateJournalItemStatus iPaymentId, iItemTypeId, iNewClassListId, "WAITLIST", "W"	' in class_global_functions

	' Update the class count for waitlist
	UpdateWaitCount iNewTimeId, "+", "waitlistsize", iQuantity

	' insert the egov_account_ledger row
	iLedgerId = MakeLedgerEntry( Session("Orgid"), "NULL", iPaymentId, CDbl(0), iItemTypeId, "credit", "+", iNewClassListId, 0, "NULL", "NULL", 0 )	' in class_global_functions
	'response.write "iLedgerId: " & iLedgerId & "<br />"

Next 

' Go to the target class roster
response.redirect("view_roster.asp?classid=" & iNewClassId & "&timeid=" & iNewTimeId )



'--------------------------------------------------------------------------------------------------
' integer getNewClassId( iTimeId )
'--------------------------------------------------------------------------------------------------
Function getNewClassId( ByVal iTimeId )
	Dim sSql, oRs

	sSql = "SELECT classid FROM egov_class_time WHERE timeid = " & iTimeId
	'response.write sSql & "<br />" & vbcrlf

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		getNewClassId = oRs("classid")
	Else
		getNewClassId = 0
	End If

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' integer getItemTypeId( sItemType )
'--------------------------------------------------------------------------------------------------
Function getItemTypeId( ByVal sItemType )
	Dim sSql, oRs

	sSql = "SELECT itemtypeid FROM egov_item_types WHERE itemtype = '" & sItemType & "'"
	'response.write sSql & "<br />" & vbcrlf

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		getItemTypeId = oRs("itemtypeid")
	Else
		getItemTypeId = 0
	End If

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' GetClassListInfo iClassListId, iUserId, iFamilyMemberId, iAttendeeUserId, iQuantity
'--------------------------------------------------------------------------------------------------
Sub GetClassListInfo( ByVal iClassListId, ByRef iUserId, ByRef iFamilyMemberId, ByRef iAttendeeUserId, ByRef iQuantity )
	Dim sSql, oRs

	sSql = "SELECT userid, familymemberid, attendeeuserid, quantity "
	sSql = sSql & "FROM egov_class_list WHERE classlistid = " & iClassListId
	'response.write sSql & "<br />" & vbcrlf

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iUserId = oRs("userid")
		iFamilyMemberId = oRs("familymemberid")
		iAttendeeUserId = oRs("attendeeuserid")
		' using the real qty causes problems when a "tickected" user with qty > 1 is copied to a "registration required" class that takes a qty of 1
		'iQuantity = oRs("quantity")
		iQuantity = 1
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' UpdateWaitCount iTimeId, sSign, sField, iQty 
'--------------------------------------------------------------------------------------------------
Sub UpdateWaitCount( ByVal iTimeId, ByVal sSign, ByVal sField, ByVal iQty )
	Dim sSql

	sSql = "UPDATE egov_class_time SET " & sField & " = " & sField & " " & sSign & " " & iQty & " WHERE timeid = " & iTimeId
	'response.write sSql & "<br />" & vbcrlf
	
	RunSQLStatement sSql

End Sub


'------------------------------------------------------------------------------
' integer AddToClassList(iUserId, iClassId, sStatus, iQuantity, iTimeId, iFamilymemberId, fAmount, iPaymentId, iAttendeeUserId )
'------------------------------------------------------------------------------
Function AddToClassList( ByVal iUserId, ByVal iClassId, ByVal sStatus, ByVal iQuantity, ByVal iTimeId, ByVal iFamilymemberId, ByVal fAmount, ByVal iPaymentId, ByVal iAttendeeUserId )
	Dim sSql

	AddToClassList = 0

	sSql = "INSERT INTO egov_class_list ( "
	sSql = sSql & "userid, "
	sSql = sSql & "classid, "
	sSql = sSql & "status, "
	sSql = sSql & "quantity, "
	sSql = sSql & "classtimeid, "
	sSql = sSql & "familymemberid, "
	sSql = sSql & "amount, "
	sSql = sSql & "paymentid, "
	sSql = sSql & "attendeeuserid "
	sSql = sSql & " ) VALUES ( " 
	sSql = sSql & iUserId                        & ", "
	sSql = sSql & iClassId                       & ", "
	sSql = sSql & "'" & sStatus                  & "', "
	sSql = sSql & iQuantity                      & ", "
	sSql = sSql & iTimeId                        & ", "
	sSql = sSql & iFamilymemberId                & ", "
	sSql = sSql & fAmount                        & ", "
	sSql = sSql & iPaymentId                     & ", "
	sSql = sSql & iAttendeeUserId                
	sSql = sSql & " )"

	'response.write sSql & "<br />" & vbcrlf

	AddToClassList = RunInsertStatement( sSql )

End Function 

%>

<!-- #include file="class_global_functions.asp" //-->

