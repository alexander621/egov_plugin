<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME:  waitlist_removal.asp
' AUTHOR: Steve Loar
' CREATED: 00/02/2007
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   05/02/2007	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPaymentId, iCitizenId, iAdminUserId, sAmount, iJournalEntryTypeID, sNotes, iMaxPaymentTypes
Dim sCheck, iCitizenAccountId, sPlusMinus, cPriorBalance, iAccountId, iPaymentLocationId, iAdminLocationId
Dim cAmount, iLedgerId, iQuantity, iFamilyMemberId, iClassId, bIsParent, iOldPaymentId

iPaymentLocationId = request("PaymentLocationId") 

' this is where the admin person is working today
If Session("LocationId") <> "" Then
	iAdminLocationId = Session("LocationId")
Else
	iAdminLocationId = 0 
End If 

iCitizenId = request("iUserId") ' Purchasing citizen (Head of Household)
iFamilyMemberId = request("familymemberid")
iAttendeeUserId = GetAttendeeUserId( iFamilymemberId )
iAdminUserId = Session("UserID")
iJournalEntryTypeID = GetJournalEntryTypeID( "refund" )
sNotes = dbsafe(request("notes"))
iQuantity = request("quantity")
sAmount = 0.00  'The waitlist is free
sStatus = "WAITLIST REMOVED"
iClassId = request("classid")
bIsParent = IsSeriesParent( iClassId )
iOldPaymentId = request("oldpaymentid")


' Insert the egov_class_payment row
iPaymentId = MakeJournalEntry( iPaymentLocationId, iAdminLocationId, iCitizenId, iAdminUserId, sAmount, iJournalEntryTypeID, sNotes, request("oldpaymentid") )

iItemTypeId = GetItemTypeId( "recreation activity" )

' Loop through the price fields and process any checked
'For Each iPriceTypeId In request("pricetypeid")
	' Pull any accountids - These may not exist, so pull seperately
'	iAccountId = GetAccountId( iPriceTypeId, request("classid") )  ' In class_global_functions.asp
'	cAmount = CDbl(request("price" & iPriceTypeId )) * CDbl(iQuantity)

	' create the ledger rows for the class accounts
	'MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iPriceTypeid )
	iLedgerId = MakeLedgerEntry( Session("Orgid"), "NULL", iPaymentId, sAmount, iItemTypeId, "debit", "-", request("iclasslistid"), 0, "NULL", "NULL", 0 )
	'AddClassLedgerRows Session("Orgid"), iPaymentId, iCartId, iItemTypeId, iClassListId, "credit", "+"
'Next  


' UPDATE THE CLASS LIST
UpdateStatus request("iclasslistid"), sStatus, request("paymenttotal"), iPaymentId

' Add to egov_journal_item_status
CreateJournalItemStatus iPaymentId, iItemTypeId, request("iclasslistid"), sStatus, "R"


' UPDATE WAITLIST count
UpdateEnrollment request("timeid"), iQuantity 

' Remove from children if necessary.  Check that they are still waitlist
If bIsParent Then
	RemoveFromChildren iClassId, sStatus, iQuantity, iOldPaymentId, iPaymentId
End If 

' RETURN TO ROSTER VIEW
response.redirect("view_roster.asp?classid=" & request("classid") & "&timeid=" & request("timeid") )


%>

<!--#Include file="class_global_functions.asp"-->  

<!-- #include file="../includes/common.asp" //-->

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function MakeJournalEntry( iPaymentLocationId, iAdminLocationId )
'--------------------------------------------------------------------------------------------------
Function MakeJournalEntry( iPaymentLocationId, iAdminLocationId, iCitizenId, iAdminUserId, sAmount, iJournalEntryTypeID, sNotes, iRelatedPaymentId )
	Dim sSql, oInsert

	MakeClassPayment = 0

	sSql = "Insert into egov_class_payment (paymentdate, paymentlocationid, orgid, adminlocationid, "
	sSql = sSql & " userid, adminuserid, paymenttotal, journalentrytypeid, notes, relatedpaymentid) Values (dbo.GetLocalDate(" & Session("orgid") & ",GetDate()), " 
	sSql = sSql & iPaymentLocationId & ", " & Session("orgid") & ", " & iAdminLocationId & ", "
	sSql = sSql & iCitizenId & ", " & iAdminUserId & ", " & sAmount & ", " & iJournalEntryTypeID & ", '" & sNotes & "', " & iRelatedPaymentId & " )"
	sSql = "SET NOCOUNT ON;" & sSql & ";SELECT @@IDENTITY AS ROWID;"
	response.write sSQL & "<br /><br />"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSql, Application("DSN"), 3, 3

	MakeJournalEntry = oInsert("ROWID")

	oInsert.close
	Set oInsert = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' SUB UPDATESTATUS(ICLASSLISTID,SSTATUS,SAMOUNT, iPaymentId)
'--------------------------------------------------------------------------------------------------
Sub UpdateStatus( iClasslistId, sStatus, sAmount, iPaymentId )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "Update egov_class_list Set status = '" & sStatus & "', amount = " & sAmount & ", paymentid = " & iPaymentId & " Where classlistid = " & iClasslistId
		.Execute
	End With
	Set oCmd = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' SUB UPDATEPAYMENTINFO( IPAYMENTID, IPAYMENTTYPE, IPAYMENTLOCATION, CURPAYMENTTOTAL )
'--------------------------------------------------------------------------------------------------
Sub UpdatePaymentInfo( iPaymentId, iPaymentType, iPaymentLocation, curPaymentTotal )
	Dim sSql, oCmd

	sSql = "Update egov_class_payment set paymenttypeid = " & iPaymentType & ", paymentlocationid = " & iPaymentLocation & ",paymenttotal = " & curPaymentTotal & " WHERE paymentid = " & iPaymentId

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' SUB UPDATEENROLLMENT(ICLASSTIMEID, iQuantity)
'--------------------------------------------------------------------------------------------------
Sub UpdateEnrollment( iclasstimeid, iQuantity )
	Dim sSql, oCmd

	sSQL = "UPDATE EGOV_CLASS_TIME SET waitlistsize = waitlistsize - " & iQuantity & " WHERE TIMEID = " & iclasstimeid 

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'  Make buffer Database 'safe'
'  Useful in building SQL Strings
'    strSQL="SELECT *....WHERE Value='" & DBSafe(strValue) & "';"
'--------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
	If Not VarType( strDB ) = vbString Then 
		DBsafe = strDB
	Else 
		DBsafe = Replace( strDB, "'", "''" )
	End If 
End Function


'--------------------------------------------------------------------------------------------------
' Sub RemoveFromChildren( iClassId, sStatus, iQuantity, iOldPaymentId )
'--------------------------------------------------------------------------------------------------
Sub RemoveFromChildren( iClassId, sStatus, iQuantity, iOldPaymentId, iNewPaymentId )
	Dim sSql, oChild, iChildClassListId

	sSql = "Select L.classlistid, L.classtimeid From egov_class C, egov_class_list L "
	sSql = sSql & " Where C.classid = L.classid and L.status = 'WAITLIST' and C.parentclassid = " & iClassId & " and paymentid = " & iOldPaymentId

	Set oChild = Server.CreateObject("ADODB.Recordset")
	oChild.Open sSQL, Application("DSN"), 0, 1

	Do While Not oChild.EOF
		UpdateStatus oChild("classlistid"), sStatus, CDbl(0.00), iNewPaymentId 
		UpdateEnrollment oChild("classtimeid"), iQuantity 
		oChild.movenext
	Loop 

	oChild.close
	Set oChild = Nothing
End Sub 
%>
