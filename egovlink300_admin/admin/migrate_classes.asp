<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
	response.write "Classes migration did not run.  PLEASE DISABLE RESPONSE.END TO RUN SCRIPT."
	response.end

'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: migrate_classes.asp
' AUTHOR: Steve Loar
' CREATED: 06/12/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module migrates classses to the new Menlo Park structure
'
' MODIFICATION HISTORY
' 1.0   06/12/07	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' USER VALUES

sLevel = "../" ' Override of value from common.asp


%>

<html>
<head>
	<title>E-Gov Classes Migration Script</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

</head>

<body>
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	<h1>E-Gov Classes Migration Script</h1>
	<p><strong>Started: <%=Now()%></strong></p>
	<p><hr /></p>

<%
	Dim sSql, oRs, iLedgerId, sBuyOrWait, iAttendeeUserId, iMembershipid, dRegistrationStart, iOldTimeId, iTimeDayId, Item
	Dim iJournalEntryTypeID, iPaymentId, iPaymentInformationId

	iOldTimeId = CLng(0)
	iTimeDayId = 0

	' Update egov_class_payments - Mark them all as purchases
'	response.write "<p>egov_class_payments</p>"
'	RunSQL( "Update egov_class_payment Set journalentrytypeid = 1" )

'	Update egov_verisign_payment_information - Public Web Purchases
'	sSql = "Select paymentgatewayinformationid, paymentid, paymenttotal, orgid From egov_class_payment "
'	sSql = sSql & " Where paymentgatewayinformationid in (Select paymentinformationid from egov_verisign_payment_information) "
'	sSql = sSql & " Order by paymentgatewayinformationid"
'	response.write "<p>egov_verisign_payment_information, egov_accounts_ledger - Public Web Purchases<br />" & sSQL & "</p><br /><br />"
'	response.flush 

'	Set oRs = Server.CreateObject("ADODB.Recordset")
'	oRs.Open sSQL, Application("DSN"), 3, 1

'	Do While Not oRs.Eof
		' Create a ledger row for this.
'		sSql = "Insert into egov_accounts_ledger ( paymentid, orgid, entrytype, amount, plusminus, ispaymentaccount, paymenttypeid ) Values ( " 
'		sSql = sSql & oRs("paymentid") & ", " & oRs("orgid") & ", 'debit', " & oRs("paymenttotal") & ", '+', 1, 5 )"
'		iLedgerId = RunIdentityInsert( sSql )

		' Update the verisign table with the new ledger id
'		sSql = "Update egov_verisign_payment_information Set paymentid = " & oRs("paymentid") & ", amount = " & oRs("paymenttotal")
'		sSql = sSql & ", ledgerid = " & iLedgerId & " Where paymentinformationid = " & oRs("paymentgatewayinformationid")
'		RunSQL( sSql )
'		oRs.MoveNext
'	Loop
'	oRs.Close
'	Set oRs = Nothing 

	' Create ledger entries - Admin Purchases
'	sSql = "Select paymentid, isnull(paymenttotal,0.00) as paymenttotal, orgid From egov_class_payment "
'	sSql = sSql & " Where paymentgatewayinformationid is null and paymentid > 400 "
'	sSql = sSql & " Order by paymentid"
'
'	response.write "<p>egov_accounts_ledger - Admin Purchases<br />" & sSQL & "</p><br /><br />"
'	response.flush
'
'	Set oRs = Server.CreateObject("ADODB.Recordset")
'	oRs.Open sSQL, Application("DSN"), 3, 1
'
'	Do While Not oRs.Eof
'		' Create a ledger row for this.
'		sSql = "Insert into egov_accounts_ledger ( paymentid, orgid, entrytype, amount, plusminus, ispaymentaccount, paymenttypeid ) Values ( " 
'		sSql = sSql & oRs("paymentid") & ", " & oRs("orgid") & ", 'debit', " & oRs("paymenttotal") & ", '+', 1, 1 )"
'		RunSQL( sSql )
'		oRs.MoveNext
'	Loop
'	oRs.Close
'	Set oRs = Nothing 

'	response.end

'	Update egov_verisign_payment_information - Admin Purchases
'	sSql = "select paymentid, ledgerid, amount, paymenttypeid from egov_accounts_ledger where paymentid not in (select paymentid from egov_verisign_payment_information where paymentid is not null) and ispaymentaccount = 1 order by paymentid"
'	response.write "<p>egov_verisign_payment_information, egov_accounts_ledger - Admin Purchases<br />" & sSQL & "</p><br /><br />"
'	response.flush 
'
'	Set oRs = Server.CreateObject("ADODB.Recordset")
'	oRs.Open sSQL, Application("DSN"), 3, 1
'
'	Do While Not oRs.Eof
'		' Insert the verisign rows
'		sSql = "Insert INTO egov_verisign_payment_information (paymentid, ledgerid, amount, paymenttypeid) VALUES ( "
'		sSql = sSql & oRs("paymentid") & ", " & oRs("ledgerid") & ", " &  oRs("amount") & ", " & oRs("paymenttypeid") & " )"
'		iPaymentInformationId = RunIdentityInsert( sSql )
'		oRs.MoveNext
'	Loop
'	oRs.Close
'	Set oRs = Nothing 
'

	' Create Ledger rows for the classes
'	sSql = "Select L.classlistid, isnull(L.amount,0.00) as amount, L.status, L.paymentid, P.orgid "
'	sSql = sSql & " From egov_class_list L, egov_class_payment P Where L.paymentid = P.paymentid Order By L.classlistid"
'
'	response.write "<p>egov_accounts_ledger, egov_journal_item_status - Class Rows<br />" & sSQL & "</p><br /><br />"
'	response.flush
'	
'	Set oRs = Server.CreateObject("ADODB.Recordset")
'	oRs.Open sSQL, Application("DSN"), 3, 1
'
'	Do While Not oRs.Eof
'		' Create a ledger row for this.
'		sSql = "Insert into egov_accounts_ledger ( paymentid, orgid, entrytype, amount, itemtypeid, plusminus, itemid ) Values ( " 
'		sSql = sSql & oRs("paymentid") & ", " & oRs("orgid") & ", 'credit', " & oRs("amount") & ", 1, '+', " & oRs("classlistid") & " )"
'		RunSQL( sSql )
'
'		' Create a journal item status row
'		If oRs("status") = "WAITLIST" Then 
'			sBuyOrWait = "W"
'			sStatus = oRs("status")
'		Else
'			sBuyOrWait = "B" 
'			sStatus = "ACTIVE"
'		End If 
'		sSql = "Insert Into egov_journal_item_status (paymentid, itemtypeid, itemid, status, buyorwait) Values ( "
'		sSql = sSql & oRs("paymentid") & ", 1, " & oRs("classlistid") & ", '" & sStatus & "', '" & sBuyOrWait &"' )"
'		RunSQL( sSql )
'		oRs.MoveNext
'	Loop
'	oRs.Close
'	Set oRs = Nothing 
'
'	response.End 

	' Create records for the Dropped status ones
'	sSql = "Select classlistid, status, L.paymentid, isnull(refundamount,0.00) as refundamount, isnull(amount,0.00) as amount, P.userid, classid, orgid "
'	sSql = sSql & " From egov_class_list L, egov_class_payment P Where status = 'DROPPED' and L.paymentid = P.paymentid Order By L.paymentid, classlistid"

'	response.write "<p>Pull records for dropped classes<br />" & sSQL & "</p><br /><br />"
'	response.flush
'	
'	Set oRs = Server.CreateObject("ADODB.Recordset")
'	oRs.Open sSQL, Application("DSN"), 3, 1
'
'	Do While Not oRs.Eof
'		' Create the Journal Entry in egov_class_payment
'		sSql = "Insert into egov_class_payment (paymentdate, paymentlocationid, orgid, userid, paymenttotal, journalentrytypeid, notes, relatedpaymentid) Values ( "
'		sSql = sSql & " dbo.GetLocalDate(" & oRs("orgid") & ",GetDate()), 2, " & oRs("orgid") & ", " & oRs("userid") & ", " & oRs("refundamount") & ", 2, 'System generated for data migration', " & oRs("paymentid") & " )"
'		iPaymentId = RunIdentityInsert( sSql )
'
'		' Update the classlist with the new paymentid
'		sSql = "Update egov_class_list Set paymentid = " & iPaymentId & " Where classlistid = " & oRs("classlistid")
'		RunSQL( sSql )
'
'		' Create the Journal Item Status
'		sSql = "Insert Into egov_journal_item_status (paymentid, itemtypeid, itemid, status, buyorwait) Values ( "
'		sSql = sSql & iPaymentId & ", 1, " & oRs("classlistid") & ", 'DROPPED', 'D' )"
'		RunSQL( sSql )
'
'		' Credit the class payment
'		sSql = "Insert into egov_accounts_ledger ( paymentid, orgid, entrytype, amount, plusminus, ispaymentaccount, paymenttypeid ) Values ( " 
'		sSql = sSql & iPaymentId & ", " & oRs("orgid") & ", 'credit', " & oRs("refundamount") & ", '-', 1, 1 )"
'		RunSQL( sSql )
'
'		' Debit the class purchase
'		sSql = "Insert into egov_accounts_ledger ( paymentid, orgid, entrytype, amount, itemtypeid, plusminus, itemid ) Values ( " 
'		sSql = sSql & iPaymentId & ", " & oRs("orgid") & ", 'debit', " & oRs("amount") & ", 1, '-', " & oRs("classlistid") & " )"
'		RunSQL( sSql )
'
'		' Credit any kept money
'		If CDbl(oRs("amount")) > CDbl(oRs("refundamount")) Then
'			iKeptAmount = CDbl(oRs("amount")) - CDbl(oRs("refundamount"))
'			sSql = "Insert into egov_accounts_ledger ( paymentid, orgid, entrytype, amount, itemtypeid, plusminus, itemid ) Values ( " 
'			sSql = sSql & iPaymentId & ", " & oRs("orgid") & ", 'credit', " & iKeptAmount & ", 1, '+', " & oRs("classlistid") & " )"
'			RunSQL( sSql )
'		End If 
'
'		oRs.MoveNext
'	Loop
'	oRs.Close
'	Set oRs = Nothing 
'
'	response.write "<p><hr /></p><p><strong>Finished: " & Now() & "</strong></p>"
'	response.End 


	' Update the class list with attendeeuserid
'	sSql = "Select classlistid, familymemberid From egov_class_list where familymemberid is not null Order By familymemberid"
'
'	response.write "<p>egov_class_list - Update AttendeeUserId<br />" & sSQL & "</p><br /><br />"
'	response.flush
'	
'	Set oRs = Server.CreateObject("ADODB.Recordset")
'	oRs.Open sSQL, Application("DSN"), 3, 1
'
'	Do While Not oRs.Eof
'		iAttendeeUserId = GetAttendeeUserId( oRs("familymemberid") )
'		sSql = "Update egov_class_list Set attendeeuserid = " & iAttendeeUserId & " Where classlistid = " & oRs("classlistid")
'		RunSQL( sSql )
'		oRs.MoveNext
'	Loop
'	oRs.Close
'	Set oRs = Nothing 
'
'	response.write "<p><hr /></p><p><strong>Finished: " & Now() & "</strong></p>"
'	response.End


	' Update egov_class_pricetype_price
'	sSql = "Select classid, isnull(membershipid,0) as membershipid, registrationstartdate From egov_class where registrationstartdate is not null order by classid"

'	response.write "<p>egov_class_pricetype_price - Update pricetypes with registration start dates<br />" & sSQL & "</p><br /><br />"
'	response.flush
	
'	Set oRs = Server.CreateObject("ADODB.Recordset")
'	oRs.Open sSQL, Application("DSN"), 3, 1
'
'	Do While Not oRs.Eof
'		If CLng(oRs("membershipid")) = CLng(0) Then
'			iMembershipid = "NULL"
'		Else
'			iMembershipid = CLng(oRs("membershipid"))
'		End If 
'		If IsNull(oRs("registrationstartdate")) Then 
'			dRegistrationStart = Date()
'		Else
'			dRegistrationStart = oRs("registrationstartdate")
'		End If 
'		sSql = "Update egov_class_pricetype_price Set registrationstartdate = " & dRegistrationStart & ", membershipid = " & iMembershipid & " Where classid = " & oRs("classid")
'		RunSQL( sSql )
'		oRs.MoveNext
'	Loop
'	oRs.Close
'	Set oRs = Nothing 

'	response.write "<p><hr /></p><p><strong>Finished: " & Now() & "</strong></p>"
'	response.End


	' Update egov_class_time
'	sSql = "Update egov_class_time set activityno = Right('0000000' + Cast(timeid as Varchar), 8)"
'	RunSQL( sSql )

'	response.write "<p><hr /></p><p><strong>Finished: " & Now() & "</strong></p>"
'	response.End


	' Create the class time days rows
'	sSql = "select T.timeid, T.starttime, T.endtime, D.dayofweek, "
'	sSql = sSql & " case D.dayofweek when 1 then 'sunday' when 2 then 'monday' when 3 then 'tuesday' when 4 then 'wednesday' when 5 then 'thursday' when 6 then 'friday' when 7 then 'saturday' end as dayname "
'	sSql = sSql & " from egov_class_time T, egov_class_dayofweek D where T.classid = D.classid order by T.classid, T.timeid, D.dayofweek"
'
'	response.write "<p>egov_class_time_days - Create time day rows<br />" & sSQL & "</p><br /><br />"
'	response.flush
'	
'	Set oRs = Server.CreateObject("ADODB.Recordset")
'	oRs.Open sSQL, Application("DSN"), 3, 1
'
'	Do While Not oRs.Eof
'		If CLng(oRs("timeid")) = iOldTimeId Then
'			' Do an update for a new day of an existing time day row
'			sSql = "Update egov_class_time_days Set " & oRs("dayname") & " = 1 Where timedayid = " & iTimeDayId
'			RunSQL( sSql )
'		Else
'			' Do an insert of a new time day row
'			iOldTimeId = CLng(oRs("timeid"))
'			sSql = "Insert into egov_class_time_days ( timeid, starttime, endtime, " & oRs("dayname") & " ) Values ( "
'			sSql = sSql & oRs("timeid") & ", '" & oRs("starttime") & "', '" & oRs("endtime") & "', 1 )"
'			iTimeDayId = RunIdentityInsert( sSql )
'		End If 
'		oRs.MoveNext
'	Loop
'	oRs.Close
'	Set oRs = Nothing 

	' Update Accounts_ledger with pricetypeid - Did this for E, R, N, price match, and this catch all
'	sSql = "select L.classlistid, L.classid, P.pricetypeid, L.userid, P.amount as price, L.amount as paid, T.pricetype, U.residenttype "
'	sSql = sSql & " from egov_class_list L, egov_class_pricetype_price P, egov_price_types T, egov_users U "
'	sSql = sSql & " where L.classid = P.classid and P.pricetypeid = T.pricetypeid and L.userid = U.userid "
'	sSql = sSql & " and classlistid in (select itemid from egov_accounts_ledger where pricetypeid is null and ispaymentaccount = 0) "
'	sSql = sSql & " order by classlistid, L.classid"
'
'	response.write "<p>egov_accounts_ledger - Update pricetypeid<br />" & sSQL & "</p><br /><br />"
'	response.flush
'
'	Dim iOldClassListId
'	iOldClassListId = CLng(0)
'	
'	Set oRs = Server.CreateObject("ADODB.Recordset")
'	oRs.Open sSQL, Application("DSN"), 3, 1
'
'	Do While Not oRs.Eof
'		If iOldClassListId <> CLng(oRs("classlistid")) then
'			sSql = "Update egov_accounts_ledger Set pricetypeid = " & oRs("pricetypeid") & " Where itemid = " & oRs("classlistid")
'			RunSQL( sSql )
'			iOldClassListId = CLng(oRs("classlistid"))
'		End If 
'		oRs.MoveNext
'	Loop
'	oRs.Close
'	Set oRs = Nothing 

	'sSql = "select classid, registrationstartdate from egov_class "
	'sSql = sSql & " where classid in (select classid from egov_class_pricetype_price where registrationstartdate = '1/1/1900 12:00:00 AM')"
	'sSql = sSql & " where classid in (select classid from egov_class_pricetype_price where registrationstartdate is null) and registrationstartdate is not null"
	sSql = "select classid, publishstartdate as registrationstartdate from egov_class "
	sSql = sSql & " where classid in (select classid from egov_class_pricetype_price where registrationstartdate is null)"

	response.write "<p>egov_class_pricetype_price - Update egov_class_pricetype_price<br />" & sSQL & "</p><br /><br />"
	response.flush

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	Do While Not oRs.Eof
		sSql = "Update egov_class_pricetype_price Set registrationstartdate = '" & oRs("registrationstartdate") & "' Where classid = " & oRs("classid")
		RunSQL( sSql )
		oRs.MoveNext
	Loop
	oRs.Close
	Set oRs = Nothing 

%>

	<p><hr /></p>
	<p><strong>Finished: <%=Now()%></strong></p>
	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Sub RunSQL( sSql )
'-------------------------------------------------------------------------------------------------
Sub RunSQL( sSql )
	Dim oCmd

	response.write "<p>" & sSql & "</p><br /><br />"
	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 


'-------------------------------------------------------------------------------------------------
' Function RunIdentityInsert( sInsertStatement )
'-------------------------------------------------------------------------------------------------
Function RunIdentityInsert( sInsertStatement )
	Dim sSQL, iReturnValue, oInsert

	iReturnValue = 0

	response.write "<p>" & sInsertStatement & "</p><br /><br />"
	response.flush

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSQL = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.Open sSQL, Application("DSN"), 3, 3
	iReturnValue = oInsert("ROWID")
	oInsert.close
	Set oInsert = Nothing

	RunIdentityInsert = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' Function GetAttendeeUserId( iFamilymemberId )
'--------------------------------------------------------------------------------------------------
Function GetAttendeeUserId( iFamilymemberId )
	Dim sSql, oUserId

	sSQL = "select userid From egov_familymembers Where familymemberid = " & iFamilymemberId

	Set oUserId = Server.CreateObject("ADODB.Recordset")
	oUserId.Open sSQL, Application("DSN"), 3, 1
	
	If Not oUserId.EOF Then 
		GetAttendeeUserId = CLng(oUserId("userid"))
	Else
		GetAttendeeUserId = 0
	End If 
	
	oUserId.close 
	Set oUserId = Nothing
End Function 


%>
