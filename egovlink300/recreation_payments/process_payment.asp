<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="../recreation/facility_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: process_payment.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2005 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module processes payments.
'
' MODIFICATION HISTORY
' 1.0	??/??/??	?????????? - Initial Version
' 1.1	06/02/09	Steve Loar - Changes for centralized PayPal processing and PayFlow Pro changes
' 1.2	09/25/09	Steve Loar - Added the session timeout check.
' 1.3	12/08/09	David Boyer - Added check for "Edit Display" for "pool pass" paragraph (line: 796)
' 2.0	06/23/2010	Steve Loar - Split name field into first and last 
' 2.1	07/28/2010	Steve Loar - Added Processing for Point and Pay
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sError, iPaymentControlNumber, sOrderId, sProcessingRoute, lcl_orghasfeature_display_cvv
Dim lcl_orghasfeature_custom_registration_craigco, lcl_facility_avail

	strResponse = request.form("g-recaptcha-response")
	strIP = request.servervariables("REMOTE_HOST")
	strSecret = "6LcVxxwUAAAAAGGp_29X6bpiJ8YsWeNXinuUz6sx"

		Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
		objWinHttp.SetTimeouts 0, 120000, 60000, 120000

		objWinHttp.Open "POST", "https://www.google.com/recaptcha/api/siteverify", False

		objWinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"

		objWinHttp.Send "secret=" & strSecret & "&response=" & strResponse & "&remoteip=" & strIP


		If objWinHttp.Status = 200 Then 
			' Get the text of the response.
			transResponse = objWinHttp.ResponseText
		End If 

		' Trash our object now that we are finished with it.
		Set objWinHttp = Nothing

		if instr(transResponse, """success"": true") = 0 then
			response.redirect "../recreation/facility_availability.asp?success=TIMEOUT"
		end if

If clng(request("iPAYMENT_MODULE")) = 3 Then 
	'If this is for a facility then verify that the facility is still available
	lcl_facility_avail = isFacilityAvail(request("iFacilityPaymentID"), "", "", "", "", "")

	If Not lcl_facility_avail Then 
		response.redirect "../recreation/facility_availability.asp?L=" & session("facilityid") & "&Y=" & year(session("D")) & "&M=" & month(session("D")) & "&success=NA"
	End If 

	'Verify that the session has not timed out while on the verisign_form page. We do not what to charge them if they have lost the reservation.
	If ReservationDataIsGone( request("iFacilityPaymentID") ) Then 
		response.redirect "../recreation/facility_availability.asp?success=TIMEOUT"
	End If 

	
	' Verify that the the person is not just hitting refresh and getting double billed. True indicates they have already paid.
	If facilityScheduleIsReserved( CLng(request("iFacilityPaymentID")) ) Then
		' send them back somewhere safe'
		response.redirect "../recreation/facility_availability.asp"
	End If 
End If 

'Check for org features
lcl_orghasfeature_display_cvv = orgHasFeature( iOrgID, "display cvv" )
lcl_orghasfeature_custom_registration_craigco = orgHasFeature( iOrgID, "custom_registration_CraigCO" )

'Get the processing route
sProcessingRoute = GetProcessingRoute()

%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<title>E-Gov Services <%=sOrgName%> - Verisign Payment Form</title>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script language="javascript" src="../scripts/modules.js"></script>

	<style>
<%
		if request.servervariables("HTTPS") = "on" then
  			response.write "body {behavior: url('https://secure.egovlink.com/" & sorgVirtualSiteName & "/csshover.htc');}" & vbcrlf
		end if
%>
	</style>
</head>

<!--#Include file="../include_top.asp"-->

	<!--BODY CONTENT-->
	<table border="0">
	<tr>
	<td valign="top">

	<!--BEGIN: INTRO TEXT-->
	<div class="title"><%=sWelcomeMessage%></div>
	
	<div class="main"><font color="#1c4aab">Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%></font></div>
	<!--END: INTRO TEXT-->

	<!--BEGIN:  DISPLAY SKIPJACK RESPONSE-->

	<% 
	reprocess = false
	Select Case clng(request.form("iPAYMENT_MODULE"))
		Case 1
			' COMMEMORATIVE GIFT -- check request("iGiftPaymentId")
			sSQL = "SELECT result as status FROM egov_gift_payment WHERE giftpaymentid = '" & request("iGiftPaymentId") & "'"

		Case 2
			' POOL PASS -- check request("itemnumber")
			sSQL = "SELECT paymentresult as status FROM egov_poolpasspurchases WHERE poolpassid = '" & request("itemnumber") & "'"
		Case 3
			'FACILITY RESERVATION -- check request("iFacilityPaymentID")
			sSQL = "SELECT result as status FROM egov_facilityschedule WHERE facilityscheduleid = '" & request("iFacilityPaymentID") & "'"
		Case Else
			sSQL = "SELECT paymentresult as status FROM egov_poolpasspurchases WHERE 1 = 2"
	End Select
	Set oRsCheck = Server.CreateObject("ADODB.RecordSet")
	oRsCheck.Open sSQL, Application("DSN"), 3, 1
	if not oRsCheck.EOF then
		if oRsCheck("status") = "APPROVED" then reprocess = true
	end if
	oRsCheck.Close
	Set oRsCheck = Nothing


	if reprocess then response.redirect "../purchases_report/purchases_list.asp"


	If OrgHasFeature( iOrgId, "skippayment" ) Then 
		'dtb_debug("Skipping")
'		response.write "Skipping" & "<br />"
		iPaymentControlNumber = CreatePaymentControlRow( "PAYMENT SCRIPT STARTED." )
		AddToPaymentLog iPaymentControlNumber, "TRANSACTION SUCCEEDED - Bypassed Authorization"
		approved  = True
		sAuthcode = "010101"
		sPNREF    = "V19F1D5C82TEST"
		sRespMsg  = "Approved"
		AddToPaymentLog iPaymentControlNumber, "AUTHCODE: " & sAuthcode
		AddToPaymentLog iPaymentControlNumber, "PNREF: " & sPNREF
		AddToPaymentLog iPaymentControlNumber, "RESPMSG: " & sRespMsg
		ProcessSuccessfulTransaction sOrderID,sAuthcode, sPNREF, sRespMsg, request("transactionamount"),  request("iPAYMENT_MODULE"), "", "NULL", "NULL", "NULL"
		'ProcessSuccessfulTransaction sOrderID, oReturnCodes.Item("AUTHCODE"), oReturnCodes.Item("PNREF"), oReturnCodes.Item("RESPMSG"), sAmount, request("iPAYMENT_MODULE"), sCVV2Match
	Else 
		'dtb_debug("sProcessingRoute= " & sProcessingRoute)
		Select Case sProcessingRoute
			Case "StandardPayPal"
				' Should be using the Old way of processing PayFlow Pro payments - This should be Bullhead City Only
				ProcessCreditCardTransaction request("firstname") & " " & request("lastname"),request("accountnumber"),request("month")&request("year"),request("transactionamount")
			Case "VerisignPayFlowPro"
				' Old way of processing PayFlow Pro payments
				ProcessCreditCardTransaction request("firstname") & " " & request("lastname"),request("accountnumber"),request("month")&request("year"),request("transactionamount")
			Case "PayPalPayFlowPro"
				' Newer way to handle PayFlow Pro payments
				ProcessPayPalTransaction request("firstname") & " " & request("lastname"), request("accountnumber"), request("month")&request("year"), request("transactionamount") 
			Case "PointAndPay"
				' Process Point and Pay payments
				ProcessPointAndPayTransaction request("firstname"), request("lastname"), request("accountnumber"), request("month")&request("year"), request("transactionamount") 
		End Select 
	End If
%>	

	<!--BEGIN: PAYMENT FOOTER-->
	<center><br /><br />
		<input type="button" class="button hideme" onClick="location.href='<%=sEgovWebsiteURL%>/';" value="Click here to return to the E-Government Website" /><br />

		<p class="smallnote">
			NOTE: Your IP address [<%=request.servervariables("REMOTE_ADDR")%>] has been logged with this transaction.<br /><br />
		</p>
		<p> </p>
	</center>

	<!--END: PAYMENT FOOTER-->
	</div>

	<!--END: DISPLAY SKIPJACK RESPONSE-->

	<!--SPACING CODE-->
	<p>&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</p>
	<!--SPACING CODE-->

	<!--#Include file="../include_bottom.asp"--> 



<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ProcessPayPalTransaction sName, sCardNumber, sExpiration, sAmount
'------------------------------------------------------------------------------
Sub ProcessPayPalTransaction( ByVal sName, ByVal sCardNumber, ByVal sExpiration, ByVal sAmount )
	Dim parmList, objWinHttp, sParameter, strLength, sResult, sPNREF, sRespMsg, sAuthcode, sCVV2Match, sDuplicate

	sDuplicate = "Start"

	iPaymentControlNumber = CreatePaymentControlRow( "PAYMENT SCRIPT STARTED." )
	'dtb_debug("PAYMENT SCRIPT STARTED.")

	parmList = "cardNum=" & sCardNumber
	parmList = parmList + "&cardExp=" & sExpiration  ' format is MMYY
	parmList = parmList + "&sjname=" + sName 
	AddToPaymentLog iPaymentControlNumber, "Name: " & sName 
	If OrgHasFeature( iOrgId, "display cvv" ) And request("cvv2") <> "" Then
		parmList = parmList + "&cvv2=" & request("cvv2")
	End If 

	parmList = parmList + "&amount=" & sAmount
	AddToPaymentLog iPaymentControlNumber, "Amount: " & FormatNumber(sAmount,2,,,0)

	'parmList = parmList + "&Email=" + request("EMAIL")
	parmList = parmList + "&StreetAddress=" & request("StreetAddress")
	if iOrgId = "228" then
		parmList = parmList + "&City=" + request("City")
		parmList = parmList + "&State=" + request("State")
	end if
	parmList = parmList + "&ZipCode=" & request("ZipCode")

	parmList = parmList + "&ordernumber=" + request("ordernumber") 

	sParameter = request("paymentname")
	strLength  = CleanAndCountForPayFlowPro( sParameter )
	parmList   = parmList + "&comment1=" & sParameter & " - " & request("ordernumber")
	AddToPaymentLog iPaymentControlNumber, "COMMENT1: " & sParameter & " - " & request("ordernumber")

	sParameter = request("details")
	strLength  = CleanAndCountForPayFlowPro( sParameter )
	parmList   = parmList + "&comment2=" & sParameter
	AddToPaymentLog iPaymentControlNumber, "COMMENT2: " & sParameter

	parmList = parmList + "&paymentcontrolnumber=" & iPaymentControlNumber
	parmList = parmList + "&orgid=" & iOrgId
	parmList = parmList + "&orgfeature=rentals"


	'response.write parmList & "<br />"
	'dtb_debug(parmlist)
	Do While sDuplicate <> "" and sDuplicate <> "-1"
		Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")

		' Set timeouts of resolve(0), connection(60000), send(30000), receive(30000) in milliseconds. 0 = infinite
		objWinHttp.SetTimeouts 0, 120000, 60000, 120000

		objWinHttp.Open "GET", sEgovWebsiteURL & "/payment_processors/paypalsend.asp?" & parmList, False

		objWinHttp.setRequestHeader "Content-Type", "text/namevalue"

		objWinHttp.Send parmList

		If objWinHttp.Status = 200 Then 
			' Get the text of the response.
			transResponse = objWinHttp.ResponseText
		End If 

		' Trash our object now that we are finished with it.
		Set objWinHttp = Nothing
		'dtb_debug("Back")

		sDuplicate = GetResponseValue(transResponse, "DUPLICATE")
		' A duplicate will pull results from the initial request with that requestid.
		' All of these will be unique purchases so duplicates are not allowed.
		' So try again, hoping for a new requestid
	Loop 

	' Continue processing the results from the PayPal call
	sResult = GetResponseValue(transResponse, "RESULT")
	sResult = clng(sResult)
	sPNREF = GetResponseValue(transResponse, "PNREF")
	sRespMsg = GetResponseValue(transResponse, "RESPMSG")
	sCVV2Match = GetResponseValue(transResponse, "CVV2MATCH")

	If sResult = clng(0) Then 
		sAuthcode = GetResponseValue(transResponse, "AUTHCODE")

		' Successful Transaction 
		AddToPaymentLog iPaymentControlNumber, "TRANSACTION SUCCEEDED"
		approved = True
		AddToPaymentLog iPaymentControlNumber, "AUTHCODE: " & sAuthcode
		AddToPaymentLog iPaymentControlNumber, "PNREF: " & sPNREF
		AddToPaymentLog iPaymentControlNumber, "RESPMSG: " & sRespMsg
		AddToPaymentLog iPaymentControlNumber, "CVV2MATCH: " & sCVV2Match
		ProcessSuccessfulTransaction sOrderID, sAuthcode, sPNREF, sRespMsg, sAmount, request("iPAYMENT_MODULE"), sCVV2Match, "NULL", "NULL", "NULL"

	ElseIf sResult < clng(0) Then 
		' Communication Error
		AddToPaymentLog iPaymentControlNumber, "Communication Error"
		AddToPaymentLog iPaymentControlNumber, "Result: " & sResult
		AddToPaymentLog iPaymentControlNumber, "PNREF: " & sPNREF
		AddToPaymentLog iPaymentControlNumber, "RESPMSG: " & sRespMsg
		approved = False 
		response.write "<div class=""payflowmsgfail"">Your credit card purchase was unable to processed because of a network communication error. Please try your transaction again later.<blockquote><font color=#000000>Payment Reference Number:</font> " & sPNREF & " <br><font color=#000000>Description:</font> (" & sResult & ") - " & sRespMsg & " </blockquote></div>"

	ElseIf sResult > clng(0) Then 
		' Catch and send an alert when the client has changed their password and not told us
		If sResult = clng(1) Then
			SendLoginFailedEmail
		End If
		' Transaction Declined
		AddToPaymentLog iPaymentControlNumber, "Transaction Declined"
		AddToPaymentLog iPaymentControlNumber, "Result: " & sResult
		AddToPaymentLog iPaymentControlNumber, "PNREF: " & sPNREF
		AddToPaymentLog iPaymentControlNumber, "RESPMSG: " & sRespMsg
		approved = False
		ProcessDeclinedTransaction "paypal", sOrderID, sResult, sPNREF, sRespMsg, request("iPAYMENT_MODULE"), "NULL", "NULL", FormatNumber(sAmount,2,,,0)

	End If  
			
	AddToPaymentLog iPaymentControlNumber, "PAYMENT PROCESSING FINISHED."

End Sub 


'--------------------------------------------------------------------------------------------------
' void ProcessPointAndPayTransaction sFirstName, sLastName, sCardNumber, sExpiration, sAmount
'--------------------------------------------------------------------------------------------------
Sub ProcessPointAndPayTransaction( ByVal sFirstName, ByVal sLastName, ByVal sCardNumber, ByVal sExpiration, ByVal sAmount )
	Dim parmList, objWinHttp, sResult, sNotes, sStatus, sErrorMsg, dFeeAmount, sOrderNumber, sSVA, sTotalCharges

	iPaymentControlNumber = CreatePaymentControlRow( "PAYMENT SCRIPT STARTED - Point and Pay." )

	parmList = "paymentcontrolnumber=" & iPaymentControlNumber
	parmList = parmList + "&chargeaccountnumber=" & sCardNumber
	parmList = parmList + "&chargeexpirationmmyy=" & sExpiration  ' format is MMYY
	parmList = parmList + "&signerfirstname=" + sFirstName 
	parmList = parmList + "&signerlastname=" + sLastName
	AddToPaymentLog iPaymentControlNumber, "Name: " & sFirstName & " " & sLastName 

	If OrgHasFeature( iOrgId, "display cvv" ) And request("cvv2") <> "" Then
		parmList = parmList + "&chargecvn=" & request("cvv2")
		AddToPaymentLog iPaymentControlNumber, "ChargeCVN: present but not stored"
	End If 

	parmList = parmList + "&chargeamount=" & sAmount
	AddToPaymentLog iPaymentControlNumber, "Amount: " & FormatNumber(sAmount,2,,,0)

	parmList = parmList + "&signeraddressline1=" & request("StreetAddress")
	parmList = parmList + "&signeraddresscity=" + request("City")
	parmList = parmList + "&signeraddressregioncode=" + request("State")
	parmList = parmList + "&signeraddresspostalcode=" & request("ZipCode")
	
	sNotes = request("paymentname") & " - " & request("details")
	AddToPaymentLog iPaymentControlNumber, "Notes: " & CleanAndCutForPNPNotes( sNotes )
	parmList = parmList + "&notes=" & CleanAndCutForPNPNotes( sNotes )

'	response.write "[" & parmList & "]<br /><br />"
'	response.End 

	Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")

	' Set timeouts of resolve(0), connection(60000), send(30000), receive(30000) in milliseconds. 0 = infinite
	objWinHttp.SetTimeouts 0, 120000, 60000, 120000

	' do not roll this log entry out to prod. it has the CC number in it
	'AddToPaymentLog iPaymentControlNumber, "URL: " & sEgovWebsiteURL & "/payment_processors/pnpsend.aspx?" & parmList

	objWinHttp.Open "GET", sEgovWebsiteURL & "/payment_processors/pnpsend.aspx?" & parmList, False

	objWinHttp.setRequestHeader "Content-Type", "text/namevalue"

	objWinHttp.Send parmList

	'response.write "objWinHttp.Status: " & objWinHttp.Status & "<br /><br />"

	If objWinHttp.Status = 200 Then 
		' Get the text of the response.
		transResponse = objWinHttp.ResponseText
	End If 

'	response.write "objWinHttp.ResponseText: [" & objWinHttp.ResponseText & "]<br /><br />"
	' Trash our object now that we are finished with it.
	Set objWinHttp = Nothing
'	response.End 

	sStatus = GetPNPResponseValue(transResponse, "status")		' in ../includes/common.asp
	AddToPaymentLog iPaymentControlNumber, "status: " & sStatus

	sErrorMsg = GetPNPResponseValue(transResponse, "errors")	' in ../includes/common.asp
	AddToPaymentLog iPaymentControlNumber, "errors: " & sErrorMsg

	sSVA = GetPNPResponseValue(transResponse, "sva")	' in ../includes/common.asp
	AddToPaymentLog iPaymentControlNumber, "sva: " & sSVA

	sOrderNumber = GetPNPResponseValue(transResponse, "orderNumber")	' in ../includes/common.asp
	AddToPaymentLog iPaymentControlNumber, "orderNumber: " & sOrderNumber

	If LCase(sStatus) <> "success" Then		
		' They were declined or there was an error
		approved = False
		' ProcessDeclinedTransaction "paypal", sResult, sPNREF, sRespMsg, dTotalAmount, sCardNumber, sReservationDetails, "NULL", "NULL"
		' ProcessDeclinedTransaction "paypal", sOrderID, sResult, sPNREF, sRespMsg, request("iPAYMENT_MODULE"), "NULL", "NULL"
		ProcessDeclinedTransaction "PointAndPay", "NULL", "declined", "NULL", sErrorMsg, request("iPAYMENT_MODULE"), sOrderNumber, sSVA, FormatNumber(sAmount,2,,,0)
	Else
		approved = True
		dFeeAmount = GetPNPResponseValue(transResponse, "fee")	' in ../includes/common.asp
		If dFeeAmount = "" Then
			dFeeAmount = "0.00"
		End If 
		dFeeAmount = CDbl(dFeeAmount)
		AddToPaymentLog iPaymentControlNumber, "Fee Amount: " & FormatNumber(dFeeAmount,2,,,0)

		sTotalCharges = GetPNPResponseValue(transResponse, "total")	' in ../includes/common.asp
		AddToPaymentLog iPaymentControlNumber, "Total Charged: " & sTotalCharges

		ProcessSuccessfulTransaction "NULL", "NULL", "NULL", "approved", FormatNumber(sAmount,2,,,0), request("iPAYMENT_MODULE"), "NULL", sOrderNumber, sSVA, FormatNumber(dFeeAmount,2,,,0)

	End If 

	AddToPaymentLog iPaymentControlNumber, "PAYMENT PROCESSING FINISHED."

End Sub 


'------------------------------------------------------------------------------
' void ProcessCreditCardTransaction( sName, sCardNumber, sExpiration, sAmount )
'------------------------------------------------------------------------------
Sub ProcessCreditCardTransaction( ByVal sName, ByVal sCardNumber, ByVal sExpiration, ByVal sAmount )
	Dim sPayflowURL, sCVV2Match

	'------------------------------------------------------------------------------
	'Comment this OUT if this is PROD!!!
	 'if 1 = 2 then
	'------------------------------------------------------------------------------

	sPayflowURL = ""
	' CREATE THE PAYFLOW COM CLIENT COMPONENT
	Set client = Server.CreateObject("PFProCOMControl.PFProCOMControl.1")
	Set oReturnCodes = CreateObject("Scripting.Dictionary")

	' BUILD PARAMETER LIST
	' TRANSACTION DETAILS
	parmList = "TRXTYPE=S" ' SALE TRANSACTION - IMMEDIATELY FUND WITHDRAWAL
	parmList = parmList + "&TENDER=C" ' CREDIT CARD TRANSACTION
	parmList = parmList + "&ACCT=" + sCardNumber ' SET CREDIT CARD NUMBER
	parmList = parmList + "&EXPDATE=" + sExpiration' SET CREDIT CARD EXP DATE

	If OrgHasFeature( iOrgId, "display cvv" ) And request("cvv2") <> "" Then
  		parmList = parmList + "&CVV2=" + request("cvv2") ' CVV Security Code
	End If 

	parmList = parmList + "&NAME=" + sName' SET CUSTOMER's FULL NAME
	parmList = parmList + "&AMT=" + sAmount ' SET AMOUNT TO BE CHARGED

	' This function gets the 4 lines above and the URL to send to
	parmList = parmList & GetVerisignOptions( iorgid, sPayflowURL )

	' PERSONNEL INFO
	parmList = parmList + "&EMAIL=" + REQUEST("EMAIL") ' SET PAYFLOW PARTNER
	parmList = parmList + "&STREET=" + REQUEST("STREETADDRESS") ' SET PAYFLOW PARTNER
	parmList = parmList + "&CITY=" + REQUEST("CITY") ' SET PAYFLOW PARTNER
	parmList = parmList + "&STATE=" + REQUEST("STATE") ' SET PAYFLOW PARTNER
	parmList = parmList + "&ZIP=" + REQUEST("ZIPCODE") ' SET PAYFLOW PARTNER

	' ORDER INFO
	parmList = parmList + "&PONUM=" + REQUEST("ordernumber") ' COMMENT
	parmList = parmList + "&COMMENT1=" + REQUEST("paymentname")' COMMENT
	parmList = parmList + "&COMMENT2=" + Replace(Replace(REQUEST("details"),"</br>",", "),", ,","") ' COMMENT

	' CREATE TRANSACTION AND COMMUNICATE WITH PAYFLOW SERVER
	oTransaction = client.CreateContext(sPayflowURL, 443, 30, "", 0, "", "")
	sReturnCodes = client.SubmitTransaction(oTransaction, parmList, Len(parmList))
	client.DestroyContext (oTransaction)

	' PROCESS RETURN CODES 
	Do While Len(sReturnCodes) <> 0
		' GET NAME VALUE PAIR
		If InStr(sReturnCodes,"&") Then
			varString = Left(sReturnCodes, InStr(sReturnCodes , "&" ) -1)
		Else 
			varString = sReturnCodes
		End If 

		' GET VALUES FOR PAIR FROM STRING
		name = Left(varString, InStr(varString, "=" ) -1) ' GET RETURN CODE NAME
		value = Right(varString, Len(varString) - (Len(name)+1)) ' GET RETURN CODE VALUE

		' ADD ITEMS TO DICTIONARY
		oReturnCodes.Add name,value
		'response.write name & value
		' SKIP PROCESSING & IN RETURN CODE STRING
		If Len(sReturnCodes) <> Len(varString) Then 
			sReturnCodes = Right(sReturnCodes, Len(sReturnCodes) - (Len(varString)+1))
		Else 
			sReturnCodes = ""
		End If 
	Loop

	'------------------------------------------------------------------------------
	'Comment this OUT if this is PROD!!!
	' else
	'		Set oReturnCodes = CreateObject("Scripting.Dictionary")
	 '   oReturnCodes.Item("RESULT")         = 0
	'  '  oReturnCodes.Item("AUTHCODE")       = ""
	'    oReturnCodes.Item("PNREF")          = ""
	'    oReturnCodes.Item("RESPMSG")        = ""
	'    oReturnCodes.Item("iPAYMENTMODULE") = ""
	' end if
	'------------------------------------------------------------------------------

	' PROCESS TRANSACTION RESULT
	If  clng(oReturnCodes.Item("RESULT")) = 0 Then
		' TRANSACTION SUCCEEDED
		approved = True
		If oReturnCodes.Item("CVV2MATCH") <> "" Then
			sCVV2Match = oReturnCodes.Item("CVV2MATCH")
		Else
			sCVV2Match = ""
		End If 
		ProcessSuccessfulTransaction sOrderID, oReturnCodes.Item("AUTHCODE"), oReturnCodes.Item("PNREF"), oReturnCodes.Item("RESPMSG"), sAmount, request("iPAYMENT_MODULE"), sCVV2Match, "NULL", "NULL", "NULL"
	
	ElseIf clng(oReturnCodes.Item("RESULT")) < 0 Then
		' COMMUNICATION ERROR
		approved = FALSE

		' DISPLAY COMMUNICATION MESSAGE TO CUSTOMER
		response.write "<div class=""payflowmsgfail"">Your credit card purchase was unable to processed because of a network communication error. Please try your transaction again later.<blockquote><font color=#000000>DSI Order Number:</font> " & sOrderID & "<br /><font color=#000000>Payment Reference Number:</font> "&oReturnCodes.Item("PNREF")&" <br /><font color=#000000>Description:</font> ("&oReturnCodes.Item("RESULT")&") - "&oReturnCodes.Item("RESPMSG")&" </blockquote></div>"

	ElseIf clng(oReturnCodes.Item("RESULT")) > 0 Then

		' TRANSACTION FAILED
		approved = False 
		'response.write oReturnCodes.Item("RESULT") & ":" & oReturnCodes.Item("PNREF")  & ":" & oReturnCodes.Item("RESPMSG")
		ProcessDeclinedTransaction "paypal", sOrderID, oReturnCodes.Item("RESULT"), oReturnCodes.Item("PNREF"), oReturnCodes.Item("RESPMSG"), request("iPAYMENT_MODULE"), "NULL", "NULL", FormatNumber(sAmount,2,,,0)
		' ProcessDeclinedTransaction "paypal", sOrderID, sResult, sPNREF, sRespMsg, request("iPAYMENT_MODULE"), "NULL", "NULL"
	End If

	' DESTORY OBJECTS
	'		Set client = Nothing
	Set oReturnCodes = Nothing

End Sub


'------------------------------------------------------------------------------
' string GetResponseValue( sResponse, sParamName )
'------------------------------------------------------------------------------
Function GetResponseValue( ByVal sResponse, ByVal sParamName )
	Dim curString, name, value, varString, MyValue

	curString = sResponse
	MyValue = ""

	Do While Len(curString) <> 0

		If InStr(curString,"&") Then
  			varString = Left(curString, InStr(curString , "&" ) -1)
 		Else 
  			varString = curString
 		End If 
 
 		name = Left(varString, InStr(varString, "=" ) -1)
 		value = Right(varString, Len(varString) - (Len(name)+1))

  		If UCase(name) = UCase(sParamName) Then 
  			MyValue = value
 			Exit Do
  		End If 
  	
  		If Len(curString) <> Len(varString) Then 
  			curString = Right(curString, Len(curString) - (Len(varString)+1))
  		Else 
  			curString = ""
  		End If 
 
	Loop

	GetResponseValue = MyValue

End Function


'------------------------------------------------------------------------------
' string fn_OrderStringtoHTML( sString )
'------------------------------------------------------------------------------
Function fn_OrderStringtoHTML( ByVal sString )

	sReturnValue = "Not Able to Parse"

	arrItems = SPLIT(sString,"||")

	For i=0 to UBOUND(arrItems)-1
		arrDetails = SPLIT(arrItems(i),"~")
		For j=0 to UBOUND(arrDetails)
			response.write arrDetails(j) & "<br />"
		Next
	Next

	fn_OrderStringtoHTML = sReturnValue

End Function


'------------------------------------------------------------------------------
' Function fn_OrderStringtoHTML( sString ) - defunct
'------------------------------------------------------------------------------
'Function fn_AddPaymentInformation(iPaymentID,sPaymentRef,sAmount,sAuthcode)
'
'	'AddtoLog(iPaymentRefID & " - UPDATING PAYMENT TABLE...")
'	
'	blnAddNew = False
'
'	' INSERT FORM INFORMATION INTO DATABASE	
'	Set oPayment = Server.CreateObject("ADODB.Recordset")
'	oPayment.CursorLocation = 3
'	sSql = "SELECT * FROM egov_payments WHERE paymentid='" & iPaymentID & "'"
'	oPayment.Open sSql, Application("DSN") , 1, 3
'	
'	' CHECK FOR NEW RECORD
'	If oPayment.EOF Then
'		' ADD NEW RECORD
'		oPayment.AddNew
'		oPayment("orgid") = iOrgID
'		oPayment("paymentrefid") = sPaymentRef
'		oPayment("userid") = AddUserInformation()
'		blnAddNew = True
'	End If
'
'	' UPDATE PAYMENT DATABASE FIELDS
'	oPayment("paymentamount") = sAmount
'	oPayment("paymentstatus") = "COMPLETED"
'	oPayment("paymentrefid") = sPaymentRef
'	oPayment("userid") = AddUserInformation()
'	oPayment("orgid") = iOrgID
'	oPayment.Update
'	oPayment.Close
'
'	
'	' ADD RAW TRANSACTION DATA
'	AddPaymentInformation iPaymentID,sAuthcode,sPaymentRef 
'
'	' Send Email
'	'On Error Resume Next
'	SendEmail iPaymentID,sAmount,sAuthcode,sPaymentRef
'	Set oPayment = Nothing
'
'End Function


'------------------------------------------------------------------------------
' Function fn_OrderStringtoHTML( sString ) - defunct 
'------------------------------------------------------------------------------
Function AddPaymentInformation(iPaymentID,sAuthcode,sPaymentRef)

	sCompleteData = ""
	sCompleteData = sCompleteData & "Payment Reference Number:  " &  sPaymentRef & "<br />"
	sCompleteData = sCompleteData & "Authorization Code:  " &  sAuthcode & "<br />"

	Set oDetails = Server.CreateObject("ADODB.Recordset")
	oDetails.CursorLocation = 3
	sSql = "SELECT * FROM egov_paymentdetails where paymentid='" & iPaymentID & "'"
	oDetails.Open sSql, Application("DSN") , 1, 3
	If oDetails.EOF Then
		oDetails.AddNew
		oDetails("paymentid") = iPaymentID
		oDetails("paymentsummary") = dbsafe(sCompleteData)
		oDetails.Update
	End If 
	oDetails.Close

End Function


'------------------------------------------------------------------------------
' Function AddUserInformation( ) - defunct
'------------------------------------------------------------------------------
Function AddUserInformation()

	'AddtoLog(Request("txn_id")& " - UPDATING USER INFORMATION...")

	' INSERT FORM INFORMATION INTO DATABASE	
	iReturnValue = 0
	
	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.CursorLocation = 3
	oUser.Open "egov_users", Application("DSN") , 1, 2, 2
	oUser.AddNew
	oUser("userfname")   = dbsafe(request("sjname"))
	'oUser("userlname")   = dbsafe(request("lastname"))
	oUser("useraddress") = dbsafe(request("streetaddress"))
	oUser("usercity")    = dbsafe(request("city"))
	oUser("userstate")   = dbsafe(request("state"))
	oUser("userzip")     = dbsafe(request("zipcode"))
	oUser.Update
	iReturnValue = oUser("userid")
	oUser.Close

	Set oUser = Nothing

	AddUserInformation = iReturnValue

End Function


'------------------------------------------------------------------------------
' string DBsafe( sString )
'------------------------------------------------------------------------------
Function DBsafe( ByVal sString )

  If Not VarType( sString ) = vbString Then DBsafe = sString : Exit Function
  DBsafe = Replace( sString, "'", "''" )

End Function


'------------------------------------------------------------------------------
' void SendPurchaseEmail iPaymentModule, sAmount, sAuthcode, sPNREF, sOrderNumber, sSVA, dFeeAmount
'------------------------------------------------------------------------------
Sub SendPurchaseEmail( ByVal iPaymentModule, ByVal sAmount, ByVal sAuthcode, ByVal sPNREF, ByVal sOrderNumber, ByVal sSVA, ByVal dFeeAmount )
	Dim sSql, oRs, oCdoAdminMail, oCdoAdminConf, oCdoCitizenMail, oCdoCitizenConf, sSubject, sThankyou

	Select Case iPaymentModule
		Case 1
			sSubject = " Gift Purchase"
		Case 2
			sSubject = " Membership Purchase"
		Case 3
			sSubject = " Facility Reservation"
		Case Else
			sSubject = " E-Gov Payment Submission"		' These do not go through this process
	End Select
	
	' CONNECT TO DATABASE AND GET PAYMENT INFORMATION
	sSql = "SELECT * FROM dbo.egov_paymentservices WHERE orgid = " & iOrgID & " AND paymentservice_type = " & iPaymentModule 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If oRs("assigned_email") = "" Or IsNull(oRs("assigned_email")) then
		'adminEmailAddr = "noreply@eclink.com" ' NEED TO HAVE A DEFAULT INSTITUTION EMAIL ADDRESS
		If sDefaultEmail <> "" Then 
			adminEmailAddr = sDefaultEmail  ' City default email
		Else 
			adminEmailAddr = "noreply@eclink.com"
		End If 
	Else 
		adminEmailAddr = oRs("assigned_email") ' ASSIGNED ADMIN USER EMAIL
	End If 

	oRs.Close
	Set oRs = Nothing 

	' Build the email message to the admin person
	sMsg2 = sMsg2 & "<p>This automated message was sent by the " & sOrgName & " E-Gov web site. Do not reply to this message.  Contact " & adminEmailAddr & " for inquiries regarding this email.</p>" & vbcrlf  & vbcrlf
	sMsg2 = sMsg2 & "<p>A" & LCase(sSubject) & "  was created on " & Date() & ".</p>" & vbcrlf  & vbcrlf 

	sMsg2 = sMsg2 & "<p><strong>Transaction Details:</strong>" & vbcrlf
	If sSVA = "NULL" Then
		' This is PayPal
		sMsg2 = sMsg2 & "<br />Amount Charged: " & FormatCurrency(sAmount,2) & vbcrlf
		sMsg2 = sMsg2 & "<br />Payment Reference Number: " & sPNREF & vbcrlf
		sMsg2 = sMsg2 & "<br />Authorization Code:" & sAUTHCODE & vbcrlf & vbcrlf
	Else
		' This is Point and Pay
		sMsg2 = sMsg2 & "<br />Purchase Amount: " & FormatCurrency(sAmount,2) & vbcrlf
		sMsg2 = sMsg2 & "<br />Processing Fee: " & FormatCurrency(dFeeAmount,2) & vbcrlf
		sMsg2 = sMsg2 & "<br />Amount Charged: " & FormatCurrency((CDbl(sAmount) + CDbl(dFeeAmount)),2) & vbcrlf
		sMsg2 = sMsg2 & "<br />Order Number: " & sOrderNumber & vbcrlf
		sMsg2 = sMsg2 & "<br />SVA:" & sSVA & vbcrlf & vbcrlf
	End If 

	sMsg2 = sMsg2 & "<br />Product Information" & vbcrlf
	sMsg2 = sMsg2 & "<br />Payment: "  & request("paymentname") & vbcrlf
	sMsg2 = sMsg2 & "</p><p><strong>Details: </strong><br />" & vbtab & request("details") & "</p>" & vbcrlf 

	sMsg2 = sMsg2 & "<p><strong>Payment Information:</strong>" & vbcrlf
	sMsg2 = sMsg2 & "<br />Credit Card: XXXXXXXXXXXX" & Right(request("accountnumber"),4)  & vbcrlf
	sMsg2 = sMsg2 & "<br />Name: "  & request("sjname") & vbcrlf
	sMsg2 = sMsg2 & "<br />Email: "  & request("email") & vbcrlf
	sMsg2 = sMsg2 & "<br />Address: "  & request("streetaddress") & vbcrlf
	sMsg2 = sMsg2 & "<br />City: "  & request("city") & vbcrlf
	sMsg2 = sMsg2 & "<br />State: "  & request("state") & vbcrlf
	sMsg2 = sMsg2 & "<br />Zip: "  & request("zipcode") & "</p>" & vbcrlf & vbcrlf
	
	If iPaymentModule = 3 Then
		' Include Lessee information'
		sMsg2 = sMsg2 & "<p><strong>Lessee Information:</strong>" & vbcrlf
		' request("iFacilityPaymentID") is the facility schedule id. We can get lessee information from that '
		sMsg2 = sMsg2 & GetLeseeInformation( request("iFacilityPaymentID") )
		sMsg2 = sMsg2 & "</p>" & vbcrlf
		
		' put a link to the reservation'
		sMsg2 = sMsg2 & "<p><strong>View Reservation:</strong><br />" & vbcrlf
		sLinkUrl = sEgovWebsiteURL & "/admin/recreation/facility_reservation_edit.asp?ireservationid=" & request("iFacilityPaymentID")
		sMsg2 = sMsg2 & "<a href=""" & sLinkUrl & """>" & sLinkUrl & "</a>"
		sMsg2 = sMsg2 & "</p>" & vbcrlf
	End If 
	
	sendEmail "", adminEmailAddr, "", sOrgName & sSubject, sMsg2, clearHTMLTags(sMsg2), "N"
	

	' Build the emails message to the citizen
	sMsg3 = sMsg3 & "<p>This automated message was sent by the " & sOrgName & " E-Gov web site. Do not reply to this message.  Contact " & adminEmailAddr & " for inquiries regarding this email.</p>" & vbcrlf  & vbcrlf
	sMsg3 = sMsg3 & "<p>Thank you for your" & LCase(sSubject) & " on " & Date() & ".</p>" & vbcrlf  & vbcrlf 
	If iPaymentModule = 3 And OrgHasDisplay( iOrgid, "facility receipt notes top" ) Then
		sMsg3 = sMsg3 & vbcrlf & vbcrlf & GetOrgDisplay( iOrgid, "facility receipt notes top" )
	End If
	sMsg3 = sMsg3 & vbcrlf & vbcrlf & "<p><strong>Transaction Details:</strong>" & vbcrlf
	If sSVA = "NULL" Then
		sMsg3 = sMsg3 & "<br />Amount Charged: " & FormatCurrency(sAmount,2) & vbcrlf
		sMsg3 = sMsg3 & "<br />Payment Reference Number: " & sPNREF & vbcrlf
		sMsg3 = sMsg3 & "<br />Authorization Code:" & sAUTHCODE & vbcrlf & vbcrlf
	Else
		' This is Point and Pay
		sMsg3 = sMsg3 & "<br />Purchase Amount: " & FormatCurrency(sAmount,2) & vbcrlf
		sMsg3 = sMsg3 & "<br />Processing Fee: " & FormatCurrency(dFeeAmount,2) & vbcrlf
		sMsg3 = sMsg3 & "<br />Amount Charged: " & FormatCurrency((CDbl(sAmount) + CDbl(dFeeAmount)),2) & vbcrlf
		sMsg3 = sMsg3 & "<br />Order Number: " & sOrderNumber & vbcrlf
		sMsg3 = sMsg3 & "<br />SVA:" & sSVA & vbcrlf & vbcrlf
	End If 

	sMsg3 = sMsg3 & "<br />Product Information" & vbcrlf
	sMsg3 = sMsg3 & "<br />Payment: "  & request("paymentname") & vbcrlf
	sMsg3 = sMsg3 & "</p><p><strong>Details: </strong><br />" & vbcrlf & request("details") & vbcrlf & vbcrlf 
	sMsg3 = sMsg3 & "</p><p><strong>Payment Information:</strong>" & vbcrlf
	sMsg3 = sMsg3 & "<br />Credit Card: XXXXXXXXXXXX" & Right(request("accountnumber"),4)  & vbcrlf
	sMsg3 = sMsg3 & "<br />Name: "  & request("sjname") & vbcrlf
	sMsg3 = sMsg3 & "<br />Address: "  & request("streetaddress") & vbcrlf
	sMsg3 = sMsg3 & "<br />City: "  & request("city") & vbcrlf
	sMsg3 = sMsg3 & "<br />State: "  & request("state") & vbcrlf
	sMsg3 = sMsg3 & "<br />Zip: "  & request("zipcode") & "</p>" & vbcrlf & vbcrlf
	If iPaymentModule = 3 And OrgHasDisplay( iOrgid, "facility receipt notes bottom" ) Then
		sMsg3 = sMsg3 & vbcrlf & vbcrlf &  GetOrgDisplay( iOrgid, "facility receipt notes bottom" )
	End If

	sendEmail "", request("email"), "", "Thank You For Your " & sOrgName & sSubject, sMsg3, clearHTMLTags(sMsg3), "N"

End Sub


'------------------------------------------------------------------------------
' void ProcessSuccessfulTransaction( sOrderID, sAUTHCODE, sPNREF, sRESPMSG, sAmount, iPaymentModule, sCVV2Match, sOrderNumber, sSVA, dFeeAmount )
'------------------------------------------------------------------------------
Sub ProcessSuccessfulTransaction( ByVal sOrderID, ByVal sAUTHCODE, ByVal sPNREF, ByVal sRESPMSG, ByVal sAmount, ByVal iPaymentModule, ByVal sCVV2Match, ByVal sOrderNumber, ByVal sSVA, ByVal dFeeAmount )

	Select Case clng(iPaymentModule)
		Case 1
			' COMMEMORATIVE GIFT
			AddToPaymentLog iPaymentControlNumber, "Commemorative Gift Purchase"
			AddToPaymentLog iPaymentControlNumber, "iGiftPaymentId = " & request("iGiftPaymentId")
			UpdateGiftPayment request("iGiftPaymentId"), sAUTHCODE, sPNREF, "APPROVED", sRESPMSG, "1","3", sOrderNumber, sSVA, dFeeAmount
			SendPurchaseEmail iPaymentModule, sAmount, sAuthcode, sPNREF, sOrderNumber, sSVA, dFeeAmount

		Case 2
			' POOL PASS
			AddToPaymentLog iPaymentControlNumber, "Membership Purchase"
			AddToPaymentLog iPaymentControlNumber, "iPoolPassId = " & request("itemnumber")
			UpdatePoolPass request("itemnumber"), sAUTHCODE, sPNREF, "APPROVED", sRESPMSG, sOrderNumber, sSVA, dFeeAmount
			SendPurchaseEmail iPaymentModule, sAmount, sAuthcode, sPNREF, sOrderNumber, sSVA, dFeeAmount

		Case 3
			'FACILITY RESERVATION
			AddToPaymentLog iPaymentControlNumber, "Facility Reservation"
			AddToPaymentLog iPaymentControlNumber, "iFacilityPaymentID = " & request("iFacilityPaymentID")
			UpdateFacilityPayment request("iFacilityPaymentID"), sAUTHCODE, sPNREF, "APPROVED", sRESPMSG, "3", "1", sOrderNumber, sSVA, dFeeAmount
			CleanupUnusableFacilities request("iFacilityPaymentID")
			SendPurchaseEmail iPaymentModule, sAmount, sAuthcode, sPNREF, sOrderNumber, sSVA, dFeeAmount

		Case Else
			AddToPaymentLog iPaymentControlNumber, "Unknown Item Purchased"
			' JUST DEFAULT LEFT TO COMPARE AGAINST JS 1/31/2006
			' ADD INFORMATION TO ADMINSTRATION DATABASE
			iPaymentID = Right(request("ordernumber"),Len(request("ordernumber"))-InStr(request("ordernumber"),"O"))
			'fn_AddPaymentInformation iPaymentID,sPNREF,formatcurrency(sAmount,2),sAuthcode

	End Select

	' DISPLAY SUCCESS MESSAGE TO CUSTOMER
	' BRANDING
	response.write "<center><p>"
	sPaymentImg = GetPaymentImage( "../" )
	If sPaymentImg <> "" Then 
		response.write "<img src=""" & sPaymentImg & """ border=""0"" /><br /><br />"
	End If 
	response.write "<strong>" & GetPaymentGatewayName( ) & "</strong> has routed, processed, and secured your payment information."
	response.write "</p>"

	'BEGIN: Transaction result details --------------------------------------------
	response.write "<p><div class=""group""><p>Your credit card purchase was <strong>approved</strong>.<br /> You will receive a confirmation email containing this receipt.  It is also recommended that you print this page as proof of your purchase." & vbcrlf

	If iPaymentModule = 2 Then 
		'Check for "edit display" to override default message.
		If OrgHasDisplay(iorgid, "poolpass_processpayment_paragraph") Then 
			lcl_processpayment_msg = GetOrgDisplay(iorgid, "poolpass_processpayment_paragraph")
		Else 
			lcl_processpayment_msg = ""
			lcl_processpayment_msg = lcl_processpayment_msg & "Please bring this receipt to the pool to receive your pass. "
			lcl_processpayment_msg = lcl_processpayment_msg & "All family members must be present to have their photograph taken. "
			lcl_processpayment_msg = lcl_processpayment_msg & "If this is a renewal from a previous year, please bring old passes so they can be validated."
		End If 

  		response.write "<br /><br />" & lcl_processpayment_msg
	End If 

	response.write "</p><blockquote>"
	If iPaymentModule = 3 And OrgHasDisplay( iOrgid, "facility receipt notes top" ) Then 
  		response.write  GetOrgDisplay( iOrgid, "facility receipt notes top" )
	End If 

	response.write "<table>"
	response.write "<tr><td colspan=""2""><strong>Transaction Details</strong></td></tr>"
	If sSVA = "NULL" Then 
		' This is Pay Pal
		response.write "<tr>"
		response.write "<td align=""right"">Amount Charged:</td>"
		response.write "<td align=""left""> " & FormatCurrency(sAmount,2) & "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td align=""right"">Payment Reference Number:</td>"
		response.write "<td align=""left""> " & sPNREF & "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td align=""right"">Authorization Code:</td>"
		response.write "<td align=""left""> " & sAUTHCODE & "</td>"
		response.write "</tr>" 
	Else
		' This is Point and Pay
		response.write "<tr>"
		response.write "<td align=""right"">Purchase Amount:</td>"
		response.write "<td align=""left""> " & FormatCurrency(sAmount,2) & "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td align=""right"">Processing Fee:</td>"
		response.write "<td align=""left""> " & FormatCurrency(CDbl(dFeeAmount),2) & "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td align=""right"">Total Charged:</td>"
		response.write "<td align=""left""> " & FormatCurrency((CDbl(sAmount) + CDbl(dFeeAmount)),2) & "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td align=""right"">Order Number:</td>"
		response.write "<td align=""left""> " & sOrderNumber & "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td align=""right"">SVA:</td>"
		response.write "<td align=""left""> " & sSVA & "</td>"
		response.write "</tr>" 
	End If 
	response.write "</table>"
	'END: Transaction result details ----------------------------------------------

	'BEGIN: Product information ---------------------------------------------------
	response.write "<p>" & vbcrlf
	response.write "<table>" & vbcrlf
	response.write "<tr><td colspan=""2""><strong>Product Information</strong></td></tr>" & vbcrlf
	'response.write "<tr><td><font color=#000000>iPAYMENT_MODULE:</font></td><td> "&request("iPAYMENT_MODULE")&" </td></tr>"
	response.write "<tr>" & vbcrlf
	response.write "<td align=""right""><font color=""#000000"">Item Number:</font></td>" & vbcrlf
	response.write "<td align=""left""> "&request("itemnumber")&" </td>" & vbcrlf
	response.write "</tr>" & vbcrlf
	response.write "<tr>" & vbcrlf
	response.write "<td align=""right"">Payment: </td>" & vbcrlf
	response.write "<td align=""left"">" & request("paymentname") & "</td>" & vbcrlf
	response.write "</tr>" & vbcrlf
	response.write "<tr>" & vbcrlf
	response.write "<td valign=""top"" align=""right"">Details: </td>" & vbcrlf
	response.write "<td valign=""top"" align=""left"">" & replace(request("details"),"</br>,","") & "</td>" & vbcrlf
	response.write "</tr>" & vbcrlf
	response.write "</table>" & vbcrlf

	If iPaymentModule = 3 And OrgHasDisplay( iOrgId, "facility receipt notes" ) Then 
  		response.write "<p align=""left"">" & vbcrlf
  		response.write GetOrgDisplay( iOrgId, "facility receipt notes" )
		  response.write "</p>" & vbcrlf
	End If 
	'END: Product Information -----------------------------------------------------

	' CREDIT CARD INFORMATION	
	response.write "<p><table>"
	response.write "<tr><td colspan=""2""><strong>User Information</strong></td></tr>"
	response.write "<tr><td align=""right"">Credit Card: </td><td align=""left"">XXXXXXXXXXXX" & Right(request("accountnumber"),4)  & "</td></tr>"
	response.write "<tr><td align=""right"">Name: </td><td align=""left"">" & request("sjname") & "</td></tr>"
	response.write "<tr><td align=""right"">Address: </td><td align=""left"">" & request("streetaddress") & "</td></tr>"
	response.write "<tr><td align=""right"">City: </td><td align=""left"">" & request("city") & "</td></tr>"
	response.write "<tr><td align=""right"">State: </td><td align=""left"">" & request("state") & "</td></tr>"
	response.write "<tr><td align=""right"">Zip: </td><td align=""left"">" & request("zipcode") & "</td></tr>"
	response.write "</table></p>"
	If iPaymentModule = 3 And OrgHasDisplay( iOrgid, "facility receipt notes bottom" ) Then
		response.write  GetOrgDisplay( iOrgid, "facility receipt notes bottom" )
	End If 
	response.write "</blockquote></div></p></center>"

	AddToPaymentLog iPaymentControlNumber, "Purchase Complete."

End Sub

'------------------------------------------------------------------------------
' void ProcessDeclinedTransaction( sProcessingPath, sOrderID, sRESULT, sPNREF, sRESPMSG, iPaymentModule, sOrderNumber, sSVA, sAmount )
'------------------------------------------------------------------------------
Sub ProcessDeclinedTransaction( ByVal sProcessingPath, ByVal sOrderID, ByVal sRESULT, ByVal sPNREF, ByVal sRESPMSG, ByVal iPaymentModule, ByVal sOrderNumber, ByVal sSVA, ByVal sAmount )
	
	Select Case iPaymentModule
		Case 1
			' COMMEMORATIVE GIFT
			' NO ACTION NEEDED

		Case 2
			' POOL PASSES
			UpdatePoolPass request("itemnumber"), "AUTHCODE", "PNREF", "DECLINED", "RESPMSG", "NULL", sSVA, "NULL"

		Case 3
			' FACILITY RESERVATIONS

		Case Else
			' JUST DEFAULT LEFT TO COMPARE AGAINST JS 1/31/2006
			' UPDATE DATABASE WITH RESULT
			'Set oUpdate = Server.CreateObject("ADODB.Recordset")
			'sSql = "UPDATE orders SET datebilled='" & now() & "', status='INCOMPLETE',payflow_authorizationcode='XXXX-XXXX-XXXX',payflow_result=0,payflow_respmsg='"&sRESPMSG&"',payflow_pnref='"&sPNREF&"' WHERE orderid=" & sOrderID
			'oUpdate.Open sSql, Application("DSN") , 3, 1
			'Set oUpdate = Nothing

	End Select
		
	' BRANDING
	'response.write "<center><p><a href=""http://seal.verisign.com/payment"" TARGET=""_VERISIGN"" ><img vspace=10 border=0 hspace=20 src=""images/verisign.gif""></a><br />VeriSign has routed, processed, and secured your payment information. <A HREF="" http://www.verisign.com/products-services/payment-processing/index.html"">More information about VeriSign</a></p>"
	response.write "<center><p>"
	sPaymentImg = GetPaymentImage( "../" )
	If sPaymentImg <> "" Then 
		response.write "<img src=""" & sPaymentImg & """ border=""0"" /><br /><br />"
	End If 
	response.write "<strong>" & GetPaymentGatewayName( ) & "</strong> has routed, processed, and secured your payment information."
	response.write "</p>"

	' TRANSACTION RESULT DETAILS
	response.write "<p><div class=""group""><p>Your credit card purchase was <strong>declined</strong> for the following reason:"
	response.write "<blockquote><strong> ( " & sRESULT & " ) - " & sRESPMSG & "</strong></blockquote>"

	response.write "<blockquote>"
	response.write "<p><table>"
	response.write "<tr><td colspan=""2""><strong>Transaction Details</strong></td></tr>"
	If sSVA = "NULL" Then 
		' This is PayPal
		response.write "<tr><td align=""right"">Purchase Amount:</td><td align=""left"">" & FormatCurrency(sAmount,2) & "</td></tr>"
		response.write "<tr><td align=""right"">Payment Reference Number:</td><td align=""left""> " & sPNREF & " </td></tr>"
		response.write "<tr><td align=""right"">Authorization Code:</td><td align=""left"">" & sAUTHCODE & "</td></tr>"
	Else
		response.write "<tr><td align=""right"">Purchase Amount:</td><td align=""left"">" & FormatCurrency(sAmount,2) & "</td></tr>"
		response.write "<tr><td align=""right"">Order Number:</td><td align=""left""> " & sOrderNumber & " </td></tr>"
		response.write "<tr><td align=""right"">SVA:</td><td align=""left"">" & sSVA & "</td></tr>"
	End If 
	response.write "</table></p>"

	' PRODUCT INFORMATION
	response.write "<p><table >"
	response.write "<tr><td colspan=""2""><strong>Product Information</strong></td></tr>"
	response.write "<tr><td align=""right"">Item Number:</td><td align=""left""> "&request("itemnumber")&" </td></tr>"
	response.write "<tr><td align=""right"">Payment: </td><td align=""left"">" & request("paymentname") & "</td></tr>"
	response.write "<tr><td valign=""top"" align=""right"">Details: </td><td valign=""top"" align=""left"">" & Replace(request("details"),"</br>,","") & "</td></tr>"
	response.write "</table></p>"

	' CREDIT CARD INFORMATION	
	response.write "<p><table>"
	response.write "<tr><td colspan=""2""><strong>User Information</strong></td></tr>"
	response.write "<tr><td>Credit Card: </td><td>XXXXXXXXXXXX" & Right(request("accountnumber"),4)  & "</td></tr>"
	response.write "<tr><td>Name: </td><td>" & request("sjname") & "</td></tr>"
	response.write "<tr><td>Address: </td><td>" & request("streetaddress") & "</td></tr>"
	response.write "<tr><td>City: </td><td>" & request("city") & "</td></tr>"
	response.write "<tr><td>State: </td><td>" & request("state") & "</td></tr>"
	response.write "<tr><td>Zip: </td><td>" & request("zipcode") & "</td></tr>"
	response.write "</table></p></blockquote></center>"

End Sub


'------------------------------------------------------------------------------
' void ProcessCommunicationError()
'------------------------------------------------------------------------------
Sub ProcessCommunicationError()

	' DISPLAY COMMUNICATION MESSAGE TO CUSTOMER
    response.write "<div class=payflowmsgfail>Your credit card purchase was unable to processed because of a network communication error:<blockquote><font color=#000000>DSI Order Number:</font> " & sOrderID & "<br /><font color=#000000>Payment Reference Number:</font> "&sPNREF&" <br /><font color=#000000>Description:</font> ("&sRESULT&") - "&sRESPMSG&" </blockquote></div>"
	
End Sub


'------------------------------------------------------------------------------
' void UpdatePoolPass( iPoolPassId, sAuthCode, sPNRef, sResult, sRespMsg, sOrderNumber, sSVA, dFeeAmount )
'------------------------------------------------------------------------------
Sub UpdatePoolPass( ByVal iPoolPassId, ByVal sAuthCode, ByVal sPNRef, ByVal sResult, ByVal sRespMsg, ByVal sOrderNumber, ByVal sSVA, ByVal dFeeAmount )
	Dim sSql 

	If sSVA <> "NULL" Then
		sSVA = "'" & dbready_string(sSVA, 50) & "'"
	End If

	If sAuthCode <> "NULL" Then
		sAuthCode = "'" & dbready_string(sAuthCode, 50) & "'"
	End If 

	If sPNRef <> "NULL" Then
		sPNRef = "'" & dbready_string(sPNRef, 50) & "'"
	End If 

	If sResult <> "NULL" Then 
		sResult = "'" & dbready_string(sResult, 50) & "'"
	End If 

	If sRespMsg <> "NULL" Then
		sRespMsg = "'" & dbready_string(sRespMsg, 50) & "'"
	End If 

	sSql = "UPDATE dbo.egov_poolpasspurchases "
	sSql = sSql & "SET paymentauthcode = " & sAuthCode 
	sSql = sSql & ", paymentpnref = " & sPNRef
	sSql = sSql & ", paymentresult = " & sResult 
	sSql = sSql & ", paymentrespmsg = " & sRespMsg  
	sSql = sSql & ", sva = " & sSVA
	sSql = sSql & ", processingfee = " & dFeeAmount
	sSql = sSql & ", ordernumber = " & sOrderNumber
	sSql = sSql & " WHERE poolpassid = " & iPoolPassId

	RunSQLStatement sSql
	
'	Dim oCmd
'	'	response.write "<p>In the UpdatePoolPass - " & iPoolPassId & "</p>"
'
'	Set oCmd = Server.CreateObject("ADODB.Command")
'	With oCmd
'		.ActiveConnection = Application("DSN")
'		.CommandText = "UpdatePoolPassPurchase"
'		.CommandType = 4
'		.Parameters.Append oCmd.CreateParameter("@PoolPassId", 3, 1, 4, iPoolPassId)
'		.Parameters.Append oCmd.CreateParameter("@PaymentAuthCode", 200, 1, 50, sAuthCode)
'		.Parameters.Append oCmd.CreateParameter("@PaymentPNRef", 200, 1, 50, sPNRef)
'		.Parameters.Append oCmd.CreateParameter("@PaymentResult", 200, 1, 50, sResult)
'		.Parameters.Append oCmd.CreateParameter("@PaymentRespMsg", 200, 1, 255, sRespMsg)
'		.Execute
'	End With
'
'	Set oCmd = Nothing

End Sub 


'------------------------------------------------------------------------------
' void UpdateGiftPayment( iGiftPaymentId, sAuthCode, sPNRef, sResult, sRespMsg, itype, ilocation, sOrderNumber, sSVA, dFeeAmount )
'------------------------------------------------------------------------------
Sub UpdateGiftPayment( ByVal iGiftPaymentId, ByVal sAuthCode, ByVal sPNRef, ByVal sResult, ByVal sRespMsg, ByVal itype, ByVal ilocation, ByVal sOrderNumber, ByVal sSVA, ByVal dFeeAmount )
	' UpdateGiftPayment request("iGiftPaymentId"), sAUTHCODE, sPNREF, "APPROVED", sRESPMSG, "1","3", sOrderNumber, sSVA, dFeeAmount
	Dim sSql

	If sSVA <> "NULL" Then
		sSVA = "'" & dbready_string(sSVA, 50) & "'"
	End If

	If sAuthCode <> "NULL" Then
		sAuthCode = "'" & dbready_string(sAuthCode, 50) & "'"
	End If 

	If sPNRef <> "NULL" Then
		sPNRef = "'" & dbready_string(sPNRef, 50) & "'"
	End If 

	If sResult <> "NULL" Then 
		sResult = "'" & dbready_string(sResult, 50) & "'"
	End If 

	If sRespMsg <> "NULL" Then
		sRespMsg = "'" & dbready_string(sRespMsg, 50) & "'"
	End If 

	sSql = "UPDATE egov_gift_payment "
	sSql = sSql & "SET authcode = " & sAuthCode
	sSql = sSql & ", pnref = " & sPNRef
	sSql = sSql & ", result = " & sResult
	sSql = sSql & ", replymsg = " & sRespMsg
	sSql = sSql & ", paymenttype = 1"		' this is a constant
	sSql = sSql & ", paymentlocation = 3"	' this is a constant
	sSql = sSql & ", sva = " & sSVA
	sSql = sSql & ", processingfee = " & dFeeAmount
	sSql = sSql & ", ordernumber = " & sOrderNumber
	sSql = sSql & " WHERE  giftpaymentid = " & iGiftPaymentId
	
	RunSQLStatement sSql
	
'	Dim oCmd
'
'	Set oCmd = Server.CreateObject("ADODB.Command")
'	With oCmd
'		.ActiveConnection = Application("DSN")
'		.CommandText = "UpdateGiftPayment"
'		.CommandType = 4
'		.Parameters.Append oCmd.CreateParameter("@iGiftPaymentID", 3, 1, 4, iGiftPaymentId)
'		.Parameters.Append oCmd.CreateParameter("@sAuthCode", 200, 1, 50, sAuthCode)
'		.Parameters.Append oCmd.CreateParameter("@sPNRef", 200, 1, 50, sPNRef)
'		.Parameters.Append oCmd.CreateParameter("@sResult", 200, 1, 50, sResult)
'		.Parameters.Append oCmd.CreateParameter("@sReplyMsg", 200, 1, 255, sRespMsg)
'		.Parameters.Append oCmd.CreateParameter("@paymenttype", 200, 1, 4, itype)
'		.Parameters.Append oCmd.CreateParameter("@paymentlocation", 200, 1, 4, ilocation)
'		.Execute
'	End With
'
'	Set oCmd = Nothing

End Sub


'------------------------------------------------------------------------------
' void UpdateFacilityPayment( iFacilityPaymentId, sAuthCode, sPNRef, sResult, sRespMsg, itype, ilocation, sOrderNumber, sSVA, dFeeAmount )
'------------------------------------------------------------------------------
Sub UpdateFacilityPayment( ByVal iFacilityPaymentId, ByVal sAuthCode, ByVal sPNRef, ByVal sResult, ByVal sRespMsg, ByVal iPaymentLocation, ByVal iPaymentType, ByVal sOrderNumber, ByVal sSVA, ByVal dFeeAmount )
'   UpdateFacilityPayment request("iFacilityPaymentID"), sAUTHCODE, sPNREF, "APPROVED", sRESPMSG, "3", "1", sOrderNumber, sSVA, dFeeAmount
	Dim sSql 

	If sSVA <> "NULL" Then
		sSVA = "'" & dbready_string(sSVA, 50) & "'"
	End If

	If sAuthCode <> "NULL" Then
		sAuthCode = "'" & dbready_string(sAuthCode, 50) & "'"
	End If 

	If sPNRef <> "NULL" Then
		sPNRef = "'" & dbready_string(sPNRef, 50) & "'"
	End If 

	If sResult <> "NULL" Then 
		sResult = "'" & dbready_string(sResult, 50) & "'"
	End If 

	If sRespMsg <> "NULL" Then
		sRespMsg = "'" & dbready_string(sRespMsg, 50) & "'"
	End If 
	
	sSql = "UPDATE dbo.egov_facilityschedule "
	sSql = sSql & "SET authcode = " & sAuthCode
	sSql = sSql & ", pnref = " & sPNRef
	sSql = sSql & ", result = " & sResult
	sSql = sSql & ", replymsg = " & sRespMsg
	sSql = sSql & ", status = 'RESERVED'" 
	sSql = sSql & ", paymenttype = 1"   ' this is a constant
	sSql = sSql & ", paymentlocation = 3"    ' this is a constant
	sSql = sSql & ", sva = " & sSVA
	sSql = sSql & ", processingfee = " & dFeeAmount
	sSql = sSql & ", ordernumber = " & sOrderNumber
	sSql = sSql & " WHERE facilityscheduleid = " & iFacilityPaymentId

	RunSQLStatement sSql

'	Dim oCmd
'
'	Set oCmd = Server.CreateObject("ADODB.Command")
'
'	With oCmd
'		.ActiveConnection = Application("DSN")
'		.CommandText = "UpdateFacilityPayment"
'		.CommandType = 4
'		.Parameters.Append oCmd.CreateParameter("@iFacilityPaymentID", 3, 1, 4, iFacilityPaymentId)
'		If sAuthCode <> "NULL" Then 
'			.Parameters.Append oCmd.CreateParameter("@sAuthCode", 200, 1, 50, sAuthCode)
'		Else
'			.Parameters.Append oCmd.CreateParameter("@sAuthCode", 200, 1, 50, Null)
'		End If 
'		If sPNRef <> "NULL" Then 
'			.Parameters.Append oCmd.CreateParameter("@sPNRef", 200, 1, 50, sPNRef)
'		Else
'			.Parameters.Append oCmd.CreateParameter("@sPNRef", 200, 1, 50, Null)
'		End If 
'		.Parameters.Append oCmd.CreateParameter("@sResult", 200, 1, 50, sResult)
'		.Parameters.Append oCmd.CreateParameter("@sReplyMsg", 200, 1, 255, sRespMsg)
'		.Parameters.Append oCmd.CreateParameter("@paymenttype", 200, 1, 4, itype)
'		.Parameters.Append oCmd.CreateParameter("@paymentlocation", 3, 1, 4, ilocation)
'		.Parameters.Append oCmd.CreateParameter("@status", 200, 1, 50, "RESERVED")
'		If sSVA <> "NULL" Then 
'			.Parameters.Append oCmd.CreateParameter("@sva", 200, 1, 50, sSVA)
'		Else
'			.Parameters.Append oCmd.CreateParameter("@sva", 200, 1, 50, Null)
'		End If 
'		If dFeeAmount <> "NULL" Then 
'			.Parameters.Append oCmd.CreateParameter("@processingfee", 6, 1, , CDbl(dFeeAmount))
'		Else
'			.Parameters.Append oCmd.CreateParameter("@processingfee", 6, 1, , Null)
'		End If 
'		If sOrderNumber <> "NULL" Then 
'			.Parameters.Append oCmd.CreateParameter("@ordernumber", 3, 1, 4, sOrderNumber)
'		Else
'			.Parameters.Append oCmd.CreateParameter("@ordernumber", 3, 1, 4, Null)
'		End If 
'		.Execute
'	End With

'	Set oCmd = Nothing

End Sub


'------------------------------------------------------------------------------
' void CleanupUnusableFacilities iFacilityPaymentID 
'------------------------------------------------------------------------------
Sub CleanupUnusableFacilities( ByVal iFacilityPaymentID )
	Dim sSql, oRs

	'Get the facility information
	sSql = "SELECT checkindate, checkoutdate, facilitytimepartid, facilityid, result, status, lesseeid, sessionid "
	sSql = sSql & " FROM egov_facilityschedule "
	sSql = sSql & " WHERE facilityscheduleid = " & iFacilityPaymentID

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		sSql = "DELETE FROM egov_facilityschedule "
		sSql = sSql & " WHERE checkindate = '"     & oRs("checkindate")  & "' "
		sSql = sSql & " AND checkoutdate = '"      & oRs("checkoutdate") & "' "
		sSql = sSql & " AND facilitytimepartid = " & oRs("facilitytimepartid")
		sSql = sSql & " AND facilityid = "         & oRs("facilityid")
		sSql = sSql & " AND orgid = "              & iorgid
		sSql = sSql & " AND facilityscheduleid <> " & iFacilityPaymentID
		sSql = sSql & " AND (result = '' OR result is null) "
		sSql = sSql & " AND (status = '' OR status is null) "

		RunSQLStatement sSql 

		oRs.MoveNext
	Loop 

	oRs.Close 
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' string  GetLeseeInformation( iFacilityPaymentID )
'------------------------------------------------------------------------------
Function GetLeseeInformation( ByVal iFacilityPaymentID )
	Dim sSql, oRs, sReturn
	
	sReturn = ""

	'Get the Lessee information
	sSql = "SELECT ISNULL(U.userfname,'') AS userfname, ISNULL(U.userlname,'') AS userlname, R.description, ISNULL(U.useraddress,'') AS useraddress, "
	sSql = sSql & "ISNULL(U.usercity,'') AS usercity, ISNULL(U.userstate,'') AS userstate, ISNULL(U.userzip,'') AS userzip "
	sSql = sSql & "FROM egov_users U, egov_facilityschedule F, egov_poolpassresidenttypes R "
	sSql = sSql & "WHERE F.lesseeid = U.userid AND U.residenttype = R.resident_type AND R.orgid = " & iOrgid
	sSql = sSql & " AND F.facilityscheduleid = " & iFacilityPaymentID
	'response.write "Lessee sSql = " & sSql & "<br /><br />"
	
	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sReturn = sReturn & "<br />Name: "  & oRs("userfname") & " " & oRs("userlname") & vbcrlf
		sReturn = sReturn & "<br />Residency: "  & oRs("description") & vbcrlf
		sReturn = sReturn & "<br />Address: "  & oRs("useraddress") & vbcrlf
		sReturn = sReturn & "<br />City: "  & oRs("usercity") & vbcrlf
		sReturn = sReturn & "<br />State: "  & oRs("userstate") & vbcrlf
		sReturn = sReturn & "<br />Zip: "  & oRs("userzip") & "</p>" & vbcrlf & vbcrlf
	End If
	
	oRs.Close 
	Set oRs = Nothing 
	
	GetLeseeInformation = sReturn

End Function 


'------------------------------------------------------------------------------
' string GetVerisignOptions( iOrgId, sPayflowURL )
'------------------------------------------------------------------------------
Function GetVerisignOptions( ByVal iOrgId, ByRef sPayflowURL )
	Dim sSql, oRs

	GetVerisignOptions = ""

	sSql = "SELECT vendor, [user], password, partner, IsLive, isnull(liveurl,'') AS liveurl, isnull(devurl,'') AS devurl "
	sSql = sSql & " FROM egov_verisign_options where orgid = "  & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetVerisignOptions = "&PWD=" & oRs("password")  ' SET PAYFLOW PASSWORD
		GetVerisignOptions = GetVerisignOptions & "&USER=" & oRs("user")  ' SET PAYFLOW USER
		GetVerisignOptions = GetVerisignOptions & "&VENDOR=" & oRs("vendor") ' SET PAYFLOW VENDOR
		GetVerisignOptions = GetVerisignOptions & "&PARTNER=" & oRs("partner") ' SET PAYFLOW PARTNER
		If oRs("islive") Then 
			sPayflowURL = oRs("liveurl")
		Else
			sPayflowURL = oRs("devurl")
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function


'------------------------------------------------------------------------------
' integer CreatePaymentControlRow( sLogEntry )
'------------------------------------------------------------------------------
Function CreatePaymentControlRow( ByVal sLogEntry )
	Dim sSql, iPaymentControlNumber

	sSql = "INSERT INTO paymentlog ( orgid, applicationside, feature, logentry ) VALUES ( "
	sSql = sSql & iOrgID & ", 'public', 'recreation', '" & sLogEntry & "' )"
	'response.write sSql & "<br /><br />"

	iPaymentControlNumber = RunIdentityInsertStatement( sSql )

	sSql = "UPDATE paymentlog SET paymentcontrolnumber = " & iPaymentControlNumber
	sSql = sSql & " WHERE paymentlogid = " & iPaymentControlNumber
	'response.write sSql & "<br /><br />"
	RunSQLStatement( sSql )

	CreatePaymentControlRow = iPaymentControlNumber

End Function 


'------------------------------------------------------------------------------
' void AddToPaymentLog( iPaymentControlNumber, sLogEntry )
'------------------------------------------------------------------------------
Sub AddToPaymentLog( ByVal iPaymentControlNumber, ByVal sLogEntry  )
	Dim sSql

	sSql = "INSERT INTO paymentlog ( paymentcontrolnumber, orgid, applicationside, feature, logentry ) VALUES ( "
	sSql = sSql & iPaymentControlNumber & ", " & iOrgID & ", 'public', 'recreation', '" & dbready_string(sLogEntry, 500) & "' )"
	'response.write sSql & "<br /><br />"
	RunSQLStatement( sSql )

End Sub 


'------------------------------------------------------------------------------
Sub dtb_debug(p_value)
'	if p_value <> "" then
'		sSqli = "INSERT INTO my_table_dtb(notes) VALUES ('" & Now() & " " & Replace(p_value,"'","''") & "')"
'		set rsi = Server.CreateObject("ADODB.Recordset")
'		rsi.Open sSqli, Application("DSN"), 3, 1
'	end If
'	Set rsi = Nothing 
end Sub


'------------------------------------------------------------------------------
' void SendLoginFailedEmail
'------------------------------------------------------------------------------
Sub SendLoginFailedEmail( )
	Dim sHTMLBody, sTextTBody

	sTextTBody = "A User Authentication error has been received for " & sOrgName & " ( " & iOrgId & " ). "
	sTextTBody = sTextTBody & vbcrlf & "The client has changed their password and needs to update our database."

	sHTMLBody = "<p>A User Authentication error has been received for " & sOrgName & " ( " & iOrgId & " ). "
	sHTMLBody = sHTMLBody & "<br />The client has changed their password and needs to update our database.</p>"

	sendEmail "noreply@eclink.com", "egovsupport@eclink.com", "", "PayPal User Authentication Error Received", sHTMLBody, sTextTBody, "Y"
End Sub 


'----------------------------------------------------------------------------------------
' string CleanAndCutForPNPNotes( sParameter )
'----------------------------------------------------------------------------------------
Function CleanAndCutForPNPNotes( ByVal sParameter )
	' cleans forbidden characters and returns the string

	sParameter = Replace(sParameter, Chr(34), "")
	sParameter = Replace(sParameter, Chr(13), " ")
	sParameter = Replace(sParameter, Chr(10), "")
	sParameter = Replace(sParameter, "'", "")
	sParameter = Replace(sParameter, "&", "and")
	sParameter = Replace(sParameter, "=", "is")
	sParameter = Replace(sParameter, "</br>", "")
	sParameter = Replace(sParameter, "<br />", "")
	sParameter = Replace(sParameter, "<br>", "")
	sParameter = Replace(sParameter, "<b>", "")
	sParameter = Replace(sParameter, "</b>", "")
	sParameter = Replace(sParameter, "<strong>", "")
	sParameter = Replace(sParameter, "</strong>", "")
	sParameter = Replace(sParameter, "<i>", "")
	sParameter = Replace(sParameter, "</i>", "")
	sParameter = Replace(sParameter, ", ,", "")
	sParameter = Trim(sParameter)
	sParameter = Left(sParameter, 255)	' 255 characters is the PNP limit for notes

	CleanAndCutForPNPNotes = sParameter

End Function 


'----------------------------------------------------------------------------------------
' string facilityScheduleIsApproved( facilityScheduleId )
'----------------------------------------------------------------------------------------
Function facilityScheduleIsReserved( ByVal facilityScheduleId )
	Dim sSql, oRs
	
	sSql = "SELECT ISNULL(status, 'PENDING') AS status "
      sSql = sSql & " FROM egov_facilityschedule "
      sSql = sSql & " WHERE facilityscheduleid = " & facilityScheduleId

      Set oRs = Server.CreateObject("ADODB.Recordset")
      oRs.Open sSql, Application("DSN"), 0, 1

      If Not oRs.EOF Then
      	If oRs("status") = "RESERVED" THEN 
      		facilityScheduleIsReserved = true  
      	Else
      		facilityScheduleIsApproved = false 
      	End If 
      Else
      	' These missing ones should be filtered out before this check'
      	facilityScheduleIsReserved = false 
      End If
      
      oRs.Close
	Set oRs = Nothing 

End Function 




%>
