<!DOCTYPE HTML PUBLIC "-//W3C//DTD XHTML 1.1 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->
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
' 1.0	??/??/????	??? ??? - Initial Version
' 1.1	05/28/2009	Steve Loar - Changes for centralized PayPal processing and PayFlow Pro changes
' 2.0	08/02/2010	Steve Loar - Added Processing for Point and Pay
' 2.1	09/20/2011	Steve Loar - Redirecting back to payments when the paymentid is missing
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sError, iPaymentControlNumber, sOrderId, sTransactionAmount, sPaymentId

If request("ordernumber") <> "" Then 
	' pull out the payment id
	sPaymentId = Right(request("ordernumber"),Len(request("ordernumber"))-InStr(request("ordernumber"),"O"))
Else
	response.redirect "../../payment.asp"
End If 

'response.Expires = 60
'response.Expiresabsolute = Now() - 1
'response.AddHeader "pragma","no-store"
'response.AddHeader "cache-control","private"
'response.CacheControl = "no-store" 'HTTP prevent back button


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
			response.redirect "../../payment.asp"
		end if


%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<title>E-Gov Services <%=sOrgName%> - Payment Form</title>

	<link rel="stylesheet" type="text/css" href="../../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../../global.css" />
	<link rel="stylesheet" type="text/css" href="../../css/style_<%=iorgid%>.css" />

	<script language="Javascript" src="../../scripts/modules.js"></script>

	<style>
	<%
	  if request.servervariables("HTTPS") = "on" then
		 response.write "	body {behavior: url('https://secure.egovlink.com/" & sorgVirtualSiteName & "/csshover.htc');}" & vbcrlf
	  end if
	%>
	</style>

</head>


<!--#Include file="../../include_top.asp"-->

<!--BODY CONTENT-->
<tr><td valign="top">

<!--BEGIN: INTRO TEXT-->
<div class="title">
	<%=sWelcomeMessage%>
</div>

<div class="main">
	<font class="datetagline">Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%></font>
</div>
<!--END: INTRO TEXT-->

<!--BEGIN:  Process the payment-->

<% 

If GetPaymentStatus( sPaymentId ) <> "PROCESSING" Then 
	' This has already been processed so do not process again as it will charge the citizen again
	ShowReceipt sPaymentId
Else 
	' Process the payment transaction 
	'sTransactionAmount = Replace(request("transactionamount"), ",", "")  ' Get rid of any commas they may have entered into the payment amount.
	sTransactionAmount = request("transactionamount")

	If OrgHasFeature( iOrgId, "skippayment" ) Then 
		iPaymentControlNumber = CreatePaymentControlRow( "VERISIGN PAYMENT SCRIPT STARTED." )
		AddToPaymentLog iPaymentControlNumber, "TRANSACTION SUCCEEDED - Bypassed Authorization"
		approved = True
		sAuthcode = "010101"
		sPNREF = "V19F1D5C82TEST"
		sRespMsg = "Approved"
		AddToPaymentLog iPaymentControlNumber, "AUTHCODE: " & sAuthcode
		AddToPaymentLog iPaymentControlNumber, "PNREF: " & sPNREF
		AddToPaymentLog iPaymentControlNumber, "RESPMSG: " & sRespMsg
		ProcessSuccessfulTransaction sOrderID, sAuthcode, sPNREF, sRespMsg, sTransactionAmount, "NULL", "NULL", "NULL" 
		AddToPaymentLog iPaymentControlNumber, "VERISIGN PAYMENT PROCESSING FINISHED."
	Else 
		sProcessingRoute = GetProcessingRoute()		' In ../include_top_functions.asp

		Select Case sProcessingRoute
			Case "PayPalPayFlowPro"
				' Can put branching logic here to handle different payment processors
				ProcessPayPalTransaction request("firstname") & " " & request("lastname"), request("accountnumber"), request("month")&request("year"), sTransactionAmount 
			Case "PointAndPay"
				ProcessPointAndPayTransaction request("firstname"), request("lastname"), request("accountnumber"), request("month")&request("year"), sTransactionAmount 
		End Select 
	End If 
End If 
	
%>


<!--BEGIN: PAYMENT FOOTER-->
<center><br /><br />
	<input type="button" class="button" onClick="location.href='<%=sEgovWebsiteURL%>/';" value="Click here to return to the E-Government Website"><br />
</center>

<center>

	<p class="smallnote">
		NOTE: Your IP address [<%=request.servervariables("REMOTE_ADDR")%>] has been logged with this transaction.<br /><br />
		Do you have questions?<br />
<%		If CLng(iOrgid) <> CLng(153) Then		%>
			Contact us at <a href="mailto:<%=sDefaultEmail%>"><%=sDefaultEmail%></a> or <%=formatphonenumber(sDefaultPhone)%>.
<%		Else		
			' this is for Rye, NY
			If request("paymentformid") <> "" Then
				iPaymentFormID = request("paymentformid")
			Else
				iPaymentFormID = -1
			End If 
			ShowContactLine iPaymentFormID
		End If		%>

	</p>

	<p>&nbsp;</p>

</center>

<!--END: PAYMENT FOOTER-->


</div>


<!--END: DISPLAY RESPONSE-->


<!--SPACING CODE-->
<p>&nbsp;<bR>&nbsp;<bR>&nbsp;<bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->


<!--#Include file="../../include_bottom.asp"--> 

<%

'--------------------------------------------------------------------------------------------------
' void ShowContactLine iPaymentFormID
'--------------------------------------------------------------------------------------------------
Sub ShowContactLine( ByVal iPaymentFormID )
	Dim sSql, oRs, sEmail, sPhone

	sSql = "SELECT ISNULL(U.email,'') AS email, ISNULL(U.businessnumber,'') AS businessnumber "
	sSql = sSql & "FROM egov_paymentservices P, users U "
	sSql = sSql & "WHERE P.assigned_userID = U.userid AND P.paymentserviceid = " & iPaymentFormID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("email") <> "" Then 
			sEmail = oRs("email")
		Else
			sEmail = sDefaultEmail
		End If 

		If oRs("businessnumber") <> "" Then 
			sPhone = oRs("businessnumber")
		Else
			sPhone = FormatPhoneNumber(sDefaultPhone)
		End If 

		response.write "Contact Customer Service: <a href=""mailto:" & sEmail & """>" & sEmail & "</a> or " & sPhone
	End If 

	oRs.Close 
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------------------------------------
' FUNCTION FN_DISPLAYPAYMENTS() -- not called anywhere
'------------------------------------------------------------------------------------------------------------
Function fn_OrderStringtoHTML( sString )

	sReturnValue = "Not Able to Parse"

	arrItems = SPLIT(sString,"||")

	For i=0 to UBOUND(arrItems)-1
		arrDetails = SPLIT(arrItems(i),"~")
		For j=0 to UBOUND(arrDetails)
			response.write arrDetails(j) & "<br>"
		Next
	Next

	fn_OrderStringtoHTML = sReturnValue

End Function


'------------------------------------------------------------------------------------------------------------
' void AddPaymentInformation iPaymentID, sPaymentRef, sAmount, sAuthcode, sOrderNumber, sSVA, dFeeAmount 
' re-written 3/3/2010 Steve Loar
'------------------------------------------------------------------------------------------------------------
Sub AddPaymentInformation( ByVal iPaymentId, ByVal sPaymentRef, ByVal sAmount, ByVal sAuthcode, ByVal sOrderNumber, ByVal sSVA, ByVal dFeeAmount )
	Dim sSql, iUserId

	If sSVA <> "NULL" Then
		sSVA = "'" & dbready_string(sSVA, 50) & "'"
	End If

	If sPaymentRef <> "NULL" Then
		sPaymentRef = "'" & dbready_string(sPaymentRef, 50) & "'"
	End If 

	If request("userid") = "" Then 
		iUserId = AddUserInformation()
	Else
		iUserId = CLng(request("userid"))
	End If 

	If PaymentRecordExixts( iPaymentId ) Then
		sSql = "UPDATE egov_payments SET orgid = " & iOrgId
		sSql = sSql & ", paymentamount = " & sAmount
		sSql = sSql & ", paymentstatus = 'COMPLETED'"
		sSql = sSql & ", paymentrefid = " & sPaymentRef
		sSql = sSql & ", userid = " & iUserId
		sSql = sSql & ", ordernumber = " & sOrderNumber
		sSql = sSql & ", sva = " & sSVA
		sSql = sSql & ", processingfee = " & dFeeAmount
		sSql = sSql & " WHERE paymentid = " & iPaymentId

		RunSQLStatement sSql

		AddToPaymentLog iPaymentControlNumber, "egov_payments updated. paymentid: " & iPaymentId
	Else
		' I think there would be a problem if there is not a row already as the serviceid would be missing
		sSql = "INSERT INTO egov_payments ( orgid, paymentamount, paymentstatus, paymentrefid, userid, "
		sSql = sSql & "ordernumber, sva, processingfee ) VALUES ( "
		sSql = sSql & iOrgId & ", " & sAmount & ", 'COMPLETED', '" & sPaymentRef & "', " & iUserId & "', " 
		sSql = sSql & sOrderNumber & ", " & sSVA & ", " & dFeeAmount & " )"

		iPaymentId = RunIdentityInsertStatement( sSql )

		AddToPaymentLog iPaymentControlNumber, "egov_payments new record added. paymentid: " & iPaymentId
	End If 


	' ADD RAW TRANSACTION DATA
	AddTransactionInformation iPaymentID, sAuthcode, Replace(sPaymentRef,"'",""), sOrderNumber, Replace(sSVA,"'",""), dFeeAmount
	AddToPaymentLog iPaymentControlNumber, "AddTransactionInformation completed."

	' Send Email
	SendPurchaseEmail iPaymentID, sAmount, sAuthcode, Replace(sPaymentRef,"'",""), sOrderNumber, Replace(sSVA,"'",""), dFeeAmount 
	AddToPaymentLog iPaymentControlNumber, "SendPurchaseEmail completed."

End Sub 


'------------------------------------------------------------------------------------------------------------
' boolean PaymentRecordExixts( iPaymentId )
' Added 3/3/2010 Steve Loar
'------------------------------------------------------------------------------------------------------------
Function PaymentRecordExixts( ByVal iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(paymentid) AS hits FROM egov_payments WHERE paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If clng(oRs("hits")) > clng(0) Then
			PaymentRecordExixts = True 
		Else
			PaymentRecordExixts = False 
		End If 
	Else
		PaymentRecordExixts = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------------------------------------
' void AddTransactionInformation( iPaymentID, sAuthcode, sPaymentRef, sOrderNumber, sSVA, dFeeAmount  )
' Re-written: 3/3/2010 Steve Loar
'------------------------------------------------------------------------------------------------------------
Sub AddTransactionInformation( ByVal iPaymentId, ByVal sAuthcode, ByVal sPaymentRef, ByVal sOrderNumber, ByVal sSVA, ByVal dFeeAmount  )
	Dim sSql, sCompleteData

	If sSVA = "NULL" Then 
		sCompleteData = "Payment Reference Number:  " &  sPaymentRef & " <br>"
		sCompleteData = sCompleteData & "Authorization Code:  " &  sAuthcode & " <br>"
	Else
		sCompleteData = "Fee Amount: " & FormatNumber(dFeeAmount,2,,,0) & " <br />"
		sCompleteData = sCompleteData & "Order Number: " & sOrderNumber & " <br />"
		sCompleteData = sCompleteData & "SVA: " & sSVA & " <br />"
	End If 

	sSql = "INSERT INTO egov_paymentdetails ( paymentid, paymentsummary ) VALUES ( " & iPaymentId
	sSql = sSQl & ", '" & DBsafe( sCompleteData ) & "' )"

	RunSQLStatement sSql 

End Sub 


'------------------------------------------------------------------------------------------------------------
' integer AddUserInformation()
'------------------------------------------------------------------------------------------------------------
Function AddUserInformation()
	Dim sSql, iUserId

	sSql = "INSERT INTO egov_users ( userfname, useraddress, usercity, userstate, userzip ) VALUES ( '"
	sSql = sSql & dbsafe(request("sjname")) & "', '" & dbsafe(request("streetaddress")) & "', '" 
	sSql = sSql & dbsafe(request("city")) & "', '" & dbsafe(request("state")) & "', '" 
	sSql = sSql & dbsafe(request("zipcode")) & "' )"

	iUserId = RunIdentityInsertStatement( sSql )

	AddUserInformation = iUserId

End Function


'------------------------------------------------------------------------------------------------------------
' string DBsafe( strDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )
	Dim sNewString
	
	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function

	sNewString = Replace( strDB, "'", "''" )
	'sNewString = Replace( sNewString, "<", "&lt;" )
	DBsafe = sNewString

End Function


'------------------------------------------------------------------------------------------------------------
' void SendPurchaseEmail iID, sAmount, sAuthcode, sPNREF, sOrderNumber, sSVA, dFeeAmount
'------------------------------------------------------------------------------------------------------------
Sub SendPurchaseEmail( ByVal iPaymentId, ByVal sAmount, ByVal sAuthcode, ByVal sPNREF, ByVal sOrderNumber, ByVal sSVA, ByVal dFeeAmount )
	Dim sSql, oRs, details, sMsg2, sMsg3, adminEmailAddr, sPayInfo, oCdoMail, oCdoConf
	Dim oCdoBuyerMail, oCdoBuyerConf

	' CONNECT TO DATABASE AND GET PAYMENT INFORMATION
	sSql = "SELECT ISNULL(assigned_email,'') AS assigned_email, payment_information FROM egov_payment_list WHERE paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If oRs("assigned_email") = "" Or isNull(oRs("assigned_email")) then
		If sDefaultEmail <> "" Then 
			adminEmailAddr = sDefaultEmail  ' City default email
		Else 
			adminEmailAddr = "noreply@eclink.com"
		End If 
	Else 
		adminEmailAddr = oRs("assigned_email") ' ASSIGNED ADMIN USER EMAIL
	End If 

	' BUILD MESSAGE 
	sPayInfo = Replace(oRs("payment_information"),"<br>",vbcrlf)
	sPayInfo = Replace(oRs("payment_information"),"</br>",vbcrlf)

	oRs.Close
	Set oRs = Nothing 
	
	' build message to Admin
	sMsg2 = "This automated message was sent by the " & sOrgName & " E-Gov web site. Do not reply to this message.  Contact " & adminEmailAddr & " for inquiries regarding this email." & vbcrlf  & vbcrlf
	sMsg2 = sMsg2 & "<br />Payment was submitted on " & Date() & "." & vbcrlf  & vbcrlf 
	
	sMsg2 = sMsg2 & "<p><strong>Transaction Details</strong>" & vbcrlf
	sMsg2 = sMsg2 & "<br />Amount: " & FormatCurrency(sAmount,2) & vbcrlf
	If sSVA = "NULL" Then 
		sMsg2 = sMsg2 & "<br />Payment Reference Number: " & sPNREF & vbcrlf
		sMsg2 = sMsg2 & "<br />Authorization Code:" & sAUTHCODE & vbcrlf & vbcrlf
	Else
		'If CitizenPaysFee( iOrgId ) Then
			sMsg2 = sMsg2 & "<br />Processing Fee: " & FormatCurrency(dFeeAmount,2) & vbcrlf
		'Else
	'		dFeeAmount = 0.00
	'	End If 
		sMsg2 = sMsg2 & "<br />Total Charged: " & FormatCurrency((CDbl(dFeeAmount) + CDbl(sAmount)), 2) & vbcrlf
		sMsg2 = sMsg2 & "<br />Order Number: " & sOrderNumber & vbcrlf
		sMsg2 = sMsg2 & "<br />SVA: " & sSVA & vbcrlf
	End If 

	sMsg2 = sMsg2 & "<p><strong>Product Information</strong><br />" & vbcrlf
	sMsg2 = sMsg2 & "Payment: "  & request("paymentname") & "<br />" & vbcrlf
	details = Replace( request("details"), ",", "<br />" & vbcrlf ) 
	'sMsg2 = sMsg2 & "<strong>Details:</strong><br />" & Replace( details, "</br>", "<br />" & vbcrlf ) & "</p>" & vbcrlf 
	sMsg2 = sMsg2 & Replace( details, "</br>", "<br />" & vbcrlf ) & "</p>" & vbcrlf 

	sMsg2 = sMsg2 & "<p><strong>User Information</strong>" & vbcrlf
	sMsg2 = sMsg2 & "<br />Credit Card: XXXXXXXXXXXX" & Right(request("accountnumber"),4)  & vbcrlf
	sMsg2 = sMsg2 & "<br />Name: "  & request("sjname") & vbcrlf
	sMsg2 = sMsg2 & "<br />Address: "  & request("streetaddress") & vbcrlf
	sMsg2 = sMsg2 & "<br />City: "  & request("city") & vbcrlf
	sMsg2 = sMsg2 & "<br />State: "  & request("state") & vbcrlf
	sMsg2 = sMsg2 & "<br />Zip: "  & request("zipcode") & "</p>" & vbcrlf & vbcrlf
	
	sendEmail "", adminEmailAddr, "", sOrgName & " E-GOV PAYMENT SUBMISSION", sMsg2, clearHTMLTags(sMsg2), "N"
	
	If request("email") <> "" Then 
		If isValidEmail( request("email") ) Then 
			' Build the email message to the public
			sMsg3 = "This automated message was sent by the " & sOrgName & " E-Gov web site. Do not reply to this message.  Contact " & adminEmailAddr & " for inquiries regarding this email." & vbcrlf  & vbcrlf
			sMsg3 = sMsg3 & "<br />Thank you for submitting a payment on " & Date() & "." & vbcrlf  & vbcrlf 

			sMsg3 = sMsg3 & "<p><strong>Transaction Details</strong>" & vbcrlf
			sMsg3 = sMsg3 & "<br />Amount: " & FormatCurrency(sAmount,2) & vbcrlf
			If sSVA = "NULL" Then
				sMsg3 = sMsg3 & "<br />Payment Reference Number: " & sPNREF & vbcrlf
				sMsg3 = sMsg3 & "<br />Authorization Code:" & sAUTHCODE & vbcrlf & vbcrlf
			Else
				sMsg3 = sMsg3 & "<br />Processing Fee: " & FormatCurrency(dFeeAmount,2) & vbcrlf
				sMsg3 = sMsg3 & "<br />Total Charged: " & FormatCurrency((CDbl(dFeeAmount) + CDbl(sAmount)), 2) & vbcrlf
				sMsg3 = sMsg3 & "<br />Order Number: " & sOrderNumber & vbcrlf
				sMsg3 = sMsg3 & "<br />SVA: " & sSVA & vbcrlf
			End If 

			sMsg3 = sMsg3 & vbcrlf & "<p><strong>Product Information</strong><br />" & vbcrlf
			sMsg3 = sMsg3 & "Payment: "  & request("paymentname") & "<br />" & vbcrlf
			details = Replace(request("details"), ",", "<br />" & vbcrlf) 
			'sMsg3 = sMsg3 & "<strong>Details:</strong>    " & Replace(details, "</br>", "<br />" & vbcrlf) & "</p>" & vbcrlf & vbcrlf
			sMsg3 = sMsg3 & Replace(details, "</br>", "<br />" & vbcrlf) & "</p>" & vbcrlf & vbcrlf

			sMsg3 = sMsg3 & "<p><strong>User Information</strong>" & vbcrlf
			sMsg3 = sMsg3 & "<br />Credit Card: XXXXXXXXXXXX" & Right(request("accountnumber"),4)  & vbcrlf
			sMsg3 = sMsg3 & "<br />Name: "  & request("sjname") & vbcrlf
			sMsg3 = sMsg3 & "<br />Address: "  & request("streetaddress") & vbcrlf
			sMsg3 = sMsg3 & "<br />City: "  & request("city") & vbcrlf
			sMsg3 = sMsg3 & "<br />State: "  & request("state") & vbcrlf
			sMsg3 = sMsg3 & "<br />Zip: "  & request("zipcode") & "</p>" & vbcrlf & vbcrlf
			
			sendEmail "", request("email"), "", "THANK YOU FOR YOUR " & UCase(sOrgName) & " E-GOV PAYMENT SUBMISSION", sMsg3, clearHTMLTags(sMsg3), "N"

			Set oCdoBuyerMail = Nothing
			Set oCdoBuyerConf = Nothing
		End If 
	End If 

End Sub 


'------------------------------------------------------------------------------
' void ProcessPayPalTransaction( sName, sCardNumber, sExpiration, sAmount )
'------------------------------------------------------------------------------
Sub ProcessPayPalTransaction( ByVal sName, ByVal sCardNumber, ByVal sExpiration, ByVal sAmount )
	Dim parmList, objWinHttp, sParameter, strLength, sResult, sPNREF, sRespMsg, sAuthcode, sDuplicate

	sDuplicate = "Start"

	iPaymentControlNumber = CreatePaymentControlRow( "PAYMENT SCRIPT STARTED." )

	parmList = "cardNum=" & Replace( Trim(sCardNumber), " ", "" )
	parmList = parmList + "&cardExp=" & sExpiration  ' format is MMYY
	parmList = parmList + "&cardType=" & request("cardtype")
	parmList = parmList + "&sjname=" + sName 
	AddToPaymentLog iPaymentControlNumber, "Name: " & sName 
	If OrgHasFeature( iOrgId, "display cvv" ) And request("cvv2") <> "" Then
		parmList = parmList + "&cvv2=" & request("cvv2")
	End If 

	parmList = parmList + "&amount=" & sAmount
	AddToPaymentLog iPaymentControlNumber, "Amount: " & FormatNumber(sAmount,2,,,0)

	parmList = parmList + "&StreetAddress=" & request("StreetAddress")
	parmList = parmList + "&ZipCode=" & request("ZipCode")

	parmList = parmList + "&ordernumber=" + request("ordernumber") 

	parmList = parmList + "&paymentcontrolnumber=" & iPaymentControlNumber
	parmList = parmList + "&orgid=" & iOrgId
	parmList = parmList + "&orgfeature=payments"

	sParameter = request("paymentname")
	strLength = CleanAndCountForPayFlowPro( sParameter )
	parmList = parmList + "&comment1=" & sParameter 
	AddToPaymentLog iPaymentControlNumber, "COMMENT1: " & sParameter

	sParameter = request("comment2")
	strLength = CleanAndCountForPayFlowPro( sParameter )
	parmList = parmList + "&comment2=" & sParameter
	AddToPaymentLog iPaymentControlNumber, "COMMENT2: " & sParameter

	' Here we look to see if the org has an alternate payment account for "Payments"
	' This function is in common.asp
	If OrgFeatureHasAlternateAccount( iOrgId, "payments" ) Then
		parmList = parmList & "&feature=payments"
		AddToPaymentLog iPaymentControlNumber, "feature: payments"
	Else
		parmList = parmList & "&feature=default"
		AddToPaymentLog iPaymentControlNumber, "feature: default"
	End If 

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

		AddToPaymentLog iPaymentControlNumber, "transResponse: " & left(transResponse,1000)
		AddToPaymentLog iPaymentControlNumber, "DUPLICATE: " & GetResponseValue(transResponse, "DUPLICATE")
		sDuplicate = GetResponseValue(transResponse, "DUPLICATE")
		' A duplicate will pull results from the initial request with that requestid.
		' All of these will be unique purchases so duplicates are not allowed.
		' So try again, hoping for a new requestid
	Loop 

	' Continue processing the results from the PayPal call
	sResult = GetResponseValue(transResponse, "RESULT")
	on error resume next
	sResult = clng(sResult)
	if err.number <> 0 then sResult = -1
	on error goto 0
	sPNREF = GetResponseValue(transResponse, "PNREF")
	sRespMsg = GetResponseValue(transResponse, "RESPMSG")

	If sResult = clng(0) Then 
		' Successful Transaction 
		AddToPaymentLog iPaymentControlNumber, "TRANSACTION SUCCEEDED"
		approved = True
		sAuthcode = GetResponseValue( transResponse, "AUTHCODE" )
		AddToPaymentLog iPaymentControlNumber, "AUTHCODE: " & sAuthcode
		AddToPaymentLog iPaymentControlNumber, "PNREF: " & sPNREF
		AddToPaymentLog iPaymentControlNumber, "RESPMSG: " & sRespMsg
		ProcessSuccessfulTransaction sOrderID, sAuthcode, sPNREF, sRespMsg, Replace(sAmount, ",", "" ), "NULL", "NULL", "NULL"

	ElseIf sResult < clng(0) Then 
		' Communication Error
		AddToPaymentLog iPaymentControlNumber, "Communication Error"
		AddToPaymentLog iPaymentControlNumber, "Result: " & sResult
		AddToPaymentLog iPaymentControlNumber, "PNREF: " & sPNREF
		AddToPaymentLog iPaymentControlNumber, "RESPMSG: " & sRespMsg
		approved = False 
		response.write "<div class=""payflowmsgfail"">Your credit card purchase was unable to processed because of a network communication error. Please try your transaction again later.<blockquote><font color=#000000>Payment Reference Number:</font> " & sPNREF & " <br><font color=#000000>Description:</font> (" & sResult & ") - " & sRespMsg & " </blockquote></div>"
		response.write "<div>" & transResponse & "</div>"

	ElseIf sResult > clng(0) Then 
		' Transaction Declined
		AddToPaymentLog iPaymentControlNumber, "Transaction Declined"
		AddToPaymentLog iPaymentControlNumber, "Result: " & sResult
		AddToPaymentLog iPaymentControlNumber, "PNREF: " & sPNREF
		AddToPaymentLog iPaymentControlNumber, "RESPMSG: " & sRespMsg
		approved = False
		ProcessDeclinedTransaction sResult, sPNREF, sRespMsg, sAmount, "", ""

	End If  
			
	AddToPaymentLog iPaymentControlNumber, "PAYMENT PROCESSING FINISHED."

End Sub 


'------------------------------------------------------------------------------
' void ProcessPointAndPayTransaction( sFirstName, sLastName, sCardNumber, sExpiration, sAmount )
'------------------------------------------------------------------------------
Sub ProcessPointAndPayTransaction( ByVal sFirstName, ByVal sLastName, ByVal sCardNumber, ByVal sExpiration, ByVal sAmount )
	Dim parmList, objWinHttp, sParameter, strLength, sResult, sPNREF, sRespMsg, sAuthcode
	Dim sStatus, sErrorMsg, dFeeAmount, sOrderNumber, sSVA, sTotalCharges, sNotes

	iPaymentControlNumber = CreatePaymentControlRow( "PAYMENT SCRIPT STARTED - Point And Pay." )

	parmList = "paymentcontrolnumber=" & iPaymentControlNumber
	parmList = parmList + "&chargeaccountnumber=" & sCardNumber
	parmList = parmList + "&cardtype=" & request("cardtype")
	parmList = parmList + "&chargeexpirationmmyy=" & sExpiration  ' format is MMYY
	parmList = parmList + "&signerfirstname=" + sFirstName 
	parmList = parmList + "&signerlastname=" + sLastName
	AddToPaymentLog iPaymentControlNumber, "Name: " & sFirstName & " " & sLastName 

	If bOrgHasCVV And request("cvv2") <> "" Then
		parmList = parmList + "&chargecvn=" & request("cvv2")
		AddToPaymentLog iPaymentControlNumber, "ChargeCVN: present but not stored"
	End If 

	parmList = parmList + "&chargeamount=" & FormatNumber(sAmount,2,,,0)
	AddToPaymentLog iPaymentControlNumber, "Amount: " & FormatNumber(sAmount,2,,,0)

	parmList = parmList + "&signeraddressline1=" & request("StreetAddress")
	parmList = parmList + "&signeraddresscity=" + request("City")
	parmList = parmList + "&signeraddressregioncode=" + request("State")
	parmList = parmList + "&signeraddresspostalcode=" & request("ZipCode")

	sNotes = request("paymentname") & " - " & request("comment2")
	sReservationDetails = sNotes
	AddToPaymentLog iPaymentControlNumber, "Notes: " & CleanAndCutForPNPNotes( sNotes )
	parmList = parmList + "&notes=" & CleanAndCutForPNPNotes( sNotes )

'	response.write "parmList: " & parmList & "<br /><br />"
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

	' Trash our object now that we are finished with it.
	Set objWinHttp = Nothing

'	response.write "transResponse: " & transResponse & "<br /><br />"
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
		ProcessDeclinedTransaction "declined", "", sErrorMsg, sAmount, sOrderNumber, sSVA
		'ProcessDeclinedTransaction sResult, sPNREF, sRespMsg, sAmount, "", ""
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

		ProcessSuccessfulTransaction sOrderID, "NULL", "NULL", "approved", FormatNumber(sAmount,2,,,0), sOrderNumber, sSVA, dFeeAmount
		'ProcessSuccessfulTransaction sOrderID, sAuthcode, sPNREF, sRespMsg, Replace(sAmount, ",", "" ), "NULL", "NULL", "NULL"
	End If 

	AddToPaymentLog iPaymentControlNumber, "PAYMENT PROCESSING FINISHED."

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
' void ProcessSuccessfulTransaction( sOrderID, sAUTHCODE, sPNREF, sRESPMSG, sAmount, sOrderNumber, sSVA, dFeeAmount )
'------------------------------------------------------------------------------
Sub ProcessSuccessfulTransaction( ByVal sOrderID, ByVal sAUTHCODE, ByVal sPNREF, ByVal sRESPMSG, ByVal sAmount, ByVal sOrderNumber, ByVal sSVA, ByVal dFeeAmount )
	Dim iPermitNumber
	
	iPermitNumber = "NULL"

	' remove this session variable for the Rye Qty permit offerings
	session.contents.remove("myCount")

	' ADD INFORMATION TO ADMINSTRATION DATABASE
	iPaymentID = Right(request("ordernumber"),Len(request("ordernumber"))-InStr(request("ordernumber"),"O"))
	AddToPaymentLog iPaymentControlNumber, "PaymentId: " & iPaymentID

	AddPaymentInformation iPaymentID, sPNREF, FormatNumber(sAmount,2,,,0), sAuthcode, sOrderNumber, sSVA, dFeeAmount

	If request("paymentservice") <> "" Then
		AddToPaymentLog iPaymentControlNumber, "Payment Service: " & request("paymentservice")
		' These are special forms that need extra things to happen
		If request("paymentservice") = "rye commuter permit renewal" Then 
			AddToPaymentLog iPaymentControlNumber, "Permit Holder Type: " & request("permitholdertype")
			AddToPaymentLog iPaymentControlNumber, "Renewal Id: " & request("renewalid")
			If LCase(request("permitholdertype")) = "current resident railroad permit holder" Then
				' Get the next number in the sequence
				iPermitNumber = GetNewPermitNumber( iPaymentID, request("renewalid"), "egov_ryepermitnumbers" )
				AddToPaymentLog iPaymentControlNumber, "Assigned Permit No: " & iPermitNumber
			End If 
			If LCase(request("permitholdertype")) = "current non-resident railroad permit holder" Then
				' Get the next number in the sequence
				iPermitNumber = GetNewPermitNumber( iPaymentID, request("renewalid"), "egov_ryepermitnumbers_nonres" )
				AddToPaymentLog iPaymentControlNumber, "Assigned Permit No: " & iPermitNumber
			End If 
			' Update the master list
			UpdatePermitMasterList request("renewalid"), iPermitNumber, iPaymentID 
		End If 
		If request("paymentservice") = "rye railroad new waitlist" Then 
			iPermitNumber = GetNewPermitNumber( iPaymentID, 0, "egov_ryepermitnumbers_mnrrnewwaitlist" )
			AddToPaymentLog iPaymentControlNumber, "Assigned Permit No: " & iPermitNumber

			sSqlpn = "UPDATE egov_payments SET ordernumber = " & iPermitNumber & " WHERE paymentid = " & iPaymentID
			RunSQLStatement sSqlpn

		End If
		If request("paymentservice") = "rye commuter waitlist renewal" Or request("paymentservice") = "rye commuter permits by name" Then 
			AddToPaymentLog iPaymentControlNumber, "Permit Holder Type: " & request("permitholdertype")
			AddToPaymentLog iPaymentControlNumber, "Renewal Id: " & request("renewalid")
			' Update the master list
			UpdatePermitMasterList request("renewalid"), iPermitNumber, iPaymentID 
		End If 
	End If 

'	ShowReceipt iPaymentID

	' DISPLAY SUCCESS MESSAGE TO CUSTOMER
    ' VERISIGN BRANDING
	'response.write "<center><p><a href=""http://seal.verisign.com/payment"" TARGET=""_VERISIGN"" ><img vspace=10 border=0 hspace=20 src=""images/verisign.gif""></a><br>VeriSign has routed, processed, and secured your payment information. <A HREF="" http://www.verisign.com/products-services/payment-processing/index.html"">More information about VeriSign</a></p>"
	response.write "<center><p>"
	sPaymentImg = GetPaymentImage( "../../" )
	If sPaymentImg <> "" Then 
		response.write "<img src=""" & sPaymentImg & """ border=""0"" /><br /><br />"
	End If 
	response.write "<strong>" & GetPaymentGatewayName( ) & "</strong> has routed, processed, and secured your payment information."
	response.write "</p>"

	' TRANSACTION RESULT DETAILS
	response.write "<div class=""group""><p>Your credit card payment was <b>approved</b>.<br> Keep the following information for your records:</p>"
	'response.write "<p><blockquote>"
	response.write "<table cellpadding=""2"" cellspacing=""0"" border=""0"">"
	response.write "<tr><td colspan=""2""><b>Transaction Details</b></td></tr>"
	
	If sSVA = "NULL" Then 
		response.write "<tr><td align=""right"">Amount Charged:</td><td align=""left""> " & FormatCurrency(sAmount,2) & "</td></tr>"
		response.write "<tr><td align=""right"">Payment Reference Number:</td><td align=""left""> " & sPNREF & " </td></tr>"
		response.write "<tr><td align=""right"">Authorization Code:</td><td align=""left""> " & sAUTHCODE & "</td></tr>"
	Else
		response.write "<tr><td align=""right"">Amount:</td><td align=""left""> " & FormatCurrency(sAmount,2) & "</td></tr>"
		'If CitizenPaysFee( iOrgId ) Then 
			response.write "<tr>"
			response.write "<td align=""right"">Processing Fee:</td>"
			response.write "<td align=""left""> " & FormatCurrency(CDbl(dFeeAmount),2) & "</td>"
			response.write "</tr>"
		'Else
		'	dFeeAmount = 0.00
		'End If
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
	response.write "<br />"
	'response.write "</p>"
	
	' PRODUCT INFORMATION
	response.write "<table cellpadding=""2"" cellspacing=""0"" border=""0"">"
	response.write "<tr><td colspan=""2""><strong>Product Information</strong></td></tr>"
	response.write "<tr><td align=""right"">Payment: </td><td align=""left"">" & request("paymentname") & "</td></tr>"
	response.write "<tr><td valign=""top"" align=""right"">Details: </td><td  valign=""top"" align=""left"">" & Replace(request("details"),"</br>,","") & "</td></tr>"
	response.write "</table>"
	response.write "<br />"

	' CREDIT CARD INFORMATION	
	response.write "<table cellpadding=""2"" cellspacing=""0"" border=""0"">"
	response.write "<tr><td colspan=""2""><strong>User Information</strong></td></tr>"
	response.write "<tr><td align=""right"">Credit Card: </td><td align=""left"">XXXXXXXXXXXX" & Right(request("accountnumber"),4)  & "</td></tr>"
	response.write "<tr><td align=""right"">Name: </td><td align=""left"">" & request("sjname") & "</td></tr>"
	response.write "<tr><td align=""right"">Address: </td><td align=""left"">" & request("streetaddress") & "</td></tr>"
	response.write "<tr><td align=""right"">City: </td><td align=""left"">" & request("city") & "</td></tr>"
	response.write "<tr><td align=""right"">State: </td><td align=""left"">" & request("state") & "</td></tr>"
	response.write "<tr><td align=""right"">Zip: </td><td align=""left"">" & request("zipcode") & "</td></tr>"
	response.write "</table>"

	'response.write "</p></blockquote>"
	response.write "</div></center><br /><br />"

End Sub 


'------------------------------------------------------------------------------
' void  ShowReceipt
'------------------------------------------------------------------------------
Sub ShowReceipt(  ByVal iPaymentID )
	Dim sPaymentImg, sSql, oRs

	sSql = "SELECT paymentservicename, paymentamount, paymentsummary, payment_information, ISNULL(userfname,'') AS userfname, "
	sSql = sSql & "ISNULL(userlname,'') AS userlname, ISNULL(useraddress,'') AS useraddress, ISNULL(usercity,'') AS usercity, "
	sSql = sSql & "ISNULL(userstate,'') AS userstate, ISNULL(userzip,'') AS userzip, paymentrefid "
	sSql = sSql & "FROM egov_payment_list WHERE paymentid = " & iPaymentID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then

		response.write "<center><p>"

		sPaymentImg = GetPaymentImage( "../../" )
		If sPaymentImg <> "" Then 
			response.write "<img src=""" & sPaymentImg & """ border=""0"" /><br /><br />"
		End If 

		response.write "<strong>" & GetPaymentGatewayName( ) & "</strong> has routed, processed, and secured your payment information."
		response.write "</p>"

		' TRANSACTION RESULT DETAILS
		response.write "<div class=""group""><p>Your credit card payment was <b>approved</b>.<br> Keep the following information for your records:</p>"
		'response.write "<p>"
		response.write "<table cellpadding=""2"" cellspacing=""0"" border=""0"">"
		response.write "<tr><td colspan=""2""><b>Transaction Details</b></td></tr>"
		
		If GetPaymentSummaryValue( "Authorization Code", oRs("paymentsummary") ) <> "" Then  
			' This is for PayPal
			response.write "<tr><td align=""right"">Amount Charged:</td><td align=""left""> " & FormatCurrency(oRs("paymentamount"),2) & "</td></tr>"
			response.write "<tr><td align=""right"">Payment Reference Number:</td><td align=""left""> " & GetPaymentSummaryValue( "Payment Reference Number", oRs("paymentsummary") ) & " </td></tr>"
			response.write "<tr><td align=""right"">Authorization Code:</td><td align=""left""> " & GetPaymentSummaryValue( "Authorization Code", oRs("paymentsummary") ) & "</td></tr>"
		Else
			' This is for Point and Pay
			dProcessingFee = GetProcessingFee( iPaymentID ) ' return back a double
			response.write "<tr><td align=""right"">Amount:</td><td align=""left""> " & FormatCurrency((CDbl(oRs("paymentamount")) - dProcessingFee),2) & "</td></tr>"
			response.write "<tr>"
			response.write "<td align=""right"">Processing Fee:</td>"
			response.write "<td align=""left""> " & FormatCurrency(dProcessingFee,2) & "</td>"
			response.write "</tr>"
			response.write "<tr>"
			response.write "<td align=""right"">Total Charged:</td>"
			response.write "<td align=""left""> " & FormatCurrency(CDbl(oRs("paymentamount")),2) & "</td>"
			response.write "</tr>"
			' Order number and SVA are PNP values returned to us. NO PNP payments as of 9/9/2011
			response.write "<tr>"
			response.write "<td align=""right"">Order Number:</td>"
			response.write "<td align=""left""> " & GetPaymentSummaryValue( "Order Number", oRs("paymentsummary") ) & "</td>"
			response.write "</tr>"
			response.write "<tr>"
			response.write "<td align=""right"">SVA:</td>"
			response.write "<td align=""left""> " & GetPaymentSummaryValue( "SVA", oRs("paymentsummary") ) & "</td>"
			response.write "</tr>" 
		End If 
		response.write "</table >"
		response.write "<br />"
		
		' PRODUCT INFORMATION
		response.write "<table cellpadding=""2"" cellspacing=""0"" border=""0"">"
		response.write "<tr><td colspan=""2""><strong>Product Information</strong></td></tr>"
		response.write "<tr><td align=""right"">Payment: </td><td align=""left"">" &oRs("paymentservicename") & "</td></tr>"
		response.write "<tr><td valign=""top"" align=""right"">Details: </td><td  valign=""top"" align=""left"">" & Replace(oRs("payment_information"),"</br>,","<br />") & "</td></tr>"
		response.write "</table><br />"

		' CREDIT CARD INFORMATION	
		response.write "<table cellpadding=""2"" cellspacing=""0"" border=""0"">"
		response.write "<tr><td colspan=""2""><strong>User Information</strong></td></tr>"
		response.write "<tr><td align=""right"">Credit Card: </td><td align=""left"">XXXXXXXXXXXX" & Right(request("accountnumber"),4)  & "</td></tr>"
		response.write "<tr><td align=""right"">Name: </td><td align=""left"">" & Trim(oRs("userfname") & " " & oRs("userlname")) & "</td></tr>"
		response.write "<tr><td align=""right"">Address: </td><td align=""left"">" & oRs("useraddress") & "</td></tr>"
		response.write "<tr><td align=""right"">City: </td><td align=""left"">" & oRs("usercity") & "</td></tr>"
		response.write "<tr><td align=""right"">State: </td><td align=""left"">" & oRs("userstate") & "</td></tr>"
		response.write "<tr><td align=""right"">Zip: </td><td align=""left"">" & oRs("userzip") & "</td></tr>"
		response.write "</table>"
		response.write "</div></center><br /><br />"
	Else
		response.write "<p><strong>No information could be found for this payment.</strong></p>"
	End If 

End Sub 


'-------------------------------------------------------------------------------------------------
' double GetProcessingFee( iPaymentID )
'-------------------------------------------------------------------------------------------------
Function GetProcessingFee( ByVal iPaymentID )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(processingfee,0.00) AS processingfee FROM egov_payments WHERE paymentid = " & iPaymentID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetProcessingFee = CDbl(oRs("processingfee"))
	Else
		GetProcessingFee = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPaymentDetailValue( sLabel, sPaymentInfo )
'-------------------------------------------------------------------------------------------------
Function GetPaymentDetailValue( ByVal sLabel, ByVal sPaymentInfo )
	Dim aDetailinfo, sRow, aCells, sValue

	sValue = ""

	aDetailinfo = Split( sPaymentInfo, "</br>")
	For Each sRow In aDetailinfo
		If InStr( sRow, sLabel ) > 0  Then 
			'sValue = sValue & "[sLabel=" & sLabel & "] "
			'sValue = sValue & "[sRow=" & sRow & "] "
			aCells = Split( sRow, ":" )
			sValue = sValue & Trim(aCells(1))
		End If 
	Next 

	GetPaymentDetailValue = sValue

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPaymentSummaryValue( sLabel, sPaymentInfo )
'-------------------------------------------------------------------------------------------------
Function GetPaymentSummaryValue( ByVal sLabel, ByVal sPaymentInfo )
	Dim aDetailinfo, sRow, aCells, sValue

	sValue = ""

	aDetailinfo = Split( sPaymentInfo, "<br>")
	For Each sRow In aDetailinfo
		If InStr( sRow, sLabel ) > 0  Then 
			'sValue = sValue & "[sLabel=" & sLabel & "] "
			'sValue = sValue & "[sRow=" & sRow & "] "
			aCells = Split( sRow, ":" )
			sValue = sValue & Trim(aCells(1))
		End If 
	Next 

	GetPaymentSummaryValue = sValue

End Function 



'------------------------------------------------------------------------------
' integer GetNewPermitNumber( iPaymentID, iRenewalId, sTableName )
'------------------------------------------------------------------------------
Function GetNewPermitNumber( ByVal iPaymentID, ByVal iRenewalId, ByVal sTableName )
	Dim oRs, sSql

	sSql = "INSERT INTO " & sTableName & " ( paymentid, renewalid ) VALUES ( "
	sSql = sSql & iPaymentID & ", " & iRenewalId & " )"

	GetNewPermitNumber = RunIdentityInsertStatement( sSql )

End Function 


'------------------------------------------------------------------------------
' void UpdatePermitMasterList iRenewalId, iPermitNumber, iPaymentID 
'------------------------------------------------------------------------------
Sub UpdatePermitMasterList( ByVal iRenewalId, ByVal iPermitNumber, ByVal iPaymentID )
	Dim sSql, oRs

	sSql = "UPDATE egov_ryepermitrenewals SET "
	sSql = sSql & "hasrenewed = 1, "
	sSql = sSql & "assignedpermitnumber = " & iPermitNumber & ", "
	sSql = sSql & "paymentid = " & iPaymentID & " "
	sSql = sSql & "WHERE renewalid = " & iRenewalId

	RunSQLStatement sSql

End Sub 


'------------------------------------------------------------------------------
' Function ProcessDeclinedTransaction( sRESULT, sPNREF, sRESPMSG, sAmount, sOrderNumber, sSVA )
'------------------------------------------------------------------------------
Function ProcessDeclinedTransaction( ByVal sRESULT, ByVal sPNREF, ByVal sRESPMSG, ByVal sAmount, ByVal sOrderNumber, ByVal sSVA )

	' DISPLAY DECLINED MESSAGE TO CUSTOMER

	response.write "<center><p>"
	sPaymentImg = GetPaymentImage( "../../" )
	If sPaymentImg <> "" Then 
		response.write "<img src=""" & sPaymentImg & """ border=""0"" /><br /><br />"
	End If 
	
	response.write "<strong>" & GetPaymentGatewayName( ) & "</strong> has routed, processed, and secured your payment information."
	response.write "</p>"

	' TRANSACTION RESULT DETAILS
	response.write "<p><div class=""group""><p>Your credit card purchase was <strong>declined</strong> for the following reason:"
	response.write "<blockquote>"
	If sSVA = "" Then 
		response.write "Payment Reference Number: " & sPNREF & " <br />"
	End If 
	response.write "<strong>Description: ( " & UCase(sRESULT) & " ) - " & sRESPMSG & "</strong></blockquote></p>"
	response.write "<p><blockquote>"
	response.write "<table>"
	response.write "<tr><td colspan=""2""><strong>Transaction Details</strong></td></tr>"
	response.write "<tr><td align=""right"">Amount:</td><td align=""left""> " & FormatCurrency(sAmount,2) & "</td></tr>"
	If sSVA = "" Then 
		response.write "<tr><td align=""right"">Reference Number:</td><td align=""left""> " & sPNREF & " </td></tr>"
		response.write "<tr><td align=""right"">Authorization Code:</td><td align=""left""> " & sAUTHCODE & "</td></tr></table>"
	Else 
		response.write "<tr><td align=""right"">Order Number:</td><td align=""left""> " & sOrderNumber & " </td></tr>"
		response.write "<tr><td align=""right"">SVA:</td><td align=""left""> " & sSVA & "</td></tr></table>"
	End If 
	response.write "</blockquote></p></div></p></center>"

End Function


'------------------------------------------------------------------------------
' Function ProcessCommunicationError()
'------------------------------------------------------------------------------
Function ProcessCommunicationError()

	' DISPLAY COMMUNICATION MESSAGE TO CUSTOMER
    response.write "<div class=""payflowmsgfail"">Your credit card purchase was unable to processed because of a network communication error:<blockquote><font color=#000000>Payment Reference Number:</font> " & sPNREF & " <br><font color=#000000>Description:</font> (" & sRESULT & ") - " & sRESPMSG & " </blockquote></div>"
	
End Function


'----------------------------------------------------------------------------------------
' void AddtoLog sText
'----------------------------------------------------------------------------------------
Sub AddtoLog( ByVal sText )

    ' WRITES SUPPLIED TEXT TO FILE WITH DATETIME
	Set oFSO = Server.Createobject("Scripting.FileSystemObject")
	Set oFile = oFSO.GetFile(Application("VerisignPaymentLog"))

	'response.write Application("VerisignPaymentLog") & "<br />"
    Set oText = oFile.OpenAsTextStream(8)
    oText.WriteLine (Now() & Chr(9) & sText)
    oText.Close
    
    Set oText = Nothing
    Set oFile = Nothing
    Set oFSO = Nothing

End Sub 


'------------------------------------------------------------------------------
' integer CreatePaymentControlRow( sLogEntry )
'------------------------------------------------------------------------------
Function CreatePaymentControlRow( ByVal sLogEntry )
	Dim sSql, iPaymentControlNumber

	sSql = "INSERT INTO paymentlog ( orgid, applicationside, feature, logentry ) VALUES ( "
	sSql = sSql & iOrgID & ", 'Public', 'Payments', '" & sLogEntry & "' )"
	'response.write sSql & "<br /><br />"

	iPaymentControlNumber = RunIdentityInsertStatement( sSql )

	sSql = "UPDATE paymentlog SET paymentcontrolnumber = " & iPaymentControlNumber
	sSql = sSql & " WHERE paymentlogid = " & iPaymentControlNumber
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql 

	CreatePaymentControlRow = iPaymentControlNumber

End Function 


'------------------------------------------------------------------------------
' void AddToPaymentLog( iPaymentControlNumber, sLogEntry )
'------------------------------------------------------------------------------
Sub AddToPaymentLog( ByVal iPaymentControlNumber, ByVal sLogEntry  )
	Dim sSql

	sSql = "INSERT INTO paymentlog ( paymentcontrolnumber, orgid, applicationside, feature, logentry ) VALUES ( "
	sSql = sSql & iPaymentControlNumber & ", " & iOrgID & ", 'public', 'payments', '" & dbready_string(sLogEntry, 500) & "' )"
	'response.write sSql & "<br /><br />"

	RunSQLStatement sSql 

End Sub 


'------------------------------------------------------------------------------
' void SendLoginFailedEmail
'------------------------------------------------------------------------------
Sub SendLoginFailedEmail( )
	Dim sHTMLBody, sTextTBody

	sTextTBody = "A User Authentication error has been received for " & sOrgName & " ( " & iOrgId & " ). "
	sTextTBody = sTextTBody & vbcrlf & "The client has changed their password and needs to update our database."

	sHTMLBody = "<p>A User Authentication error has been received for " & sOrgName & " ( " & iOrgId & " ). "
	sHTMLBody = sHTMLBody & "<br />The client has changed their password and needs to update our database.</p>"

	sendEmail "noreply@eclink.com", "noreply@eclink.com", "", "PayPal User Authentication Error Received", sHTMLBody, sTextTBody, "Y"

End Sub 


'----------------------------------------------------------------------------------------
' string CleanAndCutForPNPNotes( sParameter )
'----------------------------------------------------------------------------------------
Function CleanAndCutForPNPNotes( ByVal sParameter )
	' cleans forbidden characters and returns the string

	sParameter = Replace(sParameter, Chr(34), "")
	sParameter = Replace(sParameter, "'", "")
	sParameter = Replace(sParameter, ",", "")
	sParameter = Replace(sParameter, "&", "and")
	sParameter = Replace(sParameter, "=", "is")
	sParameter = Replace(sParameter, "</br>", ", ")
	sParameter = Replace(sParameter, "<br />", ", ")
	sParameter = Replace(sParameter, "<br>", ", ")
	sParameter = Replace(sParameter, ", ,", "")
	sParameter = Trim(sParameter)
	sParameter = Left(sParameter, 255)	' 255 characters is the PNP limit for notes

	CleanAndCutForPNPNotes = sParameter

End Function 


'----------------------------------------------------------------------------------------
' string GetPaymentStatus( sPaymentId )
'----------------------------------------------------------------------------------------
Function GetPaymentStatus( ByVal sPaymentId )
	Dim sSql, oRs

	sSql = "SELECT UPPER(ISNULL(paymentstatus,'ERROR')) AS paymentstatus FROM egov_payments WHERE paymentid = " & CLng(sPaymentId)

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetPaymentStatus = oRs("paymentstatus")
	Else
		GetPaymentStatus = "ERROR"
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


%>

