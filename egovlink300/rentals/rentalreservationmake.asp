<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="../include_top_functions.asp" //-->
<!-- #include file="rentalcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalreservationmake.asp
' AUTHOR: Steve Loar
' CREATED: 02/05/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Handles making the reservation and payment.
'
' MODIFICATION HISTORY
' 1.0   02/05/2010	Steve Loar - INITIAL VERSION
' 1.1	03/04/2010	Steve Loar - Modified to handle paid reservations with the call to PayPal
' 2.0	06/23/2010	Steve Loar - Split name field into first and last 
' 2.`	07/26/2010	Steve Loar - Modified to handle Point and Pay transactions
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iReservationTempId, sSource, dTotalAmount, bOrgHasCVV, iPaymentControlNumber, bOk
Dim sReservationTypeId, iRentalUserid, sStartDateTime, sEndDateTime, sMessage, iRentalid
Dim iReservationId, sBillingEndDateTime, sArrivalDateTime, sDepartureDateTime, sOffSeasonFlag
Dim approved, sAuthcode, sPNREF, sRespMsg, sProcessingRoute, sRentalName, sSelectedDate, sStartTime, sEndTime


iReservationTempId = CLng(request("rti"))

If request("src") = "" Then
	' if no src then take them to the non-secure category page
	LogThePage()
	response.redirect sEgovWebsiteURL & "/rentals/rentalcategories.asp"
Else
	sSource = request("src")
End If

'if iReservationTempId = 30577 then
	'sSource = "rc"
'end if

If sSource <> "rc" Then
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
			LogThePage()
			response.redirect sEgovWebsiteURL & "/rentals/rentalcategories.asp"
		end if
end if

dTotalAmount = CDbl(0.00)
sAuthcode = ""
sPNREF = ""
sRespMsg = ""


%>
<html>
<head>
	<title>E-Gov Services <%=sOrgName%> - Payment Form</title>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalstyles.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script language="Javascript" src="../scripts/modules.js"></script>

	<script language="javascript">
		function openWin2(url, name) 
		{
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
		}
	</script>

<%
	If request.servervariables("HTTPS") = "on" Then 
		response.write "<style>" & vbcrlf
		response.write "	body {behavior: url('https://secure.egovlink.com/" & sorgVirtualSiteName & "/csshover.htc');}"
		response.write "</style>" & vbcrlf
	End If 
%>

</head>

<!--#Include file="../include_top.asp"-->

<%


If sSource = "rc" Then
	' this is the no cost rentals from rental control that skip the payment form
	iPaymentControlNumber = CreatePaymentControlRow( "PAYMENT SCRIPT STARTED." )
	AddToPaymentLog iPaymentControlNumber, "TRANSACTION SUCCEEDED - No Cost Reservation"
	ProcessSuccessfulTransaction iReservationTempId, dTotalAmount, sAuthcode, sPNREF, sRespMsg, "NULL", "NULL", "0.00"

ElseIf sSource = "pf" then
	' This is from the payment form so process to PayPal

	' GET THE RESERVATION DETAILS HERE
	GetSomeReservationDetails iReservationTempId, iRentalid, sRentalName, sSelectedDate, sStartTime, sEndTime, sStartDateTime, sEndDateTime
	session("rentalid") = iRentalid
	session("SelectedDate") = sSelectedDate
	session("StartTime") = sStartTime
	session("EndTime") = sEndTime 
	session("StartDateTime") = sStartDateTime
	session("EndDateTime") = sEndDateTime

	sOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(CDate(sStartDateTime)) )
	session("OffSeasonFlag") = sOffSeasonFlag

	' Check that the time is still available
	'bOk = CheckRentalAvailability( iRentalid, sStartDateTime, sEndDateTime, sMessage )	' In rentalcommonfunctions.asp
	If CheckForExistingReservations( iRentalid, sStartDateTime, sEndDateTime, "selectedperiod", sOffSeasonFlag ) = "No" Then
		' go to time unavailable page.
		'response.write "Conflict found. startdatetime: " & sStartDateTime & " enddatetime: " & sEndDateTime & "<br /><br />"
		'response.redirect sEgovWebsiteURL & "/rentals/rentalunavailable.asp?rti=" & iReservationTempId
		LogThePage()
		response.redirect "rentalunavailable.asp?rti=" & iReservationTempId
	End If 

	' Get the total charge amount
	dTotalAmount = GetTotalCharges( iReservationTempId )
	
	'Check for org features
	bOrgHasCVV = orghasfeature(iOrgID,"display cvv")

	If OrgHasFeature( iOrgId, "skippayment" ) Then
		iPaymentControlNumber = CreatePaymentControlRow( "PAYMENT SCRIPT STARTED." )
		AddToPaymentLog iPaymentControlNumber, "TRANSACTION SUCCEEDED - Bypassed Authorization"
		approved = True
		sAuthcode = "010101"
		sPNREF = "V19F1D5C82TEST"
		sRespMsg = "Approved"
		AddToPaymentLog iPaymentControlNumber, "AUTHCODE: " & sAuthcode
		AddToPaymentLog iPaymentControlNumber, "PNREF: " & sPNREF
		AddToPaymentLog iPaymentControlNumber, "RESPMSG: " & sRespMsg
		ProcessSuccessfulTransaction iReservationTempId, dTotalAmount, sAuthcode, sPNREF, sRespMsg, "NULL", "NULL", "0.00"
	Else
		sProcessingRoute = GetProcessingRoute()		' In ../include_top_functions.asp

		' the case statement is because there used to be several, and there could be more in the future
		Select Case sProcessingRoute
			Case "PayPalPayFlowPro"
				' Newer way to handle PayFlow Pro payments - Everyone should have this right now
				ProcessPayPalTransaction iReservationTempId, Trim(request("firstname") & " " & request("lastname")), request("accountnumber"), request("month")&request("year"), dTotalAmount, sRentalName, sSelectedDate, sStartTime, sEndTime, bOrgHasCVV 
			Case "PointAndPay"
				' Point and Pay Transaction Processing
				ProcessPointAndPayTransaction iReservationTempId, request("firstname"), request("lastname"), request("accountnumber"), request("month") & request("year"), dTotalAmount, sRentalName, sSelectedDate, sStartTime, sEndTime, bOrgHasCVV 
		End Select 

	End If 
Else
	' This is something not expected, so handle gracefully
	LogThePage()
	response.redirect sEgovWebsiteURL & "/rentals/rentalcategories.asp"
End If 


'BEGIN: Payment Footer
response.write vbcrlf & "<center>"
response.write vbcrlf & "<input type=""button"" class=""button"" onClick=""location.href='" & sEgovWebsiteURL & "/';"" value=""Click here to return to the E-Government Website"" /><br />"
response.write vbcrlf & "</center>" 

response.write vbcrlf & "<center>"
response.write vbcrlf & "<p class=""smallnote"">NOTE: Your IP address [" & request.servervariables("REMOTE_ADDR") & "] has been logged with this transaction.</p>"
response.write vbcrlf & "</center>"
'END: Payment Footer

response.write vbcrlf & "</div>"
'END: Display Response

'BEGIN: Spacing Code
response.write "<p>&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</p>"
'END: Spacing Code
%>

<!--#Include file="../include_bottom.asp"--> 


<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' void ProcessPayPalTransaction iReservationTempId, sName, sCardNumber, sExpiration, dTotalAmount, sRentalName, sSelectedDate, sStartTime, sEndTime, bOrgHasCVV 
'------------------------------------------------------------------------------
Sub ProcessPayPalTransaction( ByVal iReservationTempId, ByVal sName, ByVal sCardNumber, ByVal sExpiration, ByVal dTotalAmount, ByVal sRentalName, ByVal sSelectedDate, ByVal sStartTime, ByVal sEndTime, ByVal bOrgHasCVV )
	Dim parmList, objWinHttp, sParameter, strLength, sResult, sPNREF, sRespMsg, sAuthcode, sDuplicate, sReservationDetails

	sDuplicate = "Start"

	iPaymentControlNumber = CreatePaymentControlRow( "PAYMENT SCRIPT STARTED - PayFlow Pro." )

	parmList = "cardNum=" & sCardNumber
	parmList = parmList + "&cardExp=" & sExpiration  ' format is MMYY
	parmList = parmList + "&sjname=" + sName 
	AddToPaymentLog iPaymentControlNumber, "Name: " & sName 
	If bOrgHasCVV And request("cvv2") <> "" Then
		parmList = parmList + "&cvv2=" & request("cvv2")
		AddToPaymentLog iPaymentControlNumber, "CVV2: Present but not shown" '& request("cvv2")
	End If 

	parmList = parmList + "&amount=" & dTotalAmount
	AddToPaymentLog iPaymentControlNumber, "Amount: " & FormatNumber(dTotalAmount,2,,,0)

	parmList = parmList + "&StreetAddress=" & request("StreetAddress")
	if iOrgId = "228" then
		parmList = parmList + "&City=" + request("City")
		parmList = parmList + "&State=" + request("State")
	end if
	parmList = parmList + "&ZipCode=" & request("ZipCode")

	' ordernumber is not a parameter that PayPal takes
	'parmList = parmList + "&ordernumber=" + request("ordernumber") 

	sParameter = "Reservation"
	strLength = CleanAndCountForPayFlowPro( sParameter )
	parmList = parmList + "&comment1=" & sParameter 
	AddToPaymentLog iPaymentControlNumber, "COMMENT1: " & sParameter 

	sParameter = Left(sRentalName, 89) & " on " & sSelectedDate & " from " & sStartTime & " to " & sEndTime
	sReservationDetails = sParameter
	strLength = CleanAndCountForPayFlowPro( sParameter )
	parmList = parmList + "&comment2=" & sParameter
	AddToPaymentLog iPaymentControlNumber, "COMMENT2: " & sParameter

	parmList = parmList + "&paymentcontrolnumber=" & iPaymentControlNumber
	parmList = parmList + "&orgid=" & iOrgId
	parmList = parmList + "&orgfeature=rentals"

	'response.write parmList & "<br />"
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
	sResult = clng(sResult)
	sPNREF = GetResponseValue(transResponse, "PNREF")
	sRespMsg = GetResponseValue(transResponse, "RESPMSG")

	If sResult = clng(0) Then 
		sAuthcode = GetResponseValue(transResponse, "AUTHCODE")

		' IN CASE ERROR PROCESSING AFTER PAYMENT RECIEVED DISPLAY PAYMENT INFORMATION TO USER
		response.write vbcrlf & "<p><h2>Payment Processing Failed!</h2><div>"
		response.write "<strong>Your credit card was charged, but our application failed to process your payment.</strong><br />"
		response.write "AUTHORIZATION NUMBER:" & sAuthcode & "<br />"
		response.write "PAYMENT REFERENCE NUMBER:" & sPNREF & "<br />"
		response.write "MESSAGE:" & sRespMsg & "<br>"
		response.write "AMOUNT:" & dTotalAmount & "<br>"
		response.write "Credit Card Number: xxxx-xxxx-xxxx-" & RIGHT(sCardNumber,4) & "<br />"
		response.write "</div></p>"

		session("PAYMENTPROCESSING") = "TRUE"

		' Successful Transaction 
		AddToPaymentLog iPaymentControlNumber, "TRANSACTION SUCCEEDED"
		approved = True
		AddToPaymentLog iPaymentControlNumber, "AUTHCODE: " & sAuthcode
		AddToPaymentLog iPaymentControlNumber, "PNREF: " & sPNREF
		AddToPaymentLog iPaymentControlNumber, "RESPMSG: " & sRespMsg
		
		response.write "TRANSACTION SUCCEEDED" & "<br />"
		response.write "AUTHCODE: " & sAuthcode & "<br />"
		response.write "PNREF: " & sPNREF & "<br />"
		response.write "RESPMSG: " & sRespMsg & "<br />"

		ProcessSuccessfulTransaction iReservationTempId, dTotalAmount, sAuthcode, sPNREF, sRespMsg, "NULL", "NULL", "0.00"

	ElseIf sResult < clng(0) Then 
		'response.End 
		' Communication Error
		AddToPaymentLog iPaymentControlNumber, "Communication Error"
		AddToPaymentLog iPaymentControlNumber, "Result: " & sResult
		AddToPaymentLog iPaymentControlNumber, "PNREF: " & sPNREF
		AddToPaymentLog iPaymentControlNumber, "RESPMSG: " & sRespMsg
		approved = False 
		response.write "<div class=""payflowmsgfail"">Your credit card purchase was unable to be processed because of a network communication error. Please try your transaction again later.<blockquote><font color=""#000000"">Payment Reference Number:</font> " & sPNREF & " <br><font color=""#000000"">Description:</font> (" & sResult & ") - " & sRespMsg & " </blockquote></div>"

	ElseIf sResult > clng(0) Then 
		'response.End 
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
		' ProcessDeclinedTransaction sRESULT, sPNREF, sRESPMSG, sAmount, sAccountNumber, sReservationDetails , sOrderNumber, sSVA
		ProcessDeclinedTransaction "paypal", sResult, sPNREF, sRespMsg, dTotalAmount, sCardNumber, sReservationDetails, "NULL", "NULL"

	End If  
			
	AddToPaymentLog iPaymentControlNumber, "PAYMENT PROCESSING FINISHED."

End Sub 


'------------------------------------------------------------------------------
' void ProcessPointAndPayTransaction iReservationTempId, sFirstName, sLastName, sCardNumber, sExpiration, dTotalAmount, sRentalName, sSelectedDate, sStartTime, sEndTime, bOrgHasCVV 
'------------------------------------------------------------------------------
Sub ProcessPointAndPayTransaction( ByVal iReservationTempId, ByVal sFirstName, ByVal sLastName, ByVal sCardNumber, ByVal sExpiration, ByVal dTotalAmount, ByVal sRentalName, ByVal sSelectedDate, ByVal sStartTime, ByVal sEndTime, ByVal bOrgHasCVV )
	Dim parmList, objWinHttp, sParameter, strLength, sResult, sPNREF, sRespMsg, sAuthcode, sReservationDetails
	Dim sStatus, sErrorMsg, dFeeAmount, sOrderNumber, sSVA, sTotalCharges, sNotes

	iPaymentControlNumber = CreatePaymentControlRow( "PAYMENT SCRIPT STARTED - Point And Pay." )

	parmList = "paymentcontrolnumber=" & iPaymentControlNumber
	parmList = parmList + "&chargeaccountnumber=" & sCardNumber
	parmList = parmList + "&chargeexpirationmmyy=" & sExpiration  ' format is MMYY
	parmList = parmList + "&signerfirstname=" + sFirstName 
	parmList = parmList + "&signerlastname=" + sLastName
	AddToPaymentLog iPaymentControlNumber, "Name: " & sFirstName & " " & sLastName 

	If bOrgHasCVV And request("cvv2") <> "" Then
		parmList = parmList + "&chargecvn=" & request("cvv2")
		AddToPaymentLog iPaymentControlNumber, "ChargeCVN: present but not stored"
	End If 

	parmList = parmList + "&chargeamount=" & dTotalAmount
	AddToPaymentLog iPaymentControlNumber, "Amount: " & FormatNumber(dTotalAmount,2,,,0)

	parmList = parmList + "&signeraddressline1=" & request("StreetAddress")
	parmList = parmList + "&signeraddresscity=" & request("City")
	parmList = parmList + "&signeraddressregioncode=" & request("State")
	parmList = parmList + "&signeraddresspostalcode=" & request("ZipCode")

	sNotes = Left(sRentalName, 89) & " on " & sSelectedDate & " from " & sStartTime & " to " & sEndTime
	sReservationDetails = sNotes
	AddToPaymentLog iPaymentControlNumber, "Notes: " & CleanAndCutForPNPNotes( sNotes )
	parmList = parmList + "&notes=" & CleanAndCutForPNPNotes( sNotes )

	'response.write parmList & "<br />"


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

	sStatus = GetPNPResponseValue(transResponse, "status")		' in ../includes/common.asp
	AddToPaymentLog iPaymentControlNumber, "status: " & sStatus

	sErrorMsg = GetPNPResponseValue(transResponse, "errors")	' in ../includes/common.asp
	AddToPaymentLog iPaymentControlNumber, "errors: " & sErrorMsg

	sSVA = GetPNPResponseValue(transResponse, "sva")	' in ../includes/common.asp
	AddToPaymentLog iPaymentControlNumber, "sva: " & sSVA

	sOrderNumber = GetPNPResponseValue(transResponse, "orderNumber")	' in ../includes/common.asp
	AddToPaymentLog iPaymentControlNumber, "orderNumber: " & sOrderNumber

	If LCase(sStatus) <> "success" Then		' change this to approved when in prod
		' They were declined or there was an error
		approved = False
		ProcessDeclinedTransaction "PointAndPay", "declined", "", sErrorMsg, dTotalAmount, sCardNumber, sReservationDetails, sOrderNumber, sSVA
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

		' IN CASE ERROR PROCESSING AFTER PAYMENT RECIEVED DISPLAY PAYMENT INFORMATION TO USER
		response.write vbcrlf & "<p><h2>Payment Processing Failed!</h2><div>"
		response.write "<strong>Your credit card was charged, but our application failed to process your payment.</strong><br />"
		response.write "Order Number: " & sOrderNumber & "<br />"
		response.write "SVA: " & sSVA & "<br />"
		response.write "Amount Charged:" & sTotalCharges & "<br />"
		response.write "Credit Card Number: xxxx-xxxx-xxxx-" & RIGHT(sCardNumber,4) & "<br />"
		response.write "</div></p>"

		session("PAYMENTPROCESSING") = "TRUE"
		approved = True

		'response.write "Going to ProcessSuccessfulTransaction<br /><br />"
		ProcessSuccessfulTransaction iReservationTempId, dTotalAmount, "NULL", "NULL", "approved", sOrderNumber, sSVA, dFeeAmount
	End If 

	AddToPaymentLog iPaymentControlNumber, "PAYMENT PROCESSING FINISHED."

End Sub 


'--------------------------------------------------------------------------------------------------
' void ProcessSuccessfulTransaction iReservationTempId, dTotalAmount, sAuthcode, sPNREF, sRespMsg, sOrderNumber, sSVA, dFeeAmount
'--------------------------------------------------------------------------------------------------
Sub ProcessSuccessfulTransaction( ByVal iReservationTempId, ByVal dTotalAmount, ByVal sAuthcode, ByVal sPNREF, ByVal sRespMsg, ByVal sOrderNumber, ByVal sSVA, ByVal dFeeAmount )
	Dim sUserType, iInitialStatusId, bOffSeasonFlag, iReservationDateId, iWeekday, iItemTypeID
	Dim iAdminLocationId, iJournalEntryTypeID, iPaymentLocationId, iCitizenAccountId, sCheck
	Dim sPlusMinus, cPriorBalance, bNoCostToRent, dPaymentTotal, sPurchaseNotes, iPaymentTypeId
	Dim iAccountId, iPaymentId, iLedgerId, adminEmailAddr, sCitizenEmailAddress, iRentalId, iOrgId
	Dim sStartDateTime, sEndDateTime, sBillingEndDateTime, sArrivalDateTime, sDepartureDateTime, iRentalUserid
	Dim iIncludedPriceTypeId

	' pull things from the temp table
	GetReservationTempInformation iReservationTempId, iRentalId, sStartDateTime, sEndDateTime, sBillingEndDateTime, sArrivalDateTime, sDepartureDateTime, iRentalUserid, iOrgId, iIncludedPriceTypeId
	'response.write "iOrgId = " & iOrgId & "<br /><br />"

	sReservationTypeId = GetReservationTypeIdBySelector( "public" )
	'response.write "iRentalUserid = " & iRentalUserid & "<br /><br />"

	sUserType = GetUserResidentType( iRentalUserid )

	'If they are not one of these (R, N), we have to figure which they are
	If sUserType <> "R" And sUserType <> "N" Then 
		'This leaves E and B - See if they are a resident, also
		sUserType = GetResidentTypeByAddress( iRentalUserid, iOrgId )
	End If 

	iInitialStatusId = GetInitialReservationStatusId( iOrgid )

	' Create the reservation row
	sSql = "INSERT INTO egov_rentalreservations ( orgid, reservationtypeid, reservationstatusid, rentaluserid, "
	sSql = sSql & "adminuserid, reserveddate, originalrentalid ) VALUES ( " & iOrgID & ", " & sReservationTypeId & ", "
	sSql = sSql & iInitialStatusId & ", " & iRentalUserid & ", NULL, dbo.GetLocalDate("
	sSql = sSql & iOrgID & ",getdate()), " & iRentalId & " )"
	'response.write sSql & "<br /><br />"
	iReservationId = RunIdentityInsertStatement( sSql )


	' Handle the reservation date 
	bOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(CDate(sStartDateTime)) )

	' The real end time includes the end buffer if this is not to closing time for the rental
	If EndTimeIsNotClosingTime( iRentalId, sBillingEndDateTime, bOffSeasonFlag, sStartDateTime ) Then 
		' Add on the end buffer
		sEndDateTime = AddPostBufferTime( iRentalid, bOffSeasonFlag, sBillingEndDateTime, sStartDateTime )
	Else
		sEndDateTime = sBillingEndDateTime
	End If 

	sSql = "INSERT INTO egov_rentalreservationdates ( reservationid, rentalid, orgid, statusid, reservationstarttime, "
	sSql = sSql & "reservationendtime, billingendtime, actualstarttime, actualendtime, adminuserid, reserveddate ) VALUES ( "
	sSql = sSql & iReservationId & ", " & iRentalid & ", " & iOrgId & ", " & iInitialStatusId & ", '" & sStartDateTime & "', '"
	sSql = sSql & sEndDateTime & "', '" & sBillingEndDateTime & "', '" & sArrivalDateTime & "', '" & sDepartureDateTime & "', "
	sSql = sSql & "NULL, " & "dbo.GetLocalDate(" & iOrgID & ",getdate()) )"
	'response.write sSql & "<br /><br />"
	iReservationDateId = RunIdentityInsertStatement( sSql )

	iWeekday = Weekday(sStartDateTime)

	'dHours = CalculateDurationInHours( sStartDateTime, sBillingEndDateTime )

	' create the reservation date fees rows
	CreateRentalReservationDateFees iReservationDateId, iReservationId, iRentalid, bOffSeasonFlag, iWeekday, sUserType, sStartDateTime, sBillingEndDateTime

	' Create the rows for the reservation date items - These will not have any cost, but will just be place holders
	CreateRentalReservationDateItems iReservationDateId, iReservationId, iRentalid

	' Create the rental reservation fees - These are things like deposits, alcohol surcharge, damages charges
	CreateRentalReservationFees iReservationId, iRentalid, iIncludedPriceTypeId


	' Get the itemtype for the payment
	iItemTypeID = GetItemTypeId( "rentals" )
	iAdminLocationId = 0
	iJournalEntryTypeID = GetJournalEntryTypeID( "rentalpayment" )
	iPaymentLocationId = "3" ' this is the public website
	iCitizenAccountId = "NULL"
	sCheck = "NULL"
	sPlusMinus = "+"
	cPriorBalance = "NULL"

	' now handle the payment information 
	bNoCostToRent = RentalHasNoCosts( iRentalId )

	If bNoCostToRent Then
		dPaymentTotal = CDbl(0.00) ' Payment total
		sPurchaseNotes = "No Charge Reservation from Public Site."
		iPaymentTypeId = GetRentalPaymentTypeId( iOrgid, "isothermethod" )	' in rentalscommonfunctions.asp
	Else
		dPaymentTotal = CDbl(dTotalAmount)
		sPurchaseNotes = "Reservation from Public Site."
		iPaymentTypeId = GetRentalPaymentTypeId( iOrgid, "ispublicmethod" )	' in rentalscommonfunctions.asp
	End If 

	iAccountId = GetPaymentAccountId( iOrgid, iPaymentTypeId )		' In common.asp


	'Insert the egov_class_payment row (Journal entry)
	sSql = "INSERT INTO egov_class_payment (paymentdate, paymentlocationid, orgid, adminlocationid, "
	sSql = sSql & " userid, adminuserid, paymenttotal, journalentrytypeid, notes, isforrentals, reservationid) VALUES (dbo.GetLocalDate(" & iOrgid & ",GetDate()), " 
	sSql = sSql & iPaymentLocationId & ", " & iOrgId & ", " & iAdminLocationId & ", " & iRentalUserid & ", NULL, "
	sSql = sSql & dPaymentTotal & ", " & iJournalEntryTypeID & ", '" & sPurchaseNotes & "', 1, " & iReservationId & " )"
	'response.write sSql & "<br /><br />"
	iPaymentId = RunIdentityInsertStatement( sSql )


	'Make the ledger entry for the payment
	'iLedgerId = MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iPriceTypeid )
	'iLedgerId = MakeLedgerEntry( iOrgid, iAccountId, iPaymentId, CDbl(request("paymentamount" & x)), "NULL", "debit", sPlusMinus, "NULL", 1, x, cPriorBalance, "NULL" )
	sSql = "INSERT INTO egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
	sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, pricetypeid, reservationid ) VALUES ( "
	sSql = sSql & iPaymentId & ", " & iOrgid & ", 'debit', " & iAccountId & ", " & dPaymentTotal & ", NULL, '" & sPlusMinus & "', " 
	sSql = sSql & " NULL, 1, " & iPaymentTypeId & ", " & cPriorBalance & ", NULL, " & iReservationId & " )"
	'response.write sSql & "<br /><br />"
	iLedgerId = RunIdentityInsertStatement( sSql )


	If sAUTHCODE <> "NULL" Then
		sAUTHCODE = "'" & sAUTHCODE & "'"
	End If 

	If sPNREF <> "NULL" Then
		sPNREF = "'" & sPNREF & "'"
	End If 

	If sOrderNumber <> "NULL" Then
		sOrderNumber = "'" & sOrderNumber & "'"
	End If 

	If sSVA <> "NULL" Then
		sSVA = "'" & sSVA & "'"
	End If 

	'Make the entry in the egov_verisign_payment_information table
	'InsertPaymentInformation iPaymentId, iLedgerId, x, CDbl(request("amount" & x)), "APPROVED", sCheck, iCitizenAccountId
	'InsertPaymentInformation iPaymentId, iLedgerId, iPaymentTypeId, sAmount, sStatus, sCheckNo, iAccountId
	sSql = "INSERT INTO egov_verisign_payment_information ( paymentid, ledgerid, paymenttypeid, amount, "
	sSql = sSql & "paymentstatus, checkno, citizenuserid, authorizationcode, paymentreferenceid, "
	sSql = sSql & "paymentmessage, sva, processingfee, ordernumber ) Values (" & iPaymentId & ", " & iLedgerId & ", " 
	sSql = sSql & iPaymentTypeId & ", " & dPaymentTotal & ", 'APPROVED', " & sCheck & ", " & iCitizenAccountId & ", "
	sSql = sSql & sAuthcode & ", " & sPNREF & ", '" & sRespMsg & "', " & sSVA & ", " & dFeeAmount & ", "
	sSql = sSql & sOrderNumber & " )"
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql


	' Now pull the daily rates and process them
	sSql = "SELECT reservationdatefeeid, reservationdateid, ISNULL(feeamount,0.00) AS feeamount "
	sSql = sSql & " FROM egov_rentalreservationdatefees WHERE reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF

		' Get the account for the rate
		iAccountId = GetReservationAccountId( oRs("reservationdatefeeid"), "reservationdatefeeid", "egov_rentalreservationdatefees" )	' In rentalscommonfunctions.asp

		' Add to Accounts Ledger Row
		sSql = "INSERT Into egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
		sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, reservationid, reservationfeetypeid, reservationfeetype, reservationdateid ) VALUES ( "
		sSql = sSql & iPaymentId & ", " & iOrgid & ", 'credit', " & iAccountId & ", " & CDbl(oRs("feeamount")) & ", " & iItemTypeId & ", '+', " 
		sSql = sSql & iReservationId & ", 0, NULL, NULL, " & iReservationId & ", " & oRs("reservationdatefeeid") & ", 'reservationdatefeeid', " & oRs("reservationdateid") & " )"
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql

		oRs.MoveNext
	Loop
		
	oRs.Close
	Set oRs = Nothing 

	' Now pull the reservation fees and process them 
	sSql = "SELECT reservationfeeid, ISNULL(feeamount,0.00) AS feeamount "
	sSql = sSql & " FROM egov_rentalreservationfees WHERE reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF

		' Get the account for this fee
		iAccountId = GetReservationAccountId( oRs("reservationfeeid"), "reservationfeeid", "egov_rentalreservationfees" )	' In rentalscommonfunctions.asp

		' Add to Accounts Ledger Row
		sSql = "INSERT Into egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
		sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, reservationid, reservationfeetypeid, reservationfeetype ) VALUES ( "
		sSql = sSql & iPaymentId & ", " & iOrgid & ", 'credit', " & iAccountId & ", " & CDbl(oRs("feeamount")) & ", " & iItemTypeId & ", '+', " 
		sSql = sSql & iReservationId & ", 0, NULL, NULL, " & iReservationId & ", " & oRs("reservationfeeid" & x) & ", 'reservationfeeid' )"
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql
		oRs.MoveNext

	Loop

	oRs.Close
	Set oRs = Nothing 


	' Update the total amount due and paid on the reservation
	sSql = "UPDATE egov_rentalreservations SET totalamount = " & CDbl(dTotalAmount) & ", totalpaid = " & CDbl(dPaymentTotal)
	sSql = sSql & " WHERE reservationid = " & iReservationId
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql

	' clear out the temp table row 
	ClearTempReservation iReservationTempId, iOrgId

	' Send out any emails
	If sDefaultEmail <> "" Then 
		adminEmailAddr = sDefaultEmail  ' City default email
	Else 
		adminEmailAddr = "noreply@eclink.com"
	End If 

	sCitizenEmailAddress = GetCitizenEmail( iRentalUserid )

	' Send an eamil to the admins set up to get alerts for this rental
	SendAdminEmailAlerts iPaymentId, iReservationId, iRentalId, CDbl(dPaymentTotal), Replace(sAuthcode,"'",""), Replace(sPNREF,"'",""), bNoCostToRent, adminEmailAddr, Replace(sOrderNumber,"'",""), Replace(sSVA,"'",""), dFeeAmount

	' Send an email to the renting citizen
	SendCitizenConfirmation iPaymentId, iReservationId, iRentalId, CDbl(dPaymentTotal), Replace(sAuthcode,"'",""), Replace(sPNREF,"'",""), bNoCostToRent, sCitizenEmailAddress, adminEmailAddr, Replace(sOrderNumber,"'",""), Replace(sSVA,"'",""), dFeeAmount

	AddToPaymentLog iPaymentControlNumber, "PAYMENT PROCESSING FINISHED." 


	' Take them to the receipt page 
	LogThePage()
	response.redirect "view_receipt.asp?ipaymentid=" & iPaymentId & "&userid=" & iRentalUserid

End Sub 



'--------------------------------------------------------------------------------------------------
' void GetReservationTempInformation iReservationTempId, iRentalId, sStartDateTime, sEndDateTime, sBillingEndDateTime, sArrivalDateTime, sDepartureDateTime, iRentalUserid, iOrgId 
'--------------------------------------------------------------------------------------------------
Sub GetReservationTempInformation( ByVal iReservationTempId, ByRef iRentalId, ByRef sStartDateTime, ByRef sEndDateTime, ByRef sBillingEndDateTime, ByRef sArrivalDateTime, ByRef sDepartureDateTime, ByRef iRentalUserid, ByRef iOrgId, ByRef iIncludedPriceTypeId )
	Dim sSql, oRs

	sSql = "SELECT rentalid, orgid, selecteddate, starthour, startminute, startampm, endhour, endminute, endampm, "
	sSql = sSql & " citizenuserid, arrivalhour, arrivalminute, arrivalampm, departurehour, departureminute, "
	sSql = sSql & " departureampm, ISNULL(includepricetypeid,0) AS includepricetypeid "
	sSql = sSql & " FROM egov_rentalreservationstemppublic "
	sSql = sSql & " WHERE reservationtempid = " & iReservationTempId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iRentalId = oRs("rentalid")
		iOrgId = oRs("orgid")
		'response.write "iOrgId in GetINfo: " & iOrgId & "<br /><br />"
		sStartDateTime = CDate(oRs("selecteddate") & " " & oRs("starthour") & ":" & oRs("startminute") & " " & oRs("startampm"))
		sEndDateTime = CDate(oRs("selecteddate") & " " & oRs("endhour") & ":" & oRs("endminute") & " " & oRs("endampm"))
		If sEndDateTime < sStartDateTime Then
			sEndDateTime = DateAdd("d", 1, sEndDateTime)
		End If 
		sBillingEndDateTime = sEndDateTime
		sArrivalDateTime = CDate(oRs("selecteddate") & " " & oRs("arrivalhour") & ":" & oRs("arrivalminute") & " " & oRs("arrivalampm"))
		sDepartureDateTime = CDate(oRs("selecteddate") & " " & oRs("departurehour") & ":" & oRs("departureminute") & " " & oRs("departureampm"))
		If sDepartureDateTime < sStartDateTime Then
			sDepartureDateTime = DateAdd("d", 1, sDepartureDateTime)
		End If 
		iRentalUserid = oRs("citizenuserid")
		iIncludedPriceTypeId = oRs("includepricetypeid")
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void SendAdminEmailAlerts iPaymentId, iReservationId, iRentalId, dTotalAmount, sAuthcode, sPNREF, bNoCostToRent, adminEmailAddr, sOrderNumber, sSVA, dFeeAmount
'--------------------------------------------------------------------------------------------------
Sub SendAdminEmailAlerts( ByVal iPaymentId, ByVal iReservationId, ByVal iRentalId, ByVal dTotalAmount, ByVal sAuthcode, ByVal sPNREF, ByVal bNoCostToRent, ByVal adminEmailAddr, ByVal sOrderNumber, ByVal sSVA, ByVal dFeeAmount )
	Dim sSql, oRs, sMessageBody, sSubject, sLocation

	sSubject = sOrgName & " Reservation made from Public E-Gov Site"

	sLocation = GetRentalLocation( iRentalId )

	sMessageBody = GetAdminEmailBody( iReservationId, dTotalAmount, sPNREF, sAuthcode, bNoCostToRent, adminEmailAddr, sOrderNumber, sSVA, dFeeAmount, sLocation )

	' Pull those flagged to get the alerts
	sSql = "SELECT ISNULL(U.email,'') AS email "
	sSql = sSql & " FROM egov_rentalalerts R, egov_rentalalerttypes A, Users U "
	sSql = sSql & " WHERE R.rentalalerttypeid = A.rentalalerttypeid AND R.userid = U.userid AND "
	sSql = sSql & " A.isforpublicreservations = 1 AND rentalid = " & iRentalId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF 
		If oRs("Email") <> "" Then 
			If IsValidEmail( oRs("Email") ) Then 
				'sendEmail "", oRs("email"), "tfoster@eclink.com", sSubject, sMessageBody, "", "Y"
				sendEmail "", oRs("email"), "", sSubject, sMessageBody, "", "Y"
			End If 
		End If 
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' string GetAdminEmailBody( iReservationId, dTotalAmount, sPNREF, sAuthcode, bNoCostToRent, adminEmailAddr, sOrderNumber, sSVA, dFeeAmount, sLocation )
'--------------------------------------------------------------------------------------------------
Function GetAdminEmailBody( ByVal iReservationId, ByVal dTotalAmount, ByRef sPNREF, ByVal sAuthcode, ByVal bNoCostToRent, ByVal adminEmailAddr, ByVal sOrderNumber, ByVal sSVA, ByVal dFeeAmount, ByVal sLocation )
	Dim sMessage, sRentalName, sReservationDate, sStartTime, sEndTime, sName, sAddress, sCity, sState, sZip

	sMessage = sMessage & "<p>This automated message was sent by the " & sOrgName
	sMessage = sMessage & " E-Gov web site. Do not reply to this message.  Contact " 
	sMessage = sMessage & adminEmailAddr & " for inquiries regarding this email.</p>" & vbcrlf  & vbcrlf

	sMessage = sMessage & "<p>A reservation was made on " & Date() & ".</p>" & vbcrlf  & vbcrlf 
	
	sMessage = sMessage & "<p><strong>Transaction Details:</strong>" & vbcrlf
	If Not bNoCostToRent Then 
		sMessage = sMessage & "<br />Credit Card: XXXXXXXXXXXX" & Right(request("accountnumber"),4)  & vbcrlf
		If sSVA = "NULL" Then 
			sMessage = sMessage & "<br />Amount: " & FormatCurrency(dTotalAmount,2) & vbcrlf & vbcrlf
			sMessage = sMessage & "<br />Payment Reference Number: " & sPNREF & vbcrlf
			sMessage = sMessage & "<br />Authorization Code:" & sAUTHCODE & vbcrlf
		Else 
			' This is Point and Pay
			sMessage = sMessage & "<br />Purchase Amount: " & FormatCurrency(dTotalAmount,2) & vbcrlf
			sMessage = sMessage & "<br />Processing Fee: " & FormatCurrency(dFeeAmount,2) & vbcrlf
			sMessage = sMessage & "<br />Amount Charged: " & FormatCurrency((CDbl(dTotalAmount) + CDbl(dFeeAmount)),2) & vbcrlf
			sMessage = sMessage & "<br />Order Number: " & sOrderNumber & vbcrlf
			sMessage = sMessage & "<br />SVA:" & sSVA & vbcrlf & vbcrlf
		End If 
	Else
		sMessage = sMessage & "<br />There was no charge for this reservation." & vbcrlf & vbcrlf
	End If 
	
	' show the rental name, reservation date, reservation start time and end time
	sMessage = sMessage & "</p><p><strong>Reservation Details: </strong>" & vbcrlf 
	GetRentalDetailsByReservationId iReservationId, sRentalName, sReservationDate, sStartTime, sEndTime
	sMessage = sMessage & "<br />Location: " & sLocation & " &ndash; " & sRentalName
	'sMessage = sMessage & "<br />Rental: " & sRentalName
	sMessage = sMessage & "<br />Date: " & sReservationDate
	sMessage = sMessage & "<br />Start Time: " & sStartTime
	sMessage = sMessage & "<br />End Time: " & sEndTime

	sMessage = sMessage & "</p><p><strong>Renter Information:</strong>" & vbcrlf
	If request("sjname") <> "" Then 
		sMessage = sMessage & "<br />Name: "  & request("sjname") & vbcrlf
		sMessage = sMessage & "<br />Address: "  & request("streetaddress") & vbcrlf
		sMessage = sMessage & "<br />City: "  & request("city") & vbcrlf
		sMessage = sMessage & "<br />State: "  & request("state") & vbcrlf
		sMessage = sMessage & "<br />Zip: "  & request("zipcode") & "</p>" & vbcrlf & vbcrlf
	Else
		' Get the renters info
		GetRenterInfoByReservationId iReservationId, sName, sAddress, sCity, sState, sZip
		sMessage = sMessage & "<br />Name: "  & sName & vbcrlf
		sMessage = sMessage & "<br />Address: "  & sAddress & vbcrlf
		sMessage = sMessage & "<br />City: "  & sCity & vbcrlf
		sMessage = sMessage & "<br />State: "  & sState & vbcrlf
		sMessage = sMessage & "<br />Zip: "  & sZip & "</p>" & vbcrlf & vbcrlf
	End If 
	'response.write sMessage & "<br /><br />"

	GetAdminEmailBody = sMessage
End Function


'--------------------------------------------------------------------------------------------------
' void SendCitizenConfirmation iPaymentId, iReservationId, iRentalId, dTotalAmount, sAuthcode, sPNREF, bNoCostToRent, sCitizenEmailAddress, adminEmailAddr, sOrderNumber, sSVA, dFeeAmount
'--------------------------------------------------------------------------------------------------
Sub SendCitizenConfirmation( ByVal iPaymentId, ByVal iReservationId, ByVal iRentalId, ByVal dTotalAmount, ByVal sAuthcode, ByVal sPNREF, ByVal bNoCostToRent, ByVal sCitizenEmailAddress, ByVal adminEmailAddr, ByVal sOrderNumber, ByVal sSVA, ByVal dFeeAmount )
	Dim sMessageBody, sSubject, sLocation

	sSubject = "Thank You For Your " & sOrgName & " Reservation"

	sLocation = GetRentalLocation( iRentalId )

	sMessageBody = GetCitizenEmailBody( iReservationId, adminEmailAddr, bNoCostToRent, dTotalAmount, sAuthcode, sPNREF, iRentalId, sOrderNumber, sSVA, dFeeAmount, sLocation )

	'response.write "sCitizenEmailAddress = " & sCitizenEmailAddress & "<br /><br />"

	If IsValidEmail( sCitizenEmailAddress ) Then 
		sendEmail "", sCitizenEmailAddress, "", sSubject, sMessageBody, "", "N"
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' string GetCitizenEmailBody( iReservationId, adminEmailAddr, bNoCostToRent, dTotalAmount, sAuthcode, sPNREF, iRentalId, sOrderNumber, sSVA, dFeeAmount, sLocation )
'--------------------------------------------------------------------------------------------------
Function GetCitizenEmailBody( ByVal iReservationId, ByVal adminEmailAddr, ByVal bNoCostToRent, ByVal dTotalAmount, ByVal sAuthcode, ByVal sPNREF, ByVal iRentalId, ByVal sOrderNumber, ByVal sSVA, ByVal dFeeAmount, ByVal sLocation )
	Dim sMessage, sRentalName, sReservationDate, sStartTime, sEndTime, sName, sAddress, sCity, sState, sZip

	' Build the email message to the citizen
	sMessage = sMessage & "<p>This automated message was sent by the " & sOrgName
	sMessage = sMessage & " E-Gov web site. Do not reply to this message.  Contact " & adminEmailAddr
	sMessage = sMessage & " for inquiries regarding this email.</p>" & vbcrlf  & vbcrlf

	sMessage = sMessage & "<p>Thank you for your reservation made on " & Date() & ".</p>" & vbcrlf  & vbcrlf 

	If OrgHasDisplay( iOrgid, "facility receipt notes top" ) Then
		sMessage = sMessage & vbcrlf & vbcrlf & GetOrgDisplay( iOrgid, "facility receipt notes top" )
	End If

	sMessage = sMessage & vbcrlf & vbcrlf & "<p><strong>Transaction Details:</strong>" & vbcrlf
	If Not bNoCostToRent Then 
		sMessage = sMessage & "<br />Credit Card: XXXXXXXXXXXX" & Right(request("accountnumber"),4)  & vbcrlf
		If sSVA = "NULL" Then 
			' This is PayPal
			sMessage = sMessage & "<br />Amount Charged: " & FormatCurrency(dTotalAmount,2) & vbcrlf
			sMessage = sMessage & "<br />Payment Reference Number: " & sPNREF & vbcrlf
			sMessage = sMessage & "<br />Authorization Code:" & sAuthcode & vbcrlf & vbcrlf
		Else 
			' This is Point and Pay
			sMessage = sMessage & "<br />Purchase Amount: " & FormatCurrency(dTotalAmount,2) & vbcrlf
			sMessage = sMessage & "<br />Processing Fee: " & FormatCurrency(dFeeAmount,2) & vbcrlf
			sMessage = sMessage & "<br />Amount Charged: " & FormatCurrency((CDbl(dTotalAmount) + CDbl(dFeeAmount)),2) & vbcrlf
			sMessage = sMessage & "<br />Order Number: " & sOrderNumber & vbcrlf
			sMessage = sMessage & "<br />SVA:" & sSVA & vbcrlf & vbcrlf
		End If 
	Else
		sMessage = sMessage & "<br />There was no charge for this reservation." & vbcrlf & vbcrlf
	End If 

	sMessage = sMessage & "</p><p><strong>Reservation Details: </strong>" & vbcrlf 
	GetRentalDetailsByReservationId iReservationId, sRentalName, sReservationDate, sStartTime, sEndTime
	sMessage = sMessage & "<br />Location: " & sLocation & " &ndash; " & sRentalName
	'sMessage = sMessage & "<br />Rental: " & sRentalName
	sMessage = sMessage & "<br />Date: " & sReservationDate
	sMessage = sMessage & "<br />Start Time: " & sStartTime
	sMessage = sMessage & "<br />End Time: " & sEndTime

	sMessage = sMessage & "</p><p><strong>Renter Information:</strong>" & vbcrlf
	If request("sjname") <> "" Then 
		sMessage = sMessage & "<br />Name: "  & request("sjname") & vbcrlf
		sMessage = sMessage & "<br />Address: "  & request("streetaddress") & vbcrlf
		sMessage = sMessage & "<br />City: "  & request("city") & vbcrlf
		sMessage = sMessage & "<br />State: "  & request("state") & vbcrlf
		sMessage = sMessage & "<br />Zip: "  & request("zipcode") & "</p>" & vbcrlf & vbcrlf
	Else
		' Get the renters info
		GetRenterInfoByReservationId iReservationId, sName, sAddress, sCity, sState, sZip
		sMessage = sMessage & "<br />Name: "  & sName & vbcrlf
		sMessage = sMessage & "<br />Address: "  & sAddress & vbcrlf
		sMessage = sMessage & "<br />City: "  & sCity & vbcrlf
		sMessage = sMessage & "<br />State: "  & sState & vbcrlf
		sMessage = sMessage & "<br />Zip: "  & sZip & "</p>" & vbcrlf & vbcrlf
	End If 

	If OrgHasDisplay( iOrgid, "facility receipt notes bottom" ) Then
		sMessage = sMessage & vbcrlf & vbcrlf &  GetOrgDisplay( iOrgid, "facility receipt notes bottom" )
	End If
	'response.write sMessage & "<br /><br />"

	GetCitizenEmailBody = sMessage

End Function 


'--------------------------------------------------------------------------------------------------
' void GetRentalDetailsByReservationId iReservationId, sRentalName, sReservationDate, sStartTime, sEndTime
'--------------------------------------------------------------------------------------------------
Sub GetRentalDetailsByReservationId( ByVal iReservationId, ByRef sRentalName, ByRef sReservationDate, ByRef sStartTime, ByRef sEndTime )
	Dim sSql, oRs

	sSql = "SELECT R.rentalname, D.reservationstarttime, D.billingendtime "
	sSql = sSql & " FROM egov_rentals R, egov_rentalreservationdates D "
	sSql = sSql & " WHERE R.rentalid = D.rentalid AND D.reservationid = " & iReservationId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		sRentalName = oRs("rentalname")
		sReservationDate = DateValue(CDate(oRs("reservationstarttime")))
		sStartTime = FormatTimeString(CDate(oRs("reservationstarttime")))
		sEndTime = FormatTimeString(CDate(oRs("billingendtime")))
	Else 
		sRentalName = ""
		sReservationDate = ""
		sStartTime = ""
		sEndTime = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
'void GetRenterInfoByReservationId iReservationId, sName, sAddress, sCity, sState, sZip
'--------------------------------------------------------------------------------------------------
Sub GetRenterInfoByReservationId( ByVal iReservationId, ByRef sName, ByRef sAddress, ByRef sCity, ByRef sState, ByRef sZip )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(U.userfname,'') AS userfname, ISNULL(U.userlname,'') AS userlname, "
	sSql = sSql & " ISNULL(U.useraddress,'') AS useraddress, ISNULL(U.usercity,'') AS usercity, "
	sSql = sSql & " ISNULL(userstate,'') AS userstate, ISNULL(U.userzip,'') AS userzip "
	sSql = sSql & " FROM egov_rentalreservations R, egov_users U "
	sSql = sSql & " WHERE R.rentaluserid = U.userid AND R.reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		sName = Trim(oRs("userfname") & " " & oRs("userlname"))
		sAddress = oRs("useraddress")
		sCity = oRs("usercity")
		sState = oRs("userstate")
		sZip = oRs("userzip")
	Else
		sName = "" 
		sAddress = ""
		sCity = ""
		sState = ""
		sZip = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' String GetCitizenEmail( iRentalUserid )
'--------------------------------------------------------------------------------------------------
Function GetCitizenEmail( ByVal iRentalUserid )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(useremail,'') AS useremail "
	sSql = sSql & " FROM egov_users  WHERE userid = " & iRentalUserid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetCitizenEmail = Trim(oRs("useremail"))
	Else
		GetCitizenEmail = "" 
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
	sSql = sSql & iOrgID & ", 'public', 'rentals', '" & sLogEntry & "' )"
	'response.write sSql & "<br /><br />"

	iPaymentControlNumber = RunIdentityInsertStatement( sSql )

	sSql = "UPDATE paymentlog SET paymentcontrolnumber = " & iPaymentControlNumber
	sSql = sSql & " WHERE paymentlogid = " & iPaymentControlNumber
	'response.write sSql & "<br /><br />"
	RunSQLStatement( sSql )

	CreatePaymentControlRow = iPaymentControlNumber

End Function 


'------------------------------------------------------------------------------
' void AddToPaymentLog iPaymentControlNumber, sLogEntry 
'------------------------------------------------------------------------------
Sub AddToPaymentLog( ByVal iPaymentControlNumber, ByVal sLogEntry  )
	Dim sSql

	sSql = "INSERT INTO paymentlog ( paymentcontrolnumber, orgid, applicationside, feature, logentry ) VALUES ( "
	sSql = sSql & iPaymentControlNumber & ", " & iOrgID & ", 'public', 'rentals', '" & dbready_string(sLogEntry, 500) & "' )"
	'response.write sSql & "<br /><br />"
	RunSQLStatement( sSql )

End Sub 


'------------------------------------------------------------------------------
' double GetTotalCharges( iReservationTempId )
'------------------------------------------------------------------------------
Function GetTotalCharges( ByVal iReservationTempId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(feetotal,0.00) AS feetotal "
	sSql = sSql & "FROM egov_rentalreservationstemppublic WHERE reservationtempid = " & iReservationTempId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetTotalCharges = CDbl(oRs("feetotal"))
	Else
		GetTotalCharges = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


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


'------------------------------------------------------------------------------
' void GetSomeReservationDetails iReservationTempId, iRentalid, sRentalName, sSelectedDate, sStartTime, sEndTime, sStartDateTime, sEndDateTime
'------------------------------------------------------------------------------
Sub GetSomeReservationDetails( ByVal iReservationTempId, ByRef iRentalid, ByRef sRentalName, ByRef sSelectedDate, ByRef sStartTime, ByRef sEndTime, ByRef sStartDateTime, ByRef sEndDateTime )
	Dim sSql, oRs, iStartHour, iStartMinute, sStartAmPm, iEndHour, iEndMinute, sEndAmPm

	sSql = "SELECT R.rentalid, R.rentalname, P.selecteddate, ISNULL(P.starthour,1) AS starthour, "
	sSql = sSql & "dbo.AddLeadingZeros(ISNULL(P.startminute,0),2) AS startminute, "
	sSql = sSql & "ISNULL(P.startampm,'PM') AS startampm, ISNULL(P.endhour,2) AS endhour, "
	sSql = sSql & "dbo.AddLeadingZeros(ISNULL(P.endminute,0),2) AS endminute,  ISNULL(P.endampm,'PM') AS endampm "
	sSql = sSql & "FROM egov_rentals R, egov_rentalreservationstemppublic P "
	sSql = sSql & "WHERE R.rentalid = P.rentalid AND reservationtempid = " & iReservationTempId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iRentalid = oRs("rentalid")
		sRentalName = GetRentalName( iRentalId )	' In rentalcommonfunctions.asp
		sSelectedDate = oRs("selecteddate")
		iStartHour = oRs("starthour")
		iStartMinute = oRs("startminute")
		sStartAmPm = oRs("startampm")
		iEndHour = oRs("endhour")
		iEndMinute = oRs("endminute")
		sEndAmPm = oRs("endampm")
		sStartTime = iStartHour & ":" & iStartMinute & " " & sStartAmPm
		sStartDateTime = sSelectedDate & " " & sStartTime
		sEndTime = iEndHour & ":" & iEndMinute & " " & sEndAmPm
		sEndDateTime = sSelectedDate & " " & sEndTime
		' if the end date is less than the start date then it must end the next day
		'If sEndDateTime < sStartDateTime Then
		If datediff("n", sEndDateTime , sStartDateTime) >= 0 Then
			sEndDateTime = DateAdd("d", 1, sEndDateTime)
		End If 
	Else
		iRentalid = 0
		sRentalName = ""
		sSelectedDate = ""
		sStartTime = ""
		sEndTime = ""
		sStartDateTime = ""
		sEndDateTime = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

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
' void ProcessDeclinedTransaction sProcessingPath, sRESULT, sPNREF, sRESPMSG, sAmount, sAccountNumber, sReservationDetails, sOrderNumber, sSVA
'------------------------------------------------------------------------------
Sub ProcessDeclinedTransaction( ByVal sProcessingPath, ByVal sRESULT, ByVal sPNREF, ByVal sRESPMSG, ByVal sAmount, ByVal sAccountNumber, ByVal sReservationDetails, ByVal sOrderNumber, ByVal sSVA )
	
	' BRANDING
	response.write "<center>"
	sPaymentImg = GetPaymentImage( "../" )	' in ../include_top_functions.asp
	If sPaymentImg <> "" Then 
		response.write "<img src=""" & sPaymentImg & """ border=""0"" /><br /><br />"
	End If 
	
	response.write "<strong>" & GetPaymentGatewayName( ) & "</strong> has routed, processed, and secured your payment information."
	response.write "</p>"

	' TRANSACTION RESULT DETAILS
	response.write "<p><div class=""group""><p><h3>Your credit card purchase was declined for the following reason:</h3>"

	If sPNREF <> "" Then 
		response.write "Payment Reference Number: " & sPNREF & " <br />"
	End If 
	response.write "<strong>Description: (" & sRESULT & ") - " & sRESPMSG & "</strong></p>"

	response.write "<p>"
	response.write "<table id=""declineddisplay"" cellpadding=""2"" cellspacing=""0"" border=""0"" width=""100%"">"
	response.write "<tr><td colspan=""2"" class=""declinedsectiontitle""><b>Transaction Details</b></td></tr>"
	If sProcessingPath = "PointAndPay" Then
		response.write "<tr><td align=""right"">Purchase Amount:</td><td align=""left""> " & FormatCurrency(sAmount,2) & "</td></tr>"
		response.write "<tr><td align=""right"">Order Number:</td><td align=""left""> " & sOrderNumber & " </td></tr>"
		response.write "<tr><td align=""right"">SVA:</td><td align=""left""> " & sSVA & "</td></tr>"
	Else ' PayPal
		response.write "<tr><td align=""right"">Amount:</td><td align=""left""> " & FormatCurrency(sAmount,2) & "</td></tr>"
		response.write "<tr><td align=""right"">Reference Number:</td><td align=""left""> " & sPNREF & " </td></tr>"
		response.write "<tr><td align=""right"">Authorization Code:</td><td align=""left""> " & sAUTHCODE & "</td></tr>"
	End If 
	response.write "<tr><td colspan=""2"">&nbsp;</td></tr>"

	' RESERVATION INFORMATION
	response.write "<tr><td colspan=""2"" class=""declinedsectiontitle""><b>Reservation Information</b></td></tr>"
	response.write "<tr><td align=""right"">Details:</td><td align=""left""> " & sReservationDetails & " </td></tr>"
	response.write "<tr><td colspan=""2"">&nbsp;</td></tr>"

	' CREDIT CARD INFORMATION	
	response.write "<tr><td colspan=""2"" class=""declinedsectiontitle""><strong>Billing Information</strong></td></tr>"
	response.write "<tr><td align=""right"">Credit Card: </td><td align=""left"">XXXXXXXXXXXX" & Right(sAccountNumber,4)  & "</td></tr>"
	response.write "<tr><td align=""right"">Name: </td><td align=""left"">" & request("sjname") & "</td></tr>"
	response.write "<tr><td align=""right"">Address: </td><td align=""left"">" & request("streetaddress") & "</td></tr>"
	response.write "<tr><td align=""right"">City: </td><td align=""left"">" & request("city") & "</td></tr>"
	response.write "<tr><td align=""right"">State: </td><td align=""left"">" & request("state") & "</td></tr>"
	response.write "<tr><td align=""right"">Zip: </td><td align=""left"">" & request("zipcode") & "</td></tr>"

	response.write "</table></p>"
	response.write "</div>"
	response.write "</center><br /><br />"

End Sub


'----------------------------------------------------------------------------------------
' string CleanAndCutForPNPNotes( sParameter )
'----------------------------------------------------------------------------------------
Function CleanAndCutForPNPNotes( ByVal sParameter )
	' cleans forbidden characters and returns the string

	sParameter = Replace(sParameter, Chr(34), "")
	sParameter = Replace(sParameter, "'", "")
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


Sub LogThePage( )
	Dim sSql, oCmd, sScriptName, sVirtualDirectory, aVirtualDirectory, sPage, arr, sUserAgent, sUserAgentGroup

	sScriptName = Request.ServerVariables("SCRIPT_NAME")

	If request.servervariables("http_user_agent") <> "" Then 
		sUserAgent = "'" & Track_DBsafe(Trim(Left(request.servervariables("http_user_agent"),480))) & "'"
	Else
		sUserAgent = "NULL"
	End If 

	If Len(Trim(request.servervariables("http_user_agent"))) > 0 Then 
		sUserAgentGroup = "'" & GetUserAgentGroup( LCase(request.servervariables("http_user_agent")) ) & "'"
	Else
		sUserAgentGroup = "'" & GetUntrackedUserAgentGroup( ) & "'"
	End If 

	' Get the virtual directory
	aVirtualDirectory = Split(sScriptName, "/", -1, 1) 
	sVirtualDirectory = "/" & aVirtualDirectory(1) 
	sVirtualDirectory = "'" & Replace(sVirtualDirectory,"/","") & "'"

	' Get the page
	For Each arr in aVirtualDirectory 
		sPage = arr 
	Next 

	sSql = "INSERT INTO egov_pagelog ( virtualdirectory, applicationside, page, loadtime, scriptname, querystring, "
	sSql = sSql & " servername, remoteaddress, requestmethod, orgid, userid, username, sectionid, documenttitle, useragent, useragentgroup, requestformcollection, cookiescollection, sessioncollection, sessionid  ) VALUES ( "
	sSql = sSql & sVirtualDirectory & ", "
	sSql = sSql & "'public', "
	sSql = sSql & "'" & sPage & "', "
	sSql = sSql & FormatNumber(iLoadTime,3,,,0) & ", "
	sSql = sSql & "'" & sScriptName & "', "

	If Request.ServerVariables("QUERY_STRING") <> "" Then 
		sSql = sSql & "'" & Track_DBsafe(Left(Request.ServerVariables("QUERY_STRING"),500)) & "', "
	Else
		sSql = sSql & "NULL, "
	End If 
	' our server name
	sSql = sSql & "'" & Request.ServerVariables("SERVER_NAME") & "', "

	' remote address
	sSql = sSql & "'" & Request.ServerVariables("REMOTE_ADDR") & "', "

	' request method - GET or POST
	sSql = sSql & "'" & Request.ServerVariables("REQUEST_METHOD") & "', "

	' orgid
	If iorgid <> "" Then 
		sSql = sSql & iorgid & ", "
	Else
		sSql = sSql & "NULL, "
	End If 

	' Userid
	If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" and isnumeric(request.cookies("userid")) Then
		sSql = sSql & request.cookies("userid") & ", "
	Else
		sSql = sSql & "NULL, "
		response.cookies("userid") = ""
	End If 

	' Get username
	If sUserName <> "" Then
		sSql = sSql & "'" & Track_DBsafe(sUserName) & "', "
	Else
		sSql = sSql & "NULL, "
	End If 

	' Section Id for the old LogPageVisit functionality
	If iSectionID <> "" Then 
		sSql = sSql & iSectionID & ", "
	Else
		sSql = sSql & "NULL, "
	End If 

	' Document Title for the old LogPageVisit functionality
	If sDocumentTitle <> "" Then 
		sSql = sSql & "'" & Track_DBsafe(sDocumentTitle) & "',  "
	Else
		sSql = sSql & "NULL, "
	End If 

	' User Agent
	sSql = sSql & sUserAgent & ", "

	' User Agent Group
	sSql = sSql & sUserAgentGroup & ", "

	sSql = sSql & "'" & Track_DBsafe(GetRequestformInformation()) & "',"
	sSql = sSql & "'" & GetCookiesCollection() & "',"
	sSql = sSql & "'" & GetSessionCollection() & "',"


	sSql = sSql & "'" & Session.SessionID & "'"

	sSql = sSql & " )"
	'response.write sSql

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql

	session("sSql") = sSql
	oCmd.Execute
	session("sSql") = ""

	Set oCmd = Nothing


End Sub 
Function GetCookiesCollection()
	Collection = ""
	on error resume next
	For Each Item in Request.Cookies
		Collection = Collection & Item & ":  " & request.cookies(Item) & vbcrlf
	Next
	on error goto 0
	GetCookiesCollectionCollection = track_dbsafe(Collection)
End Function
Function GetSessionCollection()
	sSessionLog = ""
	on error resume next
	For each session_name in Session.Contents
		sSessionLog = sSessionLog & session_name & ":  " & session(session_name) & vbcrlf
	Next
	on error goto 0

	GetSessionCollection = track_dbsafe(sSessionLog)
End Function


'------------------------------------------------------------------------------
Function GetUserAgentGroup( ByVal sUserAgent )
	Dim sSql, oRs, sUserAgentGroup

	sUserAgentGroup = GetUntrackedUserAgentGroup()

	sSql = "SELECT useragentgroup FROM UserAgent_Groups WHERE isuntracked = 0 AND isactive = 1 ORDER BY checkorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If clng(InStr( 1, sUserAgent, LCase(oRs("useragentgroup")), 1 )) > clng(0) Then
			sUserAgentGroup = oRs("useragentgroup")
			Exit Do 
		End If 
		oRs.MoveNext
	Loop 
	
	oRs.Close
	Set oRs = Nothing 
	
	GetUserAgentGroup = sUserAgentGroup

End Function 


'------------------------------------------------------------------------------
Function GetUntrackedUserAgentGroup( )
	Dim sSql, oRs

	sSql = "SELECT useragentgroup FROM UserAgent_Groups WHERE isuntracked = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetUntrackedUserAgentGroup = oRs("useragentgroup")
	Else
		GetUntrackedUserAgentGroup = "untracked"
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 
'--------------------------------------------------------------------------------------------------
' FUNCTION GETREQUESTFORMINFORMATION()
'--------------------------------------------------------------------------------------------------
Function GetRequestFormInformation()
	Dim sReturnValue, key
	
	sReturnValue = ""

	For each key in request.Form
		If key <> "accountnumber" And key <> "cvv2" Then 
			sReturnValue = sReturnValue & key & ":" & request.form(key) & "<br />" & vbcrlf
		End If 
	Next 
	
	GetRequestFormInformation = sReturnValue

End Function



%>
