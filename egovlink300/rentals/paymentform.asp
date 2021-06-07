<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<% 'Response.Expires = -1000 %>
<!--#include file="../includes/common.asp" //-->
<!--#include file="../includes/start_modules.asp" //-->
<!--#Include file="rentalcommonfunctions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: paymentform.asp
' AUTHOR: Steve Loar
' CREATED: 03/3/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page takes payment information for rentals.
'
' MODIFICATION HISTORY
' 1.0	03/03/2010	Steve Loar - Initial version
' 2.0	06/23/2010	Steve Loar - Split name field into first and last 
' 2.1	07/07/2010	Steve Loar - Added Fee Amount fetch for PNP
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
response.Expires = 60
response.Expiresabsolute = Now() - 1
response.AddHeader "pragma","no-store"
response.AddHeader "cache-control","private"
response.CacheControl = "no-store" 'HTTP prevent back button

Dim sError, bOrgHasCVV, iReservationTempId, sTotalAmount, iItemCount, bHasPaymentFee, dFeeAmount, sErrorMsg, dTotalCharges
Dim sProcessingRoute

dFeeAmount = CDbl(0.00)

If request("rti") = "" Then
	response.redirect sEgovWebsiteURL & "/rentals/rentalcategories.asp"
Else 
	If Not IsNumeric(request("rti")) Then
		response.redirect sEgovWebsiteURL & "/rentals/rentalcategories.asp"
	Else 
		iReservationTempId = CLng(request("rti"))
	End If 
End If 

' check if payment gateway needs a fee check for this page
If PaymentGatewayRequiresFeeCheck( iOrgId ) Then
	If CitizenPaysFee( iOrgId ) Then
		bHasPaymentFee = True 
		
		' Get total to get the fee for
		sTotalAmount = GetTotalAmount( iReservationTempId )
		
		sProcessingRoute = GetProcessingRoute()		' In ../include_top_functions.asp
		If LCase(sProcessingRoute) = "pointandpay" Then 
			' Fetch the fee for the amount to be charged.
			If Not GetPNPFee( sTotalAmount, dFeeAmount, sErrorMsg ) Then		' in ../includes/common.asp
				'If not successful, store the error, then take them to a page to display the error message.
				iGatewayErrorId = StoreGatewayError( iOrgId, sProcessingRoute, "feecheck", sErrorMsg, FormatNumber(sTotalAmount,2,,,0) )		' in ../includes/common.asp
				response.redirect Application("CARTURL") & "/" & sorgVirtualSiteName & "/payment_processors/processing_failure.asp?ge=" & iGatewayErrorId
			End If 
		End If 
	Else
		bHasPaymentFee = False 
	End If 
Else
	bHasPaymentFee = false
End If 


'Check for org features
bOrgHasCVV = OrgHasFeature( iOrgId, "display cvv" )

%>

<html>
<head>
  	<meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<title>E-Gov Services <%=sOrgName%> - Payment Form</title>

<script src='https://www.google.com/recaptcha/api.js'></script>
	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalstyles.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script language="javascript" src="../scripts/modules.js"></script>

  	<script type="text/javascript" src="../scripts/jquery-1.9.1.min.js"></script>

	<script language="javascript">
	<!--

		function processPayment()
		{
			$("#sjname").val($("#firstname").val() + ' ' + $("#lastname").val());
			// disable the pay button
			$("#complete_payment_btn").prop( "disabled", true );
			// submit the form
			document.frmPayment.submit();
		}

		// set focus on the first field when the page loads
		$('document').ready(function(){
			$("#firstname").focus();
		});

	//-->
	</script>
<%
  If request.servervariables("HTTPS") = "on" Then 
     response.write vbcrlf & "<style>"
     response.write vbcrlf & "body {behavior: url('https://secure.egovlink.com/" & sorgVirtualSiteName & "/csshover.htc');}"
     response.write vbcrlf & "</style>"
  End If 
%>
</head>

<!--#Include file="../include_top.asp"-->

<!--BODY CONTENT-->
<tr>
    <td valign="top">

		<!--BEGIN: BUILD PAYMENT FORM-->
		<form name="frmPayment" action="rentalreservationmake.asp" method="post">
			<input type="hidden" id="rti" name="rti" value="<%=iReservationTempId%>" />
			<input type="hidden" id="src" name="src" value="pf" />

		<!--BEGIN:  DISPLAY PAYMENT FORM-->

<% 
		DisplayPaymentForm iReservationTempId, bOrgHasCVV, dFeeAmount, bHasPaymentFee
%>

		<!--END: DISPLAY PAYMENT FORM-->

		<!--SPACING CODE-->
		<p>&nbsp;</p>
		<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->   

<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void DisplayPaymentForm iReservationTempId, bOrgHasCVV, dFeeAmount, bHasPaymentFee
'--------------------------------------------------------------------------------------------------
Sub DisplayPaymentForm( ByVal iReservationTempId, ByVal bOrgHasCVV, ByVal dFeeAmount, ByVal bHasPaymentFee )
	Dim sTotal, sCitizenName, sAddress, sCity, sState, sZip, iOrgId, sSelectedDate, sStartTime, sEndTime
	Dim sRentalName, sReservationDetails, sEmail, sFirstName, sLastName

	GetTempReservationInformation iReservationTempId, iOrgId, sTotal, sRentalName, sCitizenName, sAddress, sCity, sState, sZip, sSelectedDate, sStartTime, sEndTime, sEmail, sFirstName, sLastName

	sReservationDetails = sRentalName & " on " & sSelectedDate & " from " & sStartTime & " to " & sEndTime

	If iOrgId = 37 Then 
		sCreditCardNum = "5555555555554444"
		sExpMonth      = "01"
		sExpYear       = "2013"
		sCVSCode       = ""
		sPhone         = "5136814030"
	End If 

	response.write vbcrlf & "<div class=""group"" id=""paymentform"">"
	response.write vbcrlf & "<div align=""center"" style=""width:550px;"">"

	'BEGIN: Reservation Details
	response.write vbcrlf & "<fieldset id=""reservationdetails"">"
	response.write vbcrlf & "<legend><strong>Reservation Details</strong></legend>"
	response.write "<li>" & sReservationDetails & "</li>"
	response.write vbcrlf & "</fieldset>"

	' Charge Amount
	response.write vbcrlf & "<fieldset>"
	response.write vbcrlf & "<legend><strong>Charges&nbsp;</strong></legend>"
	response.write vbcrlf & "<table border=""0"" cellpadding=""2"" cellspacing=""0"" id=""chargeamounts"" align=""left"">"

	response.write vbcrlf & "<tr align=""left"">"
	response.write "<td align=""right"" class=""labelcol"" nowrap=""nowrap"">Purchase Amount:</td>"
	response.write "<td align=""right"">" & FormatNumber(sTotal,2) & "<input type=""hidden"" name=""transactionamount"" value=""" & FormatNumber(sTotal,2,,,0) & """ /></td>"
	response.write "</tr>"

	If bHasPaymentFee Then
		' Show the fee
		response.write vbcrlf & "<tr align=""left"">"
		response.write "<td align=""right"" class=""labelcol"" nowrap=""nowrap"">Processing Fee:</td>"
		response.write "<td align=""right"">" & FormatNumber(dFeeAmount,2) & "</td>"
		response.write "</tr>"

		dTotalCharges = CDbl(dFeeAmount) + CDbl(sTotal)

		' Show the total
		response.write vbcrlf & "<tr align=""left"">"
		response.write "<td align=""right"" class=""labelcol"" nowrap=""nowrap"">Total Charges:</td>"
		response.write "<td align=""right""><strong>" & FormatNumber(dTotalCharges,2) & "</strong></td>"
		response.write "</tr>"
	End If 
	response.write vbcrlf & "</table>"
	response.write "<br />"
	response.write vbcrlf & "</fieldset>"
	' End Charge Amount

	response.write vbcrlf & "<p align=""left"">"
	response.write "Please enter your billing information as it appears on your credit card statement, "
	response.write "then click the <strong>Process Payment</strong> button."
	response.write vbcrlf & "</p>"

	response.write vbcrlf & "<fieldset>"
	response.write vbcrlf & "<legend><strong>Billing Information&nbsp;</strong></legend>"
	response.write vbcrlf & "<table border=""0"" cellpadding=""2"" cellspacing=""0"" style=""font-family:Verdana"">"

	'BEGIN: Billing Information
	response.write vbcrlf & "<tr align=""left"">"

	response.write "<td align=""right"" class=""billinginfolabel"">First Name:</td>"
	response.write "<td align=""left""><input type=""text"" id=""firstname"" name=""firstname"" value=""" & sFirstName & """ maxlength=""30"" size=""30"" />"
	response.write "<input type=""hidden"" id=""sjname"" name=""sjname"" value=""" & sCitizenName & """ />"
	response.write "</td>"
	response.write "</tr>"
	response.write vbcrlf & "<tr align=""left"">"
	response.write "<td align=""right"" class=""billinginfolabel"">Last Name:</td>"
	response.write "<td align=""left""><input type=""text"" id=""lastname"" name=""lastname"" value=""" & sLastName & """ maxlength=""30"" size=""30"" /></td>"
	response.write "</tr>"
'	response.write vbcrlf & "<tr align=""left"">"
'	response.write "<td align=""right"" class=""billinginfolabel"">Name:</td>"
'	response.write "<td><input name=""sjname"" value=""" & sCitizenName & """ maxlength=""50"" size=""30"" /></td>"
'	response.write "</tr>"

	response.write vbcrlf & "<tr align=""left"">"
	response.write "<td align=""right"" class=""billinginfolabel"">E-mail:</td>"
	response.write "<td align=""left""><input type=""text"" name=""email"" value=""" & sEmail & """ maxlength=""50"" size=""50"" /></td>"
	response.write "</tr>"
	response.write vbcrlf & "<tr align=""left"">"
	response.write "<td align=""right"" class=""billinginfolabel"">Address:</td>"
	response.write "<td align=""left""><input type=""text"" name=""streetaddress"" value=""" & sAddress & """ maxlength=""50"" size=""50"" /></td>"
	response.write "</tr>"
	response.write vbcrlf & "<tr align=""left"">"
	response.write "<td align=""right"" class=""billinginfolabel"">City:</td>"
	response.write "<td align=""left""><input type=""text"" name=""city"" value=""" & sCity & """ maxlength=""20"" size=""20"" /></td>"
	response.write "</tr>"
	response.write vbcrlf & "<tr align=""left"">"
	response.write "<td align=""right"" class=""billinginfolabel"">State:</td>"
	response.write "<td align=""left"">"
	response.write "<select name=""state"" size=""1"">"

	'displayStateOptions
	ShowStatePicks sDefaultState, sState

	response.write vbcrlf & "</select>"
	response.write "</td>"
	response.write "</tr>"

	response.write vbcrlf & "<tr align=""left"">"
	response.write "<td align=""right"" class=""billinginfolabel"">Zip:</td>"
	response.write "<td align=""left""><input type=""text"" name=""zipcode"" value=""" & sZip & """ maxlength=""15"" size=""15"" /></td>"
	response.write "</tr>"
	response.write vbcrlf & "</table>"
	response.write "<br />"
	response.write vbcrlf & "</fieldset>"
	response.write "<br />"
	'END: Personal Information

	'BEGIN: Credit Card Information
	response.write vbcrlf & "<fieldset>"
	response.write vbcrlf & "<legend><strong>Credit Card Information&nbsp;</strong></legend>"
	response.write vbcrlf & "<table border=""0"" cellpadding=""2"" cellspacing=""0"">"

	'Accepted card types
	response.write vbcrlf & "<tr align=""left"">"
	response.write "<td align=""right"" class=""billinginfolabel"">Credit Card Type:</td>"
	response.write "<td>"

	ShowCreditCardPicks		' In include_top_functions.asp

	response.write "</td>"
	response.write "</tr>"
	response.write vbcrlf & "<tr align=""left"">"
	response.write "<td align=""right"" class=""billinginfolabel"">Credit Card Number:</td>"
	response.write "<td><input type=""text"" name=""accountnumber"" value=""" & sCreditCardNum & """ maxlength=""22"" size=""30"" /></td>"
	response.write "</tr>"
	response.write vbcrlf & "<tr align=""left"">"
	response.write "<td align=""right"" class=""billinginfolabel"">Expiration Month:</td>"
	response.write "<td>"
	response.write vbcrlf & "<select name=""month"">"

	displayMonthOptions()

	response.write vbcrlf & "</select>"
	response.write "</td>"
	response.write "<td colspan=""2""></td>"
	response.write "<td></td>"
	response.write "<td>&nbsp;</td>"
	response.write "</tr>"
	response.write vbcrlf & "<tr align=""left"">"
	response.write "<td align=""right"" class=""billinginfolabel"">Expiration Year:</td>"
	response.write "<td>"
	response.write vbcrlf & "<select name=""year"">"

	displayYearOptions()

	response.write vbcrlf & "</select>"
	response.write "</td>"
	response.write "<td></td>"
	response.write "<td></td>"
	response.write "</tr>"

	If bOrgHasCVV Then 
		response.write vbcrlf & "  <tr align=""left"">"
		response.write "<td align=""right"" class=""billinginfolabel"">CVV Code:</td>"
		response.write "<td><input type=""text"" name=""cvv2"" size=""4"" maxlength=""4"" value="""" /></td>"
		response.write "<td colspan=""2""></td>"
		response.write "<td></td>"
		response.write "<td></td>"
		response.write "</tr>"
	End If 

'	response.write vbcrlf & "<tr align=""left"">"
'	response.write "<td align=""right"" class=""billinginfolabel"">Amount:</td>"
'	response.write "<td>" & FormatCurrency(sTotal,2) & "<input type=""hidden"" name=""transactionamount"" value=""" & FormatNumber(sTotal,2,,,0) & """ /></td>"
'	response.write "<td colspan=""2""></td>"
'	response.write "<td></td>"
'	response.write "<td></td>"
'	response.write "</tr>"


	response.write vbcrlf & "</table>"
	response.write "<p align=""center"">Do not use dashes or spaces when entering credit card information.</p>"
	response.write vbcrlf & "</fieldset>"
	'END: Credit Card Information

	response.write vbcrlf & "<div align=""left"" style=""font-weight:bold""><small><font color=""#ff0000"">*</font><i>All Fields Required</i></small></div>"
	response.write vbcrlf & " <div class=""g-recaptcha"" data-sitekey=""6LcVxxwUAAAAAEYHUr3XZt3fghgcbZOXS6PZflD-""></div>"
	response.write vbcrlf & "<p align=""left"" class=""smallnote"">"
	response.write vbcrlf & "<font style=""font-weight:bold; color:#ff0000"">"
	response.write vbcrlf & "Press PROCESS PAYMENT button only once and please wait for the authorization page to be displayed to prevent double "
	response.write vbcrlf & "billing.  Be patient, it may take up to 2 minutes to process your transaction."
	response.write vbcrlf & "</font>"
	response.write vbcrlf & "</p>"

	'BEGIN: Process Buttons
	response.write vbcrlf & "<table border=""0"" cellpadding=""2"" cellspacing=""0"">"
	response.write vbcrlf & "<tr>"
	response.write "<td align=""center"">"
	response.write "<input type=""button"" id=""complete_payment_btn"" name=""COMPLETE_PAYMENT"" value=""PROCESS PAYMENT"" style=""width:200px;"" class=""skipjackbtn"" onclick=""processPayment();"" />"
	response.write "</td>"
	response.write "<td>&nbsp;</td>"
	response.write "</tr>"
	response.write vbcrlf & "</table>"
	'END: Process Buttons

	response.write vbcrlf & "<p class=""smallnote"">"
	response.write vbcrlf & "NOTE: Your IP address [" & request.servervariables("REMOTE_ADDR") & "] has been logged with this transaction.<br /><br />"
	response.write vbcrlf & "</p>"

	response.write vbcrlf & "</div>"
	response.write vbcrlf & "</div>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void GetTempReservationInformation iReservationTempId, sTotal, sCitizenName, sAddress, sCity, sState, sZip, sEmail, sFirstName, sLastName
'--------------------------------------------------------------------------------------------------
Sub GetTempReservationInformation( ByVal iReservationTempId, ByRef iOrgId, ByRef sTotal, ByRef sRentalName, ByRef sCitizenName, ByRef sAddress, ByRef sCity, ByRef sState, ByRef sZip, ByRef sSelectedDate, ByRef sStartTime, ByRef sEndTime, ByRef sEmail, ByRef sFirstName, ByRef sLastName )
	Dim sSql, oRs, iStartHour, iStartMinute, sStartAmPm, iEndHour, iEndMinute, sEndAmPm

	sSql = "SELECT orgid, citizenuserid, rentalid, ISNULL(feetotal,0.0000) AS feetotal, selecteddate, ISNULL(starthour,1) AS starthour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(startminute,0),2) AS startminute, "
	sSql = sSql & " ISNULL(startampm,'PM') AS startampm, ISNULL(endhour,2) AS endhour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(endminute,0),2) AS endminute,  ISNULL(endampm,'PM') AS endampm "
	sSql = sSql & " FROM egov_rentalreservationstemppublic "
	sSql = sSql & " WHERE reservationtempid = " & iReservationTempId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iRentalId = CLng(oRs("rentalid"))
		sRentalName = GetRentalName( iRentalId )	' In rentalcommonfunctions.asp
		iOrgId = CLng(oRs("orgid"))
		sSelectedDate = oRs("selecteddate")
		iStartHour = oRs("starthour")
		iStartMinute = oRs("startminute")
		sStartAmPm = oRs("startampm")
		iEndHour = oRs("endhour")
		iEndMinute = oRs("endminute")
		sEndAmPm = oRs("endampm")
		sStartTime = iStartHour & ":" & iStartMinute & " " & sStartAmPm
		sEndTime = iEndHour & ":" & iEndMinute & " " & sEndAmPm
		sTotal = CDbl(oRs("feetotal"))
		GetCitizenInformation oRs("citizenuserid"), sCitizenName, sAddress, sCity, sState, sZip, sEmail, sFirstName, sLastName
	Else
		iRentalId = 0
		sTotal = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void GetCitizenInformation iUserId, sCitizenName, sAddress, sCity, sState, sZip, sEmail, sFirstName, sLastName
'--------------------------------------------------------------------------------------------------
Sub GetCitizenInformation( ByVal iUserId, ByRef sCitizenName, ByRef sAddress, ByRef sCity, ByRef sState, ByRef sZip, ByRef sEmail, ByRef sFirstName, ByRef sLastName )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, ISNULL(useraddress,'') AS useraddress, "
	sSql = sSql & " ISNULL(usercity,'') AS usercity, ISNULL(userstate,'') AS userstate, ISNULL(userzip,'') AS userzip, "
	sSql = sSql & " ISNULL(useremail,'') AS useremail "
	sSql = sSql & " FROM egov_users WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		sCitizenName = oRs("userfname")
		sFirstName = Trim(oRs("userfname"))
		If oRs("userlname") <> "" Then 
			sCitizenName = Trim(sCitizenName) & " " & Trim(oRs("userlname"))
			sLastName = Trim(oRs("userlname"))
		Else
			sLastName = ""
		End If 
		sCitizenName = Trim(sCitizenName)
		sAddress = oRs("useraddress")
		sCity = oRs("usercity")
		sState = oRs("userstate")
		sZip = oRs("userzip")
		sEmail = oRs("useremail")
	Else
		sCitizenName = ""
		sFirstName = ""
		sLastName = ""
		sAddress = ""
		sCity = ""
		sState = ""
		sZip = ""
		sEmail = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void displayMonthOptions
'--------------------------------------------------------------------------------------------------
Sub displayMonthOptions()
	Dim i, sTemp 

	For i = 1 To 12
		If i < 10 Then 
			sTemp = "0" & i
		Else 
			sTemp = i
		End If 

		response.write vbcrlf & "<option value=""" & sTemp & """>" & sTemp & "</option>"
	Next 

End Sub 


'--------------------------------------------------------------------------------------------------
' void displayYearOptions
'--------------------------------------------------------------------------------------------------
Sub displayYearOptions()
	Dim i, sTemp

	'Draw Year selection
	sTemp = Year(Now())

	For i = 1 To 10
		If sTemp = Year(Now())+ 1 Then 
			sSelected = " selected=""selected"""
		Else 
			sSelected = ""
		End If 

		response.write vbcrlf & "<option value=""" & Right(sTemp,2) & """" & sSelected & ">" & sTemp & "</option>"
		sTemp = sTemp + 1
	Next 

End Sub




%>
