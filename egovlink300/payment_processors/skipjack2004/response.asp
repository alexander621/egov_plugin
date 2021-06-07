<!DOCTYPE HTML PUBLIC "-//W3C//DTD XHTML 1.1 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->
<% 

'	05/24/2011	Steve Loar Removing the commas in dollar amounts to stop the related SQL Server bug this causes.
' 2/7/2012 Steve Loar - changed the from email at the bottom to be from noreply@eclink.com

Dim sError, iPaymentControlNumber

iPaymentControlNumber = 0


%>
<html>
<head>
	<title>E-Gov Services <%=sOrgName%> - Skipjack Payment Form</title>

	<link rel="stylesheet" type="text/css" href="../../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../../global.css" />
	<link rel="stylesheet" type="text/css" href="../../css/style_<%=iorgid%>.css" />

	<script language="Javascript" src="../../scripts/modules.js"></script>

	<script language=javascript>
	<!--

		function openWin2(url, name) 
		{
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
		}

	//-->
	</script>

	<style>
		<%If request.servervariables("HTTPS") = "on" Then%>
			body {behavior: url('https://secure.egovlink.com/<%=sorgVirtualSiteName%>/csshover.htc');}
		<%End If%>
	</style>

</head>

<!--#Include file="../../include_top.asp"-->

<!--BODY CONTENT-->
<tr><td valign="top">

<!--BEGIN: INTRO TEXT-->
<div class="title"><%=sWelcomeMessage%></div>
<div class="main"><font color="#1c4aab">Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%></font></div>
<!--END: INTRO TEXT-->


<!--BEGIN:  DISPLAY SKIPJACK RESPONSE-->

<%
  '//================================================================================
  '// Read in posted variables
  '//================================================================================
  sForm_transactionamount  = Request("szTransactionAmount")
  sForm_authcode           = Request("szAuthorizationResponseCode")
  sForm_declinemsg         = Request("szAuthorizationDeclinedMessage")
  sForm_avsresponsecode    = Request("szAVSResponseCode")
  sForm_avsresponsemsg     = Request("szAVSResponseMessage")
  sForm_ordernumber        = Request("szOrderNumber")
  sForm_returncode         = Request("szReturnCode")
  sForm_customername       = Request("name") & Request("sjname")
  sForm_streetaddress      = Request("Streetaddress")
  sForm_city               = Request("City")
  sForm_state              = Request("State")
  sForm_zipcode            = Request("Zipcode")
  sForm_transactionamount  = CDbl(Request("Transactionamount"))
  sForm_orderstring        = Request("Orderstring")
  sForm_shiptophone        = Request("Shiptophone")
  sForm_comments           = Request("Comment")
  sForm_isapproved         = Request("szIsApproved")
  sForm_transactionfilename = Request("szTransactionFileName")
  s_errormsg               = ""  '--- Displayed error message
  s_displaytype            = ""  '--- Defines which message type to display


%>
<!--
sForm_authcode           = &LT%=sForm_authcode%>
sForm_declinemsg         = &LT%=sForm_declinemsg%>
sForm_avsresponsecode    = &LT%=sForm_avsresponsecode%>
sForm_avsresponsemsg     = &LT%=sForm_avsresponsemsg%>
sForm_ordernumber        = &LT%=sForm_ordernumber%>
sForm_returncode         = &LT%=sForm_returncode%>
sForm_customername       = &LT%=sForm_customername%>
sForm_streetaddress      = &LT%=sForm_streetaddress%>
sForm_city               = &LT%=sForm_city%>
sForm_state              = &LT%=sForm_state%>
sForm_zipcode            = &LT%=sForm_zipcode%>
sForm_transactionamount  = &LT%=sForm_transactionamount%>
sForm_orderstring        = &LT%=sForm_orderstring%>
sForm_shiptophone        = &LT%=sForm_shiptophone%>
sForm_comments           = &LT%=sForm_comments%>
sForm_isapproved         = &LT%=sForm_isapproved%>
sForm_transactionfilename = &LT%=sForm_transactionfilename%>
-->
<%
If ( sForm_returncode = "1" ) Then  '--- All information been passed correctly = 1, else 0

	If ( sForm_isapproved = "0" ) Then  '--- Has the transaction been 'Authorized' = 1. else 0

		If ( instr(sForm_transactionamount,"-") > 0 ) then  '--- Check to see if there is a negative sign in the transaction amount filed, if so then the transaction is a blind credit
			s_displaytype = "BlindCredit" '--- Blind Credit
		Else 
			s_displaytype = "Declined"    '--- Transaction has failed authorization
		End If 

	Else    '--- else, for --- if ( sForm_isapproved = "0" )
		s_displaytype = "Successful"  '--- Transaction is successful
	End If  '--- end if, for --- if ( sForm_isapproved = "0" )
  
Else '--- else, for --- if ( sForm_returncode = "1" ) --- Error message is determined from error code

  s_displaytype = "Invalid"
  Select Case sForm_returncode
    Case "-35"   s_errormsg = "You have entered an invalid credit card number"
    Case "-37"   s_errormsg = "Failed to connect to dial-up service. The system is temporarily unavailable.<br><br>Please try again in a few minutes."
    Case "-39"   s_errormsg = "The HTML serial number is empty, the incorrect length, or invalid"
    Case "-51"   s_errormsg = "The zipcode is empty, the incorrect length, or invalid"
    Case "-52"   s_errormsg = "The ship to zipcode is empty, the incorrect length, or invalid"
    Case "-53"   s_errormsg = "The expiration is empty, the incorrect length, or invalid"
    Case "-54"   s_errormsg = "The account number date empty, the incorrect length, or invalid"
    Case "-55"   s_errormsg = "The streetaddress is empty, the incorrect length, or invalid"
    Case "-56"   s_errormsg = "The ship to streetaddress is empty, the incorrect length, or invalid"
    Case "-57"   s_errormsg = "The transaction amount is empty, the incorrect length, or invalid"
    Case "-58"   s_errormsg = "The name is empty, the incorrect length, or invalid"
    Case "-59"   s_errormsg = "The location is empty, the incorrect length, or invalid"
    Case "-60"   s_errormsg = "The state is empty, the incorrect length, or invalid"
    Case "-61"   s_errormsg = "The ship to state is empty, the incorrect length, or invalid"
    Case "-62"   s_errormsg = "The orderstring is empty, the incorrect length, or invalid"
    Case "-64"   s_errormsg = "The ship to phone number is invalid"
    Case "-65"   s_errormsg = "The name is empty!"
    Case "-66"   s_errormsg = "The email is empty!"
    Case "-67"   s_errormsg = "The street address is empty!"
    Case "-68"   s_errormsg = "The city field is empty!"
    Case "-69"   s_errormsg = "The state field is empty!"
    Case "-79"   s_errormsg = "The customer name is empty, the incorrect length, or invalid"
    Case "-80"   s_errormsg = "The ship to customer name is empty, the incorrect length, or invalid"
    Case "-81"   s_errormsg = "The customer location is empty, the incorrect length, or invalid"
    Case "-82"   s_errormsg = "The customer state is empty, the incorrect length, or invalid"
    Case "-83"   s_errormsg = "The ship to phone number is either empty, has an incorrect length, or invalid"
    Case "-84"   s_errormsg = "The order number for this transaction has already been submitted.<br><br>Make sure that you haven't submitted the transaction twice."
    Case "-91"   s_errormsg = "The CVV2 field value is invalid!"
    Case "-92"   s_errormsg = "The Approval Code is incorrect!"
    Case "-93"   s_errormsg = "Blind Credits are not allowed with this account!"
    Case "-94"   s_errormsg = "Blind Credit attempt has failed!"
    Case "-95"   s_errormsg = "Voice Authorization are not allowed with this account!"
  End Select

End If  '--- end if, for --- if ( sForm_returncode = "1" )


' Add some logging to see what happens for this

iPaymentControlNumber = CreatePaymentsControlRow( "Skipjack Payment Processing Started", "'Public'", "'Payments'" )

AddToPaymentsLog iPaymentControlNumber, "Transaction Result: " & s_displaytype, "'Public'", "'Payments'"
AddToPaymentsLog iPaymentControlNumber, "Error Message: " & s_errormsg, "'Public'", "'Payments'"
AddToPaymentsLog iPaymentControlNumber, "Citizen: " & sForm_customername, "'Public'", "'Payments'"
AddToPaymentsLog iPaymentControlNumber, "Amount: " & sForm_transactionamount, "'Public'", "'Payments'"

AddToPaymentsLog iPaymentControlNumber, "szReturnCode: " & sForm_returncode, "'Public'", "'Payments'"
AddToPaymentsLog iPaymentControlNumber, "szAuthorizationResponseCode: " & sForm_authcode, "'Public'", "'Payments'"
AddToPaymentsLog iPaymentControlNumber, "szAVSResponseCode: " & sForm_avsresponsecode, "'Public'", "'Payments'"
AddToPaymentsLog iPaymentControlNumber, "szOrderNumber: " & sForm_ordernumber, "'Public'", "'Payments'"
AddToPaymentsLog iPaymentControlNumber, "szTransactionFileName: " & sForm_transactionfilename, "'Public'", "'Payments'"

If sForm_transactionamount <> "" Then 
	'--- The below inserts the decimal point back into the TransactionAmount string after processing --- ie 5000 becomes 50.00
	if InStr(sForm_transactionamount,".") then
	  'sForm_transactionamount = Left(sForm_transactionamount,InStr(sForm_transactionamount,".")-1) & Mid(sForm_transactionamount,InStr(sForm_transactionamount,".")+1,2)
	  sForm_transactionamount = FormatCurrency(sForm_transactionamount,2)
	end if
	'sForm_transactionamount = FormatCurrency(Left(sForm_transactionamount,Len(sForm_transactionamount)-2) & "." & Right(sForm_transactionamount,2),2)
	sForm_transactionamount = FormatCurrency(sForm_transactionamount,2)
End If 

AddToPaymentsLog iPaymentControlNumber, "Formatted Amount: " & sForm_transactionamount, "'Public'", "'Payments'"

%>
 
<!-- s_displaytype = &LT%=s_displaytype%>  --  &LT%=sForm_transactionamount%>  --  &LT%=sForm_isapproved%> -->
<br><br><br>
<%
response.write "<div style=""margin-left:20px;"" class=""group"">"

'//================================================================================
'// Successful Transaction
'//================================================================================
If s_displaytype = "Successful" Then 
%>

	<center>
	  <font color="blue"><b>Transaction Successfully Authorized!</b></font><br />
	  <br />
		Please keep a copy of the following order information for your records:<br /><br />
	</center>  
	  
	<p align="center" style="background-color:#FFFFFF; border: solid 1px #000000;margin: 1em;">      
				Transaction Approval Code: <%=sForm_authcode%><br />
				  AVS Response Code: <%=sForm_avsresponsecode%><br />
			   AVS Response Message: <%=sForm_avsresponsemsg%><br />
				 Transaction Amount: <%=sForm_transactionamount%><br />
					Validation Code: <%=sForm_returncode%><br />
					Transaction File Name: <%=sForm_transactionfilename%><br /><br />
	 
					  Order Number: <%=sForm_ordernumber%><br />
					Cardmember Name: <%=sForm_customername%><br />
					 Street Address: <%=sForm_streetaddress%><br />
							   City: <%=sForm_city%><br />
							  State: <%=sForm_state%><br />
							Zipcode: <%=sForm_zipcode%><br />
					  Ship-to phone: <%=sForm_shiptophone%><br /><br />

						Orderstring: <%=sForm_orderstring%><br />
							Comment: <%=sForm_comments%><br />
	</p>
	<center>

		<a href="<%=sEgovWebsiteURL%>/" >Click here to return to the E-Government Website</a><br />

	  <!--<a href="javascript:history.go(-2);">Return to the Order Form</a><br>-->
	</center>

<%


	' ADD INFORMATION TO ADMINSTRATION DATABASE
	AddToPaymentsLog iPaymentControlNumber, "order number: " & sForm_ordernumber, "'Public'", "'Payments'"
	iPaymentID = RIGHT(sForm_ordernumber,Len(sForm_ordernumber)-instr(sForm_ordernumber,"O"))
	AddToPaymentsLog iPaymentControlNumber, "paymentid: " & iPaymentID, "'Public'", "'Payments'"

	fn_AddPaymentInformation iPaymentID
	AddToPaymentsLog iPaymentControlNumber, "Back from fn_AddPaymentInformation", "'Public'", "'Payments'"


End If


'//================================================================================
'// Declined Transaction
'//================================================================================
If s_displaytype = "Declined" Then 
%>
	<center>
	<font color="red"><b>Transaction Not Authorized!</b></font><br /><br />
	<div style="background-color:#FFFFFF; border: solid 1px #000000;padding:10px;">

	  I'm sorry but your transaction has been <b>Declined</b>.<br />
	  <br>
	  The reason that we received from your Credit-Card Issuer was: <br />
	  <br />
	  <i><b><%=sForm_declinemsg%></b></i><br />
	  </div>
	  <br />
	  <font class="paymentreturnlink"><a href="javascript:history.go(-2);">Click here to return to the Order Form</a></font><br />
	</center>

<%
End If


'//================================================================================
'// BlindCredit Transaction
'//================================================================================
If s_displaytype = "BlindCredit" Then 
%>

	<center>
	<font color="red"><b>Transaction Not Authorized!</b></font><br /><br />
	<div style="background-color:#FFFFFF; border: solid 1px #000000;padding:10px;">
	  You have submitted a blind Credit to your Account.<br />
	  <br />
	  Check your register to see the current status of the transaction.<br />
	  <br />
	  </div>
	  <font class="paymentreturnlink"><a href="javascript:history.go(-2);">Click here to return to the Order Form</a></font><br />
	</center>

<%
End If 


'//================================================================================
'// Invalid Entry
'//================================================================================
If s_displaytype = "Invalid" Then 
%>
	<center>
	<font color="red"><b>Transaction Not Authorized!</b></font><br /><br />
	<div style="background-color:#FFFFFF; border: solid 1px #000000;padding:10px;">
	  Error: <%=s_errormsg%><br>
	  <br>
	  The transaction has <b>NOT</b> been completed.<br />
	  <br />
	  </div>
	 Please <a href="javascript:history.go(-2);"><font class=paymentreturnlink> Click here to return to the Order Form</font> </a> and make the appropriate corrections.<br />
	  <br />
	  If the problem persists - contact the merchant.
	</center>

<%

End If 


'//================================================================================
'//================================================================================
%>

<!--BEGIN: PAYMENT FOOTER-->
<center>
	<p class="smallnote">
		NOTE: Your IP address [<%=request.servervariables("REMOTE_ADDR")%>] has been logged with this transaction.<br /><br />
		Do you have questions?<br />
		Contact Customer Service: <a href="mailto:support@egovlink.com">support@egovlink.com</a> or 513-853-8675.
	</p>
	<p>
		<img src="images/SkipJackTNh161x63.gif">
	</p>
</center>
<!--END: PAYMENT FOOTER-->

</div>


<!--END: DISPLAY SKIPJACK RESPONSE-->


<!--SPACING CODE-->
<p>&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->


<!--#Include file="../../include_bottom.asp"--> 



<%

AddToPaymentsLog iPaymentControlNumber, "Skipjack Payment Processing Finished.", "'Public'", "'Payments'"




'------------------------------------------------------------------------------------------------------------
' void fn_AddPaymentInformation iPaymentId
'------------------------------------------------------------------------------------------------------------
Sub fn_AddPaymentInformation( ByVal iPaymentId )
	Dim sSql, iPaymentUserId

	'If request("userid") = "" Then 
		iPaymentUserId = AddUserInformation()
		AddToPaymentsLog iPaymentControlNumber, "iPaymentUserId: " & iPaymentUserId, "'Public'", "'Payments'"
	'Else
	'	iUserId = CLng(request("userid"))
	'End If

	If PaymentRecordExixts( iPaymentId ) Then
		sSql = "UPDATE egov_payments SET orgid = " & iOrgId
		sSql = sSql & ", paymentamount = " & Replace(Replace(sForm_transactionamount, ",", ""), "$","")
		sSql = sSql & ", paymentstatus = 'COMPLETED' "
		sSql = sSql & ", userid = " & iPaymentUserId
		sSql = sSql & " WHERE paymentid = " & iPaymentId
		session("paymentupdateSQL") = sSql
		AddToPaymentsLog iPaymentControlNumber, "updating paymentid: " & iPaymentId, "'Public'", "'Payments'"

		RunSQLStatement sSql

		session("paymentupdateSQL") = ""

	Else
		' I think this would be a problem as the serviceid would be missing
		sSql = "INSERT INTO egov_payments ( orgid, paymentamount, paymentstatus, userid ) VALUES ( "
		sSql = sSql & iOrgId & ", " & Replace(Replace(sForm_transactionamount, ",", ""), "$","") & ", 'COMPLETED', " & iPaymentUserId & " )"

		iPaymentId = RunIdentityInsertStatement( sSql )
		AddToPaymentsLog iPaymentControlNumber, "new paymentid generated: " & iPaymentId, "'Public'", "'Payments'"
	End If 

	
	' ADD RAW TRANSACTION DATA
	AddPaymentInformation iPaymentId
	AddToPaymentsLog iPaymentControlNumber, "Back from AddPaymentInformation", "'Public'", "'Payments'"

	' Send Email
	'On Error Resume Next
	lcl_SendEmail iPaymentId


'	Set oPayment = Nothing

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
' void AddPaymentInformation iPaymentId
'------------------------------------------------------------------------------------------------------------
Sub AddPaymentInformation( ByVal iPaymentId ) 
	Dim sSql, sCompleteData

	sCompleteData = "Transaction Approval Code:  " &  sForm_authcode & "<br />"
	sCompleteData = sCompleteData & "AVS Response Code:  " & sForm_avsresponsecode & "<br />"
	sCompleteData = sCompleteData & "AVS Response Message:  " & sForm_avsresponsemsg & "<br /> "
	sCompleteData = sCompleteData & "Transaction Amount:  " & sForm_transactionamount & "<br />"
	sCompleteData = sCompleteData & "Validation Code  " & sForm_returncode & "<br />"
	sCompleteData = sCompleteData & "Transaction File Name:  " & sForm_transactionfilename & "<br />"
	sCompleteData = sCompleteData & "Order Number:  " &  sForm_ordernumber & "<br />"
	sCompleteData = sCompleteData & "Cardmember Name:  " & sForm_customername & "<br />"
	sCompleteData = sCompleteData & "Street Address:  " & sForm_streetaddress & "<br />"
	sCompleteData = sCompleteData & "City:  "  & sForm_city & "<br />"
	sCompleteData = sCompleteData & "State:  " & sForm_state & "<br />"
	sCompleteData = sCompleteData & "Zipcode:  " & sForm_zipcode & "<br />"
	sCompleteData = sCompleteData & "Ship-to phone:  "  & sForm_shiptophone & "<br />"
	sCompleteData = sCompleteData & "Orderstring:  "  & sForm_orderstring & "<br />"
	sCompleteData = sCompleteData & "Comment: "  & sForm_comments 

	AddToPaymentsLog iPaymentControlNumber, "sCompleteData: " & sCompleteData, "'Public'", "'Payments'"

	sSql = "INSERT INTO egov_paymentdetails ( paymentid, paymentsummary ) VALUES ( " & iPaymentId
	sSql = sSql & ", '" & DBsafe( sCompleteData ) & "' )"

	RunSQLStatement sSql 

End Sub 


'------------------------------------------------------------------------------------------------------------
' integer AddUserInformation()
'------------------------------------------------------------------------------------------------------------
Function AddUserInformation()
	Dim sSql, iNewUserId

	AddToPaymentsLog iPaymentControlNumber, "new User Info: " & sForm_customername & ", " & sForm_streetaddress, "'Public'", "'Payments'"

	sSql = "INSERT INTO egov_users ( userfname, useraddress, usercity, userstate, userzip ) VALUES ( '"
	sSql = sSql & dbsafe(sForm_customername) & "', '" & dbsafe(sForm_streetaddress) & "', '" 
	sSql = sSql & dbsafe(sForm_city) & "', '" & dbsafe(sForm_state) & "', '" & dbsafe(sForm_zipcode) & "' )"

	iNewUserId = RunIdentityInsertStatement( sSql )

	AddUserInformation = iNewUserId

End Function


'------------------------------------------------------------------------------------------------------------
' string DBsafe( strDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )

	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function

	DBsafe = Replace( strDB, "'", "''" )

End Function


'------------------------------------------------------------------------------------------------------------
' void lcl_SendEmail( iPaymentId )
'------------------------------------------------------------------------------------------------------------
Sub lcl_SendEmail( ByVal iPaymentId )
	Dim sSql, oPayment, adminEmailAddr, sPayInfo, sMsg2, objMail2, ErrorCode

	' CONNECT TO DATABASE AND GET PAYMENT INFORMATION
	sSql = "SELECT * FROM dbo.egov_payment_list where paymentid = " & iPaymentId

	Set oPayment = Server.CreateObject("ADODB.Recordset")
	oPayment.Open sSql, Application("DSN"), 3, 1
	
	if oPayment("assigned_email") = "" or isNull(oPayment("assigned_email")) then
		adminEmailAddr = "webmaster@eclink.com" ' NEED TO HAVE A DEFAULT INSTITUTION EMAIL ADDRESS
	else
		adminEmailAddr = oPayment("assigned_email") ' ASSIGNED ADMIN USER EMAIL
	end if

	' BUILD MESSAGE 
	sPayInfo = replace(oPayment("payment_information"),"<br>",vbcrlf)
	sPayInfo = replace(oPayment("payment_information"),"</br>",vbcrlf)
	
	sMsg2 = "This automated message was sent by the E-Gov web site. Do not reply to this message.  Contact " & adminEmailAddr & " for inquiries regarding this email." & vbcrlf 
	sMsg2 = sMsg2 & " " & vbcrlf 
	sMsg2 = sMsg2 & "Payment was submitted on " & Date() & "." & vbcrlf 
	sMsg2 = sMsg2 & " " & vbcrlf  
	sMsg2 = sMsg2 & "FORM INFORMATION" & vbcrlf & "---------------------------------------------------------------------------------------------------" & vbcrlf 
	sMsg2 = sMsg2 & " " & vbcrlf 
	sMsg2 = sMsg2 & sPayInfo & vbcrlf & vbcrlf & vbcrlf
	sMsg2 = sMsg2 & "PAYMENT INFORMATION" & vbcrlf & "---------------------------------------------------------------------------------------------------" & vbcrlf 
	sMsg2 = sMsg2 & " " & vbcrlf 
	sMsg2 = sMsg2 & replace(oPayment("paymentsummary"),"<br>",vbcrlf)


	sendEmail "", adminEmailAddr, "", sOrgName & " E-GOV ACTION ITEM SUBMISSION", "", sMsg2, "Y"

End Sub 



%>

