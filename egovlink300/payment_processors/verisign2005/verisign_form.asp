<!DOCTYPE HTML PUBLIC "-//W3C//DTD XHTML 1.1 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<% 'Response.Expires = -1000 %>
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->
<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: verisign_form.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2005 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module processes Payments.
'
' MODIFICATION HISTORY
' 1.0	??/??/????	??? ??? - Initial Version
' 1.1	05/28/2009	Steve Loar - Changes for centralized PayPal processing and PayFlow Pro changes
' 2.0	06/23/2010	Steve Loar - Split name field into first and last 
' 2.1	08/02/2010	Steve Loar - Modified for Point and Pay Payments
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sError, sTotalAmount, iItemCount, bHasPaymentFee, dFeeAmount, sErrorMsg, dTotalCharges
Dim sProcessingRoute, iPaymentServiceId, iQtyLimit, iMyCount, sApplicantName

If request("paymentid") = "" Then
	response.redirect "../../payment.asp"
End If 

iMyCount = 0

' Here we handle the Rye payment form that only allows a fixed number of purchases
If request( "paymentservice" ) = "rye commuter permits by qty" or request( "paymentservice" ) = "rye snow field parking by qty" or request("paymentservice") = "rye halloween window painting signup" Then 
	'response.redirect "../../payment.asp"

	If session("myCount") = "" Then 
		iPaymentServiceId = CLng(request("paymentid"))
		sApplicantName = request("custom_applicantfirstname") & " " & request("custom_applicantlastname")

		' pull the qty limit for the payment service
		iQtyLimit = GetRyeQtyLimit( iPaymentServiceId )
		iCurrLimit = GetRyeCounter( iPaymentServiceId )
		'response.write iQtyLimit & "<br />"

		response.write "<!--" & iCurrLimit & "-->"

		If clng(iCurrLimit) >= clng(iQtyLimit) Then 
			TurnOffRyeForm iPaymentServiceId
			'sendEmail "","tfoster@eclink.com", "", "Payment Form " & iPaymentServiceId & " was automatically turned off", "A Rye Form has reached the limit, but might not be full yet.  Check the number sold and adjust the limit accordingly.","", "Y"
			response.redirect "../../payment.asp?paymenttype=" & iPaymentServiceId 
		else

			' insert and pull a new qty value
			iMyCount = IncrementRyeCounter( sApplicantName, iPaymentServiceId, iQtyLimit )
			'response.write iMyCount
			'response.end
	
			' if new qty value > qty limit
			If clng(iMyCount) >= clng(iQtyLimit) Then 
				' turn off the form and set the message
				TurnOffRyeForm iPaymentServiceId
	
				If iMyCount > iQtyLimit Then 
					' take the user back to the payment form so they can see the unavailable message
					response.redirect "../../payment.asp?paymenttype=" & iPaymentServiceId 
				End If 
			End If 
	
			' set this session variable for when they get declined and need to correct their CC info
			session("myCount") = iMyCount
		end if
	End If 

	iMyCount = session("myCount")

End If 

dFeeAmount = CDbl(0.00)

' check if payment gateway needs a fee check for this page
If PaymentGatewayRequiresFeeCheck( iOrgId ) Then
	If CitizenPaysFee( iOrgId ) Then
		bHasPaymentFee = True 
		
		sTotalAmount = CDbl(request("custom_paymentamount"))
		
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
	bHasPaymentFee = False 
End If 


%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
<title>E-Gov Services <%=sOrgName%> - Payment Form</title>

	<META NAME="ROBOTS" CONTENT="NOINDEX, NOFOLLOW" />

	<script src='https://www.google.com/recaptcha/api.js' async defer></script>
	<link rel="stylesheet" type="text/css" href="../../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../../global.css" />
	<link rel="stylesheet" type="text/css" href="verisign.css" />
	<link rel="stylesheet" type="text/css" href="../../css/style_<%=iorgid%>.css" />

	<script language="Javascript" src="../../scripts/modules.js"></script>

  	<script type="text/javascript" src="../../scripts/jquery-1.9.1.min.js"></script>

	<script language=javascript>
	<!--
	
		function openWin2(url, name)
		{
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
		}
		
		//--- 07.30.01: Set the form focus
		//--- 05.16.00: Added the Unique order Number
		//--- Generate Unique Order Number
		//--- (use 000658076426 for testing) 
		function GenerateOrderNumber()
		{
			tmToday = new Date();
			return tmToday.getTime();
		}

		function StartOrder() 
		{
			document.Fullorder.sjname.focus();
		}

		var sURL = unescape(window.location.pathname);

		function refresh()
		{
			window.location.replace( sURL );
		}
		/*
		Submit Once form validation- 
		*/
	 
		function submitonce(theform)
		{
			//if IE 4+ or NS 6+
			if (document.all||document.getElementById)
			{
				//screen thru every element in the form, and hunt down "submit" and "reset"
				for (i=0;i<theform.length;i++)
				{
					var tempobj=theform.elements[i]
					if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
						//disable em
					tempobj.disabled=true
				}
			}
		} 


		function processPayment()
		{
			var response = grecaptcha.getResponse();

			if(response.length == 0)
			{
    				//reCaptcha not verified
				alert("Sorry, but the CAPTCHA field is required. Please check that box before submitting again.");
			}
			else
			{
    				//reCaptch verified
				$("#sjname").val($("#firstname").val() + ' ' + $("#lastname").val());
				// disable the pay button
				document.getElementById("COMPLETE_PAYMENT").disabled = true;
				// submit the form
				document.Fullorder.submit();
			}
		}

		// set focus on the first field when the page loads
		$('document').ready(function(){
			$("#firstname").focus();
		});

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
<p>
	<font class="pagetitle">Welcome to the <%=sOrgName%> Permits and Payments Center</font> <br />
	<font class="datetagline">Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%></font>
</p>
	<!--END: INTRO TEXT-->

<div id="content">
	<div id="centercontent">

	<!--BEGIN:  DISPLAY PAYMENT FORM-->

<% 
		fnDisplayForm CLng(request("paymentid")), bHasPaymentFee, dFeeAmount
%>

	<!--END: DISPLAY PAYMENT FORM-->
 
	</div>
</div>
   
<!--#Include file="../../include_bottom.asp"-->    

<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void FNDISPLAYFORM IPAYMENTFORMID, bHasPaymentFee, dFeeAmount
'--------------------------------------------------------------------------------------------------
Sub fnDisplayForm( ByVal iPaymentFormID, ByVal bHasPaymentFee, ByVal dFeeAmount )
	Dim sSql, oRs, sFirstName, sLastName, sPaymentFormID

	sPaymentFormID = CStr(iPaymentFormID)

	' GET FORM INFORMATION
	sSql = "SELECT paymentservicename FROM egov_paymentservices "
	sSql = sSql & " WHERE paymentserviceid = " & iPaymentFormID & " AND orgid = " & iOrgId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1
		
	If Not oRs.EOF Then
		
		' FORM HEADING	
		response.write "<blockquote class=""paymentform"">"
		response.write "<font class=""formtitle"">" & oRs("paymentservicename") & " - Payment Information</font>"
		response.write "<div class=""group"" id=""paymentform"">"
		%>

		<!--BEGIN: BUILD PAYMENT FORM-->
		<form name="Fullorder" action="process_payment.asp" method="post">

		<!--BEGIN: GET PAYMENT INFORMATION-->
		<%
		Dim sOrderString,sSerialNumber,sOrderNumber
		Dim sItemNumber,sItemDescription,sItemQuantity,blnTaxTable,curItemCost

		sItemNumber = REQUEST("ITEM_NUMBER")
		sItemDescription = REQUEST("ITEM_NAME")
		curItemCost = REQUEST("custom_paymentamount")
		' CHECK FOR MEMORIAL PAYMENT
		If request("amount") <> "" Then
			curItemCost = request("amount")
		End If
		sItemQuantity = 1
		blnTaxTable = N ' Y USE TABLE | N DON'T USE TABLE
		sOrderString = sItemNumber & "~" & sItemDescription & "~" & curItemCost & "~" & sItemQuantity & "~" & blnTaxTable & "||"  
		'sSerialNumber =  GetSerialNumber(iOrgID) '"000389764111" ' <<< DEMO VALUE --- LIVE VALUE PULLED FROM DATABASE >>GetSerialNumber(iOrgID) 

		sOrderNumber = "ecC" & iOrgID & "O" & CreatePayment() ' STORE VALUES IN DATABASE

		' IF DEMO ORGANIZATIO SHOW TEST VALUES
		If iorgid = 5 Or iorgid = 145 Then 
  			'sName          = "Peter Selden"
			'sFirstName     = "Peter"
			'sLastName      = "Selden"
	  		'sEmail         = "pselden@eclink.com"
  			'sAddress       = "4303 Hamilton Avenue"
		  	'sCity          = "Cincinnati"
  			'sDefaultState  = "OH"
		  	'sZip           = "45223"
  			'sOrderNum      = "43036814030"
		  	sCreditCardNum = "5555555555554444"	' PayPal MasterCard
'			sCreditCardNum = "5462387494460372"		' PNP MasterCard
  			sExpMonth      = "03"
		  	sExpYear       = "2013"
  			sCVSCode       = ""
		  	'sPhone         = "5136814030"
		End If

		If request("userid") <> "" Then
			' Get their info for the form
			on error resume next
			GetUserInfo CLng(request("userid")), sName, sEmail, sAddress, sCity, sState, sZip, sFirstName, sLastName

			if err.number <> 0 then
				response.redirect "/payments.asp?paymenttype=" & request.form("paymentid")
			end if
			on error goto 0
		End If 

		' GET CUSTOM FIELDS 
		If CLng(iOrgid) = CLng(153) Then 
			' This attempts to get the fields in the order they are in on the posting form
			Dim sFieldName, sFieldValue, iFieldItem
			For iFieldItem = 1 To Request.Form.Count
				sFieldName = Request.Form.Key(iFieldItem)
				sFieldValue = Request.Form.Item(iFieldItem) & ""
				If Left(sFieldName,7) = "custom_" Then
					If InStr(sFieldName,"paymentamount") = False Then 
						sDetails = sDetails & Replace(sFieldName,"custom_","") & " : " & sFieldValue & " </br>"
						Select Case sPaymentFormID
							Case "382"
								If sFieldName = "custom_vehiclelicense" Or sFieldName = "custom_permitholdertype" Then 
									sComment2 = sComment2 & Replace(sFieldName,"custom_","") & ":" & sFieldValue & " </br>"
								End If 
							Case "383"
								If sFieldName = "custom_applicantfirstname" Or sFieldName = "custom_applicantlastname" Or sFieldName = "custom_applicantphone" Then 
									sComment2 = sComment2 & Replace(sFieldName,"custom_applicant","") & ":" & sFieldValue & " </br>"
								End If 
							Case Else
								sComment2 = sComment2 & Replace(sFieldName,"custom_","") & ":" & sFieldValue & " </br>"
						End Select
					End If 
				End If
			Next 
		Else
			' This approach get the fields as they are in the request collection
			For Each oField IN Request.Form
				If Left(oField,7) = "custom_" Then
					If InStr(oField,"paymentamount") = False Then 
						sDetails = sDetails & replace(oField,"custom_","") & " : " & request(oField) & " </br>"
						Select Case sPaymentFormID
							Case "270"
								' in the comment2 field there is only room for 128 characters, so limit what is input 
								If oField = "custom_controlno" Or oField = "custom_firstname" Or oField = "custom_lastname" Or oField = "custom_phone" Then 
									sComment2 = sComment2 & replace(oField,"custom_","") & ":" & request(oField) & " </br>"
								End If 
							Case "382"
								If oField = "custom_vehiclelicense" Or oField = "custom_permitholdertype" Then 
									sComment2 = sComment2 & replace(oField,"custom_","") & ":" & request(oField) & " </br>"
								End If 
							Case "383"
								If oField = "custom_applicantfirstname" Or oField = "custom_applicantlastname" Or oField = "custom_applicantphone" Then 
									sComment2 = sComment2 & replace(oField,"custom_applicant","") & ":" & request(oField) & " </br>"
								End If 
							Case Else
								sComment2 = sComment2 & replace(oField,"custom_","") & ":" & request(oField) & " </br>"
						End Select
					End If 
				End If
			Next 
		End If 

		%>

		<input type="hidden" name="orderstring" value="<%=sOrderString%>" />
		<input type="hidden" name="serialnumber" value="<%=sSerialNumber%>" />
		<input type="hidden" name="ordernumber" value="<%=sOrderNumber%>" />
		<input type="hidden" name="paymentformid" value="<%=sPaymentFormID%>" />
		<input type="hidden" name="paymentname" value="<%=sItemDescription%>" />
		
		
<%		If request("userid") <> "" Then 		%>
			<input type="hidden" name="userid" value="<%=userid%>" />
<%		End If	

		If request("paymentservice") <> "" Then 
			response.write "<input type=""hidden"" name=""paymentservice"" value=""" & request("paymentservice") & """ />"
		End If 

		If request("renewalid") <> "" Then 
			response.write "<input type=""hidden"" name=""renewalid"" value=""" & request("renewalid") & """ />"
		End If 

		If request("custom_permitholdertype") <> "" Then 
			response.write "<input type=""hidden"" name=""permitholdertype"" value=""" & request("custom_permitholdertype") & """ />"
		End If 

'		If request("paymentservice") = "rye commuter permit renewal" Then 
'			If request("custom_permitholdertype") <> "" Then 
'				response.write "<input type=""hidden"" name=""permitholdertype"" value=""" & request("custom_permitholdertype") & """ />"
'				'sDetails = "permitholdertype : " & request("custom_permitholdertype") & " </br>" & sDetails
'				'sComment2 = "permitholdertype:" & request("custom_permitholdertype") & " </br>" & sComment2
'			End If
'		Else
'			If request("custom_permitholdertype") <> "" Then 
'				response.write "<input type=""hidden"" name=""permitholdertype"" value=""" & request("custom_permitholdertype") & """ />"
'			End If 
'		End If 
%>
		<input type="hidden" name="details" value="<%=sDetails%>" />
		<input type="hidden" name="comment2" value="<%=sComment2%>" />


		<!--END: GET PAYMENT INFORMATION-->

		<div align="center" style="max-width:500px;"> 

		<!--BEGIN: PAYMENT DETAILS-->
		<fieldset>
		<legend><strong>Payment Details</strong></legend>
		<table border="0" cellpadding="2" cellspacing="0" width="100%">
		<!--BEGIN: PERSONAL INFORMATION-->
			<tr>
			
<%	
'				response.write "<td>"
'				sPaymentImg = GetPaymentImage( "../../" )
'				If sPaymentImg <> "" Then 
'					response.write "<img src=""" & sPaymentImg & """ border=""0"" />"
'				Else 
'
'					response.write "<img src=""images/verisign.gif"" border=""0"" />"
'				End If 
'				response.write "</td>"
%>
				<td align="left"><b><%="(" & sItemNumber & ") - " & sItemDescription %></b><br><br>
					

		<%	
'					response.write "<b>Details</b><br />"
					response.write "<i>" & Replace(sDetails,"</br>","<br />") & "</i>"
		%>
				</td>
			</tr>
		</table>
		<br>
		</fieldset>
		<!--END: PAYMENT DETAILS-->

<%
	' Charge Amount
	response.write vbcrlf & "<fieldset>"
	response.write vbcrlf & "<legend><strong>Charges&nbsp;</strong></legend>"
	response.write vbcrlf & "<table border=""0"" cellpadding=""2"" cellspacing=""0"" id=""chargeamounts"" align=""left"">"

	response.write vbcrlf & "<tr align=""left"">"
	response.write "<td align=""right"" class=""labelcol"" nowrap=""nowrap"">Amount:</td>"
	response.write "<td align=""right"">" & FormatNumber(curItemCost,2) & "<input type=""hidden"" name=""transactionamount"" value=""" & FormatNumber(curItemCost,2,,,0) & """ /></td>"
	response.write "</tr>"

	If bHasPaymentFee Then
		' Show the fee
		response.write vbcrlf & "<tr align=""left"">"
		response.write "<td align=""right"" class=""labelcol"" nowrap=""nowrap"">Processing Fee:</td>"
		response.write "<td align=""right"">" & FormatNumber(dFeeAmount,2) & "</td>"
		response.write "</tr>"

		dTotalCharges = CDbl(dFeeAmount) + CDbl(curItemCost)

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
%>

		<p align="left">Please enter your billing information as it appears on your credit card statement, then click the <b>Process Payment</b> button.</p>

		<fieldset>
		<legend><strong>Personal Information</strong></legend>
		<table border="0" cellpadding="2" cellspacing="0" class="respTable">
		<!--BEGIN: PERSONAL INFORMATION-->
<%
'			<tr>
'			  <td align="right"><b>Name:</b></td>
'			  <td><font face="Verdana" size="2"><input type="text" maxLength="30" size="30" id="sjname" name="sjname" value="sName" tabindex="1" /></font></td>
'			</tr>
%>
			<tr>
			  <td align="right"><b>First Name:</b></td>
			  <td align="left"><font face="Verdana" size="2"><input type="text" maxLength="30" size="30" id="firstname" name="firstname" value="<%=sFirstName%>" /></font>
					<input type="hidden" id="sjname" name="sjname" value="<%=sName%>" />
			  </td>
			</tr>
			<tr>
			  <td align="right"><b>Last Name:</b></td>
			  <td align="left"><font face="Verdana" size="2"><input type="text" maxLength="30" size="30" id="lastname" name="lastname" value="<%=sLastName%>" /></font></td>
			</tr>
			<tr>
			  <td align="right"><b>E-mail:</b></td>
			  <td align="left"><font face="Verdana" size="2"><input type="text" value="<%=sEmail%>" maxLength="50" size="50" name="email" /></font></td>
			 </tr>
			<tr>
			  <td align="right"><b>Address:</b></td>
			  <td align="left"><font face="Verdana" size="2"><input type="text" value="<%=sAddress%>" maxLength="50" size="50" name="streetaddress" /></font></td>
				 </tr>
			<tr>
			  <td align="right"><b>City:</b></td>
			  <td align="left"><font face="Verdana" size="2"><input type="text" value="<%=sCity%>" maxLength="20" size="20" name="city" /></font></td>
			
			</tr>
			<tr>
			  <td align="right"><b>State:</b></td>
			  <td align="left"><font face="Verdana" size="2">
					<select name="state" size="1">
<%						ShowStatePicks sDefaultState, ""	%>
					</select></font></td>
			 </tr>
			<tr>
			  <td align="right"><b>Zip:</b></td>
			  <td align="left"><font face="Verdana" size="2"><input type="text" value="<%=sZip%>"  maxLength="15" size="15" name="zipcode" /></font></td>
			  
			</tr>
			</table>
			<br />
			</fieldset>
			<!--END: PERSONAL INFORMATION-->
			
			<br />
			
			<!--BEGIN: CREDIT CARD INFORMATION-->
			<fieldset>
			<legend><strong>Credit Card Information</strong></legend>
			<table border="0" cellpadding="2" cellspacing="0" class="respTable">
			<tr>
				<td align="right"><strong>Credit Card Type:</strong></td>
				<td align="left">	
<%					' set up the cardtype field	
					ShowCreditCardPicks		' In include_top_functions.asp   %>
				</td>
			</tr>
			<tr>
			  <td align="right"><strong>Credit Card Number:</strong></td>
			  <td align="left"><font face="Verdana" size="2"><input type="text" value="<%=sCreditCardNum%>" maxLength="22" size="30" name="accountnumber" /></font></td>
			  
			</tr>
			<tr>
			  <td align="right"><strong>Expiration Month:</strong></td>
			  <td align="left">

				<select name="month">
				<%
				' DRAW MONTH SELECTION
				For i=1 to 12
					If i < 10 Then
						sTemp = "0" & i
					Else
						sTemp = i
					End If
					response.write "<option value=""" & sTemp & """>" & sTemp & "</option>"
				Next
				%>
			  </select>
			  
			  </td>
			  <td colspan="2"></td>
			  <td></td>
			  <td> </td>
			</tr>
			<tr>
			  <td align="right"><strong>Expiration Year: </strong></td>
			  <td align="left">
			  
				<select name="year">
				<%
				' DRAW YEAR SELECTION 
				sTemp = Year(Now())
				For i = 1 to 10

					' Preselect next year for when the page loads
					If sTemp = Year(Now())+ 3 Then
						sSelected = " selected=""selected"" "
					Else
						sSelected = ""
					End IF

					response.write vbcrlf & "<option " & sSelected & " value=""" & right(sTemp,2) & """>" & sTemp & "</option>"
					sTemp = sTemp + 1
				Next
				%>
			  </select>
			  
			  </td>
			  <td></td>
			  <td></td>
			</tr>
		<%		If OrgHasFeature( iOrgId, "display cvv" ) Then %>
				<tr>
					<td align="right"><strong>CVV Code:</strong></td>
					<td align="left"><input type="text" name="cvv2" size="4" maxlength="4" value="" /></td>
					<td colspan="2"></td>
					<td></td>
					<td></td>
				</tr>
		<%		End If %>

		  </table>
<%			response.write "<p align=""center"">Do not use dashes or spaces when entering credit card information.</p>"	%>
		 </fieldset>
		 <!--END: CREDIT CARD INFORMATION-->


		<br />
		<fieldset>
		<legend><strong>CAPTCHA</strong></legend>
		<div class="g-recaptcha" data-sitekey="6LcVxxwUAAAAAEYHUr3XZt3fghgcbZOXS6PZflD-"></div>
		 </fieldset>

		<div align="left"><small><strong><font color="red">*</font><i>All Fields Are Required</i></strong></small></div>

		<p align="left" class="smallnote"><font style="font-weight:bold;color:red">Press PROCESS PAYMENT button only once and please wait for the authorization page to be displayed to prevent double billing.  Be patient, it may take up to 2 minutes to process your transaction.</font></p> 

		<!--BEGIN: PROCESS BUTTONS-->
		 <table border="0" cellpadding="2" cellspacing="0"  >
			<tr>
			  <td ALIGN="center">
				<input style="width:200px;" class="skipjackbtn" type="button" id="COMPLETE_PAYMENT" name="COMPLETE_PAYMENT" value="PROCESS PAYMENT" onclick="processPayment();" />
				<!--<input STYLE="WIDTH:100PX;" class=skipjackbtn type="button" value="Cancel" name="buttonRefresh" onclick="process_cancel();" >-->
			</td>
			<td>&nbsp;</td>
			</tr>
		</table>
		<!--END: PROCESS BUTTONS-->

		<p class="smallnote">NOTE: Your IP address [<%=request.servervariables("REMOTE_ADDR")%>] has been logged with this transaction.<br><br>
		
<%		If CLng(iOrgid) <> CLng(153) Then		%>
			Do you have questions?<br />
			Contact Customer Service: <a href="mailto:<%=sDefaultEmail%>"><%=sDefaultEmail%></a> or <%=FormatPhoneNumber(sDefaultPhone)%>.
<%		Else
%>
			If you have a question regarding the payment process Contact 914-967-7371 during regular business hours.
<%
			'ShowContactLine iPaymentFormID
		End If		%>
		</p>

		</div>
		<!-- END REQUIRED -->

		</form>

		<%
		' END FORM
		response.write "</blockquote></div>"
	Else
		' FORM NOT FOUND REDIRECT TO COMPLETE LIST
		response.redirect("../../payment.asp")
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


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


'--------------------------------------------------------------------------------------------------
' integer CreatePayment()
'--------------------------------------------------------------------------------------------------
Function CreatePayment()
	Dim sSql, iReturnValue, iPaymentInfoId

	iReturnValue = 0

	iPaymentInfoId = AddPaymentDetail()

	sSql = "INSERT INTO egov_payments ( paymentdate, paymentstatus, paymentserviceid, paymentinfoid ) VALUES ( "
	sSql = sSql & "getdate(), 'PROCESSING', " & CLng(request("ITEM_NUMBER"))/100 & ", " & iPaymentInfoId & " )"

	iReturnValue = RunIdentityInsertStatement( sSql )

	CreatePayment = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' integer AddPaymentDetail()
'--------------------------------------------------------------------------------------------------
Function AddPaymentDetail()
	Dim sSql, iReturnValue, sDetails, sFieldName, sFieldValue, iFieldItem

	iReturnValue = 0
	
	sDetails = ""
	' GET CUSTOM FIELDS 
	If CLng(iOrgid) = CLng(153) Then 
		For iFieldItem = 1 To Request.Form.Count
			sFieldName = Request.Form.Key(iFieldItem)
			sFieldValue = Request.Form.Item(iFieldItem) & ""
			If Left(sFieldName,7) = "custom_" Then
				If sFieldName <> "custom_permitholdertypes" Then 
					sDetails = sDetails & Replace(sFieldName,"custom_","") & " : " & sFieldValue & "</br>"
				End If 
			End If
		Next 
	Else
		For Each oField IN Request.Form
			If Left(oField,7) = "custom_" Then
				If oField <> "custom_permitholdertypes" Then 
					sDetails = sDetails & replace(oField,"custom_","") & " : " & request(oField) & "</br>"
				End If 
			End If
		Next 
	End If 

	sSql = "INSERT INTO egov_paymentinformation ( payment_information ) VALUES ( '" & dbsafe(sDetails) & "' )"

	iReturnValue = RunIdentityInsertStatement( sSql )

	AddPaymentDetail = iReturnValue

End Function


'------------------------------------------------------------------------------------------------------------
' string DBsafe( strDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )

	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function

	DBsafe = Replace( strDB, "'", "''" )

End Function


'------------------------------------------------------------------------------------------------------------
' string GetSerialNumber( IID ) -- This is defunct now
'------------------------------------------------------------------------------------------------------------
Function GetSerialNumber( ByVal iID )
	Dim iReturnValue, oRs, sSql

	iReturnValue = "00000000"
	
	sSql = "SELECT * FROM egov_skipjackoptions where orgid=" & iID 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iReturnValue = oRs("serialnumber")
	End If

	oRs.Close
	Set oRs = Nothing 

	GetSerialNumber = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' void GetUserInfo iUserId, sName, sEmail, sAddress, sCity, sState, sZip, sFirstName, sLastName
'--------------------------------------------------------------------------------------------------
Sub GetUserInfo( ByVal iUserId, ByRef sName, ByRef sEmail, ByRef sAddress, ByRef sCity, ByRef sState, ByRef sZip, ByRef sFirstName, ByRef sLastName )
	Dim sSql, oRs

	sSql = "SELECT userfname, userlname, useraddress, usercity, userstate, userzip, useremail "
	sSql = sSql & " FROM egov_users "
	sSql = sSql & " WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	sName    = Proper(oRs("userfname")) & " " & Proper(oRs("userlname"))
	sFirstName = oRs("userfname")
	sLastName = oRs("userlname")
	sEmail   = oRs("useremail")
	sAddress = oRs("useraddress")
	sCity    = oRs("usercity")
	sState   = oRs("userstate")
	sZip     = oRs("userzip")

	oRs.Close
	Set oRs = Nothing
	
End Sub  


'--------------------------------------------------------------------------------------------------
' string Proper( sString )
'--------------------------------------------------------------------------------------------------
Function Proper( ByVal sString )

	Proper = sString

	If Len(sString) > 0 then
		Proper = UCase(Left(sString,1)) & Mid(sString,2)
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' int GetRyeQtyLimit( iPaymentServiceId )
'--------------------------------------------------------------------------------------------------
Function GetRyeQtyLimit( ByVal iPaymentServiceId )
	Dim sSql, oRs, iRyeQtyLimit

	iRyeQtyLimit = CLng(0)

	sSql = "SELECT ISNULL(qtylimit,0) AS qtylimit FROM egov_paymentservices WHERE paymentserviceid = " & iPaymentServiceId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		iRyeQtyLimit = CLng(oRs("qtylimit"))
	End If 

	oRs.Close
	Set oRs = Nothing

	GetRyeQtyLimit = iRyeQtyLimit

End Function 


'--------------------------------------------------------------------------------------------------
' int IncrementRyeCounter( sApplicantName, iPaymentServiceId, iQtyLimit )
'--------------------------------------------------------------------------------------------------
Function IncrementRyeCounter( ByVal sApplicantName, ByVal iPaymentServiceId, ByVal iQtyLimit )
	Dim sSql, oRs, iCounter

	sSql = "INSERT INTO egov_ryepermitqtycounter ( ApplicantName, PaymentServiceId, QuantityLimit ) VALUES ( "
	sSql = sSql & "'" & dbsafe(sApplicantName) & "'," & iPaymentServiceId & ", " & iQtyLimit & " )"

	iCounter = RunIdentityInsertStatement( sSql )
	'iCounter = CLng(iCounter)

	iCounter = GetRyeCounter(iPaymentServiceId)

	IncrementRyeCounter = iCounter

End Function 

Function GetRyeCounter(ByVal iPaymentServiceId)
	sSQL = "SELECT COUNT(CurrentCount) Num FROM egov_ryepermitqtycounter WHERE PaymentServiceId = '" & iPaymentServiceId & "'"
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQl, Application("DSN"), 3, 1
	iCounter = 0
	if not oRs.EOF then iCounter = oRs("Num")
	oRs.Close
	Set oRs = Nothing

	GetRyeCounter = iCounter
End Function


'--------------------------------------------------------------------------------------------------
' TurnOffRyeForm iPaymentServiceId
'--------------------------------------------------------------------------------------------------
Sub TurnOffRyeForm( ByVal iPaymentServiceId )
	Dim sSql

	sSql = "SELECT paymentserviceenabled FROM egov_paymentservices WHERE paymentserviceid = '" & iPaymentServiceId & "' and paymentserviceenabled = 1"
	Set oCheck = Server.CreateObject("ADODB.RecordSet")
	oCheck.Open sSql, Application("DSN"), 3, 1

	if not oCheck.EOF then
		'sSql = "UPDATE egov_paymentservices SET paymentserviceenabled = 0, disabledmessage = '<span class=""importantheader"">Open Signup has ended. All of the available permits have been sold. Check back in a few minutes as some permits may become available again when purchases are not completed.</span>' "
		if iPaymentServiceId = "424" then
			'sSql = "UPDATE egov_paymentservices SET paymentserviceenabled = 0, disabledmessage = '<span class=""importantheader""><p><b>OH NO!</b></P><p>We have reached our capacity.  To be placed on the Halloween Window Painting waiting list, please stop by Rye Recreation at 281 Midland Ave to complete your application. You will need to have your partner’s information on hand to complete the form.  We will also take payment by <u>CHECK ONLY</u> (only to be processed if you receive a window assignment)</p><p><b><u>Please note:</u></b> The waiting list is processed in the order received and it is based on the availability of windows.  </p><p>If you have any question, you may email <a href=""mailto:Halloween@ryeny.gov"">Halloween@ryeny.gov</a> or call 914-967-2535.</p><p>Rye Recreation is working hard to secure more window space so that more children can participate.  If you register for the waitlist, you''ll be notified by Wednesday, October 15 if you have received a window assignment.</p><p>Thank you for your understanding.</p><p>Rye Recreation</p></span>' "
			'sSql = "UPDATE egov_paymentservices SET paymentserviceenabled = 0, disabledmessage = '<span class=""importantheader""><p><b>OH NO!</b></P><p>Open Signup is temporarily closed.  Please check back later to see if any openings become available.  When capacity is confirmed, this page will provide information on how to join the waitlist.</p></span>' "
			sSql = "UPDATE  egov_paymentservices SET paymentserviceenabled = 0, disabledmessage = '<span class=""importantheader""> <p><b>OH NO!</b></P> <p>We have reached our capacity.  To be placed on the Halloween Window Painting waiting list, please complete the waiting list <a href=""https://www.egovlink.com/public_documents300/rye/published_documents/Parks and Recreation/Halloween Window Painting/HWP Waiting List Application 18.pdf"">form</a> and drop it off at Rye Recreation at 281 Midland Avenue with your check for $20 made payable to the ""City of Rye"".  We also have a We also have a waitlist form for single participants.  If you cannot find a partner for your child please fill out this <a href=""https://www.egovlink.com/public_documents300/rye/published_documents/Parks and Recreation/Halloween Window Painting/HWP Waiting List Application Single Participant 18.pdf"">form</a> and drop it off at Rye Recreation at 281 Midland Avenue with your check for $10 made payable to the ""City of Rye"".  We take payment by <u>CHECK ONLY</u> (only to be processed if you receive a window assignment)</p> <p><b><u>Please note:</u></b> The waiting list is processed in the order received and it is based on the availability of windows.  </p> <p>If you have any question, you may email <a href=""mailto:Halloween@ryeny.gov"">Halloween@ryeny.gov</a> or call 914-967-2535.</p> <p>Rye Recreation is working hard to secure more window space so that more children can participate.  If you register for the waitlist, you''ll be notified by Wednesday, October 17 if you have received a window assignment. </p> <p>Thank you for your understanding.</p> <p>Rye Recreation</p> </span>' " 
			sSql = sSql & "WHERE paymentserviceid = " & iPaymentServiceId
			RunSQLStatement sSql
		else
			'sSql = "UPDATE egov_paymentservices SET paymentserviceenabled = 0, disabledmessage = '<span class=""importantheader"">Open Signup has closed. All of the available permits appear to have been sold. Please check back later to see if any become available.</span>' "
			'sSql = "UPDATE egov_paymentservices SET paymentserviceenabled = 0, disabledmessage = '<span class=""importantheader""><p>Open Signup is temporarily closed.  Please check back later to see if any openings become available.  When capacity is confirmed, this page will state signup has ended.</p></span>' "
			sSql = "UPDATE egov_paymentservices SET paymentserviceenabled = 0, disabledmessage = '<span class=""imprtantheader"">All permits have been claimed and are in the process of being approved.  You may try again shortly in the event that a permit becomes available.</span>' "
			sSql = sSql & "WHERE paymentserviceid = " & iPaymentServiceId
			RunSQLStatement sSql
		end if
	end if
	oCheck.Close
	Set oCheck = Nothing


	'Send Email
	'sendEmail "","tfoster@eclink.com", "", "Payment Form " & iPaymentServiceId & " was automatically turned off", "A Rye Form has reached the limit, but might not be full yet.  Check the number sold and adjust the limit accordingly.","", "Y"

End Sub 



%>


