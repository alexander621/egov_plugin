<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->

<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: payment.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Payments.
'
' MODIFICATION HISTORY
'	?.?		12/15/2008	Steve Loar	- Making it so that the post to the verisign form goes to the secure site.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim oPayOrg, iSectionID, sDocumentTitle
Set oPayOrg = New classOrganization

' This check added so direct access to the payments page is not possible if the feature is turned off - Steve Loar - 12/27/2005
If (Not OrgHasFeature( iOrgId, "payments" )) Or (Not blnOrgPayment) Then
	response.redirect sEgovWebsiteURL & "/"
End If

Dim sError 

' catch sql intrusions here
If request("paymenttype") <> "" Then 
	If Not IsNumeric(request("paymenttype")) Then
		response.redirect "payment.asp"
	End If 
End If 
%>
<html>
<head>
	<title>E-Gov Services - <%=oPayOrg.GetOrgName()%></title>

	<link rel="stylesheet" type="text/css" href="css/styles.css" />
	<link rel="stylesheet" type="text/css" href="global.css" />
	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />
	<link rel="stylesheet" type="text/css" href="./payment_styles.css" />
	
	<script language="javascript" src="scripts/modules.js"></script>
	<script language="javascript" src="scripts/easyform.js"></script>  
	<script language="javascript" src="scripts/formatnumber.js"></script>
	<script language="javascript" src="scripts/removespaces.js"></script>
	<script language="javascript" src="scripts/removecommas.js"></script>
	<script language="javascript" src="scripts/setfocus.js"></script>
	<script language="javascript" src="scripts/formvalidation_msgdisplay.js"></script>
	<script language="javaScript" src="prototype/prototype-1.7.0.0.js"></script>
	<script language="javascript" src="scripts/selectedradiovalue.js"></script>
	<script language="JavaScript" src="scripts/ajaxLib.js"></script>

	<script language="JavaScript" src="scripts/jquery-1.7.2.min.js"></script>

	<script language="javascript">
	<!--

		function openWin2(url, name) 
		{
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
		}

		function validateIncomeTaxPayment()
		{
			if ($("skip_feesok").value == "true")
			{
				var sTotal = $("custom_paymentamount").value; 
				if (Number(sTotal) > Number(0.00))
				{
					if (validateForm('frmpayment')) 
					{ 
						document.frmpayment.submit(); 
					}
				}
				else
				{
					inlineMsg($("custom_taxamount").id,'<strong>No Payment Amount Entered: </strong>Please enter the amount you are paying.',10,$("custom_taxamount").id);
				}
			}
			else
			{
				$("skip_feesok").value = "true";
			}
		}

		function ValidatePaymentAmount()
		{
			$("custom_paymentamount").value = removeSpaces($("custom_paymentamount").value);
			$("custom_paymentamount").value = removeCommas($("custom_paymentamount").value);
			if ($("custom_paymentamount").value != "")
			{
				var rege = /^\d*\.\d{2}$/
				var Ok = rege.exec($("custom_paymentamount").value);
				if ( Ok )
				{
					if (Number($("custom_paymentamount").value) < Number("0.01"))
					{
						inlineMsg($("custom_paymentamount").id,'<strong>Invalid Value: </strong>The Payment Amount must be greater than 0.00.',10,$("custom_paymentamount").id);
						$("custom_paymentamount").focus();
						return false;
					}
					//$("custom_paymentamount").value = format_number(Number($("custom_paymentamount").value),2);
					if (validateForm('frmpayment')) 
					{ 
						document.frmpayment.submit(); 
					}
				}
				else 
				{
					inlineMsg($("custom_paymentamount").id,'<strong>Invalid Value: </strong>The Payment Amount should contain only numbers and a decimal point, such as 123.45.',10,$("custom_paymentamount").id);
					$("custom_paymentamount").focus();
					return false;
				}
			}
			else
			{
				inlineMsg($("custom_paymentamount").id,'<strong>Invalid Value: </strong>The Payment Amount cannot be blank.',10,$("custom_paymentamount").id);
				$("custom_paymentamount").focus();
				return false;
			}
		}


		function ValidateTaxAmount( oFee )
		{
			var bValid = true;
			var total = 0.00;

			// Remove any extra spaces
			oFee.value = removeSpaces(oFee.value);
			//Remove commas that would cause problems in validation
			oFee.value = removeCommas(oFee.value);

			// Validate the format of the price
			if (oFee.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(oFee.value);
				if ( Ok )
				{
					oFee.value = format_number(Number(oFee.value),2);
					if (Number(oFee.value) > Number(999.99))
					{
						oFee.value = format_number(0,2);
						bValid = false;
					}
				}
				else 
				{
					oFee.value = format_number(0,2);
					bValid = false;
				}
			}

			UpdateIncomeTaxTotal( );

			if ( bValid == false ) 
			{
				$("skip_feesok").value = "false";
				inlineMsg(oFee.id,'<strong>Invalid Value: </strong>The Payment Amount should be a number in currency format and not more than $999.99.',10,oFee.id);
				oFee.focus();
				return false;
			}

			$("skip_feesok").value = "true";
			return true;
		}

		function UpdateIncomeTaxTotal( )
		{
			var Fee = 0.00;
			var PaymentAmount = Number($("custom_taxamount").value);
			var total = 0.00;

			Fee = PaymentAmount * 0.02;
			Fee = format_number(Fee,2);
			$("custom_feeamount").innerHTML = Fee;
			$("custom_servicefee").value = Fee;
			total = Number(Fee) + PaymentAmount;
			$("custom_totalamount").innerHTML = format_number(total,2);
			$("custom_paymentamount").value = format_number(total,2);
		}
		

		function validateWaterPayment()
		{
			if ($("skip_feesok").value == "true")
			{
				var sTotal = $("custom_paymentamount").value; 
				if (Number(sTotal) > Number(0.00))
				{
					if (validateForm('frmpayment')) 
					{ 
						document.frmpayment.submit(); 
					}
				}
				else
				{
					inlineMsg($("custom_billamount").id,'<strong>No Payment Amount Entered: </strong>Please enter the amount you are paying.',10,$("custom_billamount").id);
				}
			}
			else
			{
				$("skip_feesok").value = "true";
			}
		}

		function ValidateWaterFee( oFee )
		{
			var bValid = true;
			var total = 0.00;

			// Remove any extra spaces
			oFee.value = removeSpaces(oFee.value);
			//Remove commas that would cause problems in validation
			oFee.value = removeCommas(oFee.value);

			// Validate the format of the price
			if (oFee.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(oFee.value);
				if ( Ok )
				{
					oFee.value = format_number(Number(oFee.value),2);
					if (Number(oFee.value) > Number(9999.99))
					{
						oFee.value = format_number(0,2);
						bValid = false;
					}
				}
				else 
				{
					oFee.value = format_number(0,2);
					bValid = false;
				}
			}

			UpdateWaterFeeTotal( );

			if ( bValid == false ) 
			{
				$("skip_feesok").value = "false";
				inlineMsg(oFee.id,'<strong>Invalid Value: </strong>The Payment Amount should be a number in currency format and less than $9,999.99.',10,oFee.id);
				oFee.focus();
				return false;
			}

			$("skip_feesok").value = "true";
			return true;
		}

		function UpdateWaterFeeTotal( )
		{
			var Fee = 0.00;
			var PaymentAmount = Number($("custom_billamount").value);
			var total = 0.00;

			Fee = PaymentAmount * 0.02;
			Fee = format_number(Fee,2);
			$("custom_feeamount").innerHTML = Fee;
			$("custom_servicefee").value = Fee;
			total = Number(Fee) + PaymentAmount;
			$("custom_totalamount").innerHTML = format_number(total,2);
			$("custom_paymentamount").value = format_number(total,2);
		}

		function validateCourtPayment()
		{
			if ($("skip_feesok").value == "true")
			{
				var sTotal = $("custom_paymentamount").value; 
				if (Number(sTotal) > Number(0.00))
				{
					if (validateForm('frmpayment')) 
					{ 
						document.frmpayment.submit(); 
					}
				}
				else
				{
					inlineMsg($("custom_ticketamount").id,'<strong>No Ticket Amount Entered: </strong>Please enter the amount you are paying.',10,$("custom_ticketamount").id);
				}
			}
			else
			{
				$("skip_feesok").value = "true";
			}
		}

		function ValidateTicketAmount( oFee )
		{
			var bValid = true;
			var total = 0.00;

			// Remove any extra spaces
			oFee.value = removeSpaces(oFee.value);
			//Remove commas that would cause problems in validation
			oFee.value = removeCommas(oFee.value);

			// Validate the format of the price
			if (oFee.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(oFee.value);
				if ( Ok )
				{
					oFee.value = format_number(Number(oFee.value),2);
					if (Number(oFee.value) > Number(9999.99))
					{
						oFee.value = format_number(0,2);
						bValid = false;
					}
				}
				else 
				{
					oFee.value = format_number(0,2);
					bValid = false;
				}
			}

			UpdateTicketFeeTotal( );

			if ( bValid == false ) 
			{
				$("skip_feesok").value = "false";
				inlineMsg(oFee.id,'<strong>Invalid Value: </strong>The Ticket Amount should be a number in currency format and less than $9,999.99.',10,oFee.id);
				oFee.focus();
				return false;
			}

			$("skip_feesok").value = "true";
			return true;
		}

		function UpdateTicketFeeTotal( )
		{
			var Fee = 0.00;
			var PaymentAmount = Number($("custom_ticketamount").value);
			var total = 0.00;

			Fee = PaymentAmount * 0.02;
			Fee = format_number(Fee,2);
			$("custom_feeamount").innerHTML = Fee;
			$("custom_servicefee").value = Fee;
			total = Number(Fee) + PaymentAmount;
			$("custom_totalamount").innerHTML = format_number(total,2);
			$("custom_paymentamount").value = format_number(total,2);
		}

		function validatePayment()
		{
			if ($("skip_feesok").value == "true")
			{
				var sTotal = $("custom_paymentamount").value; 
				if (Number(sTotal) > Number(0.00))
				{
					if (validateForm('frmpayment')) 
					{ 
						document.frmpayment.submit(); 
					}
				}
				else
				{
					inlineMsg($("custom_total").id,'<strong>No Fees Entered: </strong>The fee total should be more than 0.00.',10,$("custom_total").id);
				}
			}
			else
			{
				$("skip_feesok").value = "true";
			}
		}

		function ValidateFee( oFee )
		{
			var bValid = true;
			var total = 0.00;

			// Remove any extra spaces
			oFee.value = removeSpaces(oFee.value);
			//Remove commas that would cause problems in validation
			oFee.value = removeCommas(oFee.value);

			// Validate the format of the price
			if (oFee.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(oFee.value);
				if ( Ok )
				{
					oFee.value = format_number(Number(oFee.value),2);
					if (Number(oFee.value) > Number(99.99))
					{
						oFee.value = format_number(0,2);
						bValid = false;
					}
				}
				else 
				{
					oFee.value = format_number(0,2);
					bValid = false;
				}
			}

			UpdateFeeTotal( );

			if ( bValid == false ) 
			{
				$("skip_feesok").value = "false";
				inlineMsg(oFee.id,'<strong>Invalid Value: </strong>Fees should be numbers in currency format and less than $100.',10,oFee.id);
				oFee.focus();
				return false;
			}

			$("skip_feesok").value = "true";
			return true;
		}

		function UpdateFeeTotal( )
		{
			var total = 0.00;

			for (var x=1; x < 5 ; x++ )
			{
				if ($("custom_fee" + x).value != "")
				{
					total += Number($("custom_fee" + x).value);
				}
			}
			$("custom_total").value = format_number(total,2);
			$("custom_paymentamount").value = format_number(total,2);
		}

		function UpperCaseState()
		{
			var sState = $("custom_state").value;
			$("custom_state").value = sState.toUpperCase();
		}

		function FindMe()
		{
			var okToGo = true;
			if ($F("custom_applicantfirstname") == "")
			{
				inlineMsg("custom_applicantfirstname",'<strong>Missing Field: </strong>Please input a first name, then try again.',10,"custom_applicantfirstname");
				$('custom_applicantfirstname').focus();
				okToGo = false;
			}

			if ($F("custom_applicantlastname") == "")
			{
				inlineMsg("custom_applicantlastname",'<strong>Missing Field: </strong>Please input a last name, then try again.',10,"custom_applicantlastname");
				$('custom_applicantlastname').focus();
				okToGo = false;
			}

			if ( okToGo )
			{
				// Hide any old messages
				$("notfoundmsg").hide();
				$("alreadyrenewedmsg").hide();

				// Hide the input fields
				var hideRows = $$('tr.mainfields');
				hideRows.each( Element.hide );

				$("custom_permitholdertype").value = $RF('frmpayment', 'custom_permitholdertypes');

				// document.frmPermit.usetypeid.options[document.frmPermit.usetypeid.selectedIndex].value
				var sParameter = 'permitholdertype=' + encodeURIComponent($F("custom_permitholdertype"));
				sParameter += '&applicantfirstname=' + encodeURIComponent($F("custom_applicantfirstname"));
				sParameter += '&applicantlastname=' + encodeURIComponent($F("custom_applicantlastname"));
				//alert( sParameter );

				doAjax('permitrenewallookup.asp', sParameter, 'FindMeResults', 'post', '0');
			}
		}

		function FindMeInWaitlist()
		{
			var okToGo = true;
			if ($F("custom_applicantfirstname") == "")
			{
				inlineMsg("custom_applicantfirstname",'<strong>Missing Field: </strong>Please input a first name, then try again.',10,"custom_applicantfirstname");
				$('custom_applicantfirstname').focus();
				okToGo = false;
			}

			if ($F("custom_applicantlastname") == "")
			{
				inlineMsg("custom_applicantlastname",'<strong>Missing Field: </strong>Please input a last name, then try again.',10,"custom_applicantlastname");
				$('custom_applicantlastname').focus();
				okToGo = false;
			}

			if ( okToGo )
			{
				// Hide any old messages
				$("notfoundmsg").hide();
				$("alreadyrenewedmsg").hide();

				// Hide the input fields
				var hideRows = $$('tr.mainfields');
				hideRows.each( Element.hide );

				var sParameter = 'permitholdertype=' + encodeURIComponent($F('custom_permitholdertype'));
				sParameter += '&applicantfirstname=' + encodeURIComponent($F("custom_applicantfirstname"));
				sParameter += '&applicantlastname=' + encodeURIComponent($F("custom_applicantlastname"));
				//alert( sParameter );

				doAjax('permitrenewallookup.asp', sParameter, 'FindMeResults', 'post', '0');
			}
		}

		function FindMeResults( sReturnJSON )
		{
			var json = sReturnJSON.evalJSON(true);

			if (json.flag == 'success')
			{
				// populate the hidden fields
				$("renewalid").value = json.renewalid;
				$("custom_applicantfirstname").value = json.applicantfirstname;
				$("custom_applicantlastname").value = json.applicantlastname;
				$("custom_applicantaddress").value = json.applicantaddress;
				$("custom_applicantcity").value = json.applicantcity;
				$("custom_applicantstate").value = json.applicantstate;
				$("custom_applicantzip").value = json.applicantzip;
				$("custom_applicantphone").value = json.applicantphone;
				
				if ($("custom_vehiclelicense") != undefined)
				{
					$("custom_vehiclelicense").value = json.vehiclelicense;
				}

				// Hide the find me button
				$("findmebtn").hide();

				// make the applicant name fields readonly
				$("custom_applicantfirstname").readOnly = true;
				$("custom_applicantlastname").readOnly = true;
				// Disable the radio buttons on the premit renewal form, they do not have an id
				if ( $('custom_permitholdertypes') == undefined )
				{
					var form = $('frmpayment');
					var buttons = form.getInputs('radio', 'custom_permitholdertypes');
					buttons.invoke('disable');
				}

				// Show the hidden fields
				var showRows = $$('tr.mainfields');
				showRows.each( Element.show );
			}

			if (json.flag == 'notfound')
			{
				// Show the not found message
				$("notfoundmsg").show();
			}

			if (json.flag == 'duplicate')
			{
				// Show the duplicate renewal message
				$("alreadyrenewedmsg").show();
			}
		}

		jQuery(document).ready(function(){ 
			jQuery('.numbersOnly').keydown(function (event) {  
				if ( event.keyCode == 46 || event.keyCode == 8 || event.keyCode == 9 || event.keyCode == 27 || event.keyCode == 13 || (event.keyCode == 65 && event.ctrlKey === true) || (event.keyCode >= 35 && event.keyCode <= 39)) {
					return;
				}
				else {
				
					if (event.shiftKey || (event.keyCode < 48 || event.keyCode > 57) && (event.keyCode < 96 || event.keyCode > 105 )) {
						event.preventDefault();
					}
				}
			}); 

			jQuery('.zipCodeOnly').keyup(function () {  
				this.value = this.value.replace(/[^0-9 \-]/g,''); 
			}); 

		});

	//-->
	</script>

</head>

<!--#Include file="include_top.asp"-->

<%
 'BODY Content
  lcl_org_name = oPayOrg.GetOrgName()
  lcl_org_state = oPayOrg.GetState()
  lcl_org_featurename = "Permits and Payments Center"

  response.write "<p>" & vbcrlf
  oPayOrg.buildWelcomeMessage iorgid, lcl_orghasdisplay_action_page_title, lcl_org_name, lcl_org_state, lcl_org_featurename
  response.write "<br />" & vbcrlf
  'response.write "<font class=""pagetitle"">Welcome to the " & oPayOrg.GetOrgName() & ", " & oPayOrg.GetState() & ", Permits and Payments Center</font><br />" & vbcrlf

 'User Registration and User Menu
  RegisteredUserDisplay( "" )

  response.write "</p>" & vbcrlf

'BEGIN: DISPLAY PAGE CONTENT --------------------------------------------------
If trim(request("paymenttype")) = "" Then 
	
	'--------------------------------------------------------------------------------------------------
	' BEGIN: VISITOR TRACKING
	'--------------------------------------------------------------------------------------------------
	iSectionID = 3
	sDocumentTitle = "MAIN"
	sURL = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
	datDate = Date()	
	datDateTime = Now()
	sVisitorIP = request.servervariables("REMOTE_ADDR")
	Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
	'--------------------------------------------------------------------------------------------------
	' END: VISITOR TRACKING
	'--------------------------------------------------------------------------------------------------
	
	
	' DISPLAY LIST OF PAYMENTS
	response.write "<table><tr><td valign=top>"
	response.write "<div class=""box_header2"">Payment and Permit Services</div>"
	response.write "<div class=""groupSmall"">"

	' List the payment forms available
	DisplayPayments

	response.write "</div>"
	response.write "</td>"
%>

	<td width="225" style="padding-left:15px;" valign="top">

	<!--BEGIN: REGISTER/LOGIN LINKS-->
	<%If sOrgRegistration AND (request.cookies("userid") = "" OR request.cookies("userid") = "-1") Then %>
		<b>Personalized E-Gov Services</b>
		<ul>
			<li><a href="user_login.asp">Click here to Login</a>
			<li><a href="register.asp">Click here to Register</a>
		</ul>
		<hr style="width: 90%; size: 1px; height: 1px;">
	<%End If%>
	<!--END: REGISTER/LOGIN LINKS-->


	<%=sPaymentDescription%></td>
	</tr>
	</table>

	<% Set oPayOrg = Nothing %>

	<!--SPACING CODE-->
	<p><bR>&nbsp;<bR>&nbsp;</p>
	<!--SPACING CODE-->

	
<%
Else

	' DISPLAY PAYMENT FORM
	DisplayPaymentForm CLng(request("paymenttype"))

End If
' ---------------------------------------------------------------------------------------
' END DISPLAY PAGE CONTENT
' ---------------------------------------------------------------------------------------
%>


<!--SPACING CODE-->
<p>&nbsp;<bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->


<!--#Include file="include_bottom.asp"-->  

<%
' -----------------------------------------------------------------------------------------------------------
' FUNCTIONS AND SUBROUTINES
' -----------------------------------------------------------------------------------------------------------

' -----------------------------------------------------------------------------------------------------------
' void DisplayPaymentForm iID
' -----------------------------------------------------------------------------------------------------------
Sub DisplayPaymentForm( ByVal iID )
	Dim oRs, sSql, bFound, sService, bShowButton

	bFound = False 
	bShowButton = False 

	' GET FORM INFORMATION
	sSql = "SELECT * FROM egov_paymentservices WHERE paymentserviceid = " & iID & " AND orgid = " & iOrgid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
		
	If Not oRs.EOF Then

		'--------------------------------------------------------------------------------------------------
		' BEGIN: VISITOR TRACKING
		'--------------------------------------------------------------------------------------------------
		iSectionID = 33
		sDocumentTitle = oRs("paymentservicename")
		sURL = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
		datDate = Date()	
		datDateTime = Now()
		sVisitorIP = request.servervariables("REMOTE_ADDR")
		Call LogPageVisit( iSectionID, sDocumentTitle, sURL, datDate, datDateTime, sVisitorIP, iorgid )
		'--------------------------------------------------------------------------------------------------
		' END: VISITOR TRACKING
		'--------------------------------------------------------------------------------------------------
		
		' FORM HEADING	
		response.write "<blockquote>"
		response.write "<font class=""formtitle"">" & oRs("paymentservicename") & "</font>"
		response.write "<div class=""group"">"

		' FORM DESCRIPTION 
		If oRs("paymentservicedescription") <> "" Then
			response.write oRs("paymentservicedescription")
		End If

		' FORM INSTRUCTIONS
		If oRs("paymentserviceinstructions") <> "" Then
			response.write oRs("paymentserviceinstructions")
		End If


		' FORM PAYMENT OPTIONS
		fn_GetPaymentGatewayOptions iPaymentGatewayID ,iID, oRs("paymentservicename")

		' FORM REQUIRED VALUES
		response.write vbcrlf & vbcrlf & "<table id=""maintable"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		
		' FORM SPECIFIC FIELD OPTIONS
		If iID = 21 Or iID = 22 Or iID = 23 Or iID = 254 Or iID = 270 Or iID = 276 Or iID = 277 Then
			Custom_Payment_FormID iID 
			If iID = 254 Or iID = 23 Then 
				bShowButton = True 
			End If 
		Else
			' The payment service field has been created to handle special payment forms without having to hard code the form id as above - SJL 8/2011
			sService = oRs("paymentservice")
			Select Case sService
				Case "rye commuter permit renewal"
					'This is the special Rye permit renewal form
					GetRyePaymentFields iID
				Case "rye commuter waitlist renewal"
					'This is the special Rye waitlist renewal form
					GetRyeWaitlistRenewalFields iID
				Case "west carrollton court payments"
					ShowWestCarrolltonCourtPaymentForm
					bShowButton = True 
				Case Else 
					' Normal Payment services
					GetPaymentFields iID 
					bShowButton = True 
			End Select  

		End If

		' OPTIONAL FIELDS AND AMOUNT FOR PAY PAL
		GetPayPalFieldValues iID 

		' SUBMIT BUTTON
		Select Case iId
			Case 22
				response.write vbcrlf & "<tr><td colspan=""2"" align=""right""><input onclick=""vcheck();"" type=""button"" class=""paymentbtn"" name=""btnsubmit"" value=""CONTINUE"" alt=""CONTINUE""></td></tr></table>"
			Case 270
				response.write vbcrlf & "<tr><td colspan=""4"" align=""center""><input onclick=""validatePayment();"" type=""button"" class=""button"" name=""btnsubmit"" value=""CONTINUE"" alt=""CONTINUE""></td></tr></table>"
			Case 276
				response.write vbcrlf & "<tr><td colspan=""3"" align=""center""><input onclick=""validateWaterPayment();"" type=""button"" class=""button"" name=""btnsubmit"" value=""CONTINUE"" alt=""CONTINUE""></td></tr></table>"
			Case 277
				response.write vbcrlf & "<tr><td colspan=""3"" align=""center""><input onclick=""validateIncomeTaxPayment();"" type=""button"" class=""button"" name=""btnsubmit"" value=""CONTINUE"" alt=""CONTINUE""></td></tr></table>"
			Case Else
				If bShowButton Then 
					Select Case oRs("paymentservice")
						Case "west carrollton court payments"
							response.write vbcrlf & "<tr><td colspan=""3"" align=""center""><input onclick=""validateCourtPayment();"" type=""button"" class=""button"" name=""btnsubmit"" value=""CONTINUE"" alt=""CONTINUE""></td></tr></table>"
						Case Else 
							response.write vbcrlf & "<tr><td colspan=""2"" align=""right""><input onclick=""ValidatePaymentAmount()"" type=""button"" class=""paymentbtn"" name=""btnsubmit"" value=""CONTINUE"" alt=""CONTINUE"" /></td></tr></table>"
					End Select
				End If 
		End Select 

		' FORM NOTES
		If oRs("paymentservicenotes") <> "" Then
			response.write oRs("paymentservicenotes")
		End If

		' END FORM	
		response.write vbcrlf & "</form></div>"

		bFound = True 

	End If

	oRs.Close
	Set oRs = Nothing 

	If Not bFound Then 
		' FORM NOT FOUND REDIRECT TO COMPLETE LIST
		response.redirect("payment.asp")
	End If 

End Sub  


'------------------------------------------------------------------------------------------------------------
' void GetPayPalFieldValues IID 
'------------------------------------------------------------------------------------------------------------
Sub GetPayPalFieldValues( ByVal iID )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_paypalfields WHERE paymentserviceid = " & iID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
		
	If Not oRs.EOF Then
		'If oRs("on0") <> "" Then
			'response.write "<tr><td><input type=""hidden"" name=""on0"" value=""" & oRs("on0") & """ ><b>" & 'oRs("on0")& "</b></td><td><input type=""text"" name=""os0"" maxlength=""200""></td></tr>"
		'End If

		'If oRs("on1") <> "" Then
			'response.write "<tr><td><input type=""hidden"" name=""on1"" value=""" & oRs("on1") & """><b>" & 'oRs("on1") & "</b></td><td><input type=""text"" name=""os1"" maxlength=""200""></td></tr>"
		'End If

		If oRs("amount") <> "" Then
			curValue = formatnumber(oRs("amount"),2)
			sDisabled = "DISABLED"
			response.write vbcrlf & "<tr><td><b>Payment Amount: </b></td><td>" & curValue &"<input type=""hidden"" name=""amount"" value=""" & curValue &""" /></td></tr>"
		Else
			curValue = ""
			sDisabled = ""
			response.write vbcrlf & "<input type=""hidden"" name=""ef:amount-text/req"" value=""Payment Amount"" />"
			response.write vbcrlf & "<tr><td><b>Payment Amount: </b></td><td><input type=""text"" name=""AMOUNT"" maxlength=""200"" value=""" & curValue &""" /></td></tr>"
		End If
	End If

	oRs.Close 
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------------------------------------
' void GetPaymentFields IID 
'------------------------------------------------------------------------------------------------------------
Sub GetPaymentFields( ByVal iID )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_paymentfields WHERE paymentserviceid = " & iID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
		
	If Not oRs.EOF Then
		Do While Not oRs.EOF 
			' DYNAMIC VALUES
			
			' BUILD EASY VALIDATION STRING
			If ISNULL(oRs("paymentvalidation")) OR oRs("paymentvalidation") = "" Then
				sValidation = ""
			Else
				If oRs("paymentfieldtype")  = "radio" Then
					sValidation = "-radio/" & oRs("paymentvalidation") & "/req"
				Else 
					sValidation = "-text/" & oRs("paymentvalidation") & "/req"
				End If 
			End If
		 
		   ' WRITE EASY FORM HIDDEN FIELD VALIDATION VALUE
		   response.write vbcrlf & "<input type=hidden name=""ef:custom_" & oRs("paymentfieldsname") &   sValidation &  """ value=""" &  oRs("paymentfielddisplayname")  & """>"
		
		   ' WRITE ACTUAL FORM VALUE	
		   Select Case oRs("paymentfieldtype") 
				Case "radio"
					' Radio Buttons
					response.write vbcrlf & "<tr><td colspan=""3"" align=""left"" nowrap=""nowrap""><strong>" & oRs("paymentfielddisplayname") & "</strong> &ndash; Select from the following list.</td></tr>"

					arrAnswers = Split(oRs("answerlist"),Chr(10))
					
					For alist = 0 to UBound(arrAnswers)
						arrAnswers(alist) = RemoveNewLine(arrAnswers(alist))
						response.write vbcrlf & "<tr><td colspan=""3"" align=""left"" nowrap=""nowrap""><input value=""" & arrAnswers(alist) & """ name=""custom_" & oRs("paymentfieldsname") & """ class=""formradio"" type=""radio"" " 
						If clng(alist) = clng(0) Then
							response.write "checked=""checked"" "
						End If 
						response.write "/>" & arrAnswers(alist) & "</td></tr>"
					Next
					response.write vbcrlf & "<tr><td colspan=""3"">&nbsp;</td><tr>"

				Case "textarea"
				' TEXTAREA
				response.write vbcrlf & "<tr><td colspan=""2"" align=""left"" nowrap=""nowrap""><b>" & oRs("paymentfielddisplayname") & " :</b><br /><textarea " & oRs("paymentfieldattributes") & "  name=""custom_" & oRs("paymentfieldsname") & """ style=""" & oRs("paymentfieldstyle") & """ class=""formtextarea""></textarea></td><td> " & oRs("paymentdesc") & " &nbsp;</td></tr>" & vbcrlf 

				Case Else
				'SPECIAL CASE FOR TRACKING NUMBERS
				If sOrgRegistration And request.cookies("userid") <> "" And request.cookies("userid") <> "-1" And request.querystring("paymenttype") = 40 AND lcase(oRs("paymentfieldsname")) = "trackingnumber" Then
					sSql = "SELECT action_autoid, submit_date FROM egov_actionline_requests WHERE userid = '" & request.cookies("userid") & "' AND status IN ('WAITING','submitted','INPROGRESS','EVALFORM') AND category_id=295"
					Set oTrackingNumbers = Server.CreateObject("ADODB.Recordset")
					oTrackingNumbers.Open sSql, Application("DSN"), 3, 1
					response.write vbcrlf & "<tr><td align=""left"" nowrap=""nowrap""><b>" & oRs("paymentfielddisplayname") & " :</b></td>"
					response.write vbcrlf & "<td align=""left""> "
					response.write vbcrlf & "<select " & oRs("paymentfieldattributes") & "  name=""custom_" & oRs("paymentfieldsname") & """ style=""" & oRs("paymentfieldstyle") & """ >"
					Do While Not oTrackingNumbers.EOF
						lngTrackingNumber = oTrackingNumbers("action_autoid")  & replace(FormatDateTime(cdate(oTrackingNumbers("submit_date")),4),":","")
						response.write vbcrlf & "<option value=""" & lngTrackingNumber & """"
						If request.querystring(oRs("paymentfieldsname")) = lngTrackingNumber Then 
							response.write " selected "
						End If 
						response.write ">" & lngTrackingNumber & "</option>"
						oTrackingNumbers.MoveNext
					Loop
					oTrackingNumbers.Close
					Set oTrackingNumbers = Nothing 
					response.write cbcrlf & "</td><td width=""40%""> " & oRs("paymentdesc") & " &nbsp;</td></tr>" & vbcrlf 
				' DEFAULT IS TEXTBOX
				Else 
					response.write vbcrlf & "<tr><td align=""left"" nowrap=""nowrap""><strong>" & oRs("paymentfielddisplayname") & " :</strong></td><td align=""left"">"
					response.write "<input type=""text"" " & oRs("paymentfieldattributes") & " id=""custom_" & oRs("paymentfieldsname") & """ name=""custom_" & oRs("paymentfieldsname") & """ style=""" & oRs("paymentfieldstyle") & """ value=""" & request.querystring(oRs("paymentfieldsname")) & """ /></td><td width=""40%""> " & oRs("paymentdesc") & " &nbsp;</td></tr>"
				End If 

			End Select
			
			oRs.MoveNext
		Loop

	End If

	oRs.Close 
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------------------------------------
' void GetRyePaymentFields iPaymentServiceId 
'------------------------------------------------------------------------------------------------------------
Sub GetRyePaymentFields( ByVal iPaymentServiceId )
	Dim sSql, oRs, iFieldCount, sRowClass

	iFieldCount = clng(0)
	sRowClass = ""

	sSql = "SELECT * FROM egov_paymentfields WHERE paymentserviceid = " & iPaymentServiceId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
		
	If Not oRs.EOF Then
		'response.write vbcrlf & vbcrlf & "<table id=""maintable"" border=""0"" cellpadding=""0"" cellspacing=""0"">"

		Do While Not oRs.EOF 
			iFieldCount = iFieldCount + 1
			If iFieldCount > 3 Then
				sRowClass = " class=""mainfields"" style=""display:none;"""
			End If 
			' DYNAMIC VALUES
			
			' BUILD EASY VALIDATION STRING
			If ISNULL(oRs("paymentvalidation")) OR oRs("paymentvalidation") = "" Then
				sValidation = ""
			Else
				If oRs("paymentfieldtype")  = "radio" Then
					sValidation = "-radio/" & oRs("paymentvalidation") & "/req"
				Else 
					sValidation = "-text/" & oRs("paymentvalidation") & "/req"
				End If 
			End If
		 
		   ' WRITE EASY FORM HIDDEN FIELD VALIDATION VALUE
		   response.write vbcrlf & "<input type=""hidden"" name=""ef:custom_" & oRs("paymentfieldsname") &   sValidation &  """ value=""" &  oRs("paymentfielddisplayname")  & """>"
		
		   ' WRITE ACTUAL FORM VALUE	
		   Select Case oRs("paymentfieldtype") 
				Case "radio"
					' Radio Buttons
					response.write vbcrlf & "<tr" & sRowClass & "><td colspan=""3"" align=""left"" nowrap=""nowrap""><strong>" & oRs("paymentfielddisplayname") & "</strong> &ndash; Select from the following list.</td></tr>"

					arrAnswers = Split(oRs("answerlist"),Chr(10))
					
					For alist = 0 to UBound(arrAnswers)
						arrAnswers(alist) = RemoveNewLine(arrAnswers(alist))
						response.write vbcrlf & "<tr" & sRowClass & "><td colspan=""3"" align=""left"" nowrap=""nowrap""><input value=""" & arrAnswers(alist) & """ name=""custom_" & oRs("paymentfieldsname") & """ class=""formradio"" type=""radio"" " 
						If clng(alist) = clng(0) Then
							response.write "checked=""checked"" "
						End If 
						response.write "/>" & arrAnswers(alist) & "</td></tr>"
					Next
					response.write vbcrlf & "<tr" & sRowClass & "><td colspan=""3"">&nbsp;</td><tr>"

				Case "textarea"
					' TEXTAREA
					response.write vbcrlf & "<tr" & sRowClass & "><td colspan=""2"" align=""left"" nowrap=""nowrap""><b>" & oRs("paymentfielddisplayname") & " :</b><br /><textarea " & oRs("paymentfieldattributes") & "  name=""custom_" & oRs("paymentfieldsname") & """ style=""" & oRs("paymentfieldstyle") & """ class=""formtextarea""></textarea></td><td> " & oRs("paymentdesc") & " &nbsp;</td></tr>" & vbcrlf 

				Case Else
					' Text
					response.write vbcrlf & "<tr" & sRowClass & "><td align=""left"" nowrap=""nowrap""><strong>" & oRs("paymentfielddisplayname") & " :</strong></td><td align=""left"">"
					response.write "<input type=""text"" " & oRs("paymentfieldattributes") & " id=""custom_" & oRs("paymentfieldsname") & """ name=""custom_" & oRs("paymentfieldsname") & """ style=""" & oRs("paymentfieldstyle") & """ value=""" & request.querystring(oRs("paymentfieldsname")) & """ /></td>"
					If iFieldCount = 3 Then
						response.write "<td width=""40%"" align=""center""><input onclick=""FindMe()"" type=""button"" id=""findmebtn"" class=""paymentbtn"" value=""Find Me"" /></td></tr>"
					Else
						response.write "<td width=""40%""> " & oRs("paymentdesc") & " &nbsp;</td></tr>"
					End If 
					

			End Select

			oRs.MoveNext
		Loop

	End If

	oRs.Close 
	Set oRs = Nothing

	' Continue button 
	response.write vbcrlf & "<tr class=""mainfields"" style=""display:none;""><td colspan=""3"">&nbsp;</td><tr>"
	response.write vbcrlf & "<tr class=""mainfields"" style=""display:none;""><td colspan=""3"" align=""center""><input onclick=""ValidatePaymentAmount()"" type=""button"" class=""paymentbtn"" name=""btnsubmit"" value=""CONTINUE"" alt=""CONTINUE"" /></td></tr>"
	response.write vbcrlf & "</table>"
	response.write vbcrlf & "<input type=""hidden"" id=""renewalid"" name=""renewalid"" value=""0"" />"
	response.write vbcrlf & "<input type=""hidden"" id=""paymentservice"" name=""paymentservice"" value=""rye commuter permit renewal"" />"
	response.write vbcrlf & "<input type=""hidden"" id=""custom_permitholdertype"" name=""custom_permitholdertype"" value="""" />"

	'  Name not found message
	response.write vbcrlf & "<div id=""notfoundmsg"" style=""display:none;"">Your name could not be found in our master list.<br />Please contact us at 914-967-7371 during regular business hours (M-F - 9:00 AM to 5:00 PM) before proceeding.</div>"

	' Already renewed message
	response.write vbcrlf & "<div id=""alreadyrenewedmsg"" style=""display:none;"">Our records show that you have already renewed.<br />Please contact us at 914-967-7371 during regular business hours (M-F - 9:00 AM to 5:00 PM) before proceeding.</div>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' void GetRyeWaitlistRenewalFields iPaymentServiceId 
'------------------------------------------------------------------------------------------------------------
Sub GetRyeWaitlistRenewalFields( ByVal iPaymentServiceId )
	Dim sSql, oRs, iFieldCount, sRowClass

	iFieldCount = clng(0)
	sRowClass = ""

	sSql = "SELECT * FROM egov_paymentfields WHERE paymentserviceid = " & iPaymentServiceId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
		
	If Not oRs.EOF Then
		'response.write vbcrlf & vbcrlf & "<table id=""maintable"" border=""0"" cellpadding=""0"" cellspacing=""0"">"

		Do While Not oRs.EOF 
			iFieldCount = iFieldCount + 1
			If iFieldCount > 2 Then
				sRowClass = " class=""mainfields"" style=""display:none;"""
			End If 
			' DYNAMIC VALUES
			
			' BUILD EASY VALIDATION STRING
			If ISNULL(oRs("paymentvalidation")) OR oRs("paymentvalidation") = "" Then
				sValidation = ""
			Else
				If oRs("paymentfieldtype")  = "radio" Then
					sValidation = "-radio/" & oRs("paymentvalidation") & "/req"
				Else 
					sValidation = "-text/" & oRs("paymentvalidation") & "/req"
				End If 
			End If
		 
		   ' WRITE EASY FORM HIDDEN FIELD VALIDATION VALUE
		   response.write vbcrlf & "<input type=""hidden"" name=""ef:custom_" & oRs("paymentfieldsname") &   sValidation &  """ value=""" &  oRs("paymentfielddisplayname")  & """>"
		
		   ' WRITE ACTUAL FORM VALUE	
		   Select Case oRs("paymentfieldtype") 
				Case "radio"
					' Radio Buttons
					response.write vbcrlf & "<tr" & sRowClass & "><td colspan=""3"" align=""left"" nowrap=""nowrap""><strong>" & oRs("paymentfielddisplayname") & "</strong> &ndash; Select from the following list.</td></tr>"

					arrAnswers = Split(oRs("answerlist"),Chr(10))
					
					For alist = 0 to UBound(arrAnswers)
						arrAnswers(alist) = RemoveNewLine(arrAnswers(alist))
						response.write vbcrlf & "<tr" & sRowClass & "><td colspan=""3"" align=""left"" nowrap=""nowrap""><input value=""" & arrAnswers(alist) & """ name=""custom_" & oRs("paymentfieldsname") & """ class=""formradio"" type=""radio"" " 
						If clng(alist) = clng(0) Then
							response.write "checked=""checked"" "
						End If 
						response.write "/>" & arrAnswers(alist) & "</td></tr>"
					Next
					response.write vbcrlf & "<tr" & sRowClass & "><td colspan=""3"">&nbsp;</td><tr>"

				Case "textarea"
					' TEXTAREA
					response.write vbcrlf & "<tr" & sRowClass & "><td colspan=""2"" align=""left"" nowrap=""nowrap""><b>" & oRs("paymentfielddisplayname") & " :</b><br /><textarea " & oRs("paymentfieldattributes") & "  name=""custom_" & oRs("paymentfieldsname") & """ style=""" & oRs("paymentfieldstyle") & """ class=""formtextarea""></textarea></td><td> " & oRs("paymentdesc") & " &nbsp;</td></tr>" & vbcrlf 

				Case Else
					' Text
					response.write vbcrlf & "<tr" & sRowClass & "><td align=""left"" nowrap=""nowrap""><strong>" & oRs("paymentfielddisplayname") & " :</strong></td><td align=""left"">"
					response.write "<input type=""text"" " & oRs("paymentfieldattributes") & " id=""custom_" & oRs("paymentfieldsname") & """ name=""custom_" & oRs("paymentfieldsname") & """ style=""" & oRs("paymentfieldstyle") & """ value=""" & request.querystring(oRs("paymentfieldsname")) & """ /></td>"
					If iFieldCount = 2 Then
						response.write "<td width=""40%"" align=""center""><input onclick=""FindMeInWaitlist()"" type=""button"" id=""findmebtn"" class=""paymentbtn"" value=""Find Me"" /></td></tr>"
					Else
						response.write "<td width=""40%""> " & oRs("paymentdesc") & " &nbsp;</td></tr>"
					End If 
					

			End Select

			oRs.MoveNext
		Loop

	End If

	oRs.Close 
	Set oRs = Nothing

	' Continue button 
	response.write vbcrlf & "<tr class=""mainfields"" style=""display:none;""><td colspan=""3"">&nbsp;</td><tr>"
	response.write vbcrlf & "<tr class=""mainfields"" style=""display:none;""><td colspan=""3"" align=""center""><input onclick=""ValidatePaymentAmount()"" type=""button"" class=""paymentbtn"" name=""btnsubmit"" value=""CONTINUE"" alt=""CONTINUE"" /></td></tr>"
	response.write vbcrlf & "</table>"
	response.write vbcrlf & "<input type=""hidden"" id=""renewalid"" name=""renewalid"" value=""0"" />"
	response.write vbcrlf & "<input type=""hidden"" id=""custom_permitholdertype"" name=""custom_permitholdertype"" value=""waitlist renewal"" />"
	response.write vbcrlf & "<input type=""hidden"" id=""paymentservice"" name=""paymentservice"" value=""rye commuter waitlist renewal"" />"

	'  Name not found message
	response.write vbcrlf & "<div id=""notfoundmsg"" style=""display:none;"">Your name could not be found in our master list.<br />Please contact us at 914-967-7371 during regular business hours (M-F - 9:00 AM to 5:00 PM) before proceeding.</div>"

	' Already renewed message
	response.write vbcrlf & "<div id=""alreadyrenewedmsg"" style=""display:none;"">Our records show that you have already renewed.<br />Please contact us at 914-967-7371 during regular business hours (M-F - 9:00 AM to 5:00 PM) before proceeding.</div>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' string RemoveNewLine( sString )
'------------------------------------------------------------------------------------------------------------
Function RemoveNewLine( ByVal sString )

	' remove the characters that make up the newline and carriage return from the string
	RemoveNewLine = Replace(sString,Chr(10),"")
	RemoveNewLine = Replace(sString,Chr(13),"")

End Function


'------------------------------------------------------------------------------------------------------------
' void DisplayPayments
'------------------------------------------------------------------------------------------------------------
Sub DisplayPayments( )
	Dim oRs, sSql

	sSql = "SELECT S.paymentserviceid, S.paymentservicename, S.paymentservice "
	sSql = sSql & "FROM egov_paymentservices S "
	sSql = sSql & "LEFT OUTER JOIN egov_organizations_to_paymentservices O ON S.paymentserviceid = O.paymentservice_id "
	sSql = sSql & "WHERE S.paymentserviceenabled = 1 AND O.paymentservice_enabled = 1 AND S.orgid = " & iorgid
	sSql = sSql & "ORDER BY S.displayorder, S.paymentserviceid DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<ul>"
	
	If Not oRs.EOF Then
		Do While Not oRs.EOF 
			response.write vbcrlf & "<li><a href=""payment.asp?paymenttype=" & oRs("paymentserviceid") & """>" & oRs("paymentservicename") &  "</a></li>"
			oRs.MoveNext
		Loop
	End If

	oRs.Close 
	Set oRs = Nothing

	' Now add some extra payment links for various clients
	If CLng(iorgid) = CLng(8) Then 
		' This is for Loveland, OH only
		response.write vbcrlf & "<li><a href=""http://www.officialpayments.com"">Official Payments</a></li>"
	End If 

	If CLng(iorgid) = CLng(69) Then 
		' This is for West Richlands only
		response.write vbcrlf & "<li><a href=""https://westrichland.merchanttransact.com"">Utility Payment</a></li>"
	End If 

	If CLng(iorgid) = CLng(79) Then 
		' This is for Angelton only
		response.write vbcrlf & "<li><a href=""https://angletontx.municipalonlinepayments.com/site/Pages/"">Water Bill Payment</a></li>"
	End If 
	 
	response.write vbcrlf & "</ul>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' void fn_GetPaymentGatewayOptions( iPaymentGatewayID, iPaymentServiceID, sServiceName )
'------------------------------------------------------------------------------------------------------------
Sub fn_GetPaymentGatewayOptions( ByVal iPaymentGatewayID, ByVal iPaymentServiceID, ByVal sServiceName )

	Select Case iPaymentGatewayID
		Case 1 
			' PayPal Gateway
			GetPayPalValues iPaymentServiceID

		Case 2
			' SkipJack Gateway
			GetSkipJackValues iPaymentServiceID, sServiceName

		Case 3
			' ECLink PayPal Demo Gateway
			GeteclinkPPDemoValues iPaymentServiceID, sServiceName

		Case 4
			' The Old Verisign PayFlow Pro Gateway
			GetVerisignValues iPaymentServiceID, sServiceName
		
		Case 5
			' Point and Pay Gateway - route them through the verisign form
			GetVerisignValues iPaymentServiceID, sServiceName
		
		Case 6
			If CLng(iOrgId) <> CLng(11) Then 
				' PayPal PayFlow Pro Gateway
				GetVerisignValues iPaymentServiceID, sServiceName
			Else
				' PayPal Gateway for Bullhead City - They have both methods - Standard and PayFlowPro
				' Standard is used for payments and PayFlowPro is used for recreation
				GetPayPalValues iPaymentServiceID
			End If 

		Case Else
			' ECLink PayPal Demo Gateway
			GeteclinkPPDemoValues iPaymentServiceID, sServiceName

	End Select

End Sub 


'------------------------------------------------------------------------------------------------------------
' void GetSkipJackValues iPaymentServiceID, sServiceName
'------------------------------------------------------------------------------------------------------------
Sub GetSkipJackValues( ByVal iPaymentServiceID, ByVal sServiceName )

	' This is what skip jack uses
	'response.write vbcrlf & "<form  name=""frmpayment"" action=""PAYMENT_PROCESSORS/SKIPJACK2004/ECSKIPJACK_FORM.ASP"" method=""post"">"
	response.write vbcrlf & "<form name=""frmpayment"" id=""frmpayment"" action=""" & Application("PAYMENTURL") & "/" & sorgVirtualSiteName & "/payment_processors/skipjack2004/ecskipjack_form.asp"" method=""post"">"
	response.write vbcrlf & "<input type=""hidden"" name=""PAYMENTID"" value=""" & iPaymentServiceID & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""ITEM_NUMBER"" value=""" & iPaymentServiceID & "00"" />"
	response.write vbcrlf & "<input type=""hidden"" name=""ITEM_NAME"" value=""" & sServiceName & """ />"

End Sub 


'------------------------------------------------------------------------------------------------------------
' void GetVerisignValues iPaymentServiceID, sServiceName 
'------------------------------------------------------------------------------------------------------------
Sub GetVerisignValues( ByVal iPaymentServiceID, ByVal sServiceName )

	' This is what we use to send to PayPal now, and I believe Point and Pay.
	response.write vbcrlf & "<form id=""frmpayment"" name=""frmpayment"" action=""" & Application("PAYMENTURL") & "/" & sorgVirtualSiteName & "/payment_processors/verisign2005/verisign_form.asp"" method=""post"">"
	response.write vbcrlf & "<input type=""hidden"" name=""PAYMENTID"" value=""" & iPaymentServiceID & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""ITEM_NUMBER"" value=""" & iPaymentServiceID & "00"" />"
	response.write vbcrlf & "<input type=""hidden"" name=""ITEM_NAME"" value=""" & sServiceName & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""orgid"" value=""" & iOrgId & """ />"

	' If they are logged in then add their userid to the hidden inputs to send to the secure pages
	If request.cookies("userid") <> "" AND request.cookies("userid") <> "-1" Then
		response.write vbcrlf & "<input type=""hidden"" name=""userid"" value=""" & request.cookies("userid") & """ />"
	End If 

End Sub 


'------------------------------------------------------------------------------------------------------------
' void GeteclinkPPDemoValues iPaymentServiceID, sPaymentService
'------------------------------------------------------------------------------------------------------------
Sub GeteclinkPPDemoValues( ByVal iPaymentServiceID, ByVal sPaymentService )

	' This is Peter's Demo page
	response.write vbcrlf & "<form name=""frmpayment"" id=""frmpayment"" action=""paypal_demo_pages/paypal_demo_page1.asp"" method=""get"">"
	
	' FORM HIDDEN PRODUCTS VALUES
	response.write vbcrlf & "<input type=""hidden"" name=""item_number"" value=""" &iPaymentServiceID & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""item_name"" value=""" & sPaymentService & """ />"

End Sub 


' -----------------------------------------------------------------------------------------------------------
' FUNCTION GETPAYPALVALUES(IID)
' -----------------------------------------------------------------------------------------------------------
'Function GetPayPalValues( ByVal iID )
'	Dim sSql, oRs
'
'	sSql = "SELECT * FROM egov_paypaloptions WHERE paymentserviceid = " & iID
'
'	Set oRs = Server.CreateObject("ADODB.Recordset")
'	oRs.Open sSql, Application("DSN"), 3, 1
'		
'	If Not oRs.EOF Then
'		' STATIC VALUES
'		'response.write "<form name=""frmpayment"" action=""https://www.sandbox.paypal.com/cgi-bin/webscr"" method=""post"">"
'		
'		' PAYPAL GATEWAY
'		response.write vbcrlf & "<form action=""transfer_payment.asp"" method=""post"">"
'
'		response.write vbcrlf & "<input type=""hidden"" name=""cmd"" value=""_xclick"">"
'		' May need for PayPal OrgId solution, but for now orgid is in table of paypal options 1/11/2006
'		' response.write "<input type=""hidden"" name=""iorgid"" value="""& iorgid & """>"
'		
'		Do While NOT oRs.EOF 
'			' DYNAMIC VALUES
'			response.write vbcrlf &  "<input type=""hidden"" name=""" & oRs("paypaloptionname") & """ value=""" & oRs("paypaloptionvalue") & """>" & vbcrlf
'			oRs.MoveNext
'		Loop
'
'	End If
'
'	oRs.Close 
'	Set oRs = Nothing
'
'End Function


'------------------------------------------------------------------------------------------------------------
' void GetPayPalValues iPaymentServiceID 
'------------------------------------------------------------------------------------------------------------
Sub GetPayPalValues( ByVal iPaymentServiceID )
	Dim sSql, oRs
	
	sSql = "SELECT * FROM egov_paypaloptions WHERE paymentserviceid = " & iPaymentServiceID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
		
	If Not oRs.EOF Then
		' PAYPAL GATEWAY
		' This goes to the standard PayPal Interface
		response.write vbcrlf & "<form name=""frmpayment"" action=""transfer_payment.asp"" method=""post"">"
		response.write vbcrlf & "<input type=""hidden"" name=""cmd"" value=""_xclick"" />"

		Do While Not oRs.EOF 
			' DYNAMIC VALUES
			response.write vbcrlf & "<input type=""hidden"" name=""" & oRs("paypaloptionname") & """ value=""" & oRs("paypaloptionvalue") & """ />"
			oRs.MoveNext
		Loop
	End If

	oRs.Close 
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------------------------------------
'Function JSsafe( strDB )
'------------------------------------------------------------------------------------------------------------
'Function JSsafe( ByVal strDB )
'
'	If Not VarType( strDB ) = vbString Then Jsafe = strDB : Exit Function
'
'	strDB = Replace( strDB, "'", "\'" )
'	strDB = Replace( strDB, chr(34), "\'" )
'	strDB = Replace( strDB, ";", "\;" )
'	strDB = Replace( strDB, "-", "\-" )
'	strDB = Replace( strDB, "(", "\(" )
'	strDB = Replace( strDB, ")", "\)" )
'	strDB = Replace( strDB, "/", "\/" )
'
'	JSsafe = strDB
'
'End Function


'------------------------------------------------------------------------------------------------------------
'void Custom_Payment_FormID iFormID
'------------------------------------------------------------------------------------------------------------
Sub Custom_Payment_FormID( ByVal iFormID )

	Select Case iFormID

		Case "21"
			' WASTEWATER FORM
			'response.write vbcrlf & vbcrlf & "<table id=""maintable"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
			response.write "<tr><td colspan=2><table>"
			response.write "<tr><td><b>Service Address :</b></td><td><table><tr><td><input maxlength=5 name=""custom_sa1"" style=""width: 50px"" type=text maxlength=5></td><td><input maxlength=20 name=""custom_sa2"" type=text maxlength=20></td><td><input maxlength=4 name=""custom_sa3"" style=""width: 50px"" type=text maxlength=4><font class=example>Ex: 123 Main Blvd</font></td></tr></table></td></tr>"
			response.write "<tr><td><b>Account Number :</b></td><td><table><tr><td><input maxlength=9 name=""custom_an1"" style=""width: 100px"" type=text maxlength=9></td><td><b> - </b></td><td><input maxlength=9 name=""custom_an2"" style=""width: 100px"" type=text maxlength=9><font class=example>Ex: 286-1384</font></td></tr></table></td></tr>"
			response.write "<tr><td><b>Payment Amount :</b></td><td><table><tr><td colspan=2><input id=""custom_paymentamount"" name=""custom_paymentamount"" type=""text"" /></td><td>&nbsp;</td></tr>"
			response.write "<tr><td></td><td></td></tr></table></td></tr>"
			response.write "</table></td></tr>"

			' FORM VALIDATION
			response.write "<input type=hidden name=""ef:custom_sa1-text/number/req"" value=""Street Number"">"
			response.write "<input type=hidden name=""ef:custom_sa2-text/req"" value=""Street Name"">"
			response.write "<input type=hidden name=""ef:custom_sa3-text"" value=""Suffix"">"
			response.write "<input type=hidden name=""ef:custom_an1-text/req/ninedigits"" value=""Account Number Part 1"">"
			response.write "<input type=hidden name=""ef:custom_an2-text/req/ninedigits"" value=""Account Number Part 2"">"
			response.write "<input type=hidden name=""ef:custom_paymentamount-text/req"" value=""Payment Amount"">"

		Case "22"
			'response.write vbcrlf & vbcrlf & "<table id=""maintable"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
			' CITY BUILDING PERMIT PAYMENTS
			response.write "<tr><td colspan=2>"
			DrawInputTable(10)	
			response.write "</td></tr>"

		Case "23"
			' SPECIAL ASSESSMENT PAYMENTS
			response.write "<tr><td colspan=2><table>"
			response.write "<tr><td><b>Parcel Number :</b></td><td><table><tr><td><input maxlength=3 name=""custom_pn1"" style=""width: 50px"" type=text></td><td><b> - </b></td><td><input maxlength=2 name=""custom_pn2"" type=text style=""width: 50px""></td><td><b> - </b></td><td><input maxlength=4 name=""custom_pn3"" style=""width: 50px"" type=text><font class=example>Ex: 123-12-124B</font></td></tr></table></td></tr>"
			response.write "<tr><td><b>Assessment Number :</b></td><td><table><tr><td><input  maxlength=30 name=""custom_an1"" style=""width: 150px"" type=text maxlength=9><!--</td><td><b> - </b></td><td><input maxlength=9 name=""custom_an2"" style=""width: 100px"" type=text maxlength=9>--><font class=example>Ex: 12-10986</font></td></tr></table></td></tr>"
			response.write "<tr><td><b>Assessment Name :</b></td><td><table><tr><td ><input name=""custom_assessmentname"" type=text ></td><td>&nbsp;</td></tr></table></td></tr>"
			response.write "<tr><td><b>Payment Amount :</b></td><td><table><tr><td ><input id=""custom_paymentamount"" name=""custom_paymentamount"" type=""text"" /></td><td>&nbsp;</td></tr>"
			response.write "<tr><td></td><td></td></tr></table></td></tr>"
			response.write "</table></td></tr>"

			' FORM VALIDATION
			response.write "<input type=hidden name=""ef:custom_pn1-text/req/threedigits"" value=""Parcel Number Part 1"">"
			response.write "<input type=hidden name=""ef:custom_pn2-text/req/twodigits"" value=""Parcel Number Part 2"">"
			response.write "<input type=hidden name=""ef:custom_pn3-text/ppn/req"" value=""Parcel Number Part 3"">"
			response.write "<input type=hidden name=""ef:custom_an1-text/req"" value=""Assessment Number Part 1"">"
			'response.write "<input type=hidden name=""ef:custom_an2-text/number/ninedigits"" value=""Assessment Number Part 2"">"
			response.write "<input type=hidden name=""ef:custom_assessmentname-text/req"" value=""Assessment Name"">"
			response.write "<input type=hidden name=""ef:custom_paymentamount-text/req"" value=""Payment Amount"">"

		Case "254"
			' WATER BILL FORM For Skokie IL (orgid = 131)
			response.write "<tr><td nowrap=""nowrap""><b>Account Number :</b></td><td nowrap=""nowrap""><input maxlength=""5"" name=""custom_accountno1"" size=""5"" type=""text"" />&nbsp;<b>/</b>&nbsp;<input maxlength=""5"" name=""custom_accountno2"" size=""5"" type=""text"" /></td><td><font class=""example"">&nbsp;Ex: 12345 / 67890</font></td></tr>"
			response.write "<tr><td nowrap=""nowrap""><b>Service Address :</b></td><td nowrap=""nowrap""><input name=""custom_serviceaddress"" style=""width: 300px"" maxlength=""200"" type=""text"" /></td><td><font class=""example"">&nbsp;Ex: 123 Main Blvd</font></td></tr>"
			response.write "<tr><td nowrap=""nowrap""><b>Payment Amount :</b></td><td><input id=""custom_paymentamount"" name=""custom_paymentamount"" type=""text"" size=""15"" maxlength=""15"" /></td><td>&nbsp;</td></tr>"
			response.write "<tr><td nowrap=""nowrap""><b>Phone Number :</b></td><td><input name=""custom_phone"" maxlength=""20"" size=""20"" type=""text"" /></td><td>&nbsp;</td></tr>"
			response.write "<tr><td colspan=""3"">"

			' FORM VALIDATION
			response.write "<input type=""hidden"" name=""ef:custom_accountno1-text/req/fivedigits"" value=""Account Number Part 1"" />"
			response.write "<input type=""hidden"" name=""ef:custom_accountno2-text/req/fivedigits"" value=""Account Number Part 2"" />"
			response.write "<input type=""hidden"" name=""ef:custom_serviceaddress-text/req"" value=""Service Address"" />"
			response.write "<input type=""hidden"" name=""ef:custom_paymentamount-text/req"" value=""Payment Amount"" />"
			response.write "<input type=""hidden"" name=""ef:custom_phone-text/req"" value=""Phone Number"" />"

			response.write "</td></tr>"

		Case "270"
			' Vehicle License Renewal Form for Skokie IL (orgid = 131) 
			' This turns on in June and off in September via a SQL Server Job
			response.write "<tr><td nowrap=""nowrap""><b>Control#</b></td><td colspan=""3"">&nbsp;</td></tr>"
			response.write "<tr><td><input maxlength=""8"" name=""custom_controlno"" size=""8"" type=""text"" /><td colspan=""3"">&nbsp;</td></tr>"

			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;""><b>First Name</b></td>"
			response.write "<td align=""left"" nowrap=""nowrap"" style=""height: 25px; padding-left:1em;""><b>Last Name</b></td>"
			response.write "<td colspan=""2"" style=""height: 25px;"">&nbsp;</td></tr>"
			response.write "<tr><td style=""height: 25px;""><input maxlength=""16"" name=""custom_firstname""  style=""width: 150px"" type=""text"" /></td>"
			response.write "<td style=""height: 25px; padding-left:1em;""><input maxlength=""16"" name=""custom_lastname""  style=""width: 150px"" type=""text"" />"
			response.write "<td colspan=""2"" style=""height: 25px;"">&nbsp;</td></tr>"

			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;""><b>Street Address</b></td>"
			response.write "<td nowrap=""nowrap"" style=""height: 25px; padding-left:1em;""><b>City</b></td>"
			response.write "<td nowrap=""nowrap"" style=""height: 25px; padding-left:1em;""><b>State</b></td>"
			response.write "<td nowrap=""nowrap"" style=""height: 25px; padding-left:1em;""><b>Zip</b></td></tr>"
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;""><input maxlength=""50"" name=""custom_address"" style=""width: 150px"" type=""text"" />"
			response.write "<td nowrap=""nowrap"" style=""height: 25px; padding-left:1em;""><input maxlength=""50"" name=""custom_city"" style=""width: 150px"" type=""text"" />"
			response.write "<td nowrap=""nowrap"" style=""width: 100px; height: 25px; padding-left:1em;""><input maxlength=""2"" id=""custom_state"" name=""custom_state"" size=""2"" type=""text"" onchange=""UpperCaseState()"" />"
			response.write "<td nowrap=""nowrap"" style=""width: 100px; height: 25px; padding-left:1em;""><input maxlength=""10"" name=""custom_zip"" size=""10"" type=""text"" /></tr>"

			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;""><b>Daytime Phone</b></td>"
			response.write "<td nowrap=""nowrap"" style=""height: 25px; padding-left:1em;""><b>Email</b></td>"
			response.write "<td colspan=""2"" style=""height: 25px;"">&nbsp;</td></tr>"
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;""><input maxlength=""12"" name=""custom_phone"" style=""width: 150px"" type=""text"" />"
			response.write "<td nowrap=""nowrap"" style=""height: 25px; padding-left:1em;""><input maxlength=""60"" name=""custom_email"" style=""width: 150px"" type=""text"" />"
			response.write "<td colspan=""2"" style=""height: 25px;"">&nbsp;</td></tr>"

			response.write "<tr><td nowrap=""nowrap"" style=""height: 30px;"" valign=""bottom""><b><u>VEHICLE LICENSE RENEWAL</u></b></td><td colspan=""3"">&nbsp;</td></tr>"

			response.write "<tr><td nowrap=""nowrap"" colspan=""2"">"

			' plate amount table here
			'Get the 2-digit year
			lcl_display_year = right(year(now),2)

			response.write "<table border=""0"" cellpadding=""2"" cellspacing=""2"" width=""100%"">"
			response.write "<tr>" & vbcrlf
			response.write "<td><b>Number from<br />office use box</b></td>"
			response.write "<td><b>Illinois<br />License Plate</b></td>"
			response.write "<td valign=""bottom""><b>Fee</b></td>" & vbcrlf
			response.write "</tr>" & vbcrlf

			response.write "<tr>" & vbcrlf
			response.write "<td nowrap=""nowrap"" style=""height: 25px;"">" & lcl_display_year & "-<input maxlength=""6"" name=""custom_usebox1"" size=""6"" type=""text"" /></td>" & vbcrlf
			response.write "<td style=""height: 25px;""><input maxlength=""14"" name=""custom_plate1"" style=""width: 100px"" type=""text"" /></td>" & vbcrlf
			response.write "<td style=""height: 25px;""><input maxlength=""5"" id=""custom_fee1"" name=""custom_fee1"" size=""5"" type=""text"" onblur=""clearMsg('custom_fee1');clearMsg('custom_total');return ValidateFee(this);"" /></td>" & vbcrlf
			response.write "</tr>"

			response.write "<tr>" & vbcrlf
			response.write "<td nowrap=""nowrap"" style=""height: 25px;"">" & lcl_display_year & "-<input maxlength=""6"" name=""custom_usebox2"" size=""6"" type=""text"" /></td>" & vbcrlf
			response.write "<td style=""height: 25px;""><input maxlength=""14"" name=""custom_plate2"" style=""width: 100px"" type=""text"" /></td>" & vbcrlf
			response.write "<td style=""height: 25px;""><input maxlength=""5"" id=""custom_fee2"" name=""custom_fee2"" size=""5"" type=""text"" onchange=""clearMsg('custom_fee2');clearMsg('custom_total');return ValidateFee(this);"" /></td>" & vbcrlf
			response.write "</tr>" & vbcrlf

			response.write "<tr>" & vbcrlf
			response.write "<td nowrap=""nowrap"" style=""height: 25px;"">" & lcl_display_year & "-<input maxlength=""6"" name=""custom_usebox3"" size=""6"" type=""text"" /></td>" & vbcrlf
			response.write "<td style=""height: 25px;""><input maxlength=""14"" name=""custom_plate3"" style=""width: 100px"" type=""text"" /></td>" & vbcrlf
			response.write "<td style=""height: 25px;""><input maxlength=""5"" id=""custom_fee3"" name=""custom_fee3"" size=""5"" type=""text"" onchange=""clearMsg('custom_fee3');clearMsg('custom_total');return ValidateFee(this);"" /></td>" & vbcrlf
			response.write "</tr>" & vbcrlf

			response.write "<tr>" & vbcrlf
			response.write "<td nowrap=""nowrap"" style=""height: 25px;"">" & lcl_display_year & "-<input maxlength=""6"" name=""custom_usebox4"" size=""6"" type=""text"" /></td>" & vbcrlf
			response.write "<td style=""height: 25px;""><input maxlength=""14"" name=""custom_plate4"" style=""width: 100px"" type=""text"" /></td>" & vbcrlf
			response.write "<td style=""height: 25px;""><input maxlength=""5"" id=""custom_fee4"" name=""custom_fee4"" size=""5"" type=""text"" onchange=""clearMsg('custom_fee4');clearMsg('custom_total');return ValidateFee(this);"" /></td>" & vbcrlf
			response.write "</tr>"

			'response.write "<table border=""0"" cellpadding=""2"" cellspacing=""2"" width=""100%"">"
			'response.write "<tr><td><b>Number from<br />office use box</b></td>"
			'response.write "<td><b>Illinois<br />License Plate</b></td>"
			'response.write "<td valign=""bottom""><b>Fee</b></td></tr>"
			'response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"">09-<input maxlength=""6"" name=""custom_usebox1"" size=""6"" type=""text"" /></td>"
			'response.write "<td style=""height: 25px;""><input maxlength=""14"" name=""custom_plate1"" style=""width: 100px"" type=""text"" /></td>"
			'response.write "<td style=""height: 25px;""><input maxlength=""5"" id=""custom_fee1"" name=""custom_fee1"" size=""5"" type=""text"" onblur=""clearMsg('custom_fee1');clearMsg('custom_total');return ValidateFee(this);"" /></td></tr>"

			'response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"">09-<input maxlength=""6"" name=""custom_usebox2"" size=""6"" type=""text"" /></td>"
			'response.write "<td style=""height: 25px;""><input maxlength=""14"" name=""custom_plate2"" style=""width: 100px"" type=""text"" /></td>"
			'response.write "<td style=""height: 25px;""><input maxlength=""5"" id=""custom_fee2"" name=""custom_fee2"" size=""5"" type=""text"" onchange=""clearMsg('custom_fee2');clearMsg('custom_total');return ValidateFee(this);"" /></td></tr>"

			'response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"">09-<input maxlength=""6"" name=""custom_usebox3"" size=""6"" type=""text"" /></td>"
			'response.write "<td style=""height: 25px;""><input maxlength=""14"" name=""custom_plate3"" style=""width: 100px"" type=""text"" /></td>"
			'response.write "<td style=""height: 25px;""><input maxlength=""5"" id=""custom_fee3"" name=""custom_fee3"" size=""5"" type=""text"" onchange=""clearMsg('custom_fee3');clearMsg('custom_total');return ValidateFee(this);"" /></td></tr>"

			'response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"">09-<input maxlength=""6"" name=""custom_usebox4"" size=""6"" type=""text"" /></td>"
			'response.write "<td style=""height: 25px;""><input maxlength=""14"" name=""custom_plate4"" style=""width: 100px"" type=""text"" /></td>"
			'response.write "<td style=""height: 25px;""><input maxlength=""5"" id=""custom_fee4"" name=""custom_fee4"" size=""5"" type=""text"" onchange=""clearMsg('custom_fee4');clearMsg('custom_total');return ValidateFee(this);"" /></td></tr>"

			' Total fee row
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"" colspan=""2"" align=""right""><b>TOTAL&nbsp;</b></td>"
			response.write "<td style=""height: 25px;""><input id=""custom_total"" name=""custom_total"" size=""8"" type=""text"" readonly=""readonly"" value=""0.00"" />"
			response.write "<input type=""hidden"" id=""custom_paymentamount"" name=""custom_paymentamount"" value=""0.00"" />"
			response.write "</td></tr>"
			response.write "</table>"
			response.write "</td><td colspan=""2"">&nbsp;</td></tr>"

			' Validation fields
			response.write "<tr><td colspan=""4"" style=""height: 25px;"">"
			response.write "<input type=""hidden"" name=""ef:custom_controlno-text/req"" value=""Control Number"" />"
			response.write "<input type=""hidden"" name=""ef:custom_firstname-text/req"" value=""First Name"" />"
			response.write "<input type=""hidden"" name=""ef:custom_lastname-text/req"" value=""Last Name"" />"
			response.write "<input type=""hidden"" name=""ef:custom_address-text/req"" value=""Street Address"" />"
			response.write "<input type=""hidden"" name=""ef:custom_city-text/req"" value=""City"" />"
			response.write "<input type=""hidden"" name=""ef:custom_state-text/req"" value=""State"" />"
			response.write "<input type=""hidden"" name=""ef:custom_zip-text/req"" value=""Zip"" />"
			response.write "<input type=""hidden"" name=""ef:custom_phone-text/req"" value=""Daytime Phone"" />"
			response.write "<input type=""hidden"" name=""ef:custom_email-text/req"" value=""Email"" />"
			response.write "<input type=""hidden"" name=""ef:custom_usebox1-text/req"" value=""Use Box"" />"
			response.write "<input type=""hidden"" name=""ef:custom_plate1-text/req"" value=""License Plate"" />"
			response.write "<input type=""hidden"" name=""ef:custom_fee1-text/req"" value=""Fee"" />"
			response.write "<input type=""hidden"" id=""skip_feesok"" name=""skip_feesok"" value=""true"" />"
			response.write "</td></tr>"

		Case "276"
			' Water/Sever/Refuse for West Carrollton, OH (orgid = 151)
			'<b><u>Instructions:</u></b><p><ol><li>Enter the Account Number.</li><li>Enter the Name on the Account.</li><li>Enter the Service Address.</li><li>Select the Payment Type.</li><li>Enter the Payment Amount.<br /><small><b><font color="red">(Warning! Entering the wrong payment amount may slow or prevent your payment from being successfully completed.)</font></b></small></li><li>Enter Additional Information/Comments</li></ol></p>
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Account Number :</b></td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;"" colspan=""2""><input maxlength=""10"" name=""custom_accountno"" size=""10"" type=""text"" /><b></td></tr>"
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Name on Account :</b></td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;"" colspan=""2""><input maxlength=""50"" name=""custom_accountname"" style=""width: 300px"" type=""text"" /></td></tr>"
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Service Address :</b></td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;""><input name=""custom_serviceaddress"" style=""width: 300px"" maxlength=""200"" type=""text"" /></td><td><font class=""example"">&nbsp;Ex: 123 Main Blvd</font></td></tr>"
			response.write "<tr><td nowrap=""nowrap"" valign=""top"" align=""right""><b>Payment Type :</b></td><td style=""height: 25px; padding-left:.5em;"" colspan=""2"">"
			response.write "<input value=""current"" name=""custom_paymenttype"" class=""formradio"" type=""radio"" checked=""checked"" /> Current<br />"
			response.write "<input value=""priorbalancedue"" name=""custom_paymenttype"" class=""formradio"" type=""radio"" /> Prior Balance Due<br />"
			response.write "<input value=""prepayments"" name=""custom_paymenttype"" class=""formradio"" type=""radio"" /> Prepayments"
			response.write "</td></tr>"
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Payment Amount :</b></td><td style=""height: 25px; padding-left:.5em;"" colspan=""2""><input id=""custom_billamount"" name=""custom_billamount"" type=""text"" size=""7"" maxlength=""7"" onchange=""clearMsg('custom_billamount');return ValidateWaterFee(this);"" /></td></tr>"
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Service Fee (2%) :</b></td><td style=""height: 25px; padding-left:.5em;""><span id=""custom_feeamount"">0.00</span></td><td>&nbsp;</td></tr>"
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Total Amount Due :</b></td><td style=""height: 25px; padding-left:.5em;"" colspan=""2""><span id=""custom_totalamount"">0.00</span>"
			response.write "<input type=""hidden"" id=""custom_paymentamount"" name=""custom_paymentamount"" value=""0.00"" />"
			response.write "<input type=""hidden"" id=""custom_servicefee"" name=""custom_servicefee"" value=""0.00"" />"
			response.write "</td></tr>"
			response.write "<tr><td nowrap=""nowrap"" colspan=""3"" style=""height: 25px;""><b>Additional Information/Comments :</b></td></tr>"
			response.write "<tr><td nowrap=""nowrap"" colspan=""3""style=""padding-left:2em;"">"
			response.write "<textarea name=""custom_comments"" class=""formtextarea""></textarea>"
			response.write "</td></tr>"
			
			' FORM VALIDATION
			response.write "<tr><td colspan=""3"" style=""height: 25px;"">&nbsp;"
			response.write "<input type=""hidden"" name=""ef:custom_accountno-text/req"" value=""Account Number"" />"
			response.write "<input type=""hidden"" name=""ef:custom_serviceaddress-text/req"" value=""Service Address"" />"
			response.write "<input type=""hidden"" name=""ef:custom_billamount-text/req"" value=""Payment Amount"" />"
			response.write "<input type=""hidden"" id=""skip_feesok"" name=""skip_feesok"" value=""true"" />"
			response.write "</td></tr>"

		Case "277"
			' Income Tax for West Carrollton, OH (orgid = 151)
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Account Number :</b></td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;"" colspan=""2""><input maxlength=""10"" name=""custom_accountno"" size=""10"" type=""text"" /><b></td></tr>"
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Name on Account :</b></td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;"" colspan=""2""><input maxlength=""50"" name=""custom_accountname"" style=""width: 300px"" type=""text"" /></td></tr>"
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Address :</b></td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;""><input name=""custom_serviceaddress"" style=""width: 300px"" maxlength=""200"" type=""text"" /></td><td><font class=""example"">&nbsp;Ex: 123 Main Blvd</font></td></tr>"
			response.write "<tr><td nowrap=""nowrap"" valign=""top"" align=""right""><b>Income Type :</b></td><td style=""height: 25px; padding-left:.5em;"" colspan=""2"">"
			response.write "<input value=""individual"" name=""custom_incometype"" class=""formradio"" type=""radio"" checked=""checked"" /> Individual<br />"
			response.write "<input value=""business"" name=""custom_incometype"" class=""formradio"" type=""radio"" /> Business<br />"
			response.write "<input value=""withholding"" name=""custom_incometype"" class=""formradio"" type=""radio"" /> Withholding"
			response.write "</td></tr>"
			response.write "<tr><td nowrap=""nowrap"" colspan=""3"" style=""height: 25px;""><b>For withholding payments, enter the month or quarter :</b></td></tr>"
			response.write "<tr><td style=""height: 25px;"">&nbsp;</td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;"" colspan=""2""><input maxlength=""50"" name=""custom_withholdingperiod"" style=""width: 300px"" type=""text"" /></td></tr>"
			response.write "<tr><td nowrap=""nowrap"" valign=""top"" align=""right""><b>Payment Type :</b></td><td style=""height: 25px; padding-left:.5em;"" colspan=""2"">"
			response.write "<input value=""current"" name=""custom_paymenttype"" class=""formradio"" type=""radio"" checked=""checked"" /> Current<br />"
			response.write "<input value=""prior"" name=""custom_paymenttype"" class=""formradio"" type=""radio"" /> Prior<br />"
			response.write "<input value=""estimate"" name=""custom_paymenttype"" class=""formradio"" type=""radio"" /> Estimate"
			response.write "</td></tr>"
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Tax Year :</b></td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;"" colspan=""2""><input maxlength=""4"" name=""custom_taxyear"" size=""4"" type=""text"" /><b></td></tr>"
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Payment Amount :</b></td><td style=""height: 25px; padding-left:.5em;"" colspan=""2""><input id=""custom_taxamount"" name=""custom_taxamount"" type=""text"" size=""8"" maxlength=""8"" onchange=""clearMsg('custom_taxamount');return ValidateTaxAmount(this);"" /></td></tr>"
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Service Fee (2%) :</b></td><td style=""height: 25px; padding-left:.5em;""><span id=""custom_feeamount"">0.00</span></td><td>&nbsp;</td></tr>"
			response.write "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Total Amount Due :</b></td><td style=""height: 25px; padding-left:.5em;"" colspan=""2""><span id=""custom_totalamount"">0.00</span>"
			response.write "<input type=""hidden"" id=""custom_paymentamount"" name=""custom_paymentamount"" value=""0.00"" />"
			response.write "<input type=""hidden"" id=""custom_servicefee"" name=""custom_servicefee"" value=""0.00"" />"
			response.write "</td></tr>"
			response.write "<tr><td nowrap=""nowrap"" colspan=""3"" style=""height: 25px;""><b>Additional Information such as further describing prior balance due :</b></td></tr>"
			response.write "<tr><td nowrap=""nowrap"" colspan=""3""style=""padding-left:2em;"">"
			response.write "<textarea name=""custom_comments"" class=""formtextarea""></textarea>"
			response.write "</td></tr>"

			' FORM VALIDATION
			response.write "<tr><td colspan=""3"" style=""height: 25px;"">&nbsp;"
			response.write "<input type=""hidden"" name=""ef:custom_accountno-text/req"" value=""Account Number"" />"
			response.write "<input type=""hidden"" name=""ef:custom_taxamount-text/req"" value=""Payment Amount"" />"
			response.write "<input type=""hidden"" id=""skip_feesok"" name=""skip_feesok"" value=""true"" />"
			response.write "</td></tr>"


		Case Else
			' Unknown Payment Service

	End Select

End Sub


'------------------------------------------------------------------------------------------------------------
' void ShowWestCarrolltonCourtPaymentForm
'------------------------------------------------------------------------------------------------------------
Sub ShowWestCarrolltonCourtPaymentForm( )

	' Court Payments for West Carrollton, OH (orgid = 151)
	
	response.write vbcrlf & "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Defendant's Name :</b></td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;"" colspan=""2""><input maxlength=""50"" name=""custom_defendantname"" style=""width: 300px"" type=""text"" /></td></tr>"
	
	response.write vbcrlf & "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Driver's License Number :</b></td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;"" colspan=""2""><input maxlength=""50"" name=""custom_driverslicense"" style=""width: 300px"" type=""text"" /></td></tr>"
	response.write vbcrlf & "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Phone Number :</b></td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;"" colspan=""2""><input maxlength=""12"" name=""custom_phonenumber""  size=""12"" type=""text"" /></td></tr>"

	response.write vbcrlf & "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Address :</b></td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;""><input name=""custom_address"" style=""width: 300px"" maxlength=""200"" type=""text"" /></td><td><font class=""example"">Ex: 123 Main Blvd</font></td></tr>"

	response.write vbcrlf & "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Ticket Number :</b></td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;"" colspan=""2""><input maxlength=""6"" name=""custom_ticketnumber"" size=""6"" type=""text"" /><b></td></tr>"
	response.write vbcrlf & "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Case Number :</b><br /><font class=""example"">(if available)</font></td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;"" colspan=""2""><input maxlength=""10"" name=""custom_casenumber"" size=""10"" type=""text"" /><b></td></tr>"
	response.write vbcrlf & "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Description of Violation :</b></td><td nowrap=""nowrap"" style=""height: 25px; padding-left:.5em;"" colspan=""2""><input maxlength=""50"" name=""custom_violation"" style=""width: 300px"" type=""text"" /></td></tr>"

	response.write vbcrlf & "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Ticket Amount :</b></td><td style=""height: 25px; padding-left:.5em;"" colspan=""2""><input id=""custom_ticketamount"" name=""custom_ticketamount"" type=""text"" size=""8"" maxlength=""8"" onchange=""clearMsg('custom_ticketamount');return ValidateTicketAmount(this);"" /></td></tr>"
	response.write vbcrlf & "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Service Fee (2%) :</b></td><td style=""height: 25px; padding-left:.5em;""><span id=""custom_feeamount"">0.00</span></td><td>&nbsp;</td></tr>"
	response.write vbcrlf & "<tr><td nowrap=""nowrap"" style=""height: 25px;"" align=""right""><b>Total Amount Due :</b></td><td style=""height: 25px; padding-left:.5em;"" colspan=""2""><span id=""custom_totalamount"">0.00</span>"
	response.write vbcrlf & "<input type=""hidden"" id=""custom_paymentamount"" name=""custom_paymentamount"" value=""0.00"" />"
	response.write vbcrlf & "<input type=""hidden"" id=""custom_servicefee"" name=""custom_servicefee"" value=""0.00"" />"
	response.write vbcrlf & "</td></tr>"
	response.write vbcrlf & "<tr><td nowrap=""nowrap"" colspan=""3"" style=""height: 25px;""><b>Additional Information/Comments :</b></td></tr>"
	response.write vbcrlf & "<tr><td nowrap=""nowrap"" colspan=""3""style=""padding-left:2em;"">"
	response.write vbcrlf & "<textarea name=""custom_comments"" class=""formtextarea""></textarea>"
	response.write vbcrlf & "</td></tr>"

	' FORM VALIDATION
	response.write vbcrlf & "<tr><td colspan=""3"" style=""height: 25px;"">&nbsp;"
	response.write vbcrlf & "<input type=""hidden"" name=""ef:custom_defendantname-text/req"" value=""Defendant's Name"" />"
	response.write vbcrlf & "<input type=""hidden"" name=""ef:custom_driverslicense-text/req"" value=""Driver's License Number"" />"
	response.write vbcrlf & "<input type=""hidden"" name=""ef:custom_phonenumber-text/req"" value=""Phone Number"" />"
	response.write vbcrlf & "<input type=""hidden"" name=""ef:custom_address-text/req"" value=""Address"" />"
	response.write vbcrlf & "<input type=""hidden"" name=""ef:custom_ticketnumber-text/req"" value=""Ticket Number"" />"
	response.write vbcrlf & "<input type=""hidden"" name=""ef:custom_violation-text/req"" value=""Description of Violation"" />"
	response.write vbcrlf & "<input type=""hidden"" name=""ef:custom_ticketamount-text/req"" value=""Ticket Amount"" />"
	response.write vbcrlf & "<input type=""hidden"" id=""skip_feesok"" name=""skip_feesok"" value=""true"" />"
	response.write vbcrlf & "</td></tr>"

End Sub 


'------------------------------------------------------------------------------------------------------------
'void DrawInputTable( iRows )
'------------------------------------------------------------------------------------------------------------
Sub DrawInputTable( ByVal iRows )
	
	response.write "<table cellspacing=0 Cellpadding=5 style=""border-top:solid 1px #000000;border-left:solid 1px #000000;border-right:solid 1px #000000""> "
	
	' HEADER ROW WITH CAPTIONS
	response.write "<tr>"
	response.write "<td style=""BACKGROUND-COLOR:#2E1999; COLOR: #FFFFFF; font-family: verdana,sans-serif; font-size: 10px; font-weight:bold;border-bottom:solid 1px #000000;border-right:solid 1px #000000;"" >Permit Type<br><FONT SIZE=-2>(B/Z/ROW/ENC/FP/ENG)</FONT></TD>"
	response.write "<td style=""BACKGROUND-COLOR:#2E1999; COLOR: #FFFFFF; font-family: verdana,sans-serif; font-size: 10px; font-weight:bold;border-bottom:solid 1px #000000;border-right:solid 1px #000000;"" >Permit Number</TD>"
	response.write "<td style=""BACKGROUND-COLOR:#2E1999; COLOR: #FFFFFF; font-family: verdana,sans-serif; font-size: 10px; font-weight:bold;border-bottom:solid 1px #000000;border-right:solid 1px #000000;"" >Project Address</TD>"
	response.write "<td style=""BACKGROUND-COLOR:#2E1999; COLOR: #FFFFFF; font-family: verdana,sans-serif; font-size: 10px; font-weight:bold;border-bottom:solid 1px #000000;"" >Permit Amount<br><FONT SIZE=-2>(Dont enter , $) </FONT></TD>"
	response.write "</tr>"
	
	' LOOP AND DRAW DATA INPUT ROWS
	For iRow = 1 to iRows 
		response.write "<tr>"
		
		' DRAW PERMIT TYPE
		response.write "<td ALIGN=CENTER style=""border-right:solid 1px #000000;border-bottom:solid 1px #000000;"" ><b>" & iRow & ".</b> <SELECT NAME=""CUSTOM_PT" & iRow & """  style=""width:100px;"">"
		response.write "<option value=""NONE"">Select...</option>"
		response.write "<option value=""B"">B</option>"
		response.write "<option value=""Z"">Z</option>"
		response.write "<option value=""ROW"">ROW</option>"
		response.write "<option value=""ENC"">ENC</option>"
		response.write "<option value=""FP"">FP</option>"
		response.write "<option value=""ENG"">ENG</option>"
		response.write "</SELECT></TD>"
		' DRAW PERMIT NUMBER INPUT
		response.write "<td ALIGN=CENTER style=""border-bottom:solid 1px #000000;border-right:solid 1px #000000;"" ><INPUT MAXLENGTH=7 NAME=""CUSTOM_PN" & iRow & """ TYPE=TEXT style=""width:75px;""></TD>"
		' DRAW PROJECT ADDRESS INPUT
		response.write "<td style=""border-bottom:solid 1px #000000;border-right:solid 1px #000000;"" ><INPUT MAXLENGTH=20 NAME=""CUSTOM_PA" & iRow & """ TYPE=TEXT style=""width:250px;""></TD>"
		' DRAW PERMIT AMOUNT
		response.write "<td style=""border-bottom:solid 1px #000000;"" >$<INPUT NAME=""PA" & iRow & """  TYPE=TEXT style=""width:100px;""></TD>" & vbcrlf


		response.write "</tr>"
	Next

	' FOOTER ROW WITH TOTAL
	response.write "<tr><td colspan=3 align=right style=""border-bottom:solid 1px #000000;"" ><b>Total Amount Paid:</b></TD><td colspan=3 align=right style=""border-bottom:solid 1px #000000;"" >$<input id=""custom_paymentamount"" name=""custom_paymentamount"" type=""text"" style=""width:100px;"" /></TD></tr>"

	response.write "</table>"

	' WARNING
	response.write "<tr><td colspan=5  ><b><font color=red>*</font><i>You must select Permit Type and complete full row or permit for that row will be ignored.</i></b></TD></tr>"


	' DRAW JAVASCRIPT FUNCTION TO COMPUTE TOTAL
	response.write vbcrlf & vbcrlf & "<script language=""Javascript"">" & vbcrlf 

	response.write "function compute(){" & vbcrlf 
	response.write "// List Variables" & vbcrlf 
	response.write "var iCount" & vbcrlf 
	response.write "var strInputName" & vbcrlf 
	response.write "var iDraftTotal" & vbcrlf 
	response.write "iDraftTotal = 0" & vbcrlf 

	response.write "// Remove check numbers from calculations" & vbcrlf 
	response.write "for (iCount=1; iCount < " & (iRows * 4)  & "; iCount++)" & vbcrlf 
	response.write "  {strInputName = document.frmpayment.elements[iCount].name;" & vbcrlf 
	response.write "   "
	response.write "   // Removes check number from totaling" & vbcrlf 
	response.write "  if (strInputName.charAt(0) == 'P'){" & vbcrlf 
	response.write "   "    & vbcrlf 
	response.write "	   // If field is empty puts a zero in the value field " & vbcrlf 
	response.write "	   if (document.frmpayment.elements[iCount].value == """"){" & vbcrlf 
	response.write "	   document.frmpayment.elements[iCount].value= ""0.00"";}" & vbcrlf 
	response.write "   " & vbcrlf 
	response.write "	   // Totals Register values" & vbcrlf 
	response.write "	   iDraftTotal = iDraftTotal + eval(document.frmpayment.elements[iCount].value);" & vbcrlf 
				  
	response.write "   }" & vbcrlf 
	response.write "}" & vbcrlf 
	response.write "document.frmpayment.custom_paymentamount.value = iDraftTotal;"

	response.write "}" & vbcrlf 
%>

	// VALIDATE INPUT
	function vcheck(){
		var blnRowValid;

		for (iCount=1; iCount <= <%=iRows%>; iCount++){ 
			iValue = eval('document.frmpayment.CUSTOM_PT' + iCount +'.selectedIndex');
			sValue = eval('document.frmpayment.CUSTOM_PT' + iCount +'.options[iValue].value');

			//IF VALUES ENTERED VALIDATE THE ROW
			if (sValue != 'NONE'){
				// VALUES
				sPermitNumber = eval('document.frmpayment.CUSTOM_PN' + iCount +'.value');
				sPermitAddress = eval('document.frmpayment.CUSTOM_PA' + iCount +'.value');
				sPermitPrice = eval('document.frmpayment.PA' + iCount +'.value');
				blnRowValid = true;

				//CHECK PERMIT NUMBER
				var regexpseven = new RegExp(/^\d{1,7}$/); // FIND ANY 1-7 DIGIT NUMBER
				if (sPermitNumber.match(regexpseven)){
				}
				else
				{
					blnRowValid = false;
				}

				//CHECK PERMIT ADDRESS
				if (sPermitAddress != ''){
				}
				else
				{
					blnRowValid = false;
				}

				//CHECK PERMIT PRICE
				var regexpseven = new RegExp(/^[0-9.]+$/); // FIND ANY DIGITS
				if (sPermitPrice.match(regexpseven)){
				}
				else
				{
					blnRowValid = false;
				}

				// DISPLAY MESSAGE IF ANY FIELDS WERE INVALID
				if (blnRowValid){
				}
				else
				{alert('The permit on line ' + iCount + ' has a missing or incorrect permit number, address, or amount!');}

			}
		}


			// CHECK TOTAL AMOUNT AGAINST ENTERED AMOUNT
			var iCount 
			var strInputName 
			var iDraftTotal 
			iDraftTotal = 0 

			// LOOP THRU ROWS TO TOTAL AMOUNTS
			for (iCount=1; iCount < <%=(68)%>; iCount++){
				strInputName = document.frmpayment.elements[iCount].name; 
			   
				// Removes check number from totaling 
				if (strInputName.charAt(0) == 'P'){ 
				   // If field is empty puts a zero in the value field  
				   if (document.frmpayment.elements[iCount].value == ''){ 
						document.frmpayment.elements[iCount].value= '0.00';} 
				
				   // Totals Register values 
				   iDraftTotal = iDraftTotal + eval(document.frmpayment.elements[iCount].value); 
				   iDraftTotal = Math.round(iDraftTotal*100)/100
				} 
			} 
			
			if (document.frmpayment.custom_paymentamount.value != iDraftTotal){
				alert('The total amount entered does not match the total calculated amount \(' + iDraftTotal + '\) please check your permit entries!');
			}

			else {
			
				if (blnRowValid){
				document.frmpayment.submit();}

			}
			
	}

	<%
	response.write "</script>"

End Sub 


%>






