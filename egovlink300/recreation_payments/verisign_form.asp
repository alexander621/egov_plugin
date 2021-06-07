<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<% Response.Expires = -1000 %>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="../recreation/facility_global_functions.asp" //-->
<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: verisign_tform.asp
' AUTHOR: John Stullenberger
' CREATED: 03/3/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page takes payment information for rentals.
'
' MODIFICATION HISTORY
' 1.0	03/03/2005	John Stullenberger - Initial version
' 2.0	06/23/2010	Steve Loar - Split name field into first and last 
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
response.Expires = 60
response.Expiresabsolute = Now() - 1
response.AddHeader "pragma","no-store"
response.AddHeader "cache-control","private"
response.CacheControl = "no-store" 'HTTP prevent back button

Dim sError, iReservationTempId, sTotalAmount, iItemCount, bHasPaymentFee, dFeeAmount, sErrorMsg, dTotalCharges
Dim sProcessingRoute

dFeeAmount = CDbl(0.00)

' check if payment gateway needs a fee check for this page
If PaymentGatewayRequiresFeeCheck( iOrgId ) Then
	bHasPaymentFee = True 
	
	sTotalAmount = CDbl(request("amount"))
	
	sProcessingRoute = GetProcessingRoute()		' In ../include_top_functions.asp
	If LCase(sProcessingRoute) = "pointandpay" Then 
		' Fetch the fee for the amount to be charged.
		If Not GetPNPFee( sTotalAmount, dFeeAmount, sErrorMsg ) Then		' in ../includes/common.asp
			'If not successful, store the error, then take them to a page to display the error message.
			iGatewayErrorId = StoreGatewayError( iOrgId, sProcessingRoute, "feecheck", sErrorMsg, FormatNumber(sTotalAmount,2,,,0) )		' in ../includes/common.asp
			response.redirect Application("CARTURL") & "/" & sorgVirtualSiteName & "/payment_processors/processing_failure.asp?ge=" & iGatewayErrorId
		End If 
	End If 
elseif iOrgId = 228 then
	bHasPaymentFee = True
	sTotalAmount = CDbl(request("amount"))
	dFeeAmount = sTotalAmount * .035
Else
	bHasPaymentFee = False 
End If 


%>
<html>
<head>
  	<meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<title>E-Gov Services <%=sOrgName%> - Payment Form</title>

<script src='https://www.google.com/recaptcha/api.js'></script>
	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="verisign.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script language="javascript" src="../scripts/modules.js"></script>

  	<script type="text/javascript" src="../scripts/jquery-1.9.1.min.js"></script>

	<script language="javascript">
	<!--
		function openWin2(url, name) 
		{
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
		}

		<!--
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
					var tempobj=theform.elements[i];
					if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
					//disable em
					tempobj.disabled=true;
				}
			}
		} 

		function processPayment()
		{
			$("#sjname").val($("#firstname").val() + ' ' + $("#lastname").val());
			// disable the pay button
			document.getElementById("COMPLETE_PAYMENT").disabled = true;
			// submit the form
			document.Fullorder.submit();
		}

		// set focus on the first field when the page loads
		$('document').ready(function(){
			$("#firstname").focus();
		});

	function iframecheck()
	{
 		if (window.top!=window.self)
		{
 			document.body.classList.add("iframeformat") // In a Frame or IFrame
 			//var element = document.getElementById("egovhead");
 			//element.classList.add("iframeformat");
		}
	}

	//-->
	</script>

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

          <!--BEGIN: BUILD PAYMENT FORM-->
          <form name="Fullorder" action="process_payment.asp" method="post">

          <!--BEGIN:  DISPLAY PAYMENT FORM-->
<% 
			fnDisplayForm request("ITEM_NAME"), GetDetails(request("iPAYMENT_MODULE"),iorgid), request("amount"), dFeeAmount, bHasPaymentFee
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
' void fnDisplayForm sPaymentName, sDetails, sAmount, dFeeAmount, bHasPaymentFee
'--------------------------------------------------------------------------------------------------
Sub fnDisplayForm( ByVal sPaymentName, ByVal sDetails, ByVal sAmount, ByVal dFeeAmount, ByVal bHasPaymentFee )

	'BEGIN: GET PAYMENT INFORMATION
 	Dim sOrderString, sSerialNumber, sOrderNumber, sName, sEmail, sAddress, sCity, sState, sZip
 	Dim sItemNumber, sItemDescription, sItemQuantity, blnTaxTable, curItemCost, sFirstName, sLastName
	Dim dTotalCharges

 	sName            = ""
 	sEmail           = ""
 	sAddress         = ""
 	sCity            = ""
 	sState           = ""
 	sZip             = ""
 	sItemNumber      = request("ITEM_NUMBER")
 	sItemDescription = request("ITEM_NAME")
 	sItemQuantity    = 1
 	blnTaxTable      = "N"  '(Y) USE TABLE | (N) DON'T USE TABLE
 	sOrderString     = sItemNumber & "~" & sItemDescription & "~" & sAmount & "~" & sItemQuantity & "~" & blnTaxTable & "||"  
 	sSerialNumber    = GetSerialNumber(5) 
 	'sOrderNumber = "ecC" & iOrgID & "O" & CreatePayment() ' STORE VALUES IN DATABASE

	'IF DEMO ORGANIZATION SHOW TEST VALUES
	If iorgid = 37 Then 
		'sName    = "John Stullenberger"
		'sEmail   = "jstullenberger@eclink.com"
		'sAddress = "4303 Hamilton Avenue"
		'sCity    = "Cincinnati"
		'sState   = "OH"
		'sZip     = "45223"
		sOrderNum      = "43036814030"
		sCreditCardNum = "5555555555554444"
		sExpMonth      = "09"
		sExpYear       = "2013"
		sCVSCode       = ""
		sPhone         = "5136814030"
	End If 

	'FORM HEADING	
 	response.write "<blockquote class=""paymentform"">" & vbcrlf
 	response.write "<font class=""formtitle"">" & sPaymentName & " - Payment Information</font>" & vbcrlf
 	response.write "<div class=""group"" id=""paymentform"">" & vbcrlf

 	GetUserInfo CLng(request("iuserid")), sName, sEmail, sAddress, sCity, sState, sZip, sFirstName, sLastName

	'GET CUSTOM FIELDS
	%>
	<input type="hidden" name="orderstring" value="<%=sOrderString%>" />
	<input type="hidden" name="serialnumber" value="<%=sSerialNumber%>" />
	<input type="hidden" name="ordernumber" value="<%=sOrderNumber%>" />
	<input type="hidden" name="itemnumber" value="<%=sItemNumber%>" />
	<input type="hidden" name="paymentname" value="<%=sItemDescription%>" />
	<input type="hidden" name="paymenttype" value="<%=request("paymenttype")%>" />
	<input type="hidden" name="paymentlocation" value="<%=request("paymentlocation")%>" />
	<input type="hidden" name="details" value="<%=sDetails%>" />
	<input type="hidden" name="display_membershipname" value="<%=request("display_membershipname")%>" />
	<!--END: GET PAYMENT INFORMATION-->

	<div align="center" style="width:500px;">

	<!--BEGIN: PAYMENT DETAILS-->
	<fieldset>
  	<legend><strong>Payment Service</strong></legend>
	<table border="0" cellpadding="2" cellspacing="0" width="100%">
  	<!--BEGIN: PERSONAL INFORMATION-->
		 <tr>
<!--       <td><img hspace="20" src="images/verisign.gif"></td> -->
   			<td align="left">
	           <strong><%="(" & sItemNumber & ") - " & sItemDescription %><br /><br />
        			<u>Details</u></strong><br />
          	<%	response.write sDetails	%>
    		</td>
  	</tr>
	</table>
	<br />
	</fieldset>
	<!--END: PAYMENT DETAILS-->

<%
	' Charge Amount
	response.write vbcrlf & "<fieldset>"
	response.write vbcrlf & "<legend><strong>Charges&nbsp;</strong></legend>"
	response.write vbcrlf & "<table border=""0"" cellpadding=""2"" cellspacing=""0"" id=""chargeamounts"" align=""left"">"

	response.write vbcrlf & "<tr align=""left"">"
	response.write "<td align=""right"" class=""labelcol"" nowrap=""nowrap"">Purchase Amount:</td>"
	if iorgid = "228" then
		response.write "<td align=""right"">" & FormatNumber(sAmount,2) & "<input type=""hidden"" name=""transactionamount"" value=""" & FormatNumber(CDbl(sAmount) + CDbl(dFeeAmount),2,,,0) & """ /></td>"
	else
		response.write "<td align=""right"">" & FormatNumber(sAmount,2) & "<input type=""hidden"" name=""transactionamount"" value=""" & FormatNumber(sAmount,2,,,0) & """ /></td>"
	end if
	response.write "</tr>"

	If bHasPaymentFee Then
		' Show the fee
		response.write vbcrlf & "<tr align=""left"">"
		response.write "<td align=""right"" class=""labelcol"" nowrap=""nowrap"">Processing Fee:</td>"
		response.write "<td align=""right"">" & FormatNumber(dFeeAmount,2) & "</td>"
		response.write "</tr>"

		dTotalCharges = CDbl(dFeeAmount) + CDbl(sAmount)

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

	<p align="left">Please enter your billing information as it appears on your credit card statement, then click the <strong>Process Payment</strong> button.</p>

	<fieldset>
  	<legend><strong>Personal Information</strong></legend>
	<table border="0" cellpadding="2" cellspacing="0">
  	<!--BEGIN: PERSONAL INFORMATION-->
<%
'		 <tr>
'		     <td align="right"><strong>Name:</strong></td>
'  		  <td align="left"><font face="Verdana" size="2"><input type="text" maxlength="30" size="30" name="sjname" value="=sName" /></font></td>
' 		</tr>
'		<tr>
%>
		     <td align="right"><strong>First Name:</strong></td>
   		  <td align="left"><font face="Verdana" size="2"><input type="text" maxlength="30" size="30" id="firstname" name="firstname" value="<%=sFirstName%>" /></font>
							<input type="hidden" id="sjname" name="sjname" value="<%=sName%>" />
		  </td>
 		</tr>
		<tr>
		     <td align="right"><strong>Last Name:</strong></td>
   		  <td align="left"><font face="Verdana" size="2"><input type="text" maxlength="30" size="30" id="lastname" name="lastname" value="<%=sLastName%>" /></font></td>
 		</tr>
	 	<tr>
		     <td align="right"><strong>E-mail:</strong></td>
   		  <td align="left"><font face="Verdana" size="2"><input type="text" value="<%=sEmail%>" maxlength="50" size="50" name="email" /></font></td>
		 </tr>
		<tr>
  		  <td align="right"><strong>Address:</strong></td>
		    <td align="left"><font face="Verdana" size="2"><input type="text" value="<%=sAddress%>" maxlength="50" size="50" name="streetaddress" /></font></td>
		</tr>
		<tr>
  		  <td align="right"><strong>City:</strong></td>
		    <td align="left"><font face="Verdana" size="2"><input type="text" value="<%=sCity%>" maxlength="20" size="20" name="city" /></font></td>
		</tr>
		<tr>
  		  <td align="right"><strong>State:</strong></td>
		    <td align="left">
          <font face="Verdana" size="2">
			<select name="state" size="1">
<%				ShowStatePicks sDefaultState, sState	%>
			</select>
          </font>
      </td>
	 </tr>
		<tr>
  		  <td align="right"><strong>Zip:</strong></td>
		    <td align="left"><font face="Verdana" size="2"><input type="text" value="<%=sZip%>"  maxlength="15" size="15" name="zipcode" /></font></td>
		</tr>
	</table>
	<br />
	</fieldset>
	<!--END: PERSONAL INFORMATION-->
			
	<br />
			
	<!--BEGIN: CREDIT CARD INFORMATION-->
	<fieldset>
			<legend><strong>Credit Card Information</strong></legend>
	<table border="0" cellpadding="2" cellspacing="0">
	  <!-- Put a drop down of accepted card types here -->
		<tr>
		   	<td align="right"><strong>Credit Card Type:</strong></td>
   			<td align="left">
<%						ShowCreditCardPicks		' In include_top_functions.asp		%>
   			</td>
		</tr>
		<tr>
		    <td align="right"><strong>Credit Card Number:</strong></td>
		    <td align="left"><font face="Verdana" size="2"><input type="text" value="<%=sCreditCardNum%>" maxlength="22" size="30" name="accountnumber" /></font></td>
		</tr>
		<tr>
  		  <td align="right"><strong>Expiration Month:</strong></td>
		    <td align="left">
       			<select name="month">
       			<%
       			 'DRAW MONTH SELECTION
         			for i=1 to 12
             			if i < 10 then
             				sTemp = "0" & i
           				else
             				sTemp = i
           				end if
           				response.write vbcrlf & "<option value=""" & sTemp & """>" & sTemp & "</option>"
          			next
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
        			'DRAW YEAR SELECTION
         			sTemp = Year(Now())
         			for i = 1 to 10
        			   'GENERATE SELECTED VALUE
           				if sTemp = Year(Now())+ 1 then
             				sSelected = " selected=""selected"" "
						else
              				sSelected = ""
						end if
            			response.write vbcrlf & "<option value=""" & Right(sTemp,2) & """" & sSelected & ">" & sTemp & "</option>"
           				sTemp = sTemp + 1
         			next
       			%>
      		  </select>
  		  </td>
  		  <td></td>
		    <td></td>
		</tr>
<%
		If OrgHasFeature( iOrgId, "display cvv" ) Then 
			 response.write vbcrlf & "<tr>"
			 response.write "<td align=""right""><strong>CVV Code:</strong></td>"
			 response.write "<td align=""left""><input type=""text"" name=""cvv2"" size=""4"" maxlength=""4"" value="""" /></td>"
			 response.write "<td colspan=""2""></td>" 
			 response.write "<td></td>" 
			 response.write "<td></td>" 
			 response.write "</tr>"
		End If 
%>

<!--		
		<tr>
		    <td align="right"><strong>Amount: </strong></td>
		    <td align="left"><font face="Verdana" size="2"><%'=formatcurrency(sAmount,2)%><input type="hidden" style="background-color:#e0e0e0;" value="<%'=formatnumber(sAmount,2)%>" maxlength="15" size="15" name="transactionamount"></font></td>
  		  <td colspan="2"></td>
		    <td></td>
		    <td></td>
		</tr>
-->

	</table>
<%	response.write "<p align=""center"">Do not use dashes or spaces when entering credit card information.</p>" %>
	</fieldset>
	<!--END: CREDIT CARD INFORMATION-->

	<div align="left"><small><strong><font color="red">*</font><i>All Fields Required</i></strong></small></div>

<div class="g-recaptcha" data-sitekey="6LcVxxwUAAAAAEYHUr3XZt3fghgcbZOXS6PZflD-"></div>

	<p align="left" class="smallnote">
		<font style="font-weight:bold;color:red">Press PROCESS PAYMENT button only once and please wait for the authorization page to be 
			displayed to prevent double billing.  Be patient, it may take up to 2 minutes to process your transaction.</font>
	</p>

	<!--BEGIN: PROCESS BUTTONS-->
	<table border="0" cellpadding="2" cellspacing="0">
		<tr>
		    <td align="center">
			       <input style="width:200px;" class="skipjackbtn" type="button" id="COMPLETE_PAYMENT" name="COMPLETE_PAYMENT" value="PROCESS PAYMENT" onclick="processPayment();" />
       			<!--<input STYLE="WIDTH:100PX;" class=skipjackbtn type="button" value="Cancel" name="buttonRefresh" onclick="process_cancel();" >-->
    		</td>
    		<td>&nbsp;</td>
		</tr>
	</table>
	<!--END: PROCESS BUTTONS-->

	<p class="smallnote">NOTE: Your IP address [<%=request.servervariables("REMOTE_ADDR")%>] has been logged with this transaction.<br /><br />
		<!--Do you have questions?<br />
		 Contact Customer Service: <a href="mailto:support@egovlink.com">support@egovlink.com</a> or 513-853-8675.-->
	</p>

	</div>
	<!-- END REQUIRED -->

	</div>
	</form>

<%
End Sub  


'------------------------------------------------------------------------------
' string dbsafe( strDB )
'------------------------------------------------------------------------------
Function dbsafe( ByVal strDB )
 	If Not VarType(strDB) = vbString Then dbsafe = strDB : Exit Function 
	    dbsafe = replace( strDB, "'", "''" )
End Function 


'------------------------------------------------------------------------------
' string GetSerialNumber( iId )
'------------------------------------------------------------------------------
Function GetSerialNumber( ByVal iID )
	Dim iReturnValue, sSql

	iReturnValue = "00000000"

	sSql = "SELECT * FROM egov_skipjackoptions where orgid = " & iID 

	Set oSerialNumber = Server.CreateObject("ADODB.Recordset")
	oSerialNumber.Open sSql, Application("DSN"), 3, 1

	If Not oSerialNumber.EOF Then 
		iReturnValue = oSerialNumber("serialnumber")
	End If 
	oSerialNumber.Close
	Set oSerialNumber = Nothing 

	GetSerialNumber = iReturnValue

End Function 


'------------------------------------------------------------------------------
' string GetDetails( iPaymentModule, iOrgID )
'------------------------------------------------------------------------------
function GetDetails( ByVal iPaymentModule, ByVal iOrgID )
 	sReturnValue = ""

 	Select Case iPaymentModule 
		'COMMEMORATIVE GIFTS
		Case 1
			iGiftPaymentID = StoreGiftInformation(iOrgID)
			sReturnValue   = GetFieldValues(iGiftPaymentID)

		'POOL PASSES
		Case 2
			response.write "<input type=""hidden"" name=""iPAYMENT_MODULE"" value=""" & request("iPAYMENT_MODULE") & """ />" & vbcrlf
			response.write "<input type=""hidden"" name=""isPunchcard"" value="""     & request("custom_Punchcard") & """ />" & vbcrlf
			response.write "<input type=""hidden"" name=""punchcard_limit"" value=""" & request("custom_Punchcard Limit") & """ />" & vbcrlf

			if request("display_membershipname") <> "" then
				lcl_membershipname = request("display_membershipname")
			else
				lcl_membershipname = "Pool Pass"
			end if

			For each oField In request.Form 
				if Left(oField,7) = "custom_" then
					If Replace(oField,"custom_","") = "Punchcard" Or Replace(oField,"custom_","") = "Punchcard Limit" Then 
						If Replace(oField,"custom_","") = "Punchcard" Then 
							If request(oField) Then 
								sReturnValue = sReturnValue & Replace(oField,"custom_","") & " : " & replace(request(oField),1,"Yes") & "<br />" & vbcrlf
							Else 
								sReturnValue = sReturnValue
							End If 
						Else 
							sReturnValue = sReturnValue
						End If 
						If Replace(oField,"custom_","") = "Punchcard Limit" Then 
							If request("custom_Punchcard") Then 
								sReturnValue = sReturnValue & Replace(oField,"custom_","") & " : " & request(oField) & "<br />" & vbcrlf
							Else 
								sReturnValue = sReturnValue
							End If 
						Else 
							sReturnValue = sReturnValue
						End If 
					Else 
						sReturnValue = sReturnValue & replace(replace(oField,"custom_",""),"Pool Pass",lcl_membershipname) & " : " & request(oField) & "<br />" & vbcrlf
					End If 
				End If 
			Next 

		'FACILITY RESERVATIONS
		Case 3
			'Check to see if this facility has been reserved while user has been filling out reservation form for same facility.
			lcl_facility_avail = isFacilityAvail("", request("checkindate"), request("checkoutdate"), request("timepartid"), request("facilityid"), request("D"))

			If lcl_facility_avail Then 
				'Determine if a record already exists for this facility, for the current date, for the user that has NOT been purchased yet.
				'If so then simply use the existing record and update the sessionid
				iFacilityPaymentID = updateFacilitySessionID("", request("checkindate"), request("checkoutdate"), request("timepartid"), request("facilityid"), request.cookies("userid"), request("S"), request("S"), request("D"))

				'If the record does not exist then create the new record.
				If iFacilityPaymentID = 0 Then 
					iFacilityPaymentID = StoreFacilityInformation(iOrgID)
				End If 

				sReturnValue = GetFacilityFieldValues(iFacilityPaymentID)

				'Set session variables
				session("facilityid") = request("facilityid")
				session("D")          = request("checkindate")
			Else 
				response.redirect "../recreation/facility_reserve_summary.asp?L=" & request("facilityid") & "&D=" & request("checkindate") & "&success=NA"
			End If 

	End Select 

 	GetDetails = sReturnValue

End  Function 


'------------------------------------------------------------------------------
' integer StoreGiftInformation( iOrgID )
'------------------------------------------------------------------------------
function StoreGiftInformation( ByVal iOrgID )
	Dim oCmd, iReturnValue, iGiftPaymentID

	iReturnValue = 0

	set oCmd = Server.CreateObject("ADODB.Command")
	with oCmd
	.ActiveConnection = Application("DSN")
	.CommandText = "StoreGiftInformation"
	.CommandType = 4

	'INITIATATOR INFORMATION
	.Parameters.Append oCmd.CreateParameter("sFirstName", 200, 1, 50, request("txtfirstname"))
	.Parameters.Append oCmd.CreateParameter("smiddle", 200, 1, 50, request("txtMI"))
	.Parameters.Append oCmd.CreateParameter("slastname", 200, 1, 50, request("txtlastname"))
	.Parameters.Append oCmd.CreateParameter("saddress1", 200, 1, 50, request("txthome_address1"))
	.Parameters.Append oCmd.CreateParameter("saddress2", 200, 1, 50, request("txthome_address2"))
	.Parameters.Append oCmd.CreateParameter("scity", 200, 1, 50, request("txthome_city"))
	.Parameters.Append oCmd.CreateParameter("sstate", 200, 1, 50, request("cbohome_state"))
	.Parameters.Append oCmd.CreateParameter("szip", 200, 1, 50, request("txthome_zip"))
	.Parameters.Append oCmd.CreateParameter("sphone", 200, 1, 50, request("txtPhone1")&"-"&request("txtPhone2")&"-"&request("txtPhone3"))
	.Parameters.Append oCmd.CreateParameter("semail", 200, 1, 50, request("txtEmail"))

	'ACKNOWLEDGEMENT
	.Parameters.Append oCmd.CreateParameter("sack_name", 200, 1, 300, request("txtack_name"))

	'CHECK TO SEE IF THE ADDRESS ARE THE SAME
	If request("chkSameAs") = "TRUE" Then 
		'USE SAME VALUES AS ABOVE
		.Parameters.Append oCmd.CreateParameter("sack_address1", 200, 1, 300, request("txthome_address1"))
		.Parameters.Append oCmd.CreateParameter("sack_address2", 200, 1, 300, request("txthome_address2"))
		.Parameters.Append oCmd.CreateParameter("sack_city", 200, 1, 300, request("txthome_city"))
		.Parameters.Append oCmd.CreateParameter("sack_state", 200, 1, 300, request("cbohome_state"))
		.Parameters.Append oCmd.CreateParameter("sack_zip", 200, 1, 300, request("txthome_zip"))
	Else 
		'USE VALUES ENTERED 
		.Parameters.Append oCmd.CreateParameter("sack_address1", 200, 1, 300, request("txtack_address1"))
		.Parameters.Append oCmd.CreateParameter("sack_address2", 200, 1, 300, request("txtack_address2"))
		.Parameters.Append oCmd.CreateParameter("sack_city", 200, 1, 300, request("txtack_city"))
		.Parameters.Append oCmd.CreateParameter("sack_state", 200, 1, 300, request("txtack_state"))
		.Parameters.Append oCmd.CreateParameter("sack_zip", 200, 1, 300,request("txtAcknoledgeZip"))
	End If 

	'GIFT INFORMATION
	.Parameters.Append oCmd.CreateParameter("decgiftamount", 6, 1,4 , request("amount"))
	.Parameters.Append oCmd.CreateParameter("igiftid", 3, 1, 4, request("GIFTID"))
	.Parameters.Append oCmd.CreateParameter("orgid", 3, 1, 4, iOrgID)
	.Parameters.Append oCmd.CreateParameter("giftpaymentid", 3, 2, 4)

	'GIFT FIELD INFORMATION
	'CALL TO STORE VALUE INFORMATION
	.Execute

	iGiftPaymentID = .Parameters("giftpaymentid")

	If iGiftPaymentID <> "" Then 
		iReturnValue = iGiftPaymentID
	End If 

	'STORE FIELD VALUES
	StoreFieldValues( iGiftPaymentID )

	End with

	Set oCmd = Nothing 

	StoreGiftInformation = iReturnValue

End Function 


'------------------------------------------------------------------------------
' void StoreFieldValues iGiftPaymentID
'------------------------------------------------------------------------------
Sub StoreFieldValues( ByVal iGiftPaymentID )
	Dim oField, arrValues, iFieldID, iFieldGroup

	'LOOP THRU EACH OF THE FIELDS AND ENTER VALUES SUBMITTED
	For Each oField In request.Form 
		If Left(oField,7) = "custom_" Then 
			'GET VALUES
			arrValues   = Split(oField,"_")
			iFieldID    = clng(arrValues(1))
			iFieldGroup = clng(arrValues(2))

			Set oCmd = Server.CreateObject("ADODB.Command")
			with oCmd
			.ActiveConnection = Application("DSN")
			.CommandText = "StoreGiftFieldValues"
			.CommandType = 4

			'STORE VALUES
			.Parameters.Append oCmd.CreateParameter("iFieldID", 3, 1, 4, iFieldID)
			.Parameters.Append oCmd.CreateParameter("iGiftPaymentID", 3, 1, 4, iGiftPaymentID)
			.Parameters.Append oCmd.CreateParameter("sValue", 200, 1, 2000, request(oField))
			.Execute

			End with

			Set oCmd = Nothing 
		End If 
	Next 

	'WRITE PAYMENTID INPUT VALUE
	response.write "<input type=""hidden"" name=""iGiftPaymentID"" value=""" & iGiftPaymentID & """ />" 
	response.write "<input type=""hidden"" name=""iPAYMENT_MODULE"" value=""" & request("iPAYMENT_MODULE") & """ />"

End Sub 


'------------------------------------------------------------------------------
' string GetFieldValues( iGiftPaymentID )
'------------------------------------------------------------------------------
function GetFieldValues( ByVal iGiftPaymentID )
	Dim sReturnValue, sSql

 	sReturnValue = ""

	sSql = "SELECT V.giftvalue, V.giftpaymentid, V.fieldid, F.fieldprompt "
	sSql = sSql & " FROM egov_gift_value V, egov_gift_fields F "
	sSql = sSql & " WHERE V.fieldid = F.fieldid AND V.giftpaymentid = " & iGiftPaymentID 
	sSql = sSql & " ORDER BY V.giftpaymentid, V.fieldid"

	Set oGiftDetails = Server.CreateObject("ADODB.Recordset")
	oGiftDetails.Open sSql, Application("DSN"), 0, 1

	Do While Not oGiftDetails.EOF
		sReturnValue = sReturnValue & "<strong>" & oGiftDetails("fieldprompt") & "</strong> : <i>" & oGiftDetails("giftvalue") & "</i><br />"
		oGiftDetails.MoveNext
	Loop 

	oGiftDetails.Close
	Set oGiftDetails = Nothing 

	GetFieldValues = sReturnValue

End Function 


'------------------------------------------------------------------------------
' integer StoreFacilityInformation( iOrgID )
'------------------------------------------------------------------------------
function StoreFacilityInformation( ByVal iOrgID )
	Dim iReturnValue, oCmd

	iReturnValue = 0

	Set oCmd = Server.CreateObject("ADODB.Command")
	with oCmd
	.ActiveConnection = Application("DSN")
	.CommandText = "StoreFacilityInformation"
	.CommandType = 4

	'RESERVATION INFORMATION
	.Parameters.Append oCmd.CreateParameter("checkintime", 200, 1, 50, request("checkintime"))
	.Parameters.Append oCmd.CreateParameter("checkindate", 200, 1, 50, request("checkindate"))
	.Parameters.Append oCmd.CreateParameter("checkouttime", 200, 1, 50, request("checkouttime"))
	.Parameters.Append oCmd.CreateParameter("checkoutdate", 200, 1, 50, request("checkoutdate"))
	.Parameters.Append oCmd.CreateParameter("lesseeid", 3, 1, 4, request("lesseeid"))
	.Parameters.Append oCmd.CreateParameter("facilitytimepartid", 3, 1, 4, request("timepartid"))
	.Parameters.Append oCmd.CreateParameter("facilityid", 3, 1, 4, request("facilityid"))
	.Parameters.Append oCmd.CreateParameter("orgid", 3, 1, 4, iOrgID)
	.Parameters.Append oCmd.CreateParameter("amount", 3, 1, 4, request("amount"))
	.Parameters.Append oCmd.CreateParameter("internalnote", 200, 1, 50, "")
	.Parameters.Append oCmd.CreateParameter("sessionid", 3, 1, 4, session.sessionid)
	.Parameters.Append oCmd.CreateParameter("facilitypaymentid", 3, 2, 4)

	'CALL TO STORE VALUE INFORMATION
	.Execute

	iFacilityPaymentID = .Parameters("facilitypaymentid")

	If iFacilityPaymentID <> "" Then 
		iReturnValue = iFacilityPaymentID
	End If 

	'STORE FIELD VALUES
	StoreFacilityFieldValues( iFacilityPaymentID )

	End with

	Set oCmd = Nothing 

	StoreFacilityInformation = iReturnValue

End Function 


'------------------------------------------------------------------------------
' void StoreFacilityFieldValues iFacilityPaymentID
'------------------------------------------------------------------------------
sub StoreFacilityFieldValues( ByVal iFacilityPaymentID )
	Dim oField, arrValues, iFieldID, oCmd

	'LOOP THRU EACH OF THE FIELDS AND ENTER VALUES SUBMITTED
	For Each oField In request.form
		If Left(oField,7) = "custom_" Then 
			'GET VALUES
			arrValues = Split(oField,"_")
			iFieldID  = clng(arrValues(2))

			Set oCmd = Server.CreateObject("ADODB.Command")
			with oCmd
			.ActiveConnection = Application("DSN")
			.CommandText = "StoreFacilityFieldValues"
			.CommandType = 4

			'STORE VALUES
			.Parameters.Append oCmd.CreateParameter("iFieldID", 3, 1, 4, iFieldID)
			.Parameters.Append oCmd.CreateParameter("iFacilityPaymentID", 3, 1, 4, iFacilityPaymentID)
			.Parameters.Append oCmd.CreateParameter("sValue", 200, 1, 2000, request(oField))
			.Execute

			End with

			Set oCmd = Nothing 
		End If 
	Next 

	'WRITE PAYMENTID INPUT VALUE
	response.write "<input type=""hidden"" name=""iFacilityPaymentID"" value=""" & iFacilityPaymentID & """ />" 
	response.write "<input type=""hidden"" name=""iPAYMENT_MODULE"" value=""" & request("iPAYMENT_MODULE") & """ />"

End Sub 


'------------------------------------------------------------------------------
' string GetFacilityFieldValues( iFacilityPaymentID )
'------------------------------------------------------------------------------
Function GetFacilityFieldValues( ByVal iFacilityPaymentID )
	Dim sSql, sReturnValue, oRs

	sReturnValue = ""
	sSql = "SELECT V.facilityvalueid, V.fieldid, V.fieldvalue, V.paymentid, "
	sSql = sSql & " F.fieldid AS Expr1, F.fieldprompt, F.fieldtype, F.facilityid, "
	sSql = sSql & " F.sequence, F.isrequired, F.fieldchoices "
	sSql = sSql & " FROM egov_facility_field_values V, egov_facility_fields F "
	sSql = sSql & " WHERE V.fieldid = F.fieldid AND V.paymentid = " & iFacilityPaymentID
	sSql = sSql & " ORDER BY V.paymentid, V.fieldid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 0, 1

	Do While Not oRs.EOF
		sReturnValue = sReturnValue & "<strong>" & oRs("fieldprompt") & "</strong> : <i>" & oRs("fieldvalue") & "</i><br />"
		oRs.MoveNext
	Loop 

	oRs.Close 
	Set oRs = Nothing 

	sTemp = "<strong>Facility Name</strong>: " & request("LodgeName") & "<br />" & vbcrlf
	sTemp = sTemp & "<strong>Check In Date</strong>: " & request("checkindate") & "<br />" & vbcrlf
	sTemp = sTemp & "<strong>Check In Time</strong>: " & request("checkintime") & "<br />" & vbcrlf
	sTemp = sTemp & "<strong>Check Out Date</strong>: " & request("checkoutdate") & "<br />" & vbcrlf
	sTemp = sTemp & "<strong>Check Out Time</strong>: " & request("checkouttime") & "<br />" & vbcrlf

	GetFacilityFieldValues = sTemp & sReturnValue

End Function 


'------------------------------------------------------------------------------
' void GetUserInfo( iUserId, sName, sEmail, sAddress, sCity, sState, sZip, sFirstName, sLastName
'------------------------------------------------------------------------------
Sub GetUserInfo( ByVal iUserId, ByRef sName, ByRef sEmail, ByRef sAddress, ByRef sCity, ByRef sState, ByRef sZip, ByRef sFirstName, ByRef sLastName )
	Dim sSql, oRs

	sName    = ""
	sEmail   = ""
	sAddress = ""
	sCity    = ""
	sState   = ""
	sZip     = ""

	If IsNumeric(iUserID) Then 
		sSql = "SELECT userfname, userlname, useraddress, usercity, userstate, userzip, useremail "
		sSql = sSql & " FROM egov_users "
		sSql = sSql & " WHERE userid = " & iUserId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			sName    = Proper(oRs("userfname")) & " " & Proper(oRs("userlname"))
			sFirstName = oRs("userfname")
			sLastName = oRs("userlname")
			sEmail   = oRs("useremail")
			sAddress = oRs("useraddress")
			sCity    = oRs("usercity")
			sState   = oRs("userstate")
			sZip     = oRs("userzip")
		End If 

		oRs.Close
		Set oRs = Nothing 
	End If 

End Sub 


'------------------------------------------------------------------------------
' string Proper( sString )
'------------------------------------------------------------------------------
function Proper( ByVal sString )
	
	Proper = sString
	If Len(sString) > 0 Then 
		Proper = UCase(Left(sString,1)) & Mid(sString,2)
	End If 

End Function 


'------------------------------------------------------------------------------
Sub dtb_debug(p_value)
'	if p_value <> "" then
'		sSqli = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
'		set rsi = Server.CreateObject("ADODB.Recordset")
'		rsi.Open sSqli, Application("DSN"), 3, 1
'	end If
'	Set rsi = Nothing 
end Sub


%>
