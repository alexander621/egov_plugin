<% 'Response.Expires = -1000 %>
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->

<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: ecskipjack_form.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2005 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module processes Payments for Pikeville via SkipJack.
'
' MODIFICATION HISTORY
' 1.0	??/??/????	??? ??? - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sError, iPaymentFormId

%>
<html>
<head>
	<title>E-Gov Services <%=sOrgName%> - Skipjack Payment Form</title>

	<style>
		<%If request.servervariables("HTTPS") = "on" Then%>
			body {behavior: url('https://secure.egovlink.com/pikeville/csshover.htc');}
		<%End If%>
	</style>

	<link rel="stylesheet" type="text/css" href="../../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../../global.css" />
	<link rel="stylesheet" type="text/css" href="skipjack.css" />
	<link rel="stylesheet" type="text/css" href="../../css/style_<%=iorgid%>.css" />

	<script language="Javascript" src="../../scripts/modules.js"></script>

	<script language=javascript>
		function openWin2(url, name) 
		{
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
		}
	</script>

	<script language="javascript">
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
					var tempobj = theform.elements[i];
					if(tempobj.type.toLowerCase() == "submit" || tempobj.type.toLowerCase() == "reset")
					//disable em
						tempobj.disabled=true;
				}
			}
		} 

	//-->
	</script>

</head>

<!--#Include file="../../include_top.asp"-->

<!--BODY CONTENT-->
<tr><td valign="top">

	<!--BEGIN: INTRO TEXT-->
<p>
<font class="pagetitle">Welcome to the <%=sOrgName%> Permits and Payments Center</font> <br />
<font class="datetagline">Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%>
</font>
</p>
	<!--END: INTRO TEXT-->


	<!--BEGIN:  DISPLAY PAYMENT FORM-->
<% 

	iPaymentFormId = CLng(request("paymentid"))

	fnDisplayForm( iPaymentFormId ) 
%>

	<!--END: DISPLAY PAYMENT FORM-->
 
   
<!--#Include file="../../include_bottom.asp"-->    


<%
'--------------------------------------------------------------------------------------------------
' FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void FNDISPLAYFORM(IPAYMENTFORMID)
'--------------------------------------------------------------------------------------------------
Sub fnDisplayForm( ByVal iPaymentFormID )
	Dim sSql, bNothing, oRs

	'iPaymentFormID = 1

	' GET FORM INFORMATION
	sSql = "SELECT * FROM egov_paymentservices WHERE paymentserviceid = " & iPaymentFormID & " AND orgid = " & iorgid & ""

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
		
	If Not oRs.EOF Then
	
		' FORM HEADING	
		response.write "<blockquote>"
		response.write "<font class=formtitle>" & oRs("paymentservicename") & " - Payment Information</font>"
		response.write "<div class=group>"
	

	%>

	<!--BEGIN: BUILD PAYMENT FORM-->
	<%If iorgid = 5 Then %>
		<form name="Fullorder" action="https://developer.skipjackic.com/scripts/EvolvCC.dll?Authorize" method="post">
	<%Else%>
		<form name="Fullorder" action="https://www.skipjackic.com/scripts/EvolvCC.dll?Authorize" method="post">
	<%End If%>



	<!--BEGIN: GET PAYMENT INFORMATION-->
	<%
	Dim sOrderString,sSerialNumber,sOrderNumber
	Dim sItemNumber,sItemDescription,sItemQuantity,blnTaxTable,curItemCost

	sItemNumber = REQUEST("ITEM_NUMBER")
	sItemDescription = REQUEST("ITEM_NAME")
	curItemCost = REQUEST("custom_paymentamount")
	sItemQuantity = 1
	blnTaxTable = N ' Y USE TABLE | N DON'T USE TABLE
	sOrderString = sItemNumber & "~" & sItemDescription & "~" & curItemCost & "~" & sItemQuantity & "~" & blnTaxTable & "||"  
	sSerialNumber =  GetSerialNumber(iOrgID) '"000389764111" ' <<< DEMO VALUE --- LIVE VALUE PULLED FROM DATABASE >>GetSerialNumber(iOrgID) 

	sOrderNumber = "ecC" & iOrgID & "O" & CreatePayment() ' STORE VALUES IN DATABASE

	' IF DEMO ORGANIZATIO SHOW TEST VALUES
	If iorgid=5 Then 
		'sName = "John Stullenberger"
		'sEmail = "jstullenberger@eclink.com"
		sAddress = "4303 Hamilton Avenue"
		sCity = "Cincinnati"
		sState = "OH"
		sZip = "45223"
		sOrderNum = "43036814030"
		sCreditCardNum = "5121212121212124"
		sExpMonth = "09"
		sExpYear = "2007"
		sCVSCode = ""
		sPhone = "5136814030"
	End If
	%>

	<input type="hidden" name="orderstring" value="<%=sOrderString%>" />
	<input type="hidden" name="serialnumber" value="<%=sSerialNumber%>" />
	<input type="hidden" name="ordernumber" value="<%=sOrderNumber%>" />
	<!--END: GET PAYMENT INFORMATION-->


	<div align="center" style="width:500px;">

	<!--BEGIN: PAYMENT DETAILS-->
	<fieldset>
	<legend><B>Payment Service</B></legend>
	<table border="0" cellpadding="2" cellspacing="0" width=100% >
	<!--BEGIN: PERSONAL INFORMATION-->
		<tr><td><img hspace=20 src="images/SkipJackTNh121x48.gif"></td>
			<td><b><%="(" & sItemNumber & ") - " & sItemDescription %></b><br><br>
			<b>Details</b><br>

	<%	
			' GET CUSTOM FIELDS 
	For Each oField IN Request.Form
		If Left(oField,7) = "custom_" Then
			sDetails = sDetails & replace(oField,"custom_","") & " : " & request(oField) & "<br />"
		End If
	Next 

	response.write "<i>" & sDetails & "</i>"
	%>

			
			</td>
		</tr>
	</table>
	<br>
	</fieldset>
	<!--END: PAYMENT DETAILS-->


	<P align=left>Please enter your billing information as it appears on your credit card statement, then click the <b>Process Payment</b> button.</P>



	<fieldset>
	<legend><B>Personal Information</B></legend>
	<table border="0" cellpadding="2" cellspacing="0"  >
	<!--BEGIN: PERSONAL INFORMATION-->
		<tr>
		  <td align="right"><B>Name:</B></td>
		  <td><font face="Verdana" size="2"><input maxLength=50 size=30 name="sjname" value="<%=sName%>" tabindex="1"></font></td>
		</tr>
		<tr>
		  <td align="right"><B>E-mail:</B></td>
		  <td><font face="Verdana" size="2"><input value="<%=sEmail%>" maxLength=50 size=30 name="email" tabindex="2"></font></td>
		 </tr>
		<tr>
		  <td align="right"><b>Address:</b></td>
		  <td><font face="Verdana" size="2"><input value="<%=sAddress%>" maxLength=20 size=30 name="streetaddress" tabindex="3"></font></td>
			 </tr>
		<tr>
		  <td align="right"><b>City:</b></td>
		  <td><font face="Verdana" size="2"><input value="<%=sCity%>" maxLength=14 size=30 name="city" tabindex="4"></font></td>
		
		</tr>
		<tr>
		  <td align="right"><b>State:</b></td>
		  <td><font face="Verdana" size="2">            <select name="state" size="1">
				  <option value selected> </option>
				  <option value>          
				  United States       </option>
				  <option value="AL">Alabama </option>
				  <option value="AK">Alaska </option>
				  <option value="AZ">Arizona </option>
				  <option value="AR">Arkansas </option>
				  <option value="CA">California </option>
				  <option value="CO">Colorado </option>
				  <option value="CT">Connecticut </option>
				  <option value="DE">Delaware </option>
				  <option value="DC">District of Columbia </option>
				  <option value="FL">Florida </option>
				  <option value="GA">Georgia </option>
				  <option value="HI">Hawaii </option>
				  <option value="ID">Idaho </option>
				  <option value="IL">Illinois </option>
				  <option value="IN">Indiana </option>
				  <option value="IA">Iowa </option>
				  <option value="KS">Kansas </option>
				  <option selected="selected" value="KY">Kentucky </option>
				  <option value="LA">Louisiana </option>
				  <option value="ME">Maine </option>
				  <option value="MD">Maryland </option>
				  <option value="MA">Massachusetts </option>
				  <option value="MI">Michigan </option>
				  <option value="MN">Minnesota </option>
				  <option value="MS">Mississippi </option>
				  <option value="MO">Missouri </option>
				  <option value="MT">Montana </option>
				  <option value="NE">Nebraska </option>
				  <option value="NV">Nevada </option>
				  <option value="NH">New Hampshire </option>
				  <option value="NJ">New Jersey </option>
				  <option value="NM">New Mexico </option>
				  <option value="NY">New York </option>
				  <option value="NC">North Carolina </option>
				  <option value="ND">North Dakota </option>
				  <option value="OH">Ohio </option>
				  <option value="OK">Oklahoma </option>
				  <option value="OR">Oregon </option>
				  <option value="PA">Pennsylvania </option>
				  <option value="RI">Rhode Island </option>
				  <option value="SC">South Carolina </option>
				  <option value="SD">South Dakota </option>
				  <option value="TN">Tennessee </option>
				  <option value="TX">Texas </option>
				  <option value="UT">Utah </option>
				  <option value="VT">Vermont </option>
				  <option value="VA">Virginia </option>
				  <option value="WA">Washington </option>
				  <option value="WV">West Virginia </option>
				  <option value="WI">Wisconsin </option>
				  <option value="WY">Wyoming </option>
				  <option value> </option>
				  <option value>          
				  U.S. Territories       </option>
				  <option value="GU">Guam </option>
				  <option value="AS">American Samoa </option>
				  <option value="FM">Federated States of Micronesia </option>
				  <option value="MP">Northern Mariana Islands </option>
				  <option value="MH">Marshall Islands </option>
				  <option value="PW">Palau Islands </option>
				  <option value="PR">Puerto Rico </option>
				  <option value="VI">US Virgin Islands </option>
				  <option value> </option>
				  <option value>          
				  Canadian Provinces       </option>
				  <option value="AB">Alberta </option>
				  <option value="BC">British Columbia </option>
				  <option value="MB">Manitoba </option>
				  <option value="NB">New Brunswick </option>
				  <option value="NF">Newfoundland </option>
				  <option value="NT">Northwest Territories </option>
				  <option value="NS">Nova Scotia </option>
				  <option value="ON">Ontario </option>
				  <option value="PE">Prince Edward Island </option>
				  <option value="PQ">Quebec </option>
				  <option value="SK">Saskatchewan </option>
				  <option value="YT">Yukon Territory </option>
				  <option value> </option>
				  <option value>          </option>
				  <option value="XX">Other or None</option>
				</select></font></td>
		 </tr>
		<tr>
		  <td align="right"><b>Zip:</b></td>
		  <td><font face="Verdana" size="2"><input value="<%=sZip%>"  maxLength=15 size=15 name="zipcode" tabindex="6"></font></td>
		  
		</tr>
		</table>
		<br>
		</fieldset>
		<!--END: PERSONAL INFORMATION-->
		
		<br>
		
		<!--BEGIN: CREDIT CARD INFORMATION-->
		<fieldset>
		<legend><B>Credit Card Information</B></legend>
		<table border="0" cellpadding="2" cellspacing="0"  >
		<tr>
		  <td align="right"><b>Credit Card Number:</b></td>
		  <td><font face="Verdana" size="2"><input value="<%=sCreditCardNum%>" maxLength=22 size=30 name="accountnumber" tabindex="8"></font></td>
		  
		</tr>
		<tr>
		  <td align="right"><b>Expiration Month:</b></td>
		  <td>

			<select name="month">
			<%
			' DRAW MONTH SELECTION
			For i=1 to 12
				If i < 10 Then
					sTemp = "0" & i
				Else
					sTemp = i
				End If
				response.write "<OPTION VALUE=""" & sTemp & """>" & sTemp
			Next
			%>
		  </select>
		  
		  </td>
		  <td colspan="2"></td>
		  <td></td>
		  <td> </td>
		</tr>
		<tr>
		  <td align="right"><b>Expiration Year: </b></td>
		  <td>
		  
			<select name="year">
			<%
			' DRAW YEAR SELECTION - CURRENT YEAR MINUS 5 AND PLUS 5
			sTemp = Year(Now())
			For i = 1 to 10

				' GENERATE SELECTED VALUE
				If sTemp = Year(Now())+ 1 Then
					sSelected = "selected=""selected"" "
				Else
					sSelected = ""
				End IF

				response.write "<option " & sSelected & " value=""" & sTemp & """>" & sTemp & "</option>"
				sTemp = sTemp + 1
			Next
			%>
		  </select>
		  
		  </td>
		  <td></td>
		  <td></td>
		</tr>
		<tr>
		  <td align="right"><b>Amount: </b></td>
		  <td><font face="Verdana" size="2"><%=formatcurrency(curItemCost)%><input type="hidden" style="background-color:#e0e0e0;" value="<%=formatnumber(curItemCost,2)%>" maxLength="15" size="15" name="transactionamount" tabindex="11" /></font></td>
		  <td colspan="2"></td>
		  <td></td>
		  <td></td>
		</tr>
		<tr>
		  <td align="right"><b>Phone:</b></td>
		  <td><font face="Verdana" size="2"><input value="<%=sPhone%>" maxLength="15" size="15" name="shiptophone" tabindex="12" /></font></td>
		  <td colspan="2"></td>
		  <td></td>
		  <td></td>
		</tr>
	  </table>
	 </FIELDSET>
	 <!--END: CREDIT CARD INFORMATION-->


	<div align=left><small><b><font color=red>*</font></b><b><i>All Fields Required</I></b></small></div>

	<P align=left class=smallnote><font style="font-weight:bold;color:red">Press PROCESS PAYMENT button only once and please wait for the authorization page to be displayed to prevent double billing.  Be patient, it may take up to 2 minutes to process your transaction.</font></p> 

	<!--BEGIN: PROCESS BUTTONS-->
	 <table border="0" cellpadding="2" cellspacing="0"  >
		<tr>
		  <td ALIGN="center">
			<INPUT STYLE="WIDTH:200PX;" class=skipjackbtn TYPE=SUBMIT NAME="COMPLETE_PAYMENT" VALUE="PROCESS PAYMENT" >
			<!--<input STYLE="WIDTH:100PX;" class=skipjackbtn type="button" value="Cancel" name="buttonRefresh" onclick="process_cancel();" >-->
		</td>
		<td>&nbsp;</td>
		</tr>
	</table>
	<!--END: PROCESS BUTTONS-->

	<P class=smallnote>NOTE: Your IP address [<%=request.servervariables("REMOTE_ADDR")%>] has been logged with this transaction.<br><br>
	Do you have questions?<br>
	Contact Customer Service: <a href="mailto:support@egovlink.com">support@egovlink.com</a> or 513-853-8675.
	</P>


	</div>
	<!-- END REQUIRED -->

	</form>



	<%
		' END FORM
		response.write "</blockquote></div>"
		bNothing = False 

	Else
		 bNothing = True 

	End If 

	oRs.Close 
	Set oRs = Nothing

	If bNothing Then 
		' FORM NOT FOUND REDIRECT TO COMPLETE LIST
		response.redirect("../../payment.asp")
	End If 


End Sub 


'--------------------------------------------------------------------------------------------------
' integer Function CreatePayment()
'--------------------------------------------------------------------------------------------------
Function CreatePayment()
	Dim sSql, iPaymentInfoId, iReturnValue

	iReturnValue = 0

	iPaymentInfoId = AddPaymentDetail()

	sSql = "INSERT INTO egov_payments ( paymentdate, paymentstatus, paymentserviceid, paymentinfoid, orgid ) VALUES ( "
	sSql = sSql & "getdate(), 'PROCESSING', " & CLng(request("ITEM_NUMBER"))/100 & ", " & iPaymentInfoId & ", " & iOrgId & " )"

	iReturnValue = RunIdentityInsertStatement( sSql )

'	Set oDetails = Server.CreateObject("ADODB.Recordset")
'	oDetails.CursorLocation = 3
'	oDetails.Open "egov_payments", Application("DSN") , 1, 2, 2
'	oDetails.AddNew
'	oDetails("paymentdate") = Now()
'	oDetails("paymentstatus") = "PROCESSING"
'	oDetails("paymentserviceid") = CLng(REQUEST("ITEM_NUMBER"))/100 
'	oDetails("paymentinfoid") = AddPaymentDetail()
'	oDetails.Update
'	iReturnValue = oDetails("paymentid")
'	oDetails.Close

'	Set oDetails = Nothing

	CreatePayment = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' integer Function AddPaymentDetail()
'--------------------------------------------------------------------------------------------------
Function AddPaymentDetail()
	Dim sSql, iReturnValue, sDetails

	iReturnValue = 0
	
	sDetails = ""
	' GET CUSTOM FIELDS 
	For Each oField IN Request.Form
		If Left(oField,7) = "custom_" Then
			sDetails = sDetails & replace(oField,"custom_","") & " : " & request(oField) & "</br>"
		End If
	Next 

	sSql = "INSERT INTO egov_paymentinformation ( payment_information ) VALUES ( '" & DBsafe(sDetails) & "' )"

	iReturnValue = RunIdentityInsertStatement( sSql )

'	Set oDetails = Server.CreateObject("ADODB.Recordset")
'	oDetails.CursorLocation = 3
'	oDetails.Open "egov_paymentinformation", Application("DSN") , 1, 2, 2
'	oDetails.AddNew
'	oDetails("payment_information") = dbsafe(sDetails)
'	oDetails.Update
'	iReturnValue = oDetails("paymentinfoid")
'	oDetails.Close

'	Set oDetails = Nothing

	AddPaymentDetail = iReturnValue

End Function


'------------------------------------------------------------------------------------------------------------
' string DBSAFE( STRDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )

  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function

  DBsafe = Replace( strDB, "'", "''" )

End Function


'------------------------------------------------------------------------------------------------------------
' string GETSERIALNUMBER( IID )
'------------------------------------------------------------------------------------------------------------
Function GetSerialNumber( ByVal iID )
	Dim sSql, oRs, iReturnValue

	iReturnValue = "00000000"
	
	sSql = "SELECT * FROM egov_skipjackoptions WHERE orgid = " & iID 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1


	If Not oRs.EOF Then
		iReturnValue = oRs("serialnumber")
	End If

	oRs.Close
	Set oRs = Nothing 

	GetSerialNumber = iReturnValue

End Function




%>








