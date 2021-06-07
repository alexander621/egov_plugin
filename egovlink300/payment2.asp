<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<% Dim sError %>
<html>
<head>
<title>E-Gov Services - <%=sOrgName%></title>
<link rel="stylesheet" href="css/styles.css" type="text/css">
<link rel="stylesheet" href="css/style_<%=iorgid%>.css" type="text/css">

	<link href="global.css" rel="stylesheet" type="text/css">
	<script language="Javascript" src="scripts/modules.js"></script>
<script language="Javascript" src="scripts/easyform.js"></script>  
<script language=javascript>
function openWin2(url, name) {
  popupWin = window.open(url, name,"resizable,width=500,height=450");
}
</script>

</head>

<!--#Include file="include_top.asp"-->


<!--BODY CONTENT-->
<p>
<font class=pagetitle>Welcome to the <%=sOrgName%> Permits and Payments Center</font> <BR>

<% If sOrgRegistration AND trim(request("paymenttype")) <> "" Then %>
		<%  If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then
				RegisteredUserDisplay()
			Else %>
				<a href="user_login.asp">Click here to Login</a> |
				<a href="register.asp">Click here to Register</a>
		<% End If %>
<% Else %>
	<font class=datetagline>Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%> </font>
<% End If%>


</font>
</p>


<%
' ---------------------------------------------------------------------------------------
' BEGIN DISPLAY PAGE CONTENT
' ---------------------------------------------------------------------------------------
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
	response.write "<div style=""margin-left:20px; "" class=box_header2>Payment and Permit Services</div>"
	response.write "<div style=""margin-left:20px; "" class=groupsmall>"
	fn_DisplayPayments()
	response.write "</div>"
	response.write "</td>"
%>

<td width=225 style="padding-left:15px;" valign="top">

	  <%If sOrgRegistration Then %>
			  <b>Personalized E-Gov Services</b>
			  <ul>
				<li><a href="user_login.asp">Click here to Login</a>
				<li><a href="register.asp">Click here to Register</a>
			  </ul>
			  <hr style="width: 90%; size: 1px; height: 1px;">
	 <%End If%>

		<%=sPaymentDescription%>
	</td>
</tr>
</table>



<!--SPACING CODE-->
<p><bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->

	
<%
Else
	' DISPLAY PAYMENT FORM
	fn_DisplayPaymentForm(request("paymenttype"))
End If
' ---------------------------------------------------------------------------------------
' END DISPLAY PAGE CONTENT
' ---------------------------------------------------------------------------------------
%>


<!--SPACING CODE-->
<p>&nbsp;<bR>&nbsp;<bR>&nbsp;<bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->


<!--#Include file="include_bottom.asp"-->  





<%
' -----------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
' -----------------------------------------------------------------------------------------------------------

' -----------------------------------------------------------------------------------------------------------
' FUNCTION FN_DISPLAYPAYMENTFORM(IID)
' -----------------------------------------------------------------------------------------------------------
Function fn_DisplayPaymentForm(iID)

' GET FORM INFORMATION
sSQL = "SELECT * FROM egov_paymentservices WHERE paymentserviceid=" & iID

Set oPaymentServices = Server.CreateObject("ADODB.Recordset")
oPaymentServices.Open sSQL, Application("DSN") , 3, 1
	
If NOT oPaymentServices.EOF Then

	'--------------------------------------------------------------------------------------------------
	' BEGIN: VISITOR TRACKING
	'--------------------------------------------------------------------------------------------------
		iSectionID = 33
		sDocumentTitle = oPaymentServices("paymentservicename")
		sURL = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
		datDate = Date()	
		datDateTime = Now()
		sVisitorIP = request.servervariables("REMOTE_ADDR")
		Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
	'--------------------------------------------------------------------------------------------------
	' END: VISITOR TRACKING
	'--------------------------------------------------------------------------------------------------
	
	' FORM HEADING	
	response.write "<blockquote>"
	response.write "<font class=formtitle>" & oPaymentServices("paymentservicename") & "</font>"
	response.write "<div class=group>"

	' FORM DESCRIPTION 
	If oPaymentServices("paymentservicedescription") <> "" Then
		response.write oPaymentServices("paymentservicedescription")
	End If

	' FORM INSTRUCTIONS
	If oPaymentServices("paymentserviceinstructions") <> "" Then
		response.write oPaymentServices("paymentserviceinstructions")
	End If


	' FORM PAYMENT OPTIONS
	fn_GetPaymentGatewayOptions iPaymentGatewayID ,iID, oPaymentServices("paymentservicename")
	

	' FORM REQUIRED VALUES
	response.write "<table>"

	
	' FORM SPECIFIC FIELD OPTIONS
	GetPaymentFields(iID)


	' OPTIONAL FIELDS AND AMOUNT FOR PAY PAL
	GetPayPalFieldValues(iID)


	' SUBMIT BUTTON
	response.write "<tr><td colspan=2 align=right><input onclick=""if (validateForm('frmpayment')) { document.frmpayment.submit(); }"" type=""button"" class=paymentbtn name=""btnsubmit"" value=""CONTINUE"" alt=""CONTINUE""></td></tr></table>"


	' FORM NOTES
	If oPaymentServices("paymentservicenotes") <> "" Then
		response.write oPaymentServices("paymentservicenotes")
	End If

	' END FORM	
	response.write "</form></div>"

Else

	' FORM NOT FOUND REDIRECT TO COMPLETE LIST
	response.redirect("payment.asp")

End If

End Function 


' -----------------------------------------------------------------------------------------------------------
' FUNCTION GETPAYPALVALUES(IID)
' -----------------------------------------------------------------------------------------------------------
Function GetPayPalValues(iID)

sSQL = "SELECT * FROM egov_paypaloptions WHERE paymentserviceid=" & iID

Set oPayPalOptions = Server.CreateObject("ADODB.Recordset")
oPayPalOptions.Open sSQL, Application("DSN") , 3, 1
	
If NOT oPayPalOptions.EOF Then


		' STATIC VALUES
		'response.write "<form name=""frmpayment"" action=""https://www.sandbox.paypal.com/cgi-bin/webscr"" method=""post"">"
		
		' PAYPAL GATEWAY
		response.write "<form action=""transfer_payment.asp"" method=""post"">"

		response.write "<input type=""hidden"" name=""cmd"" value=""_xclick"">"
	
	Do While NOT oPayPalOptions.EOF 

		
		' DYNAMIC VALUES
		response.write "<input type=""hidden"" name=""" & oPayPalOptions("paypaloptionname") & """ value=""" & oPayPalOptions("paypaloptionvalue") & """>" & vbcrlf
		
		oPayPalOptions.MoveNext
	Loop

End If

Set oPayPalOptions = Nothing

End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION GETPAYPALFIELDVALUES(IID)
'------------------------------------------------------------------------------------------------------------
Function GetPayPalFieldValues(iID)

sSQL = "SELECT * FROM egov_paypalfields WHERE paymentserviceid=" & iID

Set oPayPalFields = Server.CreateObject("ADODB.Recordset")
oPayPalFields.Open sSQL, Application("DSN") , 3, 1
	
If NOT oPayPalFields.EOF Then
		
		'If oPayPalFields("on0") <> "" Then
			'response.write "<tr><td><input type=""hidden"" name=""on0"" value=""" & oPayPalFields("on0") & """ ><b>" & 'oPayPalFields("on0")& "</b></td><td><input type=""text"" name=""os0"" maxlength=""200""></td></tr>"
		'End If
		
		'If oPayPalFields("on1") <> "" Then
			'response.write "<tr><td><input type=""hidden"" name=""on1"" value=""" & oPayPalFields("on1") & """><b>" & 'oPayPalFields("on1") & "</b></td><td><input type=""text"" name=""os1"" maxlength=""200""></td></tr>"
		'End If
		
		If oPayPalFields("amount") <> "" Then
			curValue = formatnumber(oPayPalFields("amount"),2)
			sDisabled = "DISABLED"
			response.write "<tr><td><b>Payment Amount: </b></td><td>" & curValue &"<input type=""hidden"" name=""amount"" maxlength=""200"" value=""" & curValue &"""></td></tr>"
		Else
			curValue = ""
			sDisabled = ""
			response.write "<input type=hidden name=""ef:amount-text/req"" value=""Payment Amount"">"
			response.write "<tr><td><b>Payment Amount: </b></td><td><input type=""text"" name=""AMOUNT"" maxlength=""200"" value=""" & curValue &"""></td></tr>"
		End If


		
End If

Set oPayPalFields = Nothing

End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION GETPAYMENTFIELDS(IID)
'------------------------------------------------------------------------------------------------------------
Function GetPaymentFields(iID)

sSQL = "SELECT * FROM egov_paymentfields WHERE paymentserviceid=" & iID

Set oPaymentFields = Server.CreateObject("ADODB.Recordset")
oPaymentFields.Open sSQL, Application("DSN") , 3, 1
	
If NOT oPaymentFields.EOF Then

	Do While NOT oPaymentFields.EOF 
		' DYNAMIC VALUES
		
		' BUILD EASY VALIDATION STRING
		If ISNULL(oPaymentFields("paymentvalidation")) OR oPaymentFields("paymentvalidation") = "" Then
			sValidation = ""
		Else
			sValidation = "-text/" & oPaymentFields("paymentvalidation") & "/req"
		End If
	 
	   ' WRITE EASY FORM HIDDEN FIELD VALIDATION VALUE
	   response.write "<input type=hidden name=""ef:custom_" & oPaymentFields("paymentfieldsname") &   sValidation &  """ value=""" &  oPaymentFields("paymentfielddisplayname")  & """>"
	
	   ' WRITE ACTUAL FORM VALUE	
	   Select Case oPaymentFields("paymentfieldtype") 
	
			Case "textarea"
			' TEXTAREA
			response.write "<tr><td COLSPAN=2 align=LEFT><b>" & oPaymentFields("paymentfielddisplayname") & " :</b><BR><TEXTAREA " & oPaymentFields("paymentfieldattributes") & "  name=""custom_" & oPaymentFields("paymentfieldsname") & """ style=""" & oPaymentFields("paymentfieldstyle") & """ class=""formtextarea""></TEXTAREA></td><TD> " & oPaymentFields("paymentdesc") & " </td></tr>" & vbcrlf 

			Case Else
			' DEFAULT IS TEXTBOX
			response.write "<tr><td align=LEFT><b>" & oPaymentFields("paymentfielddisplayname") & " :</b></td><td> <input type=""text"" " & oPaymentFields("paymentfieldattributes") & "  name=""custom_" & oPaymentFields("paymentfieldsname") & """ style=""" & oPaymentFields("paymentfieldstyle") & """></td><TD> " & oPaymentFields("paymentdesc") & " </td></tr>" & vbcrlf 

		End Select
		
		oPaymentFields.MoveNext
	Loop

End If

Set oPaymentFields = Nothing

End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION FN_DISPLAYPAYMENTS()
'------------------------------------------------------------------------------------------------------------
Function fn_DisplayPayments()
	response.write "<ul>"
	sSQL = "SELECT DISTINCT TOP 5 paymentserviceid,paymentservicename FROM egov_paymentservices LEFT OUTER JOIN egov_organizations_to_paymentservices ON egov_paymentservices.paymentserviceid=egov_organizations_to_paymentservices.paymentservice_id where (egov_organizations_to_paymentservices.paymentservice_enabled <> 0 and (egov_organizations_to_paymentservices.orgid=" & iorgid & " OR egov_organizations_to_paymentservices.orgid=0))"
	Set oPaymentServices = Server.CreateObject("ADODB.Recordset")
	oPaymentServices.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oPaymentServices.EOF Then
		Do while NOT oPaymentServices.EOF 
			response.write "<li><a href=""payment.asp?paymenttype=" & oPaymentServices("paymentserviceid") & """>" & oPaymentServices("paymentservicename") &  "</a>"
			oPaymentServices.MoveNext
		Loop

	End If
	Set oPaymentServices = Nothing 

	response.write "</ul>"
End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION FN_GETPAYMENTGATEWAYOPTIONS(IPAYMENTGATEWAYID,IPAYMENTSERVICEID)
'------------------------------------------------------------------------------------------------------------
Function fn_GetPaymentGatewayOptions(iPaymentGatewayID,iPaymentServiceID, sServiceName)

	Select Case iPaymentGatewayID

		Case 1 
			' PAYPAL GATEWAY
			GetPayPalValues iPaymentServiceID

		Case 2
			' SKIPJACK GATEWAY
			GetSkipJackValues iPaymentServiceID,sServiceName

		Case 3
			' ECLINK PAYPAL DEMO GATEWAY
			GeteclinkPPDemoValues iPaymentServiceID 

		Case 4
			' VERISIGN GATEWAY
			GetVerisignValues iPaymentServiceID,sServiceName

		Case Else
			' ECLINK PAYPAL DEMO GATEWAY
			GeteclinkPPDemoValues iPaymentServiceID

	End Select

End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION GETSKIPJACKVALUES(IPAYMENTSERVICEID)
'------------------------------------------------------------------------------------------------------------
Function GetSkipJackValues(iPaymentServiceID,sServiceName)
	response.write  "<FORM  name=""frmpayment"" ACTION=""PAYMENT_PROCESSORS/SKIPJACK2004/ECSKIPJACK_FORM.ASP"" METHOD=""POST"">"
	response.write "<input type=hidden name=""PAYMENTID"" value=""" & iPaymentServiceID & """>"
	response.write "<input type=hidden name=""ITEM_NUMBER"" value=""" & iPaymentServiceID & "00"">"
	response.write "<input type=hidden name=""ITEM_NAME"" value=""" & sServiceName & """>"
End Function



'------------------------------------------------------------------------------------------------------------
' FUNCTION GETVERISIGNVALUES(IPAYMENTSERVICEID,SSERVICENAME)
'------------------------------------------------------------------------------------------------------------
Function GetVerisignValues(iPaymentServiceID,sServiceName)
	response.write  "<FORM  name=""frmpayment"" ACTION=""PAYMENT_PROCESSORS/VERISIGN2005/VERISIGN_FORM.ASP"" METHOD=""POST"">"
	response.write "<input type=hidden name=""PAYMENTID"" value=""" & iPaymentServiceID & """>"
	response.write "<input type=hidden name=""ITEM_NUMBER"" value=""" & iPaymentServiceID & "00"">"
	response.write "<input type=hidden name=""ITEM_NAME"" value=""" & sServiceName & """>"
End Function



'------------------------------------------------------------------------------------------------------------
' FUNCTION GeteclinkPPDemoValues(IPAYMENTSERVICEID)
'------------------------------------------------------------------------------------------------------------
Function GeteclinkPPDemoValues(iPaymentServiceID)
	response.write "<form name=""frmpayment"" action=""paypal_demo_pages/paypal_demo_page1.asp"" method=""GET"">"
End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION GetPayPalValues(iPaymentServiceID)
'------------------------------------------------------------------------------------------------------------
Function GetPayPalValues(iPaymentServiceID)
	
	sSQL = "SELECT * FROM egov_paypaloptions WHERE paymentserviceid=" & iPaymentServiceID

	Set oPayPalOptions = Server.CreateObject("ADODB.Recordset")
	oPayPalOptions.Open sSQL, Application("DSN") , 3, 1
		
	If NOT oPayPalOptions.EOF Then


			' STATIC VALUES
			'response.write "<form name=""frmpayment"" action=""https://www.sandbox.paypal.com/cgi-bin/webscr"" method=""post"">"
			
			' PAYPAL GATEWAY
			response.write "<form name=""frmpayment"" action=""transfer_payment.asp"" method=""post"">"

			response.write "<input type=""hidden"" name=""cmd"" value=""_xclick"">"
		
		Do While NOT oPayPalOptions.EOF 

			
			' DYNAMIC VALUES
			response.write "<input type=""hidden"" name=""" & oPayPalOptions("paypaloptionname") & """ value=""" & oPayPalOptions("paypaloptionvalue") & """>" & vbcrlf
			
			oPayPalOptions.MoveNext
		Loop

	End If

	Set oPayPalOptions = Nothing

End Function


'------------------------------------------------------------------------------------------------------------
'Function JSsafe( strDB )
'------------------------------------------------------------------------------------------------------------
Function JSsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  strDB = Replace( strDB, "'", "\'" )
  strDB = Replace( strDB, chr(34), "\'" )
  strDB = Replace( strDB, ";", "\;" )
  strDB = Replace( strDB, "-", "\-" )
  strDB = Replace( strDB, "(", "\(" )
  strDB = Replace( strDB, ")", "\)" )
  strDB = Replace( strDB, "/", "\/" )
  JSsafe = strDB
End Function
%>




