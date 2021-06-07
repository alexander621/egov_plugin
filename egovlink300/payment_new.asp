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
	<!--BEGIN:  USER REGISTRATION - USER MENU-->
	<% If sOrgRegistration Then %>
			<%  If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then
					RegisteredUserDisplay()
				Else %>
					<font class=datetagline>Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%> </font>
			<% End If %>
	<% Else %>
		<font class=datetagline>Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%> </font>
	<% End If%>
	<!--END:  USER REGISTRATION - USER MENU-->
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
	'fn_GetPaymentGatewayOptions iPaymentGatewayID ,iID, oPaymentServices("paymentservicename")
	fn_GetPaymentGatewayOptions 3 ,iID, oPaymentServices("paymentservicename") ' TO TEST FORMS

	' FORM REQUIRED VALUES
	response.write "<table>"

	
	' FORM SPECIFIC FIELD OPTIONS
	'GetPaymentFields(iID)
	Custom_Payment_FormID(iID)

	' OPTIONAL FIELDS AND AMOUNT FOR PAY PAL
	GetPayPalFieldValues(iID)


	' SUBMIT BUTTON
	If iID <> 22 Then
		response.write "<tr><td colspan=2 align=right><input onclick=""if (validateForm('frmpayment')) { document.frmpayment.submit(); }"" type=""button"" class=paymentbtn name=""btnsubmit"" value=""CONTINUE"" alt=""CONTINUE""></td></tr></table>"
	Else
		response.write "<tr><td colspan=2 align=right><input onclick=""vcheck();"" type=""button"" class=paymentbtn name=""btnsubmit"" value=""CONTINUE"" alt=""CONTINUE""></td></tr></table>"
	End If


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


'------------------------------------------------------------------------------------------------------------
'Function Custom_Payment_FormID(iFormID)
'------------------------------------------------------------------------------------------------------------
Function Custom_Payment_FormID(iFormID)

	Select Case iFormID

		Case "21"

				
			' WASTERWATER FORM
			response.write "<tr><td colspan=2><table>"
			response.write "<tr><td><b>Service Address :</b></td><td><table><tr><td><input maxlength=5 name=""custom_sa1"" style=""width: 50px"" type=text maxlength=5></td><td><input maxlength=20 name=""custom_sa2"" type=text maxlength=20></td><td><input maxlength=4 name=""custom_sa3"" style=""width: 50px"" type=text maxlength=4><font class=example>Ex: 123 Main Blvd</font></td></tr></table></td></tr>"
			response.write "<tr><td><b>Account Number :</b></td><td><table><tr><td><input maxlength=9 name=""custom_an1"" style=""width: 100px"" type=text maxlength=9></td><td><b> - </b></td><td><input maxlength=9 name=""custom_an2"" style=""width: 100px"" type=text maxlength=9><font class=example>Ex: 286-1384</font></td></tr></table></td></tr>"
			response.write "<tr><td><b>Payment Amount :</b></td><td><table><tr><td colspan=2><input name=""custom_paymentamount"" type=text ></td><td>&nbsp;</td></tr>"
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
			response.write "<tr><td><b>Payment Amount :</b></td><td><table><tr><td ><input name=""custom_paymentamount"" type=text ></td><td>&nbsp;</td></tr>"
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

		Case Else

	End Select

End Function


'------------------------------------------------------------------------------------------------------------
'FUNCTION DRAWINPUTTABLE(IROWS)
'------------------------------------------------------------------------------------------------------------
Function DrawInputTable(iRows)
	
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
	response.write "<tr><td colspan=3 align=right style=""border-bottom:solid 1px #000000;"" ><b>Total Amount Paid:</b></TD><td colspan=3 align=right style=""border-bottom:solid 1px #000000;"" >$<INPUT name=""custom_paymentamount"" TYPE=TEXT style=""width:100px;""></TD></tr>"

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
			for (iCount=1; iCount < <%=(iRows * 4)%>; iCount++){
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
End Function
%>




