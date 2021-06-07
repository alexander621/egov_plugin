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
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim oPayOrg
Set oPayOrg = New classOrganization

' This check added so direct access to the payments page is not possible if the feature is turned off - Steve Loar - 12/27/2005
If (Not OrgHasFeature( iOrgId, "payments" )) Or (Not blnOrgPayment) Then
   	response.redirect sEgovWebsiteURL & "/"
End If

Dim sError 

'catch sql intrusions here
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
	
	<script language="Javascript" src="scripts/modules.js"></script>
	<script language="Javascript" src="scripts/easyform.js"></script>  

	<script language="Javascript">
	<!--
		function openWin2(url, name) 
		{
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
		}
	//-->
	</script>

</head>

<!--#Include file="include_top.asp"-->

<!--BODY CONTENT-->
<p>
<font class="pagetitle">Welcome to the <%=oPayOrg.GetOrgName()%>, <%=oPayOrg.GetState()%>, Permits and Payments Center</font><br>
	<!--BEGIN:  USER REGISTRATION - USER MENU-->
<%	RegisteredUserDisplay( "" ) %>
</p>
<%
' ---------------------------------------------------------------------------------------
' BEGIN DISPLAY PAGE CONTENT
' ---------------------------------------------------------------------------------------
if trim(request("paymenttype")) = "" then
	
	 '--------------------------------------------------------------------------------------------------
 	' BEGIN: VISITOR TRACKING
	 '--------------------------------------------------------------------------------------------------
 		iSectionID     = 3
	 	sDocumentTitle = "MAIN"
		 sURL           = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
 		datDate        = Date()	
	 	datDateTime    = Now()
		 sVisitorIP     = request.servervariables("REMOTE_ADDR")
 		Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
 	'--------------------------------------------------------------------------------------------------
	 ' END: VISITOR TRACKING
 	'--------------------------------------------------------------------------------------------------
	
  'DISPLAY LIST OF PAYMENTS
  	response.write "<table>" & vbcrlf
   response.write "  <tr>" & vbcrlf
   response.write "      <td valign=""top"">" & vbcrlf
   response.write "          <div class=""box_header2"">Payment and Permit Services</div>" & vbcrlf
  	response.write "          <div class=""groupSmall"">" & vbcrlf

    	fn_DisplayPayments()

  	response.write "          </div>" & vbcrlf
  	response.write "      </td>" & vbcrlf
   response.write "      <td width=""225"" style=""padding-left:15px;"" valign=""top"">" & vbcrlf

  'BEGIN: REGISTER/LOGIN LINKS
   if sOrgRegistration AND (request.cookies("userid") = "" OR request.cookies("userid") = "-1") then
%>
          <b>Personalized E-Gov Services</b>
    				  <ul>
         					<li><a href="user_login.asp">Click here to Login</a>
         					<li><a href="register.asp">Click here to Register</a>
    				  </ul>
    				  <hr style="width: 90%; size: 1px; height: 1px;">
<%
   end if
  'END: REGISTER/LOGIN LINKS

   sPaymentDescription

   response.write "      </td>" & vbcrlf
   response.write "  </tr>" & vbcrlf
   response.write "</table>" & vbcrlf

   set oPayOrg = nothing

  'SPACING CODE
   response.write "<p><br>&nbsp;<br>&nbsp;</p>" & vbcrlf
  'SPACING CODE

else
 	'DISPLAY PAYMENT FORM
  	fn_DisplayPaymentForm(CLng(request("paymenttype")))
end if
'---------------------------------------------------------------------------------------
' END DISPLAY PAGE CONTENT
'---------------------------------------------------------------------------------------

'SPACING CODE
 response.write "<p>&nbsp;<br>&nbsp;<br>&nbsp;</p>" & vbcrlf
'SPACING CODE
%>
<!--#Include file="include_bottom.asp"-->  
<%
' -----------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
' -----------------------------------------------------------------------------------------------------------
' -----------------------------------------------------------------------------------------------------------
' FUNCTION FN_DISPLAYPAYMENTFORM(IID)
' -----------------------------------------------------------------------------------------------------------
function fn_DisplayPaymentForm(iID)

'GET FORM INFORMATION
 sSQL = "SELECT * "
 sSQL = sSQL & " FROM egov_paymentservices "
 sSQL = sSQL & " WHERE paymentserviceid = " & iID
 sSQL = sSQL & " AND orgid = " & iorgid & " "

 set oPaymentServices = Server.CreateObject("ADODB.Recordset")
 oPaymentServices.Open sSQL, Application("DSN"), 3, 1
	
 if NOT oPaymentServices.eof then
  	'--------------------------------------------------------------------------------------------------
  	' BEGIN: VISITOR TRACKING
  	'--------------------------------------------------------------------------------------------------
  		iSectionID     = 33
  		sDocumentTitle = oPaymentServices("paymentservicename")
		  sURL           = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
  		datDate        = Date()	
		  datDateTime    = Now()
  		sVisitorIP     = request.servervariables("REMOTE_ADDR")
		  Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)

  	'--------------------------------------------------------------------------------------------------
  	' END: VISITOR TRACKING
  	'--------------------------------------------------------------------------------------------------

  	'FORM HEADING	
   	response.write "<blockquote>" & vbcrlf
   	response.write "<font class=""formtitle"">" & oPaymentServices("paymentservicename") & "</font>" & vbcrlf
   	response.write "<div class=""group"">" & vbcrlf

  	'FORM DESCRIPTION 
   	if oPaymentServices("paymentservicedescription") <> "" then
     		response.write oPaymentServices("paymentservicedescription") & vbcrlf
   	end if

  	'FORM INSTRUCTIONS
   	if oPaymentServices("paymentserviceinstructions") <> "" then
     		response.write oPaymentServices("paymentserviceinstructions") & vbcrlf
    end if

  	'FORM PAYMENT OPTIONS
   	fn_GetPaymentGatewayOptions iPaymentGatewayID ,iID, oPaymentServices("paymentservicename")

  	'FORM REQUIRED VALUES
   	response.write "<table border=""0"">" & vbcrlf

  	'FORM SPECIFIC FIELD OPTIONS
   	if iID=21 OR iID=22 OR iID=23 then
     		Custom_Payment_FormID(iID)
   	else
     		GetPaymentFields(iID)
    end if

  	'OPTIONAL FIELDS AND AMOUNT FOR PAY PAL
   	GetPayPalFieldValues(iID)

  	'SUBMIT BUTTON
   	if iID <> 22 then
     		response.write "  <tr>" & vbcrlf
       response.write "      <td colspan=""2"" align=""right"">" & vbcrlf
       response.write "          <input onclick=""if (validateForm('frmpayment')) { document.frmpayment.submit(); }"" type=""button"" class=""paymentbtn"" name=""btnsubmit"" value=""CONTINUE"" alt=""CONTINUE"">" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
       response.write "</table>" & vbcrlf
   	else
     		response.write "  <tr>" & vbcrlf
       response.write "      <td colspan=""2"" align=""right"">" & vbcrlf
       response.write "          <input onclick=""vcheck();"" type=""button"" class=""paymentbtn"" name=""btnsubmit"" value=""CONTINUE"" alt=""CONTINUE"">" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
       response.write "</table>" & vbcrlf
    end if

  	'FORM NOTES
   	if oPaymentServices("paymentservicenotes") <> "" then
     		response.write oPaymentServices("paymentservicenotes") & vbcrlf
    end if

  	'END FORM	
   	response.write "</form></div>" & vbcrlf

 else

   'FORM NOT FOUND REDIRECT TO COMPLETE LIST
   	response.redirect("payment.asp")

 end if

end function

' -----------------------------------------------------------------------------------------------------------
' FUNCTION GETPAYPALVALUES(IID)
' -----------------------------------------------------------------------------------------------------------
Function GetPayPalValues(iID)

sSQL = "SELECT * FROM egov_paypaloptions WHERE paymentserviceid=" & iID

Set oPayPalOptions = Server.CreateObject("ADODB.Recordset")
oPayPalOptions.Open sSQL, Application("DSN"), 3, 1
	
If NOT oPayPalOptions.EOF Then


		'STATIC VALUES
		'response.write "<form name=""frmpayment"" action=""https://www.sandbox.paypal.com/cgi-bin/webscr"" method=""post"">"
		
		'PAYPAL GATEWAY
		 response.write "<form action=""transfer_payment.asp"" method=""post"">" & vbcrlf

		 response.write "<input type=""hidden"" name=""cmd"" value=""_xclick"">" & vbcrlf
		' May need for PayPal OrgId solution, but for now orgid is in table of paypal options 1/11/2006
		' response.write "<input type=""hidden"" name=""iorgid"" value="""& iorgid & """>"
	
	Do While NOT oPayPalOptions.EOF 

		
		 'DYNAMIC VALUES
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
		
		 if oPayPalFields("amount") <> "" then
			   curValue  = formatnumber(oPayPalFields("amount"),2)
   			sDisabled = "DISABLED"
			   response.write "  <tr>" & vbcrlf
      response.write "      <td>" & vbcrlf
      response.write "          <b>Payment Amount: </b>" & vbcrlf
      response.write "      </td>" & vbcrlf
      response.write "      <td>" & vbcrlf
      response.write            curValue &"<input type=""hidden"" name=""amount"" maxlength=""200"" value=""" & curValue &""">" & vbcrlf
      response.write "      </td>" & vbcrlf
      response.write "  </tr>" & vbcrlf
  	else
		   	curValue  = ""
   			sDisabled = ""
			   response.write "  <input type=""hidden"" name=""ef:amount-text/req"" value=""Payment Amount"">" & vbcrlf
   			response.write "  <tr>" & vbcrlf
      response.write "      <td>" & vbcrlf
      response.write "          <b>Payment Amount: </b>" & vbcrlf
      response.write "      </td>" & vbcrlf
      response.write "      <td>" & vbcrlf
      response.write "          <input type=""text"" name=""AMOUNT"" maxlength=""200"" value=""" & curValue &""">" & vbcrlf
      response.write "      </td>" & vbcrlf
      response.write "  </tr>" & vbcrlf
 		end if

end if

Set oPayPalFields = Nothing

End Function

'------------------------------------------------------------------------------------------------------------
' FUNCTION GETPAYMENTFIELDS(IID)
'------------------------------------------------------------------------------------------------------------
Function GetPaymentFields(iID)

sSQL = "SELECT * "
sSQL = sSQL & " FROM egov_paymentfields "
sSQL = sSQL & " WHERE paymentserviceid = " & iID

Set oPaymentFields = Server.CreateObject("ADODB.Recordset")
oPaymentFields.Open sSQL, Application("DSN") , 3, 1
	
if NOT oPaymentFields.EOF then

  	do while NOT oPaymentFields.eof
	   	'DYNAMIC VALUES
   		'BUILD EASY VALIDATION STRING
    		if ISNULL(oPaymentFields("paymentvalidation")) OR oPaymentFields("paymentvalidation") = "" then
      			sValidation = ""
    		else
      			if oPaymentFields("paymentfieldtype")  = "radio" then
        				sValidation = "-radio/" & oPaymentFields("paymentvalidation") & "/req"
       		else
          		sValidation = "-text/" & oPaymentFields("paymentvalidation") & "/req"
         end if
      end if
	 
 	   'WRITE EASY FORM HIDDEN FIELD VALIDATION VALUE
  	   response.write "<input type=""hidden"" name=""ef:custom_" & oPaymentFields("paymentfieldsname") & sValidation & """ value=""" & oPaymentFields("paymentfielddisplayname") & """>" & vbcrlf

 	   'WRITE ACTUAL FORM VALUE	
  	   select Case oPaymentFields("paymentfieldtype") 
       			Case "radio"
          			'Radio Buttons
          				response.write "  <tr>" & vbcrlf
              response.write "      <td colspan=""3"" align=""left"" nowrap=""nowrap"">" & vbcrlf
              response.write "          <strong>" & oPaymentFields("paymentfielddisplayname") & "</strong> &ndash; Select from the following list." & vbcrlf
              response.write "      </td>" & vbcrlf
              response.write "  </tr>" & vbcrlf

          				arrAnswers = split(oPaymentFields("answerlist"),chr(10))

          				for alist = 0 to ubound(arrAnswers)
             					arrAnswers(alist) = RemoveNewLine(arrAnswers(alist))
             					response.write "  <tr>" & vbcrlf
                  response.write "      <td colspan=""3"" align=""left"" nowrap=""nowrap"">" & vbcrlf
                  response.write "          <input value=""" & arrAnswers(alist) & """ name=""custom_" & oPaymentFields("paymentfieldsname") & """ class=""formradio"" type=""radio"" />" & arrAnswers(alist) & vbcrlf
                  response.write "      </td>" & vbcrlf
                  response.write "  </tr>" & vbcrlf
          				next

				          response.write "  <tr>" & vbcrlf
              response.write "      <td colspan=""3"">&nbsp;</td>" & vbcrlf
              response.write "  <tr>" & vbcrlf

          Case "textarea"
          			'TEXTAREA
           			response.write "  <tr>" & vbcrlf
              response.write "      <td colspan=""2"" align=""left"" nowrap=""nowrap"">" & vbcrlf
              response.write "          <b>" & oPaymentFields("paymentfielddisplayname") & " :</b><br />" & vbcrlf
              response.write "          <textarea " & oPaymentFields("paymentfieldattributes") & " name=""custom_" & oPaymentFields("paymentfieldsname") & """ style=""" & oPaymentFields("paymentfieldstyle") & """ class=""formtextarea""></textarea>" & vbcrlf
              response.write "      </td>" & vbcrlf
              response.write "      <td>" & vbcrlf
              response.write            oPaymentFields("paymentdesc") & " &nbsp;" & vbcrlf
              response.write "      </td>" & vbcrlf
              response.write "  </tr>" & vbcrlf 

          Case else
          			'SPECIAL CASE FOR TRACKING NUMBERS
           			if sOrgRegistration AND request.cookies("userid") <> "" AND request.cookies("userid") <> "-1" AND request.querystring("paymenttype") = 40 AND lcase(oPaymentFields("paymentfieldsname")) = "trackingnumber" then
             				sSQL = "SELECT action_autoid, submit_date "
                 sSQL = sSQL & " FROM egov_actionline_requests "
                 sSQL = sSQL & " WHERE userid = '" & request.cookies("userid") & "' "
                 sSQL = sSQL & " AND status IN ('WAITING','submitted','INPROGRESS','EVALFORM') AND category_id=295"

             				set oTrackingNumbers = Server.CreateObject("ADODB.Recordset")
         								oTrackingNumbers.Open sSQL, Application("DSN"), 3, 1

         								response.write "  <tr>" & vbcrlf
                 response.write "      <td align=""left"" nowrap=""nowrap"">" & vbcrlf
                 response.write "          <b>" & oPaymentFields("paymentfielddisplayname") & " :</b>" & vbcrlf
                 response.write "      </td>" & vbcrlf
                 response.write "      <td align=""left"">" & vbcrlf
                 response.write "          <select " & oPaymentFields("paymentfieldattributes") & " name=""custom_" & oPaymentFields("paymentfieldsname") & """ style=""" & oPaymentFields("paymentfieldstyle") & """>" & vbcrlf

         								do while NOT oTrackingNumbers.eof
                				lngTrackingNumber = oTrackingNumbers("action_autoid") & replace(FormatDateTime(cdate(oTrackingNumbers("submit_date")),4),":","")
               					response.write "            <option value=""" & lngTrackingNumber & """"

               					if request.querystring(oPaymentFields("paymentfieldsname")) = lngTrackingNumber then
                 						response.write " selected "
               					end if

               					response.write ">" & lngTrackingNumber & "</option>" & vbcrlf

                				oTrackingNumbers.MoveNext
             				loop

             				response.write "      </td>" & vbcrlf
                 response.write "      <td width=""40%"">" & vbcrlf
                 response.write            oPaymentFields("paymentdesc") & " &nbsp;" & vbcrlf
                 response.write "      </td>" & vbcrlf
                 response.write "  </tr>" & vbcrlf

              		'DEFAULT IS TEXTBOX
           			else
             				response.write "  <tr>" & vbcrlf
                 response.write "      <td align=""left"" nowrap=""nowrap"">" & vbcrlf
                 response.write "          <b>" & oPaymentFields("paymentfielddisplayname") & " :</b>" & vbcrlf
                 response.write "      </td>" & vbcrlf
                 response.write "      <td align=""left"">" & vbcrlf
                 response.write "          <input type=""text"" " & oPaymentFields("paymentfieldattributes") & " name=""custom_" & oPaymentFields("paymentfieldsname") & """ style=""" & oPaymentFields("paymentfieldstyle") & """ value=""" & request.querystring(oPaymentFields("paymentfieldsname")) & """ />" & vbcrlf
                 response.write "      </td>" & vbcrlf
                 response.write "      <td width=""40%"">" & vbcrlf
                 response.write            oPaymentFields("paymentdesc") & " &nbsp;" & vbcrlf
                 response.write "      </td>" & vbcrlf
                 response.write "  </tr>" & vbcrlf 
              end if

  	   end select

    		oPaymentFields.MoveNext
   loop

end if

set oPaymentFields = nothing

end function

'------------------------------------------------------------------------------------------------------------
' FUNCTION RemoveNewLine(sString)
'------------------------------------------------------------------------------------------------------------
 function RemoveNewLine( sString)
	 'remove the characters that make up the newline and carriage return from the string
	  RemoveNewLine = replace(sString,chr(10),"")
	  RemoveNewLine = replace(sString,chr(13),"")
 end function

'------------------------------------------------------------------------------------------------------------
' FUNCTION FN_DISPLAYPAYMENTS()
'------------------------------------------------------------------------------------------------------------
Function fn_DisplayPayments()
	response.write "<ul>" & vbcrlf

	sSQL = "SELECT DISTINCT TOP 5 paymentserviceid,paymentservicename "
 sSQL = sSQL & " FROM egov_paymentservices "
 sSQL = sSQL & " LEFT OUTER JOIN egov_organizations_to_paymentservices "
 sSQL = sSQL & " ON egov_paymentservices.paymentserviceid=egov_organizations_to_paymentservices.paymentservice_id "
 sSQL = sSQL & " WHERE (egov_organizations_to_paymentservices.paymentservice_enabled <> 0 "
 sSQL = sSQL & " AND (egov_organizations_to_paymentservices.orgid=" & iorgid
 sSQL = sSQL & " OR egov_organizations_to_paymentservices.orgid=0)) "
 sSQL = sSQL & " AND paymentserviceenabled = 1 "

	set oPaymentServices = Server.CreateObject("ADODB.Recordset")
	oPaymentServices.Open sSQL, Application("DSN") , 3, 1
	
	if NOT oPaymentServices.eof then
  		do while NOT oPaymentServices.eof
    			response.write "<li><a href=""payment.asp?paymenttype=" & oPaymentServices("paymentserviceid") & """>" & oPaymentServices("paymentservicename") &  "</a>" & vbcrlf
 	   		oPaymentServices.MoveNext
  		loop
	end if
	set oPaymentServices = nothing

	response.write "</ul>" & vbcrlf
end function

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
			GeteclinkPPDemoValues iPaymentServiceID,sServiceName

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
function GetSkipJackValues(iPaymentServiceID,sServiceName)
	 response.write "<form name=""frmpayment"" action=""PAYMENT_PROCESSORS/SKIPJACK2004/ECSKIPJACK_FORM.ASP"" method=""post"">" & vbcrlf
	 response.write "  <input type=""hidden"" name=""PAYMENTID"" value=""" & iPaymentServiceID & """>" & vbcrlf
	 response.write "  <input type=""hidden"" name=""ITEM_NUMBER"" value=""" & iPaymentServiceID & "00"">" & vbcrlf
	 response.write "  <input type=""hidden"" name=""ITEM_NAME"" value=""" & sServiceName & """>" & vbcrlf
end function

'------------------------------------------------------------------------------------------------------------
' FUNCTION GETVERISIGNVALUES(IPAYMENTSERVICEID,SSERVICENAME)
'------------------------------------------------------------------------------------------------------------
function GetVerisignValues(iPaymentServiceID,sServiceName)
 	response.write "<form name=""frmpayment"" action=""PAYMENT_PROCESSORS/VERISIGN2005/VERISIGN_FORM.ASP"" method=""post"">" & vbcrlf
	 response.write "  <input type=hidden name=""PAYMENTID"" value=""" & iPaymentServiceID & """>" & vbcrlf
 	response.write "  <input type=hidden name=""ITEM_NUMBER"" value=""" & iPaymentServiceID & "00"">" & vbcrlf
	 response.write "  <input type=hidden name=""ITEM_NAME"" value=""" & sServiceName & """>" & vbcrlf
end function

'------------------------------------------------------------------------------------------------------------
' FUNCTION GeteclinkPPDemoValues(IPAYMENTSERVICEID)
'------------------------------------------------------------------------------------------------------------
function GeteclinkPPDemoValues(iPaymentServiceID,sPaymentService)
 	response.write "<form name=""frmpayment"" action=""paypal_demo_pages/paypal_demo_page1.asp"" method=""GET"">" & vbcrlf
	
	'FORM HIDDEN PRODUCTS VALUES
 	response.write "  <input type=""hidden"" name=""item_number"" value=""" &iPaymentServiceID & """>" & vbcrlf
	 response.write "  <input type=""hidden"" name=""item_name"" value=""" & sPaymentService & """>" & vbcrlf

end function

'------------------------------------------------------------------------------------------------------------
' FUNCTION GetPayPalValues(iPaymentServiceID)
'------------------------------------------------------------------------------------------------------------
function GetPayPalValues(iPaymentServiceID)
	
	sSQL = "SELECT * FROM egov_paypaloptions WHERE paymentserviceid=" & iPaymentServiceID

	set oPayPalOptions = Server.CreateObject("ADODB.Recordset")
	oPayPalOptions.Open sSQL, Application("DSN") , 3, 1
		
	if NOT oPayPalOptions.eof then
			'STATIC VALUES
			 'response.write "<form name=""frmpayment"" action=""https://www.sandbox.paypal.com/cgi-bin/webscr"" method=""post"">"
			
			'PAYPAL GATEWAY
			 response.write "<form name=""frmpayment"" action=""transfer_payment.asp"" method=""post"">" & vbcrlf

			 response.write "  <input type=""hidden"" name=""cmd"" value=""_xclick"">" & vbcrlf

   	do while NOT oPayPalOptions.eof
   			'DYNAMIC VALUES
    			response.write "  <input type=""hidden"" name=""" & oPayPalOptions("paypaloptionname") & """ value=""" & oPayPalOptions("paypaloptionvalue") & """>" & vbcrlf

    			oPayPalOptions.MoveNext
  		loop
 end if

	set oPayPalOptions = nothing

end function

'------------------------------------------------------------------------------------------------------------
'Function JSsafe( strDB )
'------------------------------------------------------------------------------------------------------------
Function JSsafe( strDB )
  If Not VarType( strDB ) = vbString Then Jsafe = strDB : Exit Function
  strDB = replace( strDB, "'", "\'" )
  strDB = replace( strDB, chr(34), "\'" )
  strDB = replace( strDB, ";", "\;" )
  strDB = replace( strDB, "-", "\-" )
  strDB = replace( strDB, "(", "\(" )
  strDB = replace( strDB, ")", "\)" )
  strDB = replace( strDB, "/", "\/" )
  JSsafe = strDB
End Function

'------------------------------------------------------------------------------------------------------------
'Function Custom_Payment_FormID(iFormID)
'------------------------------------------------------------------------------------------------------------
Function Custom_Payment_FormID(iFormID)

	Select Case iFormID

		Case "21"

			'WASTERWATER FORM
			 response.write "  <tr>" & vbcrlf
    response.write "      <td colspan=""2"">" & vbcrlf
    response.write "          <table border=""0"">" & vbcrlf

    response.write "            <tr>" & vbcrlf
    response.write "                <td><b>Service Address :</b></td>" & vbcrlf
    response.write "                <td>" & vbcrlf
    response.write "                    <table border=""0"">" & vbcrlf
    response.write "                      <tr>" & vbcrlf
    response.write "                          <td><input maxlength=""5"" name=""custom_sa1"" style=""width: 50px"" type=""text"" maxlength=""5""></td>" & vbcrlf
    response.write "                          <td><input maxlength=""20"" name=""custom_sa2"" type=""text"" maxlength=""20""></td>" & vbcrlf
    response.write "                          <td>" & vbcrlf
    response.write "                              <input maxlength=""4"" name=""custom_sa3"" style=""width: 50px"" type=""text"" maxlength=""4"">" & vbcrlf
    response.write "                              <font class=""example"">Ex: 123 Main Blvd</font>" & vbcrlf
    response.write "                          </td>" & vbcrlf
    response.write "                      </tr>" & vbcrlf
    response.write "                    </table>" & vbcrlf
    response.write "                </td>" & vbcrlf
    response.write "            </tr>" & vbcrlf

    response.write "            <tr>" & vbcrlf
    response.write "                <td><b>Account Number :</b></td>" & vbcrlf
    response.write "                <td>" & vbcrlf
    response.write "                    <table border=""0"">" & vbcrlf
    response.write "                      <tr>" & vbcrlf
    response.write "                          <td><input maxlength=""9"" name=""custom_an1"" style=""width: 100px"" type=""text"" maxlength=""9""></td>" & vbcrlf
    response.write "                          <td><b> - </b></td>" & vbcrlf
    response.write "                          <td>" & vbcrlf
    response.write "                              <input maxlength=""9"" name=""custom_an2"" style=""width: 100px"" type=""text"" maxlength=""9"">" & vbcrlf
    response.write "                              <font class=""example"">Ex: 286-1384</font>" & vbcrlf
    response.write "                          </td>" & vbcrlf
    response.write "                      </tr>" & vbcrlf
    response.write "                    </table>" & vbcrlf
    response.write "                </td>" & vbcrlf
    response.write "            </tr>" & vbcrlf

    response.write "            <tr>" & vbcrlf
    response.write "                <td><b>Payment Amount :</b></td>" & vbcrlf
    response.write "                <td>" & vbcrlf
    response.write "                    <table border=""0"">" & vbcrlf
    response.write "                      <tr>" & vbcrlf
    response.write "                          <td colspan=""2""><input name=""custom_paymentamount"" type=""text""></td>" & vbcrlf
    response.write "                      </tr>" & vbcrlf
    response.write "                      <tr>" & vbcrlf
    response.write "                          <td></td>" & vbcrlf
    response.write "                          <td></td>" & vbcrlf
    response.write "                      </tr>" & vbcrlf
    response.write "                    </table>" & vbcrlf
    response.write "                </td>" & vbcrlf
    response.write "            </tr>" & vbcrlf

    response.write "          </table>" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf


			'FORM VALIDATION
			 response.write "  <input type=""hidden"" name=""ef:custom_sa1-text/number/req"" value=""Street Number"">" & vbcrlf
			 response.write "  <input type=""hidden"" name=""ef:custom_sa2-text/req"" value=""Street Name"">" & vbcrlf
			 response.write "  <input type=""hidden"" name=""ef:custom_sa3-text"" value=""Suffix"">" & vbcrlf
			 response.write "  <input type=""hidden"" name=""ef:custom_an1-text/req/ninedigits"" value=""Account Number Part 1"">" & vbcrlf
			 response.write "  <input type=""hidden"" name=""ef:custom_an2-text/req/ninedigits"" value=""Account Number Part 2"">" & vbcrlf
			 response.write "  <input type=""hidden"" name=""ef:custom_paymentamount-text/req"" value=""Payment Amount"">" & vbcrlf

		Case "22"

			'CITY BUILDING PERMIT PAYMENTS
			 response.write "  <tr>" & vbcrlf
    response.write "      <td colspan=""2"">" & vbcrlf

  			 DrawInputTable(10)

			 response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf

		Case "23"

			'SPECIAL ASSESSMENT PAYMENTS
			 response.write "  <tr>" & vbcrlf
    response.write "      <td colspan=""2"">" & vbcrlf
    response.write "          <table border=""0"">" & vbcrlf

    response.write "            <tr>" & vbcrlf
    response.write "                <td><b>Parcel Number :</b></td>" & vbcrlf
    response.write "                <td>" & vbcrlf
    response.write "                    <table border=""0"">" & vbcrlf
    response.write "                      <tr>" & vbcrlf
    response.write "                          <td><input maxlength=""3"" name=""custom_pn1"" style=""width: 50px"" type=""text""></td>" & vbcrlf
    response.write "                          <td><b> - </b></td>" & vbcrlf
    response.write "                          <td><input maxlength=""2"" name=""custom_pn2"" type=""text"" style=""width: 50px""></td>" & vbcrlf
    response.write "                          <td><b> - </b></td>" & vbcrlf
    response.write "                          <td>" & vbcrlf
    response.write "                              <input maxlength=""4"" name=""custom_pn3"" style=""width: 50px"" type=""text"">" & vbcrlf
    response.write "                              <font class=""example"">Ex: 123-12-124B</font>" & vbcrlf
    response.write "                          </td>" & vbcrlf
    response.write "                      </tr>" & vbcrlf
    response.write "                    </table>" & vbcrlf
    response.write "                </td>" & vbcrlf
    response.write "            </tr>" & vbcrlf

			 response.write "            <tr>" & vbcrlf
    response.write "                <td><b>Assessment Number :</b></td>" & vbcrlf
    response.write "                <td>" & vbcrlf
    response.write "                    <table border=""0"">" & vbcrlf
    response.write "                      <tr>" & vbcrlf
    response.write "                          <td>" & vbcrlf
    response.write "                              <input maxlength=""30"" name=""custom_an1"" style=""width: 150px"" type=""text"" maxlength=""9"">" & vbcrlf
    response.write "                               <!--</td><td><b> - </b></td><td><input maxlength=9 name=""custom_an2"" style=""width: 100px"" type=text maxlength=9>-->" & vbcrlf
    response.write "                              <font class=""example"">Ex: 12-10986</font>" & vbcrlf
    response.write "                          </td>" & vbcrlf
    response.write "                      </tr>" & vbcrlf
    response.write "                    </table>" & vbcrlf
    response.write "                </td>" & vbcrlf
    response.write "            </tr>" & vbcrlf

			 response.write "            <tr>" & vbcrlf
    response.write "                <td><b>Assessment Name :</b></td>" & vbcrlf
    response.write "                <td>" & vbcrlf
    response.write "                    <table border=""0"">" & vbcrlf
    response.write "                      <tr>" & vbcrlf
    response.write "                          <td><input name=""custom_assessmentname"" type=""text""></td>" & vbcrlf
    response.write "                          <td>&nbsp;</td>" & vbcrlf
    response.write "                      </tr>" & vbcrlf
    response.write "                    </table>" & vbcrlf
    response.write "                </td>" & vbcrlf
    response.write "            </tr>" & vbcrlf

			 response.write "            <tr>" & vbcrlf
    response.write "                <td><b>Payment Amount :</b></td>" & vbcrlf
    response.write "                <td>" & vbcrlf
    response.write "                    <table border=""0"">" & vbcrlf
    response.write "                      <tr>" & vbcrlf
    response.write "                          <td><input name=""custom_paymentamount"" type=""text""></td>" & vbcrlf
    response.write "                          <td>&nbsp;</td>" & vbcrlf
    response.write "                      </tr>" & vbcrlf
    response.write "                      <tr>" & vbcrlf
    response.write "                          <td></td>" & vbcrlf
    response.write "                          <td></td>" & vbcrlf
    response.write "                      </tr>" & vbcrlf
    response.write "                    </table>" & vbcrlf
    response.write "                </td>" & vbcrlf
    response.write "            </tr>" & vbcrlf

			 response.write "          </table>" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf

			'FORM VALIDATION
			 response.write "  <input type=""hidden"" name=""ef:custom_pn1-text/req/threedigits"" value=""Parcel Number Part 1"">" & vbcrlf
			 response.write "  <input type=""hidden"" name=""ef:custom_pn2-text/req/twodigits"" value=""Parcel Number Part 2"">" & vbcrlf
			 response.write "  <input type=""hidden"" name=""ef:custom_pn3-text/ppn/req"" value=""Parcel Number Part 3"">" & vbcrlf
			 response.write "  <input type=""hidden"" name=""ef:custom_an1-text/req"" value=""Assessment Number Part 1"">" & vbcrlf
			 'response.write "  <input type=""hidden"" name=""ef:custom_an2-text/number/ninedigits"" value=""Assessment Number Part 2"">" & vbcrlf
			 response.write "  <input type=""hidden"" name=""ef:custom_assessmentname-text/req"" value=""Assessment Name"">" & vbcrlf
			 response.write "  <input type=""hidden"" name=""ef:custom_paymentamount-text/req"" value=""Payment Amount"">" & vbcrlf

		Case else

	end select

end function

'------------------------------------------------------------------------------------------------------------
'FUNCTION DRAWINPUTTABLE(IROWS)
'------------------------------------------------------------------------------------------------------------
function DrawInputTable(iRows)
	
	 response.write "<table cellspacing=""0"" cellpadding=""5"" style=""border-top: solid 1px #000000; border-left: solid 1px #000000; border-right: solid 1px #000000"">" & vbcrlf
	
 'HEADER ROW WITH CAPTIONS
	 response.write "  <tr>" & vbcrlf
	 response.write "      <td style=""background-color: #2E1999; color: #FFFFFF; font-family: verdana,sans-serif; font-size: 10px; font-weight: bold; border-bottom: solid 1px #000000; border-right: solid 1px #000000;"">Permit Type<br><font size=""-2"">(B/Z/ROW/ENC/FP/ENG)</font></td>" & vbcrlf
	 response.write "      <td style=""background-color: #2E1999; color: #FFFFFF; font-family: verdana,sans-serif; font-size: 10px; font-weight: bold; border-bottom: solid 1px #000000; border-right: solid 1px #000000;"">Permit Number</td>" & vbcrlf
	 response.write "      <td style=""background-color: #2E1999; color: #FFFFFF; font-family: verdana,sans-serif; font-size: 10px; font-weight: bold; border-bottom: solid 1px #000000; border-right: solid 1px #000000;"">Project Address</td>" & vbcrlf
	 response.write "      <td style=""background-color: #2E1999; color: #FFFFFF; font-family: verdana,sans-serif; font-size: 10px; font-weight: bold; border-bottom: solid 1px #000000;"">Permit Amount<br><font size=""-2"">(Dont enter , $) </font></td>" & vbcrlf
	 response.write "  </tr>" & vbcrlf
	
	'LOOP AND DRAW DATA INPUT ROWS
	 for iRow = 1 to iRows 
		    response.write "  <tr>" & vbcrlf

   		'DRAW PERMIT TYPE
    		response.write "      <td align=""center"" style=""border-right: solid 1px #000000; border-bottom: solid 1px #000000;"">" & vbcrlf
      response.write "          <b>" & iRow & ".</b> "
      response.write "          <select name=""CUSTOM_PT" & iRow & """ style=""width:100px;"">" & vbcrlf
  				response.write "            <option value=""NONE"">Select...</option>" & vbcrlf
  				response.write "            <option value=""B"">B</option>" & vbcrlf
  				response.write "            <option value=""Z"">Z</option>" & vbcrlf
  				response.write "            <option value=""ROW"">ROW</option>" & vbcrlf
  				response.write "            <option value=""ENC"">ENC</option>" & vbcrlf
  				response.write "            <option value=""FP"">FP</option>" & vbcrlf
  				response.write "            <option value=""ENG"">ENG</option>" & vbcrlf
  				response.write "          </select>" & vbcrlf
  				response.write "      </td>" & vbcrlf

  			'DRAW PERMIT NUMBER INPUT
  				response.write "      <td align=""center"" style=""border-bottom: solid 1px #000000; border-right: solid 1px #000000;"">" & vbcrlf
      response.write "          <input maxlength=""7"" name=""CUSTOM_PN" & iRow & """ type=""text"" style=""width: 75px;"">" & vbcrlf
      response.write "      </td>" & vbcrlf

  			'DRAW PROJECT ADDRESS INPUT
  				response.write "      <td style=""border-bottom: solid 1px #000000; border-right: solid 1px #000000;"">" & vbcrlf
      response.write "          <input maxlength=""20"" name=""CUSTOM_PA" & iRow & """ type=""text"" style=""width: 250px;"">" & vbcrlf
      response.write "      </td>" & vbcrlf

  			'DRAW PERMIT AMOUNT
  				response.write "      <td style=""border-bottom: solid 1px #000000;"">" & vbcrlf
      response.write "          $<input name=""PA" & iRow & """ type=""text"" style=""width: 100px;"">" & vbcrlf
      response.write "      </td>" & vbcrlf
  		  response.write "  </tr>" & vbcrlf
  next

	'FOOTER ROW WITH TOTAL
	 response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""3"" align=""right"" style=""border-bottom: solid 1px #000000;"">" & vbcrlf
  response.write "          <b>Total Amount Paid:</b>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td colspan=""3"" align=""right"" style=""border-bottom: solid 1px #000000;"">" & vbcrlf
  response.write "          $<input name=""custom_paymentamount"" type=""text"" style=""width: 100px;"">" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
	 response.write "</table>" & vbcrlf

	'WARNING
	 response.write "<tr><td colspan=5  ><b><font color=red>*</font><i>You must select Permit Type and complete full row or permit for that row will be ignored.</i></b></TD></tr>"


	'DRAW JAVASCRIPT FUNCTION TO COMPUTE TOTAL
 	response.write "<script language=""Javascript"">" & vbcrlf

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
				sPermitNumber  = eval('document.frmpayment.CUSTOM_PN' + iCount +'.value');
				sPermitAddress = eval('document.frmpayment.CUSTOM_PA' + iCount +'.value');
				sPermitPrice   = eval('document.frmpayment.PA' + iCount +'.value');
				blnRowValid    = true;

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
response.write "</script>" & vbcrlf
end function
%>