<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
  Dim sError

 'Set Timezone information into session
  session("iUserOffset") = request.cookies("tz")

 'Override of value from common.asp
  sLevel = "../"

  controlid = DBSafe(request.querystring("control"))

  updatevalue = ""
  if request.servervariables("REQUEST_METHOD") = "POST" then
	'FORMAT IS "FIELDNAME : FIELDVALUE</br>"

	'for each item in request.form
		'if item <> "control" and item <> "paymentinfoid" then updatevalue = updatevalue & item & " : " & request.form(item) & "</br>"
'
	'next

	arrFields = split(request.form("fieldNames"),"|")

	for x = 0 to UBOUND(arrFields)
		updatevalue = updatevalue & arrFields(x) & " : " & request.form(arrFields(x)) & "</br>"
	next

	'UPDATE THE DB RECORD
	'value = applicantfirstname : Diane</br>applicantlastname : Bourassa</br>applicantaddress : 720 Milton Road M-4</br>applicantcity : Rye</br>applicantstate : NY</br>applicantzip : 10580</br>applicantphone : 917-701-4820</br>applicantemail : dbourassa@optonline.net</br>vehicle1license : GXJ1819</br>vehicle2license : </br>ownerfirstname : Diane </br>ownerlastname : Bourassa</br>owneraddress : 720Milton Road M-4</br>ownercity : Rye </br>ownerstate : NY</br>ownerzip : 10580</br>paymentamount : 760.00</br>permitholdertype : Current Railroad Permit Holder</br>
	sSQL = "UPDATE egov_paymentinformation SET payment_information = '" & DBSafeWithHTML(updatevalue) & "' WHERE paymentinfoid = '" & request.form("paymentinfoid") & "'"
	'response.write sSQL
	'response.end
	RunSQLStatement sSQL

	response.redirect "action_respond.asp?control=" & controlid
	  
  end if


%>
<html>
<head>
  <title><%=langBSHome%></title>

  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />

  <script language="javascript" src="../scripts/modules.js"></script>
  <script language="javascript" src="../scripts/ajaxLib.js"></script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content" style="width:auto;">
 	<div id="centercontent">

<table id="bodytable" border="0" cellpadding="0" cellspacing="0" class="start">
  <tr valign="top">
    	<td>
	<h2>Edit Payment Receipt</h2>

	<form action="#" method="POST">
	<input type="hidden" name="control" value="<%=controlid%>" />
	<table cellpadding="2">
	<%
		sSQL = "SELECT info.* " _
			& " FROM egov_payments pay " _
			& " INNER JOIN  egov_paymentinformation info ON info.paymentinfoid = pay.paymentinfoid  " _
			& " WHERE orgid = '" & session("orgid") & "' AND paymentid = '" & controlid & "'"
		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSQL, Application("DSN"), 3, 1
		If not oRs.EOF then
			response.write "<input type=""hidden"" name=""paymentinfoid"" value=""" & oRs("paymentinfoid") & """ />"
			sPaymentInfo =  oRs("payment_information")
			session("oldrecordvalues") = sPaymentInfo

			sFieldNames = ""
			If sPaymentInfo <> "" Then 
				sPayment = Replace( sPaymentInfo, "</br>", ":" )
				aPayment = Split(sPayment, ":" )

				fieldType = "text"
				if trim(aPayment(0)) = "paymentamount" or trim(aPayment(0)) = "permitholdertype" then fieldType = "hidden"

				if fieldType = "text" then response.write "<tr><td><b>" & aPayment(0) & "</b></td><td>"
				response.write "<input name=""" & trim(aPayment(0)) & """ type=""" & fieldType & """ size=""40"" value=""" & trim(aPayment(1)) & """ />"
				if fieldType = "text" then response.write "</td></tr>"
				
				sFieldNames = sFieldNames & "|" & trim(aPayment(0))

				iFieldCount = iFieldCount + 1
				ReDim PRESERVE aFieldNames( iFieldCount )
				aFieldNames( iFieldCount ) = aPayment(0)
				For x = 2 To UBound(aPayment)
					If x Mod 2 = 0 Then 
						If aPayment(x) <> "" Then 
							fieldType = "text"
							if trim(aPayment(x)) = "paymentamount" or trim(aPayment(x)) = "permitholdertype" then fieldType = "hidden"
				
							sFieldNames = sFieldNames & "|" & trim(aPayment(x))

							if fieldType = "text" then response.write "<tr><td><b>" & aPayment(x) & "</b></td><td>"
							response.write "<input name=""" & trim(aPayment(x)) & """ type=""" & fieldType & """ size=""40"" value=""" & trim(aPayment(x+1)) & """ />"
							if fieldType = "text" then response.write "</td></tr>"
							iFieldCount = iFieldCount + 1
							ReDim PRESERVE aFieldNames( iFieldCount )
							aFieldNames( iFieldCount ) = aPayment(x)
						End If 
					End If 
				Next 
			End If 
		end if
		oRs.Close
		Set oRs = Nothing

	%>
	</table>
	<input type="hidden" name="fieldNames" value="<%=Right(sFieldNames,len(sFieldNames)-1)%>" />
	<input type="submit" value="Save" style="font-size:14px;padding:5px;" />
	</form>

      </td>
  </tr>
</table>

  </div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
