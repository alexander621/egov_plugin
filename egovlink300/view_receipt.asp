<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->


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
	<div align=left style="padding-bottom:20px;"> <% RegisteredUserDisplay("") %> </div>



<!--BEGIN: CONTENT MAIN-->
<div style=""margin-left:20px; "" class=box_header5 >Payment Receipt</div>

<% Call DisplayReceipt(CLng(request("PAYMENT_ID"))) %>


<!--END: CONTENT MAIN-->


<!--SPACING CODE-->
<p><bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->


<!--#Include file="include_bottom.asp"-->  




<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' FUNCTION DISPLAYRECEIPT(IPAYMENTID)
'--------------------------------------------------------------------------------------------------
Function DisplayReceipt(iPaymentID)

	' DISPLAY CONTACT INFORMATION
	sSQL = "SELECT * FROM PAYMENT_RECEIPT WHERE paymentid='" & iPaymentID & "'"
		
	Set oPayment = Server.CreateObject("ADODB.Recordset")
	oPayment.Open sSQL, Application("DSN") , 3, 1

	' CHECK FOR INFORMATION
	If NOT oPayment.EOF Then


		' DISPLAY PAYMENT TRANSACTION DETAILS
		response.write "<P><B>Payment Transaction Details</B><div class=box>"

		' USED TO STORE DICTIONARY DATA
		Set oDictionary=Server.CreateObject("Scripting.Dictionary")

		' MAKE SURE THERE IS INFORMATION TO PARSE
		sText = oPayment("paymentsummary")
		If sText <> "" Then
		
			' BREAK LIST INTO SEPARATE LINES
			arrInfo = SPLIT(sText, "</br>")

			' BREAK LINES INTO FIELD NAME AND VALUE
			For w = 0 to UBOUND(arrInfo)
				
				arrNamedPair = SPLIT(arrInfo(w),":")

				' MATCHED SETS ARE ADDED TO DICTIONARY
				If UBOUND(arrNamedPair) > 0 Then
					oDictionary.Add TRIM(arrNamedPair(0)),Trim(arrNamedPair(1))
				End If 
			Next

		End If

		response.write "<table>"
		response.write "<tr><td class=receipt>Payment Date:<td>" & oPayment("paymentdate") & "</td></tr>"
		response.write "<tr><td class=receipt>Product:<td>" & UCASE(oDictionary.Item("item_name")) & "</td></tr>"
		response.write "<tr><td class=receipt>Amount:<td>" & formatcurrency(oDictionary.Item("payment_gross")) & "</td></tr>"
		response.write "<tr><td class=receipt>Receipt ID:<td>" & UCASE(oDictionary.Item("receipt_id")) & "</td></tr>"
		response.write "<tr><td class=receipt>Transaction ID:<td>" & UCASE(oDictionary.Item("txn_id")) & "</td></tr>"
		response.write "<tr><td class=receipt>Payment Status:<td>" & UCASE(oDictionary.Item("payment_status")) & "</td></tr>"
		response.write "</table></p></div>"


		' DISPLAY PAYMENT FORM DETAILS
		response.write "<P><B>Payment Form Details</B>"
		response.write "<div class=box><table>"
		
		arrDetails = split(oPayment("payment_information"),"</br>")
		
		For d = 0 to ubound(arrDetails)
			If instr(UCASE(replace(arrDetails(d),"custom_","")),"PAYMENTAMOUNT") = 0 Then
				response.write "<tr><td>" & UCASE(replace(arrDetails(d),"custom_","")) & "</td></tr>"
			End If
		Next
		
		response.write "</table></div></p>"


			' DISPLAY PERSONAL INFORMATION
		response.write "<P><B>Personal Information Details</B>"
		response.write "<div class=box><table>"
		response.write "<tr><td class=receipt >FullName: </td><td>" & Ucase(oPayment("userfname") & " " & oPayment("userlname")) & "</td></tr>"
		response.write "<tr><td  class=receipt >Email: </td><td>" & ucase(oPayment("useremail")) & "</td></tr>"
		response.write "<tr><td  class=receipt >Address: </td><td>" & ucase(oPayment("useraddress")) & "</td></tr>"
		response.write "<tr><td class=receipt >City: </td><td>" & ucase(oPayment("usercity")) & "</td></tr>"
		response.write "<tr><td class=receipt >State: </td><td>" & ucase(oPayment("userstate")) & "</td></tr>"
		response.write "<tr><td class=receipt >Zip: </td><td>" & ucase(oPayment("userzip")) & "</td></tr>"
		response.write "</table></div></p>"

	Else
		' PAYMENT NOT FOUND REDIRECT TO HISTORY
		response.redirect("user_home.asp?trantype=1")
	End If

	Set oPayment = Nothing 

End Function
%>










