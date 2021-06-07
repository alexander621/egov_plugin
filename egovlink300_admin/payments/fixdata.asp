<%@LANGUAGE="VBScript"%>
<%
	' This was created by Steve Loar - 5 January 2005
	' It fixes data in a table so that the data will be formatted correctly in a CSV file for Bullhead City

Dim sPaymentSummary, sSQL, oData

' BUILD QUERYSTRING FROM SEARCH PARAMETERS PASSED
'sSQL = "SELECT paymentdetailid, paymentsummary FROM dbo.egov_paymentdetails where paymentdetailid > 1233 order by paymentdetailid" 
sSQL = "SELECT paymentinfoid, payment_information FROM dbo.egov_paymentinformation where paymentinfoid > 1975 order by paymentinfoid" 


' OPEN RECORDSET
Set oData = Server.CreateObject("ADODB.Recordset")
oData.Open sSQL, Application("DSN"), 1, 2

' IF NOT EMPTY PROCESS RESULT SET
If NOT oData.EOF Then

    ' LOOP THRU RECORDSET ADDING DATA TO FILE
	Do while NOT oData.EOF 
		'response.write "DetailId = " & oData("paymentdetailid") & "<br />"
		response.write "InfoId = " & oData("paymentinfoid") & "<br />"
		sPaymentSummary = replace(oData("payment_information"),"<br />","</br>")
		oData("payment_information") = sPaymentSummary
		oData.Update
		oData.MoveNext
	Loop

End If

oData.close
Set oData = nothing



%>