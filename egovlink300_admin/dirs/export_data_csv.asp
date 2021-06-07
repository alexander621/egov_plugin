<%
' BUILD QUERYSTRING 
sSQL = "SELECT * FROM dbo.egov_payment_list WHERE orgid=11"


' OPEN RECORDSET
Set oData = Server.CreateObject("ADODB.Recordset")
oData.Open sSQL, Application("DSN"), 3, 1

' LOOP TRHU DATA 
If NOT oData.EOF Then

	Do while NOT oData.EOF 
		' BUILD CSV ROW
		response.write chr(34) & oData("paymentserviceid") & chr(34) & ","
		response.write chr(34) & oData("paymentservicename") & chr(34) & ","
		
		' CODE FOR CUSTOM FIELDS GO HERE

		response.write chr(34) & oData("") & chr(34) & ","
		
		' TRANSACTION DETAILS
		response.write chr(34) & oData("paymentdate") & chr(34) & ","
		response.write chr(34) & oData("paymentid") & replace(FormatDateTime(Data("paymentdate"),4),":","") & chr(34) & ","
		oData.MoveNext
	Loop

Else

	' NO DATA FOUND MATCHING CRITERIA SPECIFIED

End If
%>



