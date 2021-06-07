<!-- #include file="includes/inc_dbfunction.asp" //-->
<!-- #include file="includes/common.asp" //-->
<!--#Include file="include_top_functions.asp"-->
<%
'--------------------------------------------------------------------------------------------------
' 
'--------------------------------------------------------------------------------------------------
' FILENAME:  TRANSFER_PAYMENT.ASP
' WRITTEN BY:  JOHN STULLENBERGER
' CREATED: 9/2/2004
' COPYRIGHT:  COPYRIGHT 2004 EC LINK, INC.  ALL RIGHTS RESERVED.
'
' DESCRIPTION:  
' THIS PAGE STORES ADDITIONAL PAYMENT INFORMATION TO SYNC WITH PAY PAL.
' 
' MODIFICATION HISTORY:
' 1.0  09/01/04   JOHN STULLENBERGER - INITIAL VERSION.
' 2.0  12/10/04   JOHN STULLENBERGER - REVISIONS TO ADD SKIPJACK LOGIC.
' 2.1	10/03/2008	Steve Loar	- Changed to insert payment data from open all then insert
' 2.2	4/13/2009	Steve Loar	- Added logging to paymentlog table
'--------------------------------------------------------------------------------------------------
' 
'--------------------------------------------------------------------------------------------------
Dim sInFields, iPaymentControlNumber
sInFields = ""

' May need for PayPal OrgId solution, but for now orgid is in table of paypal options 1/11/2006
' sInFields = "Orgid : " & request("iorgid") & "<br />"

' GET CUSTOM FIELDS 
For Each oField IN Request.Form
	If Left(oField,7) = "custom_" Then
		sDetails = sDetails & oField & " : " & request(oField) & "</br>"
	End If
	' This is for the log
	sInFields = sInFields & oField & " : " & request(oField) & "<br />"
Next 

iPaymentControlNumber = CreatePaymentsControlRow( "PayPal Payment Being Sent", "'Public'", "'Payments'" )

' POST DATA TO PAYPAL FOR PROCESSING
If sDetails <> "" Then
' IF THERE ARE CUSTOM FIELDS THEN STORE IN DATABASE AND PROVIDE PAY PAL WITH REFERENCE
	iReference = AddPaymentInformation( sDetails )

	'AddtoLog(iReference & " || " & sInFields)
	AddToPaymentsLog iPaymentControlNumber, iReference & " || " & sInFields, "'Public'", "'Payments'"

	sData = Request.Form
	sData =  sData & "&on0=Reference&os0=" & iReference
	' POST PAYPAL SYSTEM 
	response.write "<html>"& vbcrlf
	response.write "<head>" & vbcrlf
	response.write "<SCRIPT language=""JavaScript"">"& vbcrlf
	response.write "function submitform()"& vbcrlf
	response.write "{"& vbcrlf
	response.write "document.frmPaypal.submit();"& vbcrlf
	response.write "}"& vbcrlf
	response.write "</SCRIPT>" & vbcrlf
	response.write "</head>"& vbcrlf
	response.write "<body onload=""submitform();"">"& vbcrlf
	'response.write "<form name=""frmPaypal"" action=""https://www.sandbox.paypal.com/cgi-bin/webscr"" method=""post"">"& vbcrlf
	response.write "<form name=""frmPaypal"" action=""https://www.paypal.com/cgi-bin/webscr"" method=""post"">"& vbcrlf
	sOS1 = ""
	For Each oField IN Request.Form
		If oField <> "submit" Then
			response.write "<input type=hidden name=""" &  oField & """ value=""" & request(oField) & """></br>"& vbcrlf
			If request(oField) <> "" And LCase(Left(oField,6)) = "custom" And oField <> "custom_paymentamount" Then 
				'AddtoLog( iReference & " || " & oField & " = " & request(oField) )
				AddToPaymentsLog iPaymentControlNumber, iReference & " || " & oField & " = " & request(oField), "'Public'", "'Payments'"
				If sOS1 <> "" Then
					sOS1 = sOS1 & "."
				End If 
				sOS1 = sOS1 & Replace(Replace(Trim(request(oField)), """", "" )," ", "." )
			End If 
		End If
	Next
	response.write "<input type=hidden name=on0 value=""Reference"">"
	response.write "<input type=hidden name=os0 value=" & iReference & ">"
	response.write "<input type=hidden name=on1 value=""Payment Service Information"">"
	response.write "<input type=hidden name=os1 value=" & Left(sOS1, 200) & ">"
	response.write "<input type=hidden name=amount value=" & request("custom_paymentamount") & ">"
	response.write "</form>" & vbcrlf
	response.write "</body></html>"

Else
	
	' POST FORM DATA AS IS TO PAYPAL

	'AddtoLog("No PaymentInfoId || " & sInFields)
	AddToPaymentsLog iPaymentControlNumber, "No PaymentInfoId || " & sInFields, "'Public'", "'Payments'"

	response.write "<html>"& vbcrlf
	response.write "<head>" & vbcrlf
	response.write "<SCRIPT language=""JavaScript"">"& vbcrlf
	response.write "function submitform()"& vbcrlf
	response.write "{"& vbcrlf
	response.write "document.frmPaypal.submit();"& vbcrlf
	response.write "}"& vbcrlf
	response.write "</SCRIPT>" & vbcrlf
	response.write "</head>"& vbcrlf
	response.write "<body onload=""submitform();"">"& vbcrlf
	'response.write "<form name=""frmPaypal"" action=""https://www.sandbox.paypal.com/cgi-bin/webscr"" method=""post"">"& vbcrlf
	response.write "<form name=""frmPaypal"" action=""https://www.paypal.com/cgi-bin/webscr"" method=""post"">"& vbcrlf
	For Each oField IN Request.Form
		If oField <> "submit" Then
			response.write "<input type=hidden name=""" &  oField & """ value=""" & request(oField) & """></br>"& vbcrlf
		End If
	Next
		response.write "<input type=hidden name=amount value=" & request("custom_paymentamount") & ">"
	response.write "</form>" & vbcrlf
	response.write "</body></html>"
End If



'----------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'----------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Function AddPaymentInformation( sValue )
'-------------------------------------------------------------------------------------------------
Function AddPaymentInformation( ByRef sValue )
	Dim iReturnValue, oDetails, sInsertStatement
	
	iReturnValue = 0

'	Set oDetails = Server.CreateObject("ADODB.Recordset")
'	oDetails.CursorLocation = 3
'	oDetails.Open "egov_paymentinformation", Application("DSN"), 1, 2, 2

'	oDetails.AddNew
'	oDetails("payment_information") = paymentDBsafe(sValue)
'	oDetails.Update
'	iReturnValue = oDetails("paymentinfoid")

'	oDetails.Close
'	Set oDetails = Nothing

	sInsertStatement = "INSERT INTO egov_paymentinformation ( payment_information ) VALUES ( '" & paymentDBsafe(sValue) & "' )"
	iReturnValue = RunIdentityInsert( sInsertStatement )

	AddPaymentInformation = iReturnValue

End Function


'-------------------------------------------------------------------------------------------------
' Function paymentDBsafe( strDB )
'-------------------------------------------------------------------------------------------------
Function paymentDBsafe( strDB )
	Dim sNewString
	If Not VarType( strDB ) = vbString Then paymentDBsafe = strDB : Exit Function
	sNewString = Replace( strDB, "'", "''" )
	'sNewString = Replace( sNewString, "<", "&lt;" )
	paymentDBsafe = sNewString
End Function


'----------------------------------------------------------------------------------------
' FUNCTION ADDTOLOG(STEXT)
'----------------------------------------------------------------------------------------
Function AddtoLog(sText)
    ' WRITES SUPPLIED TEXT TO FILE WITH DATETIME
	Set oFSO = Server.Createobject("Scripting.FileSystemObject")
    Set oFile = oFSO.GetFile("C:\wwwroot\www.cityegov.com\egovlink300\transfertopaypal.log")
    Set oText = oFile.OpenAsTextStream(8)
    oText.WriteLine (Now() & Chr(9) & sText)
    oText.Close
    
    Set oText = Nothing
    Set oFile = Nothing
    Set oFSO = Nothing
End Function

%>
