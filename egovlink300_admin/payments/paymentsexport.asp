<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: paymentsexport.asp
' AUTHOR: Steve Loar
' CREATED: 04/17/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  payments dumped to excel. This only does one payment service at a time.
'
' MODIFICATION HISTORY
' 1.0   04/17/2009	Steve Loar - INITIAL VERSION
' 2.0	03/16/2012	Steve Loar - modified to pull details by field name to correct sometimes random order of the detail data
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim oRs, sSearch, sDate, iFieldCount, aFieldNames(), x

	' SET UP PAGE OPTIONS
	sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
	server.scripttimeout = 9000
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment;filename=Payments_" & sDate & ".xls"

	sSearch = session("sPaymentSql")

	If sSearch <> "" Then 

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSearch, Application("DSN"), 3, 1

		If Not oRs.EOF Then
			oRs.MoveFirst
			response.write vbcrlf & "<html>"
			response.write vbcrlf & "<style>  "
			response.write " .moneystyle "
			response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
			response.write vbcrlf & "</style>"
			response.write vbcrlf & "<body><table border=""1"" cellpadding=""4"" cellspacing=""0"">"
			response.write vbcrlf & "<tr>"
			response.write "<td align=""center""><b>Payment Id</b></td>"
			response.write "<td align=""center""><b>Payment Service Name</b></td>"
			response.write "<td align=""center""><b>Payment Date</b></td>"
			response.write "<td align=""center""><b>Payment Amount</b></td>"
			response.write "<td align=""center""><b>Payment Ref</b></td>"
			response.write "<td align=""center""><b>Name</b></td>"
			response.write "<td align=""center""><b>Address</b></td>"
			response.write "<td align=""center""><b>City</b></td>"
			response.write "<td align=""center""><b>State</b></td>"
			response.write "<td align=""center""><b>Zip</b></td>"
			If Not IsNull(oRs("payment_information")) Then 
				iFieldCount = -1
				ShowPaymentInfoHeader oRs("payment_information"), iFieldCount
			End If 
			'For x = 0 To iFieldCount 
			'	response.write "<td align=""center"">" + aFieldNames( x ) + "</td>"
			'Next 
			response.write "</tr>"
			response.flush
			
			' Loop here to get the payment data
			Do While Not oRs.EOF
				response.write "<tr>"
				response.write "<td align=""center"">" & oRs("paymentid") & "</td>"
				response.write "<td align=""center"">" & oRs("paymentservicename") & "</td>"
				response.write "<td align=""center"">" & FormatDateTime(oRs("paymentdate"),2) & "</td>"
				response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oRs("paymentamount"),2,,,0) & "</td>"
				response.write "<td align=""center"">" & oRs("paymentrefid") & "</td>"
				response.write "<td align=""center"">" & Trim(oRs("userfname") & " " & oRs("userlname")) & "</td>"
				response.write "<td align=""center"">" & oRs("useraddress") & "</td>"
				response.write "<td align=""center"">" & oRs("usercity") & "</td>"
				response.write "<td align=""center"">" & oRs("userstate") & "</td>"
				response.write "<td align=""center"">" & oRs("userzip") & "</td>"
				'ShowPaymentInfoDetails oRs("payment_information")
				For x = 0 To iFieldCount 
					response.write "<td align=""center"">" + GetPaymentDetailValue( aFieldNames( x ), oRs("payment_information") ) + "</td>"
				Next 
				response.write "</tr>"
				response.flush
				oRs.MoveNext
			Loop
			
			response.write vbcrlf & "</table></body></html>"
			response.flush

		End If 
	End If 


'-------------------------------------------------------------------------------------------------
' void ShowPaymentInfoHeader sPaymentInfo, iFieldCount
'-------------------------------------------------------------------------------------------------
Sub ShowPaymentInfoHeader( ByVal sPaymentInfo, ByRef iFieldCount )
	Dim aPayment, sPayment

	If sPaymentInfo <> "" Then 
		sPayment = Replace( sPaymentInfo, "</br>", ":" )
		aPayment = Split(sPayment, ":" )
		response.write "<td align=""center""><b>" & aPayment(0) & "</b></td>"
		iFieldCount = iFieldCount + 1
		ReDim PRESERVE aFieldNames( iFieldCount )
		aFieldNames( iFieldCount ) = aPayment(0)
		For x = 2 To UBound(aPayment)
			If x Mod 2 = 0 Then 
				If aPayment(x) <> "" Then 
					response.write "<td align=""center""><b>" & aPayment(x) & "</b></td>"
					iFieldCount = iFieldCount + 1
					ReDim PRESERVE aFieldNames( iFieldCount )
					aFieldNames( iFieldCount ) = aPayment(x)
				End If 
			End If 
		Next 
	End If 
End Sub 


'-------------------------------------------------------------------------------------------------
' Sub ShowPaymentInfoDetails( sPaymentInfo )
'-------------------------------------------------------------------------------------------------
Sub ShowPaymentInfoDetails( ByVal sPaymentInfo )
	Dim aPayment, sPayment

	If sPaymentInfo <> "" Then 
		sPayment = Replace( sPaymentInfo, "</br>", ":" )
		aPayment = Split(sPayment, ":" )
		For x = 1 To UBound(aPayment)
			If x Mod 2 <> 0 Then 
				response.write "<td align=""center"">" & aPayment(x) & "</td>"
			End If 
		Next 
	End If 
End Sub 


'-------------------------------------------------------------------------------------------------
' string GetPaymentDetailValue( sLabel, sPaymentInfo )
'-------------------------------------------------------------------------------------------------
Function GetPaymentDetailValue( ByVal sLabel, ByVal sPaymentInfo )
	Dim aDetailinfo, sRow, aCells, sValue

	sValue = ""

	aDetailinfo = Split( sPaymentInfo, "</br>")
	For Each sRow In aDetailinfo
		If InStr( sRow, sLabel ) > 0  Then 
			aCells = Split( sRow, ":" )
			sValue = sValue & Trim(aCells(1))
			Exit For 
		End If 
	Next 

	GetPaymentDetailValue = sValue

End Function 



%>
