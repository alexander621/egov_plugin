<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: simpleexport.asp
' AUTHOR: Steve Loar
' CREATED: 09/28/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Export of select fields for Rye, NY form 382 only. This is a simplified version of the 
'				normal export.
'
' MODIFICATION HISTORY
' 1.0   09/28/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	' SET UP PAGE OPTIONS
	sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
	server.scripttimeout = 9000
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment;filename=SimplePayments_" & sDate & ".xls"

	Dim sSearch, oRs

	sSearch = session("sPaymentSql")

	If sSearch <> "" Then 

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSearch, Application("DSN"), 3, 1

		If Not oRs.EOF Then
			oRs.MoveFirst
			response.write vbcrlf & "<html><body>"
			
			'response.write sSearch & "<br /><br />"

			response.write vbcrlf & "<table border=""1"" cellpadding=""4"">"

			response.write vbcrlf & "<tr>"
			'response.write "<td align=""center""><b>Payment Service Name</b></td>"
			response.write "<td align=""center""><b>Payment By</b></td>"
			response.write "<td align=""center""><b>Payment Date</b></td>"
			response.write "<td align=""center""><b>Permit Holder Type</b></td>"
			response.write "<td align=""center""><b>Applicant First Name</b></td>"
			response.write "<td align=""center""><b>Applicant Last Name</b></td>"
			response.write "<td align=""center""><b>Applicant Email</b></td>"
			response.write "<td align=""center""><b>Applicant Email 2</b></td>"
			response.write "<td align=""center""><b>Applicant Address</b></td>"
			response.write "<td align=""center""><b>Applicant City</b></td>"
			response.write "<td align=""center""><b>Applicant State</b></td>"
			response.write "<td align=""center""><b>Applicant Zip</b></td>"
			response.write "<td align=""center""><b>Applicant Phone</b></td>"
			response.write "<td align=""center""><b>Applicant Phone 2</b></td>"
			response.write "<td align=""center""><b>Vehicle 1 License</b></td>"
			response.write "<td align=""center""><b>Vehicle 2 License</b></td>"
			response.write "<td align=""center""><b>permit no.</b></td>"
			response.write "</tr>"
			response.flush
			
			' Loop here to get the payment data
			Do While Not oRs.EOF
				response.write "<tr>"
				'response.write "<td align=""center"">" & oRs("paymentservicename") & "</td>"
				response.write "<td align=""center"">" & Trim(oRs("userfname") & " " & oRs("userlname")) & "</td>"
				response.write "<td align=""center"">" & DateValue(CDate(oRs("paymentdate"))) & "</td>"
				'ShowPaymentInfoDetails oRs("payment_information")
				'response.write "<td align=""center"">" & oRs("payment_information") & "</td>"
				response.write "<td align=""center"">" & GetPaymentDetailValue( "permitholdertype", oRs("payment_information") ) & "</td>"
				response.write "<td align=""center"">" & GetPaymentDetailValue( "applicantfirstname", oRs("payment_information") ) & "</td>"
				response.write "<td align=""center"">" & GetPaymentDetailValue( "applicantlastname", oRs("payment_information") ) & "</td>"
				response.write "<td align=""center"">" & GetPaymentDetailValue( "applicantemail", oRs("payment_information") ) & "</td>"
				response.write "<td align=""center"">" & GetPaymentDetailValue( "applicantemail2", oRs("payment_information") ) & "</td>"
				response.write "<td align=""center"">" & GetPaymentDetailValue( "applicantaddress", oRs("payment_information") ) & "</td>"
				response.write "<td align=""center"">" & GetPaymentDetailValue( "applicantcity", oRs("payment_information") ) & "</td>"
				response.write "<td align=""center"">" & GetPaymentDetailValue( "applicantstate", oRs("payment_information") ) & "</td>"
				response.write "<td align=""center"">" & GetPaymentDetailValue( "applicantzip", oRs("payment_information") ) & "</td>"
				response.write "<td align=""center"">" & GetPaymentDetailValue( "applicantphone", oRs("payment_information") ) & "</td>"
				response.write "<td align=""center"">" & GetPaymentDetailValue( "applicantphone2", oRs("payment_information") ) & "</td>"
				response.write "<td align=""center"">" & GetPaymentDetailValue( "vehicle1license", oRs("payment_information") ) & GetPaymentDetailValue( "vehiclelicense", oRs("payment_information") ) & "</td>"
				response.write "<td align=""center"">" & GetPaymentDetailValue( "vehicle2license", oRs("payment_information") ) & "</td>"
				'response.write "<td align=""center"">" & aPayment(x) & "</td>"
				response.write "<td align=""center"">" & GetPermitNumber( oRs("paymentid") ) & "</td>"
				response.write "</tr>"
				response.flush
				oRs.MoveNext
			Loop
			
			response.write vbcrlf & "</table></body></html>"
			response.flush

		End If 

		oRs.Close
		Set oRs = Nothing 

	End If 


'-------------------------------------------------------------------------------------------------
' Sub ShowPaymentInfoHeader sPaymentInfo 
'-------------------------------------------------------------------------------------------------
Sub ShowPaymentInfoHeader( ByVal sPaymentInfo )
	Dim aPayment, sPayment

	If sPaymentInfo <> "" Then 
		sPayment = Replace( sPaymentInfo, "</br>", ":" )
		aPayment = Split(sPayment, ":" )
		response.write "<td align=""center""><b>" & aPayment(0) & "</b></td>"
		For x = 2 To UBound(aPayment)
			If x Mod 2 = 0 Then 
				If aPayment(x) <> "" And x <> 16 Then 
					response.write "<td align=""center""><b>" & aPayment(x) & "</b></td>"
				End If 
			End If 
			If x > 18 Then
				Exit For 
			End If 
		Next 
	End If 
End Sub 


'-------------------------------------------------------------------------------------------------
' Sub ShowPaymentInfoDetails sPaymentInfo 
'-------------------------------------------------------------------------------------------------
Sub ShowPaymentInfoDetails( ByVal sPaymentInfo )
	Dim aPayment, sPayment

	If sPaymentInfo <> "" Then 
		sPayment = Replace( sPaymentInfo, "</br>", ":" )
		aPayment = Split(sPayment, ":" )
		For x = 1 To UBound(aPayment)
			If x Mod 2 <> 0 And x <> 17 Then 
				response.write "<td align=""center"">" & aPayment(x) & "</td>"
			End If 
			If x > 18 Then
				Exit For 
			End If 
		Next 
	End If 
End Sub 


'-------------------------------------------------------------------------------------------------
' string GetPermitNumber( iPaymentid )
'-------------------------------------------------------------------------------------------------
Function GetPermitNumber( ByVal iPaymentid )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(assignedpermitnumber,0) AS assignedpermitnumber "
	sSql = sSql & "FROM egov_ryepermitrenewals WHERE paymentid = " & iPaymentid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("assignedpermitnumber")) > CLng(0) Then 
			GetPermitNumber = oRs("assignedpermitnumber")
		Else
			GetPermitNumber = ""
		End If 
	Else 
		oRs.Close
		Set oRs = Nothing 

		sSql = "SELECT ordernumber "
		sSql = sSql & "FROM egov_payments WHERE paymentid = '" & iPaymentid & "' and ordernumber IS NOT NULL"
		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1
		If Not oRs.EOF Then
			If CLng(oRs("ordernumber")) > CLng(0) Then 
				GetPermitNumber = oRs("ordernumber")
			Else
				GetPermitNumber = ""
			End If 
		Else 
			GetPermitNumber = ""
		end if
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function


'-------------------------------------------------------------------------------------------------
' string GetPaymentDetailValue( sLabel, sPaymentInfo )
'-------------------------------------------------------------------------------------------------
Function GetPaymentDetailValue( ByVal sLabel, ByVal sPaymentInfo )
	Dim aDetailinfo, sRow, aCells, sValue

	sValue = ""

	aDetailinfo = Split( sPaymentInfo, "</br>")
	For Each sRow In aDetailinfo
		If InStr( sRow & " :", sLabel & " :" ) > 0  Then 
			'sValue = sValue & "[sLabel=" & sLabel & "] "
			'sValue = sValue & "[sRow=" & sRow & "] "
			aCells = Split( sRow, ":" )
			sValue = sValue & Trim(aCells(1))
		End If 
	Next 

	GetPaymentDetailValue = sValue

End Function 


%>
