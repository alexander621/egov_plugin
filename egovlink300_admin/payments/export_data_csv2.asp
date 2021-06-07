<%
' BUILD QUERYSTRING FROM SEARCH PARAMETERS PASSED
sSQL = "SELECT * FROM dbo.egov_payment_list" & Decode(request("options"))

'sSQL = replace(sSQL,"orgid='11'","orgid='999'")


' OPEN RECORDSET
Set oData = Server.CreateObject("ADODB.Recordset")
oData.Open sSQL, Application("DSN"), 3, 1

' IF NOT EMPTY PROCESS RESULT SET
If NOT oData.EOF Then

	' SET UP PAGE HEADER INFORMATION AS CSV FILE HEADER
	Response.Clear
	Response.ContentType = "text/csv"
	Response.AddHeader "Content-Disposition", "filename=EXPORT.CSV;"

    ' LOOP THRU RECORDSET ADDING DATA TO FILE
	Do while NOT oData.EOF 
		' BASIC PAYMENT FORM INFORMATION
		response.write "BEGIN-FORM NAME" & ","
		response.write oData("paymentserviceid")  & ","
		response.write oData("paymentservicename") & ","
		response.write "END-FORM NAME" & ","
		
		' PAYMENT FORM FIELD INFORMATION
		response.write "BEGIN-FORM FIELDS" & ","
		If oData("paymentserviceid")=22 Then
			Call subWriteFormValues(oData("paymentsummary"),oData("paymentserviceid")) 
		Else
			Call subWriteFormValues(oData("payment_information"),oData("paymentserviceid")) 
		End If
		response.write "END-FORM FIELDS" & ","

		' TRANSACTION DETAILS
		response.write "BEGIN-TRANSACTION DETAILS" & ","
		response.write chr(34) & fnPaymentFormat(oData("paymentamount")) & chr(34) & ","
		response.write chr(34) & oData("paymentdate") & chr(34) & ","
		response.write chr(34) & oData("paymentid") & replace(FormatDateTime(oData("paymentdate"),4),":","") & chr(34) & ","
		response.write chr(34) & oData("paymentstatus") & chr(34) & ","
		response.write "END-TRANSACTION DETAILS" & ","

		' PAYPAL FIELDS (NEED TO ADD CODE TO CHECK FOR PAYMENT VENDOR)
		response.write "BEGIN-VENDOR INFORMATION" & ","
		Call subSeparatePayPalInformation(oData("paymentsummary"))
		response.write "END-VENDOR INFORMATION"

		' END ROW
		response.write  vbcrlf

		oData.MoveNext
	Loop

Else

	' NO DATA FOUND MATCHING CRITERIA SPECIFIED

End If



'----------------------------------------------------------------------------------------------------------------------
' CUSTOM FUNCTIONS AND SUBROUTINES
'----------------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB SUBSEPARATEPAYPALINFORMATION(STEXT)
'--------------------------------------------------------------------------------------------------
Sub subSeparatePayPalInformation(sText)

	' USED TO STORE DICTIONARY DATA
	Set oDictionary=Server.CreateObject("Scripting.Dictionary")

	' MAKE SURE THERE IS INFORMATION TO PARSE
	If sText <> "" Then
	
		' BREAK LIST INTO SEPARATE LINES
		arrInfo = SPLIT(sText, "</br>")

		' BREAK LINES INTO FIELD NAME AND VALUE
		For w = 0 to UBOUND(arrInfo)
			
			arrNamedPair = SPLIT(arrInfo(w),":")

			' MATCHED SETS ARE ADDED TO DICTIONARY
			If UBOUND(arrNamedPair) > 0 Then
				oDictionary.Add TRIM(arrNamedPair(0)),replace(Trim(arrNamedPair(1)),vbcrlf," ") 
			End If 
		Next

	End If


	'WRITE ONLY THESE SPECIFIC FIELDS
	response.write  fnPaymentFormat(TRIM(oDictionary("mc_gross"))) & ","
	response.write  oDictionary("payer_email") & ","
	response.write  oDictionary("option_name1") & ","
	response.write  oDictionary("address_status") &  ","
	response.write  oDictionary("address_street") &  ","
	response.write  oDictionary("charset") &  ","
	response.write  oDictionary("address_zip") &  ","
	response.write  oDictionary("address_country_code") &  ","
	response.write  oDictionary("address_name") &  ","
	response.write  oDictionary("payer_status") &  ","
	response.write  oDictionary("address_country") &  ","
	response.write  oDictionary("address_city") &  ","
	response.write  oDictionary("verify_sign") &  ","
	response.write  oDictionary("address_state") &  ","
	response.write  oDictionary("item_number") &  ","
	response.write  oDictionary("payer_id") &  ","
	response.write  oDictionary("first_name") &  ","
	response.write  fnPaymentFormat(TRIM(oDictionary("mc_fee"))) &  ","
	response.write  oDictionary("custom") &  ","
	response.write  oDictionary("receiver_email") &  ","
	response.write  oDictionary("receiver_id") &  ","
	response.write  oDictionary("tax") &  ","
	response.write  oDictionary("option_selection1") &  ","
	response.write  oDictionary("notify_version") &  ","
	response.write  oDictionary("payment_date") &  ","
	response.write  oDictionary("payment_status") &  ","
	response.write  oDictionary("payment_type") &  ","
	response.write  oDictionary("last_name") &  ","
	response.write  fnPaymentFormat(TRIM(oDictionary("payment_fee"))) &  ","
	response.write  fnPaymentFormat(TRIM(oDictionary("payment_gross"))) &  ","
	response.write  oDictionary("quantity") &  ","
	response.write  oDictionary("item_name") &  ","
	response.write  oDictionary("txn_id") &  ","
	response.write  oDictionary("txn_type") &  ","
	response.write  oDictionary("mc_currency") &  ","
	response.write  oDictionary("quantity") &  ","
	response.write  fnPaymentFormat(TRIM(oDictionary("shipping"))) & ","
	
	Set oDictionary = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' Function Decode(sIn)
'--------------------------------------------------------------------------------------------------
Function Decode(sIn)
    dim x, y, abfrom, abto
    Decode="": ABFrom = ""

    For x = 0 To 25: ABFrom = ABFrom & Chr(65 + x): Next 
    For x = 0 To 25: ABFrom = ABFrom & Chr(97 + x): Next 
    For x = 0 To 9: ABFrom = ABFrom & CStr(x): Next 

    abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
    For x=1 to Len(sin): y=InStr(abto, Mid(sin, x, 1))
        If y = 0 then
            Decode = Decode & Mid(sin, x, 1)
        Else
            Decode = Decode & Mid(abfrom, y, 1)
        End If
    Next
End Function



'--------------------------------------------------------------------------------------------------
' Function fnPaymentFormat(sValue)
'--------------------------------------------------------------------------------------------------
Function fnPaymentFormat(sValue)
	
	If NOT ISNULL(sValue) Then
		' REMOVE NON NUMBERIC CHARACTERS
		sValue = replace(sValue,chr(44),"")	' REMOVE COMMA
		sValue = replace(sValue,chr(36),"") ' REMOVE DOLLAR SIGN
		
		' ADD IMPLIED ZEROS AFTER DECIMAL POINT
		If instr(sValue,chr(46)) = 0 OR (LEN(sValue) = instr(sValue,chr(46))) Then
			sValue = sValue & "00"
		End If
		
		sValue = replace(sValue,chr(46),"") ' REMOVE DECIMAL POINT

		' ADD LEADING ZEROS
		sValue = RIGHT("0000000000" & sValue,7)
	End If

	' RETURN VALUE
	fnPaymentFormat = sValue

End Function 


'--------------------------------------------------------------------------------------------------
' Sub subWriteFormValues(sText,iFormID)
'--------------------------------------------------------------------------------------------------
Sub subWriteFormValues(sText,iFormID)

	' MAKE SURE THERE IS INFORMATION TO PARSE
	If sText <> "" Then
	
		Select Case iFormID

			Case "21"
				' WASTER WATER PAYMENT FORM
				' USED TO STORE DICTIONARY DATA
				Set oDictionary=Server.CreateObject("Scripting.Dictionary")

				' BREAK LIST INTO SEPARATE LINES
				arrInfo = SPLIT(sText, "</br>")

				' BREAK LINES INTO FIELD NAME AND VALUE
				For w = 0 to UBOUND(arrInfo)
					
					arrNamedPair = SPLIT(arrInfo(w),":")

					' MATCHED SETS ARE ADDED TO DICTIONARY
					If UBOUND(arrNamedPair) > 0 Then
						oDictionary.Add UCASE(TRIM(arrNamedPair(0))),Trim(arrNamedPair(1))
					End If 
				Next

			
				' WRITE OUT VALUES
				response.write ZeroFilledNumber(TRIM(oDictionary("CUSTOM_SA1")),5) & "," ' LOCATION PREFIX (5 DIGITS MAX)
				response.write RIGHT(TRIM(oDictionary("CUSTOM_SA2")),20) & ","	' LOCATION STREET (20 CHARS MAX)
				response.write RIGHT(TRIM(oDictionary("CUSTOM_SA3")),4)  & ","	' LOCATION SUFFIX (4 CHARS MAX)
				response.write ZeroFilledNumber(TRIM(oDictionary("CUSTOM_AN1")),9) & ","	' ACCOUNT # PART 1 (9 DIGITS MAX)
				response.write ZeroFilledNumber(TRIM(oDictionary("CUSTOM_AN2")),9) & ","	' ACCOUNT # PART 2 (9 DIGITS MAX)
				response.write fnPaymentFormat(TRIM(oDictionary("CUSTOM_PAYMENTAMOUNT"))) & "," ' PAYMENT AMOUNT

				Set oDictionary = Nothing

				' FINISH PAYMENT COLUMNS
				response.write STRING(35,chr(44))


			Case "23"
				' SPECIAL ASSESSMENT PAYMENTS
				' USED TO STORE DICTIONARY DATA
				Set oDictionary=Server.CreateObject("Scripting.Dictionary")

				' BREAK LIST INTO SEPARATE LINES
				arrInfo = SPLIT(sText, "</br>")

				' BREAK LINES INTO FIELD NAME AND VALUE
				For w = 0 to UBOUND(arrInfo)
					
					arrNamedPair = SPLIT(arrInfo(w),":")

					' MATCHED SETS ARE ADDED TO DICTIONARY
					If UBOUND(arrNamedPair) > 0 Then
						oDictionary.Add UCASE(TRIM(arrNamedPair(0))),Trim(arrNamedPair(1))
					End If 
				Next
			

				' WRITE OUT VALUES
				response.write ZeroFilledNumber(TRIM(oDictionary("CUSTOM_PN1")),3) & "," ' PARCEL NUMBER (=3 DIGITS)
				response.write ZeroFilledNumber(TRIM(oDictionary("CUSTOM_PN2")),2) & "," ' PARCEL NUMBER (=2 DIGITS)
				response.write TRIM(oDictionary("CUSTOM_PN3")) & "," ' PARCEL NUMBER (3-4 COMBINATION DIGITS AND CHAR EX. 123 OR 123B)
				response.write TRIM(oDictionary("CUSTOM_AN1")) & ","	' ASSESSMENT NUMBER PART 1 (30 CHAR ALPHANUMERIC WITH -)
				'response.write ZeroFilledNumber(TRIM(oDictionary("CUSTOM_AN2")),9)  & ","	' ASSESSMENT NUMBER PART 2 (9 DIGITS MAX)
				response.write RIGHT(TRIM(oDictionary("CUSTOM_ASSESSMENTNAME")),30) & ","	' ASSESSMENT NAME
				response.write fnPaymentFormat(TRIM(oDictionary("CUSTOM_PAYMENTAMOUNT"))) & "," ' PAYMENT AMOUNT

				Set oDictionary = Nothing

				' FINISH PAYMENT COLUMNS
				response.write STRING(35,chr(44))
			
			Case "22"
				' PERMIT PAYMENTS
				' USED TO STORE DICTIONARY DATA
				Set oDictionary=Server.CreateObject("Scripting.Dictionary")

				' BREAK LIST INTO SEPARATE LINES
				arrInfo = SPLIT(sText, "</br>")

				' BREAK LINES INTO FIELD NAME AND VALUE
				For w = 0 to UBOUND(arrInfo)
					
					arrNamedPair = SPLIT(arrInfo(w),":")

					' MATCHED SETS ARE ADDED TO DICTIONARY
					If UBOUND(arrNamedPair) > 0 Then
						oDictionary.Add UCASE(TRIM(arrNamedPair(0))),Trim(arrNamedPair(1))
					End If 
				Next
			
			
				' TEST WRITE DICTIONARY
				'arKeys = oDictionary.keys
				'For z=0 to oDictionary.count -1 
					'response.write "<BR>KEY = " & arKeys(z) & " -- VALUE = " & oDictionary.Item(arKeys(z))
				'Next

				
				' WRITE OUT VALUES
				For iFieldsList = 1 TO 10
					response.write RIGHT(TRIM(oDictionary("CUSTOM_PT"&iFieldsList)),4) & "," ' PERMIT TYPE (4 CHAR MAX)
					response.write ZeroFilledNumber(TRIM(oDictionary("CUSTOM_PN"&iFieldsList)),7) & ","	' PERMIT NUMBER (7 DIGIT MAX)
					response.write RIGHT(TRIM(oDictionary("CUSTOM_PA"&iFieldsList)),30)  & ","	' PROJECT ADDRESS (20 CHAR MAX)
					response.write fnPaymentFormat(TRIM(oDictionary("PA"&iFieldsList))) & "," ' PERMIT AMOUNT
				Next

				
				' FINISH PAYMENT COLUMNS
				response.write fnPaymentFormat(oDictionary("CUSTOM_PAYMENTAMOUNT")) & "," ' TOTAL AMOUNT PAID

				Set oDictionary = Nothing

			Case Else

				' BREAK LIST INTO SEPARATE LINES
				arrInfo = SPLIT(sText, "</br>")

				' BREAK LINES INTO FIELD NAME AND VALUE
				For w = 0 to 49
				
					If UBOUND(arrInfo) >= w Then
						arrNamedPair = SPLIT(arrInfo(w),":")
					Else
						 response.write chr(34) & chr(34) & ","
					End If
					
					If ISARRAY(arrNamedPair) Then
						If UBOUND(arrNamedPair) > 0 Then
							response.write chr(34) &  trim(arrNamedPair(1)) & chr(34) & ","
						End If
					End If

				Next

		End Select

	End If

End Sub


'--------------------------------------------------------------------------------------------------
' Function ZeroFilledNumber(sValue,iDigits)
'--------------------------------------------------------------------------------------------------
Function ZeroFilledNumber(sValue,iDigits)
	sPad = String(iDigits,"0")
	sValue = Right(sPad & sValue,iDigits)
	ZeroFilledNumber = sValue
End Function
%>



