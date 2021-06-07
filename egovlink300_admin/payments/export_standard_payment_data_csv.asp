<%
' BUILD QUERYSTRING FROM SEARCH PARAMETERS PASSED
'sSQL = "SELECT * FROM dbo.egov_payment_list" & Decode(request("options"))
sSQL = "SELECT * FROM dbo.egov_standard_payment_export WHERE [PAYMENT DATE] > '3/1/2006' AND ORGID=5"

' OPEN RECORDSET
Set oData = Server.CreateObject("ADODB.Recordset")
oData.Open sSQL, Application("DSN"), 3, 1

' IF NOT EMPTY PROCESS RESULT SET
If NOT oData.EOF Then

	' SET UP PAGE HEADER INFORMATION AS CSV FILE HEADER
	Response.Clear
	Response.ContentType = "text/csv"
	Response.AddHeader "Content-Disposition", "filename=EXPORT.CSV;"

	' ADD HEADERS TO FILE
	For Each fldLoop in oData.Fields
		' FILTER SELECTED FIELDS
		Select Case UCASE(fldLoop.Name)

			Case "PAYMENTSUMMARY", "PAYMENT_INFORMATION", "ORGID"
				' SKIP DISPLAYING THIS COLUMN HEADING
			Case Else
				' DISPLAY COLUMN HEADING
				response.write  fldLoop.Name & ","
		End Select
	Next

    ' LOOP THRU RECORDSET ADDING DATA TO FILE
	Do while NOT oData.EOF 
		
		' FOR EACH COLUMN DISPLAY DATA
		For Each fldLoop in oData.Fields
			Select Case UCASE(fldLoop.Name)

			Case "PAYMENTSUMMARY"
				Call subSeparatePayPalInformation(oData("paymentsummary"))

			Case "PAYMENT_INFORMATION"
				Call subSeparateFormFields(oData("payment_information"))

			Case "ORGID"
				' SKIP

			Case Else
				' DISPLAY VALUE IN DATABASE
				response.write fldLoop.Value & ","

			End Select
		Next
		
		' NEXT ROW
		response.write vbcrlf
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
				oDictionary.Add TRIM(arrNamedPair(0)),Trim(arrNamedPair(1))
			End If 
		Next

	End If

	' WRITEOUT DATA
	oCollection = oDictionary.Items
	For i=0 to oDictionary.Count-1
		Response.Write chr(34) & oCollection(i) & chr(34) & ","
	Next 

	Set oDictionary = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' SUB SUBSEPARATEFORMFIELDS
'--------------------------------------------------------------------------------------------------
Sub subSeparateFormFields(sText)

	' MAKE SURE THERE IS INFORMATION TO PARSE
	If sText <> "" Then
	
		' BREAK LIST INTO SEPARATE LINES
		arrInfo = SPLIT(sText, "</br>")

		' BREAK LINES INTO FIELD NAME AND VALUE
		For w = 0 to UBOUND(arrInfo)-1
		
			arrNamedPair = SPLIT(arrInfo(w),":")
			
			If ISARRAY(arrNamedPair) Then
				response.write  chr(34) & arrNamedPair(1) & chr(34) & ","
			End If

		Next

	End If

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
%>



