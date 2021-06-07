<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: include_excel_export.asp
' AUTHOR: Steve Loar
' CREATED: 07/19/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This takes a passed recordset and creates an excel file from it.
'				See reporting/receipt_payment_export.asp as an example.
'
' MODIFICATION HISTORY
' 1.0	7/19/2007	Steve Loar	- Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

Sub CreateExcelDownload( ByVal sRtpTitle, ByVal sTotalRow  )

	' Include the next lines in your calling page. Change the filename generated.
'	sDate = Month(Date()) & Day(Date()) & Year(Date())
'	server.scripttimeout = 9000
'	Response.ContentType = "application/vnd.ms-excel"
'	Response.AddHeader "Content-Disposition", "attachment;filename=Receipt_payments_" & sDate & ".xls"


	If Not oSchema.EOF Then
		response.write "<html>"

		response.write vbcrlf & "<style>  "
		response.write " .moneystyle "
		response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
		response.write vbcrlf & "</style>"

		response.write "<body><table border=""0"">"

		' Write the title
		If sRtpTitle <> "" Then 
			response.write sRtpTitle
		End If 
		response.flush

		response.write "<tr>"
		' WRITE COLUMN HEADINGS
		For Each fldLoop in oSchema.Fields
			If fldLoop.Name <> "accountid" And fldLoop.Name <> "ispaymentaccount" And fldLoop.Name <> "iscitizenaccount" Then
				response.write  "<th>" & fldLoop.Name & "</th>"
			End If 
		Next
		response.write "</tr>"
		response.flush

		' WRITE DATA
		Do While NOT oSchema.EOF
			response.write "<tr>"
			For Each fldLoop in oSchema.Fields
				sFieldValue = trim(fldLoop.Value)
				
				' REMOVE LINE BREAKS
				If NOT IsNull(sFieldValue) Then
					sFieldValue = Replace(sFieldValue,chr(10),"")
					sFieldValue = Replace(sFieldValue,chr(13),"")
				End If

'				response.write "<td>" & sFieldValue & "</td>"

				If UCase(fldLoop.Name) = "CLASSSEASONID" Then 
					response.write "<td>" & getSeasonName(sFieldValue) & "</td>"
				Else 
					If fldLoop.Name <> "accountid" And fldLoop.Name <> "ispaymentaccount" And fldLoop.Name <> "iscitizenaccount" Then
						response.write "<td"
						If (InStr(sFieldValue,".") > 0 And IsNumeric(sFieldValue))  Or (fldLoop.Type = 6) Then 
							' If this is a money field then show the decimals 
							response.write " class=""moneystyle"""
						End If 	
						response.write ">" & sFieldValue & "</td>"
					End If 
				End If 

			Next
			response.write "</tr>"
			response.flush

			oSchema.MoveNext
		Loop
		
		' Total Row
		If sTotalRow <> "" Then 
			response.write sTotalRow
			response.flush
		End If 

		response.write "</table></body></html>"
		response.flush

	End If

End Sub


'--------------------------------------------------------------------------------------------------
' string 
'--------------------------------------------------------------------------------------------------
Function getSeasonName( ByVal iClassSeasonId )
	Dim sSql, oRs

	If iClassSeasonId <> "" Then 
		sSql = "SELECT seasonname FROM egov_class_seasons WHERE classseasonid = " & iClassSeasonId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then 
			getSeasonName = oRs("seasonname")
		Else
			getSeasonName = ""
		End If 

		oRs.Close
		Set oRs = Nothing 

	End If 

End Function 


%>
