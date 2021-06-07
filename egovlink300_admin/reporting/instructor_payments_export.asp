<%
	Dim sSql, oRequests, oSchema, iOldAccountId, dTotal, dTotalCredit, dTotalDebit, dGrandTotal
	Dim iLocationId, toDate, fromDate, sDateRange, iPaymentLocationId, iReportType, sAdminlocation
	Dim sFile, sRptTitle, sWhereClause

	' SET UP PAGE OPTIONS
	sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
	sWhereClause = ""

	If request("reporttype") = "" Then 
		iReportType = CLng(1)
	Else
		iReportType = CLng(request("reporttype"))
	End If 

	If iReportType = CLng(1) Then
		sRptTitle = vbcrlf & "<tr><th></th><th>Instructor Payments Summary</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
		sFile = "Summary_"
		sRptType = "Summary"
	Else
		sRptTitle = vbcrlf & "<tr><th></th><th>Instructor Payments Detail</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
		sFile = "Detail_"
		sRptType = "Detail"
	End If 

	server.scripttimeout = 9000
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment;filename=Instructor_Payments_" & sFile & sDate & ".xls"

	iSeasonId = CLng(request("seasonid"))
	sRptTitle = sRptTitle & vbcrlf & "<tr><th>Season: " & GetSeasonName( iSeasonId )  & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	sWhereClause = " AND classseasonid = " & iSeasonId

	If request("instructorid") = "" Then
		iInstructorId = 0
	Else
		iInstructorId = CLng(request("instructorid"))
	End If 

	If request("reporttype") = "" Then 
		iReportType = CLng(1)
	Else
		iReportType = CLng(request("reporttype"))
	End If 

	If iReportType = CLng(1) Then
		sRptType = "Summary"
	Else
		sRptType = "Detail"
	End If 


	' BUILD SQL WHERE CLAUSE
	If iInstructorId > CLng(0) Then
		sWhereClause = sWhereClause & " AND instructorid = " & iInstructorId
		sRptTitle = sRptTitle & vbcrlf & "<tr><th>Instructor: " & GetInstructorName( iInstructorId )  & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	Else
		sRptTitle = sRptTitle & vbcrlf & "<tr><th>Instructor: All Instructors</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	End If 

	If sRptType = "Detail" Then
		DisplayDetails sWhereClause, sRptTitle
	Else
		DisplaySummary sWhereClause, sRptTitle
	End If 



'--------------------------------------------------------------------------------------------------
' Sub DisplaySummary( varWhereClause, sRptTitle )
'--------------------------------------------------------------------------------------------------
Sub DisplaySummary( ByVal sWhereClause, ByVal sRptTitle )
	Dim sSql, oPayments, iOldInstructorId, dGrandTotal, dSubTotal

	iOldInstructorId = CLng(0)
	dGrandTotal = CDbl(0.00)
	dSubTotal = CDbl(0.00)

	sSql = "SELECT instructorid, firstname, lastname, classname, activityno, startdate, enddate, SUM(instructorpay) AS instructorpay "
	sSql = sSql & " FROM egov_instructor_payment_details "
	sSql = sSql & " WHERE orgid = " & session("orgid") & sWhereClause
	sSql = sSql & " GROUP BY instructorid, firstname, lastname, classname, activityno, startdate, enddate "
	sSql = sSql & " ORDER BY lastname, firstname, classname, activityno"
'	response.write sSql & "<br />"

	Set oPayments = Server.CreateObject("ADODB.Recordset")
	oPayments.Open sSQL, Application("DSN"), 3, 1

	If Not oPayments.EOF Then

		response.write "<html>"
		
		response.write vbcrlf & "<style>  "
		response.write " .moneystyle "
		response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
		response.write vbcrlf & "</style>"

		response.write "<body><table border=""1"">"
		response.write sRptTitle
		response.write vbcrlf & "<tr><th>Instructor</th><th>Class</th><th>Activity No.</th><th>Start<br />Date</th><th>End<br />Date</th><th>Payment</th></tr>"
		response.flush

		Do While Not oPayments.EOF

			If iOldInstructorId <> CLng(oPayments("instructorid")) Then 
				If iOldInstructorId <> CLng(0) Then
					' Sub total row for instructor
					response.write vbcrlf & "<tr><td></td><td></td><td></td><td></td><td>" & sInstructorName & " Total:</td>"
					response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dSubTotal, 2,,,0) & "</td>"
					response.write "</tr>"
					response.flush
				End If 
				iOldInstructorId = CLng(oPayments("instructorid"))
				dSubTotal = CDbl(0.00)
			End If 
			' Print out line
			response.write vbcrlf & "<tr>"
			sInstructorName = oPayments("firstname") & " " & oPayments("lastname")
			response.write "<td>" & sInstructorName & "</td>"
			response.write "<td>" & oPayments("classname") & "</td>"
			response.write "<td>" & oPayments("activityno") & "</td>"
			response.write "<td>" & oPayments("startdate") & "</td>"
			response.write "<td>" & oPayments("enddate") & "</td>"
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oPayments("instructorpay"),2,,,0) & "</td>"
			response.write "</tr>"

			dSubTotal = dSubTotal +  CDbl(FormatNumber(oPayments("instructorpay"),2,,,0))
			dGrandTotal = dGrandTotal +  CDbl(FormatNumber(oPayments("instructorpay"),2,,,0))
			response.flush
			
			oPayments.MoveNext
		Loop 
		' Sub total row for final instructor
		If iOldInstructorId <> CLng(0) Then
			response.write vbcrlf & "<tr><td></td><td></td><td></td><td></td><td>" & sInstructorName & " Total:</td>"
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dSubTotal,2,,,0) & "</td>"
			response.write "</tr>"
			response.flush
		End If 

		' Total for all instructors
		response.write vbcrlf & "<tr><td></td><td></td><td></td><td></td><td>Totals:</td>"
		response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dGrandTotal,2,,,0) & "</td>"
		response.write "</tr>"
		response.flush

		response.write vbcrlf & "</table></body></html>"
	End If 

	oPayments.Close
	Set oPayments = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub DisplayDetails( sWhereClause, sRptTitle )
'--------------------------------------------------------------------------------------------------
Sub DisplayDetails( ByVal sWhereClause, ByVal sRptTitle )
	Dim oDisplay, sSql, iOldInstructorId

	iOldInstructorId = CLng(0)
	dSubTotal = CDbl(0.00)
	dGrandTotal = CDbl(0.00)

	sSql = "SELECT  instructorid, firstname, lastname, classname, activityno, startdate, enddate, CASE isdropin WHEN 1 THEN 'Yes' WHEN 0 THEN '   ' END AS isdropin, "
	sSql = sSql & " paymentid, paymentdate, pricetypename, instructorpercent, amount, entrytype, instructorpay "
	sSql = sSql & " FROM egov_instructor_payment_details "
	sSql = sSql & " WHERE orgid = " & session("orgid") & sWhereClause
	sSql = sSql & " ORDER BY lastname, firstname, instructorid, classname, activityno"
'	response.write sSql & "<br />"

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open sSQL, Application("DSN"), 3, 1

	If Not oDisplay.EOF Then

		response.write "<html>"
		
		response.write vbcrlf & "<style>  "
		response.write " .moneystyle "
		response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
		response.write vbcrlf & "</style>"

		response.write "<body><table border=""1"">"
		response.write sRptTitle
		response.write vbcrlf & "<tr class=""tablelist""><th>Instructor</th><th>Class</th><th>Activity No.</th><th>Start<br />Date</th><th>End<br />Date</th><th>Receipt No.</th><th>Purchase<br />Date</th><th>Drop In</th><th>Pricing</th><th>Amount</th><th>Instr. %</th><th>Payment</th></tr>"
		response.flush

		Do While Not oDisplay.EOF
			If iOldInstructorId <> CLng(oDisplay("instructorid")) Then 
				' Put out a sub total row
				If iOldInstructorId <> CLng(0) Then
					response.write vbcrlf & "<tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>" & sInstructorName & " Total:</td>"
					response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dSubTotal, 2,,,0) & "</td>"
					response.write "</tr>"
					response.flush
				End If 
				iOldInstructorId = CLng(oDisplay("instructorid"))
				dSubTotal = CDbl(0.00)
			End If 

			response.write vbcrlf & "<tr>"
			sInstructorName = oDisplay("firstname") & " " & oDisplay("lastname")
			response.write "<td>" & sInstructorName & "</td>"
			response.write "<td>" & oDisplay("classname") & "</td>"
			response.write "<td>" & oDisplay("activityno") & "</td>"
			response.write "<td>" & FormatDateTime(oDisplay("startdate"),2) & "</td>"
			response.write "<td>" & FormatDateTime(oDisplay("enddate"),2) & "</td>"
			response.write "<td>" & oDisplay("paymentid") & "</td>"
			response.write "<td>" & FormatDateTime(oDisplay("paymentdate"),2) & "</td>"
			response.write "<td>" & oDisplay("isdropin") & "</td>"
			response.write "<td>" & oDisplay("pricetypename") & "</td>"
			
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oDisplay("amount"),2,,,0) & "</td>"
			response.write "<td>" & oDisplay("instructorpercent") & "%</td>"
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oDisplay("instructorpay"),2,,,0) & "</td>"
			response.write "</tr>"

			dSubTotal = dSubTotal +  CDbl(FormatNumber(oDisplay("instructorpay"),2,,,0))
			dGrandTotal = dGrandTotal +  CDbl(FormatNumber(oDisplay("instructorpay"),2,,,0))
			response.flush
			
			oDisplay.MoveNext
		Loop 

		' Put out a sub total row
		If iOldInstructorId <> CLng(0) Then
			response.write vbcrlf & "<tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>" & sInstructorName & " Total:</td>"
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dSubTotal,2,,,0) & "</td>"
			response.write "</tr>"
			response.flush
		End If 
		' Totals Row
		response.write vbcrlf & "<tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>Total:</td>"
		response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dGrandTotal,2,,,0) & "</td>"
		response.write "</tr>"
		response.flush

		response.write vbcrlf & "</table></body></html>"

	End If 

	oDisplay.Close
	Set oDisplay = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetLocationName( iLocationid )
'--------------------------------------------------------------------------------------------------
Function GetLocationName( ByVal iLocationid )
	Dim sSql, oLocation

	sSql = "Select name from egov_class_location where locationid = " & iLocationId

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open sSQL, Application("DSN"), 3, 1
	
	If Not oLocation.EOF Then 
		GetLocationName = oLocation("name")
	Else
		GetLocationName = ""
	End If 

	oLocation.Close 
	Set oLocation = Nothing

End Function 




'--------------------------------------------------------------------------------------------------
' Sub CreateDetailExcelDownload( sRtpTitle, sTotalRow )
'--------------------------------------------------------------------------------------------------
Sub CreateDetailExcelDownload( ByVal sRtpTitle, ByVal sTotalRow )
	' Pulled this in to make sub-totals

	iOldAccountId = CLng(0)
	iOldPaymentId = CLng(0)
	dTotal = CDbl(0.00)
	dTotalCredit = CDbl(0.00)
	dTotalDebit = CDbl(0.00)
	dGrandTotal = CDbl(0.00)
	dCreditSubTotal = CDbl(0.00)
	dDebitSubTotal = CDbl(0.00)
	dSubTotal = CDbl(0.00)

	If NOT oSchema.EOF Then
		response.write "<html>"
		
		response.write vbcrlf & "<style>  "
		response.write " .moneystyle "
		response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
		response.write vbcrlf & "</style>"

		response.write "<body><table border=""1"">"

		' Write the title
		If sRtpTitle <> "" Then 
			response.write sRtpTitle
		End If 

		response.write "<tr>"
		' WRITE COLUMN HEADINGS
		For Each fldLoop in oSchema.Fields
			If fldLoop.Name <> "accountid" Then 
				response.write  "<th>" & fldLoop.Name & "</th>"
			End If 
		Next
		response.write "</tr>"
		response.flush

		' WRITE DATA
		Do While NOT oSchema.EOF
			If CLng(oSchema("accountid")) <> iOldAccountId Then
				If iOldAccountId <> CLng(0) Then 
					' Sub Total Row
					response.write vbcrlf & "<tr><td></td><td></td><td>Sub-Total:</td>"
					response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dCreditSubTotal,2) & "</td>"
					response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(-dDebitSubTotal,2) & "</td>"
					response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dSubTotal,2) & "</td>"
					response.write "</tr>"
					response.flush
				End If 
				dCreditSubTotal = CDbl(0.00)
				dDebitSubTotal = CDbl(0.00)
				dSubTotal = CDbl(0.00)
				iOldAccountId = oSchema("accountid")
			End If 
			' Normal Row
			response.write "<tr>"
			For Each fldLoop in oSchema.Fields
				sFieldValue = trim(fldLoop.Value)
				
				' REMOVE LINE BREAKS
				If NOT ISNULL(sFieldValue) Then
					sFieldValue = replace(sFieldValue,chr(10),"")
					sFieldValue = replace(sFieldValue,chr(13),"")
				End If
				
				If fldLoop.Name = "creditamt" Then
					dCreditSubTotal = dCreditSubTotal + CDbl(sFieldValue)
					dSubTotal = dSubTotal + CDbl(sFieldValue)
				End If 
				If fldLoop.Name = "debitamt" Then
					dDebitSubTotal = dDebitSubTotal - CDbl(sFieldValue)
					dSubTotal = dSubTotal + CDbl(sFieldValue)
				End If 

				If fldLoop.Name <> "accountid" Then
					response.write "<td>" & sFieldValue & "</td>"
				End If 
			Next
			response.write "</tr>"
			response.flush
			 

			oSchema.MoveNext
		Loop
		
		' Sub Total Row
		response.write vbcrlf & "<tr><td></td><td></td><td>Sub-Total:</td>"
		response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dCreditSubTotal,2) & "</td>"
		response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(-dDebitSubTotal,2) & "</td>"
		response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dSubTotal,2) & "</td>"
		response.write "</tr>"
		response.flush

		' Total Row
		If sTotalRow <> "" Then 
			response.write sTotalRow
		End If 

		response.write "</table></body></html>"
	Else

		' NO DATA

	End If

End Sub


'--------------------------------------------------------------------------------------------------
' Sub GetSeasonName( iClassSeasonId )
'--------------------------------------------------------------------------------------------------
Function GetSeasonName( ByVal iClassSeasonId )
	Dim sSql, oSeasons

	sSQL = "Select seasonname From egov_class_seasons C Where classseasonid = " & iClassSeasonId
 
	Set oSeasons = Server.CreateObject("ADODB.Recordset")
	oSeasons.Open sSQL, Application("DSN"), 0, 1
	
	If Not oSeasons.EOF Then
		GetSeasonName = oSeasons("seasonname")
	Else
		GetSeasonName = ""
	End If

	oSeasons.close
	Set oSeasons = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetInstructorName( iInstrudtorId )
'--------------------------------------------------------------------------------------------------
Function GetInstructorName( ByVal iInstructorId )
	Dim sSql, oName

	sSQL = "select lastname, firstname From egov_class_instructor Where instructorid = " & iInstructorId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 3, 1
	
	If Not oName.EOF Then 
		GetInstructorName = oName("firstname") & " " & oName("lastname")
	Else
		GetInstructorName = ""
	End If 
	
	oName.close 
	Set oName = Nothing

End Function 


%>

<!-- #include file="../includes/adovbs.inc" -->

