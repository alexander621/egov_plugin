<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: pagelogsummaryexport.asp
' AUTHOR: Steve Loar
' CREATED: 10/15/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of page stats, dumped to excel
'
' MODIFICATION HISTORY
' 1.0   10/15/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	' SET UP PAGE OPTIONS
	sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
	server.scripttimeout = 9000
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment;filename=Page_Log_Summary_" & sDate & ".xls"

	Dim sSearch, sRptTitle, sTodaySearch

	sSearch = session("sSql")
	sTodaySearch = session("sTodaySql")
	
	sRptTitle = vbcrlf & "<tr><th>Page Log Summary</th><th></th><th></th><th></th><th></th></tr>"

	DisplayPageLog sSearch, sRptTitle, sTodaySearch

	
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void DisplayPageLog sSearch, sRptTitle, sTodaySearch 
'--------------------------------------------------------------------------------------------------
Sub DisplayPageLog( ByVal sSearch, ByVal sRptTitle, ByVal sTodaySearch )
	Dim sSql, oRs, iRowCount, iTotalPageViews, iMaxLoadTime, iAvgLoadTime

	iTotalPageViews = CLng(0)
	iMaxLoadTime = CDbl(0.00)
	iAvgLoadTime = CDbl(0.00)

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSearch, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<html><body><table border=""1"">"
		response.write sRptTitle
		response.write "<tr><th>Log Date</th><th>Application Side</th><th>Page Views</th><th>Max Load Time(secs)</th><th>Avg Load Time(secs)</th></tr>"
		response.flush

		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr>"
			response.write "<td>" & FormatDateTime(oRs("logdate"),2) & "</td>"
			response.write "<td>" & oRs("applicationside") & "</td>"
			response.write "<td>" & oRs("pageviews") & "</td>"
			iTotalPageViews = iTotalPageViews + CLng(oRs("pageviews"))
			response.write "<td>" & FormatNumber(oRs("maxloadtime"),3,,,0) & "</td>"
			If iMaxLoadTime < CDbl(oRs("maxloadtime")) Then
				iMaxLoadTime = CDbl(oRs("maxloadtime"))
			End If 
			response.write "<td>" & FormatNumber(oRs("avgloadtime"),3,,,0) & "</td>"
			iAvgLoadTime = iAvgLoadTime + CDbl(FormatNumber(oRs("avgloadtime"),3,,,0))
			response.write "</tr>"
			response.flush
			oRs.MoveNext 
		Loop

		' If this is for the current month and year, then get a row for today
		If sTodaySearch <> "" Then
			ShowTodaysLog sTodaySearch, iTotalPageViews, iMaxLoadTime, iAvgLoadTime, iRowCount
		End If 

		' Show a Totals Row
		response.write vbcrlf & "<tr>"
		response.write "<td colspan=""2"">Totals</td>"
		response.write "<td>" & iTotalPageViews & "</td>"
		response.write "<td>" & iMaxLoadTime & "</td>"
		iAvgLoadTime = iAvgLoadTime / iRowCount
		response.write "<td>" & FormatNumber(iAvgLoadTime,3,,,0) & "</td>"
		response.write "</tr>"

		response.write "</table></body></html>"
		response.flush

	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowTodaysLog sSql, iTotalPageViews, iMaxLoadTime, iAvgLoadTime, iRowCount 
'--------------------------------------------------------------------------------------------------
Sub ShowTodaysLog( ByVal sSql, ByRef iTotalPageViews, ByRef iMaxLoadTime, ByRef iAvgLoadTime, ByRef iRowCount )
	Dim oRs

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRowCount = iRowCount + 1
		response.write vbcrlf & "<tr>"
		response.write "<td>" & FormatDateTime(oRs("logdate"),2) & "</td>"
		response.write "<td>" & oRs("applicationside") & "</td>"
		response.write "<td>" & oRs("pageviews") & "</td>"
		iTotalPageViews = iTotalPageViews + CLng(oRs("pageviews"))
		response.write "<td>" & FormatNumber(oRs("maxloadtime"),3,,,0) & "</td>"
		If iMaxLoadTime < CDbl(oRs("maxloadtime")) Then
			iMaxLoadTime = CDbl(oRs("maxloadtime"))
		End If 
		response.write "<td>" & FormatNumber(oRs("avgloadtime"),3,,,0) & "</td>"
		iAvgLoadTime = iAvgLoadTime + CDbl(FormatNumber(oRs("avgloadtime"),3,,,0))
		response.write "</tr>"
		oRs.MoveNext 
		response.flush
	Loop 

	oRs.Close
	Set oRs = Nothing 
End Sub

%>
