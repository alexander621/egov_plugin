<%
'------------------------------------------------------------------------------------------------------------
' FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' void DISPLAY_RESULTS sSql,SOPTIONS
'------------------------------------------------------------------------------------------------------------
Sub Display_Results( ByVal sSql, ByVal sOptions )
	Dim sProcessingRoute

	' INITIALIZE VALUES
	iPageSize = 1000
	iCacheSize = 1000
	iCursorLocation = 3
	sScriptName = request.servervariables("SCRIPT_NAME")
	session("DISPLAYQUERY") = sSql
	
	
	' BUILD RECORDSET
	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.PageSize = iPageSize
	oRequests.CacheSize = iCacheSize
	oRequests.CursorLocation = iCursorLocation

	'response.write sSql & "<br /><br />"

	' OPEN RECORDSET
	oRequests.Open sSql, Application("DSN"), 3, 1
	
	' NAVIGATION VARIABLES
	abspage = oRequests.AbsolutePage
	pagecnt = oRequests.PageCount
	
	' CHECK FOR EMPTY RECORDSET
	If oRequests.EOF then
		' EMPTY
		response.write "<p><b>No records found</p>"
	Else
		' FOUND RECORDS, DISPLAY TO USER

		 ' SET PAGE TO VIEW
		If Len(Request("pagenum")) = 0 OR clng(Request("pagenum")) < 1  Then
			' DEFAULT FIRST PAGE VIEW
			oRequests.AbsolutePage = 1
		Else
			' DISPLAY CURRENTLY SELECTED PAGE
			If clng(Request("pagenum")) <= clng(oRequests.PageCount) Then
				oRequests.AbsolutePage = request("pagenum")
			Else
				' DEFAULT TO FIRST PAGE VIEW
				oRequests.AbsolutePage = 1
			End If
		End If

		' NAVIGATION BUTTONS
		Response.write "<table style=""margin-bottom: 5px;"">"
		response.write "<tr>"
		response.write "<td valign=""bottom""><a href=""" & sScriptName & "?pagenum=1" & sOptions & """><img border=0 src=""../images/nav_first.gif""></a><a href=""" & sScriptName & "?pagenum="&oRequests.AbsolutePage-1 & sOptions & """><img border=0 src=""../images/nav_back.gif""></a></td>"
		Response.Write "<td valign=""bottom""><b>Current page: <font style=""color:blue;"" > " & oRequests.AbsolutePage  & "</font> | " & vbcrlf
		Response.Write "Number of pages: <font  style=""color:blue;"">" &  oRequests.PageCount & "</font>  " & vbcrlf
		'Response.Write "Number of Rows: <font  style=""color:blue;"">" & oRequests.RecordCount
		Response.write "</b></td>"
		response.write "<td valign=""bottom""><a href=""" & sScriptName & "?pagenum="&oRequests.AbsolutePage+1 & sOptions & """><img border=0 src=""../images/nav_forward.gif"" valign=bottom></a><a href=""" & sScriptName & "?pagenum="&oRequests.PageCount & sOptions & """><img border=0 src=""../images/nav_last.gif"" valign=""bottom""></a></td>"
		response.write "<td align=""right"" valign=""bottom""><input type=""button"" class=""button excelexport"" value=""Export To Excel"" onClick=""location.href='csv_export.asp'"" /></td>"
		response.write "</tr>"
		response.write "</table>"
			
					
		' DISPLAY DATA	
		Response.Write "<table cellspacing=""0"" cellpadding=""2"" border=""0"" class=""excel style-alternate sortable-onload-1"" width=""100%"">"

		sProcessingRoute = LCase(GetProcessingRoute())
		' WRITE COLUMN HEADINGS
		response.write "<tr class=""excel""><th class=""excelheaderleft sortable"">&nbsp;</th>"
		For Each fldLoop in oRequests.Fields 
			If LEFT(fldLoop.Name,3) <> "ecG" Then
				response.write "<th class=""excelheader sortable"">"
				If LCase(fldLoop.Name) = "transaction id" Then
					If sProcessingRoute = "pointandpay" Then
						response.write "Order Number"
					Else
						response.write fldLoop.Name
					End If 
				Else 
					response.write fldLoop.Name
				End If 
				response.write "</th>"
			End If
		Next
		response.write "</tr >"

		' SET BASE RECORD COUNT
		iRecordNumber = (oRequests.AbsolutePage * iPageSize) - iPageSize

		' LOOP AND DISPLAY THE RECORDS
		For irows = 1 to oRequests.PageSize

			 If NOT oRequests.EOF Then
				
				bgcolor = "#eeeeee"
				iRecordNumber = iRecordNumber + 1
				blnSkipCheck = False
				
				response.write "<tr>"
				response.write "<td class=""exceldataleft"">" & iRecordNumber & "</td>"
				
				For Each fldLoop in oRequests.Fields 

					' NORMAL EXCEL STYLE
					If LEFT(fldLoop.Name,3) <> "ecG" AND blnSkipCheck <> True Then
						sClass = "exceldata"
					Else
						
						' ADD STYLE FOR SUBTOTAL ROW - IF ROW CONTAINS COLUMN NAMED ecG WITH 1 OR GREATER IT IS AT LEAST A SUBTOTAL ROW
						If LEFT(fldLoop.Name,3) = "ecG" AND LEFT(fldLoop.Name,4) <> "ecGT"  Then
							If Trim(fldLoop.Value) > 0 Then
								If sClass <> "excelgrandtotalrow" Then
									sClass = "exceltotalrow"
									blnSkipCheck = True
								End If
							End If
						End If

						' ADD STYLE FOR GRAND TOTAL ROW - IF ROW CONTAINS COLUMN NAMED ecGT WITH 1 OR GREATER IT IS THE GRAND TOTAL ROW
						If LEFT(fldLoop.Name,4) = "ecGT" Then
							If Trim(fldLoop.Value) > 0 Then
								sClass = "excelgrandtotalrow"
								blnSkipCheck = True
							End If
						End If
			
					End If 

					' SKIP DISPLAYING COLUMN FOR TOTAL AND SUBTOTAL RETURNED GROUPING COLUMNS
					If LEFT(fldLoop.Name,3) <> "ecG" Then
						iType = fldLoop.type
						response.write "<td class=""" & sClass &""">" & FormatData(iType,trim(fldLoop.Value)) & "&nbsp;</td>"
					End If

				Next
				
				response.write "</tr>"
				blnSkipCheck = False

				oRequests.MoveNext 
			End If
		Next	
		response.write vbcrlf & "</table>"
	End If

	oRequests.Close
	Set oRequests = Nothing 

End Sub


'------------------------------------------------------------------------------------------------------------
' string FORMATDATA(STYPE,SDATA)
'------------------------------------------------------------------------------------------------------------
Function FormatData( ByVal iType, ByVal sData )
	Dim sReturnValue
	
	' DEFAULT RETURN UNALTERED DATA
	sReturnValue = sData

	' FORMAT ACCORDING TO DATA TYPE
	Select Case iType
		Case 6
			' FORMAT DISPLAY AS CURRENCY
			If sData = "" OR IsNull(sData) Then
				' IF NULL OR EMPTY SET TO 0
				sData = 0
			End If
			sReturnValue = FormatCurrency(sData,2)
		
		Case 129
			' FORMAT DATE AS MONTH SHORT NAME AND YEAR
			If NOT isnull(sData) Then
				sReturnValue = Left(MonthName(RIGHT(sData,2)),3) & " " & LEFT(sData,4)
			End If

		Case Else
			' DO NOTHING
	End Select

	FormatData = sReturnValue 

End Function


'------------------------------------------------------------------------------------------------------------
' void SUBDRAWGROUPBYOPTIONS(sSql)
'------------------------------------------------------------------------------------------------------------
Sub subDrawGroupByOptions( ByVal sSql )
	Dim oSelectList

	Set oSelectList = Server.CreateObject("ADODB.Recordset")
	oSelectList.Open sSql, Application("DSN"), 3, 1
	
	If Not oSelectList.EOF Then
	
		' WRITE COLUMN HEADINGS
		response.write vbcrlf & "<select name=""groupbyfield"" multiple=""multiple"">"
		response.write vbcrlf & "<option value=""-1"">Do Not Summarize</option>" & "</option>"
		
		i = 0

		For Each fldLoop in oSelectList.Fields
			response.write vbcrlf & "<option value=""" & fldLoop.Name & """>" & UCASE(fldLoop.Name) & "</option>"
			i = i + 1
		Next
		response.write "</select>"
	End If

	oSelectList.Close
	Set oSelectList = Nothing

End Sub


'------------------------------------------------------------------------------------------------------------
' void DRAWDATECHOICES(SNAME)
'------------------------------------------------------------------------------------------------------------
Sub DrawDateChoices( ByVal sName )

	response.write vbcrlf & "<select onChange=""getDates(document.frmPFilter." & sName & ".value);"" class=""calendarinput"" name=""" & sName & """>"
	response.write vbcrlf & "<option value=""0"">Or Select Date Range from Dropdown...</option>"
	response.write vbcrlf & "<option value=""11"">This Week</option>"
	response.write vbcrlf & "<option value=""12"">Last Week</option>"
	response.write vbcrlf & "<option value=""1"">This Month</option>"
	response.write vbcrlf & "<option value=""2"">Last Month</option>"
	response.write vbcrlf & "<option value=""3"">This Quarter</option>"
	response.write vbcrlf & "<option value=""4"">Last Quarter</option>"
	response.write vbcrlf & "<option value=""6"">Year to Date</option>"
	response.write vbcrlf & "<option value=""5"">Last Year</option>"
	response.write vbcrlf & "<option value=""7"">All Dates to Date</option>"
	response.write vbcrlf & "</select>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' void GETFIELDCHOICES(sSql)
'------------------------------------------------------------------------------------------------------------
Sub GetFieldChoices( ByVal sSql )
	Dim oRs, sSqlSub

	sSqlSub = Right(sSql,Len(sSql) - instr(sSql," FROM "))

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<table>"
		For Each oColumn In oRs.Fields 
			response.write vbcrlf & "<tr>"
			response.write "<td align=""right""><b>" &  UCASE(oColumn.Name)  & ":</b></td>"
			response.write "<td>"
			response.write vbcrlf & "<select name=""filter_" &  UCASE(oColumn.Name) & """>"
			response.write vbcrlf & "<option value=""NO FILTER"">Do Not Filter</option>"
			ListFieldChoices "SELECT DISTINCT [" & oColumn.Name & "] " & sSqlSub 
			response.write "</select>"
			response.write "</td>"
			response.write "</tr>"
		Next
		response.write vbdrlf & "</table>"
	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------------------------------------
' void LISTFIELDCHOICES(sSql)
'------------------------------------------------------------------------------------------------------------
Sub ListFieldChoices( ByVal sSql )
	Dim oRs

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs.Fields(0).Value & """>" & oRs.Fields(0).Value & "</option>"
			oRs.MoveNext
		Loop
	End If

	oRs.Close 
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------------------------------------
' void LISTREPORTCHOICES(IREPORT)
'------------------------------------------------------------------------------------------------------------
Sub ListReportChoices( ByVal ireport )
	Dim sSql, oRs

	ireport = CLng(ireport + 0)
	sSql = "SELECT reportid, reportname, iscustomreport FROM egov_reports "
	sSql = sSql & "WHERE iscustomreport <> 1 order by iscustomreport desc, reportsequence"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	response.write vbcrlf & "<select name=""ireport"">"

	If Not oRs.EOF Then
		
		' DRAW STANDARD REPORTS
		response.write vbcrlf & "<option value=""0"" class=""optionheader""> -- Standard Reports --" & "</option>"
		
		Do While Not oRs.EOF
			If Not oRs("iscustomreport") Then
				
				If clng(ireport) = clng(oRs("reportid")) Then
					sSelected = " selected=""selected"" "
				Else
					sSelected = "" 
				End If

				response.write vbcrlf & "<option  " & sSelected & " class=""optionnormal"" value=""" & oRs("reportid") & """>" & oRs("reportname") & "</option>"
			
			End If
			oRs.MoveNext
		Loop

		' DRAW CUSTOM REPORTS
		oRs.MoveFirst
		response.write vbcrlf & "<option value=""0"" class=""optionheader""> -- Custom Reports --" & "</option>"
		
		Do While Not oRs.EOF
			If oRs("iscustomreport") Then

				If clng(ireport) = clng(oRs("reportid")) Then
					sSelected = " selected=""selected"" "
				Else
					sSelected = "" 
				End If
				
				response.write vbcrlf & "<option " & sSelected & " class=""optionnormal"" value=""" & oRs("reportid") & """>" & oRs("reportname") & "</option>"
			
			End If
			oRs.MoveNext
		Loop

	End If

	oRs.Close
	Set oRs = Nothing


	response.write vbcrlf & "</select>"

End Sub


'------------------------------------------------------------------------------------------------------------
' string FNGETSQLQUERY(IREPORTID) 
'------------------------------------------------------------------------------------------------------------
Function fnGetSQLQuery( ireportid ) 
	Dim sSql, oRs

	' GET SELECTED REPORT
	sSql = "SELECT reportsqlsyntax FROM egov_reports WHERE reportid = " & ireportid 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If NOT oRs.EOF Then
		sReturnValue = oRs("reportsqlsyntax")
	Else
		sReturnValue = "SELECT 'No report query found for this report id!' as [Error Message]"
	End If

	oRs.Close
	Set oRs = Nothing

	' RETURN VALUE
	fnGetSQLQuery = sReturnValue

End Function



%>
