<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: activity_total_export.asp
' AUTHOR: SteveLoar
' CREATED: 10/09/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This pulls together the activity totals report Export. Part of a Menlo Park Project.
'
' MODIFICATION HISTORY
' 1.0   10/09/2007		Steve Loar - INITIAL VERSION
' 1.1	09/30/2011	Steve Loar - Changine the OPEN column to Drop Ins for Menlo Park
' 1.2   10/15/2013	Steve Loar - Adding sort and filter for Class End Date
'
'--------------------------------------------------------------------------------------------------
'

	' USER SECURITY CHECK
	If Not UserHasPermission( Session("UserId"), "activity totals rpt" ) Then
		response.redirect sLevel & "../permissiondenied.asp"
	End If 

	Dim oSchema, iOldAccountId, dTotal, dTotalCredit, dTotalDebit, dGrandTotal
	Dim iLocationId, toDate, fromDate, sDateRange, iPaymentLocationId, iReportType, sAdminlocation
	Dim sFile, sRptTitle, sWhereClause, fromEndDate, toEndDate

	' SET UP PAGE OPTIONS
	sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
	sWhereClause = ""
	sFrom = ""

	sRptTitle = vbcrlf & "<tr><th></th><th>Activity Totals Report</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"

	server.scripttimeout = 9000
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment;filename=Activity_totals_" & sDate & ".xls"

	iClassSeasonId = clng(request("classseasonid"))

	If request("categoryid") = "" or CLng(request("categoryid")) = CLng(0) Then 
		iCategoryid = CLng(0)
	Else
		iCategoryid = CLng(request("categoryid"))
	End If 

	If request("locationid") = "" Then
		iLocationId = CLng(0)
	Else
		iLocationId = CLng(request("locationid"))
	End If 

	If request("instructorid") = "" Then
		iInstructorId = CLng(0)
	Else
		iInstructorId = CLng(request("instructorid"))
	End If 

	If request("supervisorid") = "" Then
		iSupervisorId = CLng(0)
	Else
		iSupervisorId = CLng(request("supervisorid"))
	End If 

	If request("searchname") <> "" Then 
		sSearchName = dbsafe(request("searchname"))
	End If 

	fromDate = Request("fromDate")
	toDate = Request("toDate")
	today = Date()

	fromEndDate = Request("fromEndDate")
	toEndDate = Request("toEndDate")

	iOrderBy = clng(request("orderby"))

	' BUILD SQL WHERE CLAUSE
	If iClassSeasonId > clng(0) Then
		sWhereClause = sWhereClause & " AND C.classseasonid = " & iClassSeasonId
		sRptTitle = sRptTitle & vbcrlf & "<tr><th>Season: " & GetSeasonName( iClassSeasonId )  & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	End If 

	If iCategoryid > CLng(0) Then
		sFrom = ", egov_class_category_to_class G "
		sWhereClause = sWhereClause & " AND C.classid = G.classid AND G.categoryid = " & iCategoryid
		sRptTitle = sRptTitle & vbcrlf & "<tr><th>Category: " & GetCategoryName( iCategoryid )  & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	End If 

	If iLocationId > CLng(0) Then
		sWhereClause = sWhereClause & " AND C.locationid = " & iLocationId
		sRptTitle = sRptTitle & vbcrlf & "<tr><th>Location: " & GetLocationName( iLocationId )  & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	End If 

	If iInstructorId > CLng(0) Then
		sWhereClause = sWhereClause & " AND T.instructorid = " & iInstructorId
		sRptTitle = sRptTitle & vbcrlf & "<tr><th>Instructor: " & GetInstructorName( iInstructorId )  & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	End If 

	If iSupervisorId > CLng(0) Then
		sWhereClause = sWhereClause & " AND C.supervisorid = " & iSupervisorId
		sRptTitle = sRptTitle & vbcrlf & "<tr><th>Supervisor: " & GetSupervisorName( iSupervisorId )  & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	End If 

	If sSearchName <> "" Then
		sWhereClause = sWhereClause & " AND C.classname LIKE '%" & dbsafe(sSearchName) & "%' "
		sRptTitle = sRptTitle & vbcrlf & "<tr><th>Class Name Like: " & sSearchName  & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	End If 

	If fromDate <> "" Then 
		sWhereClause = sWhereClause & " AND C.startdate >= '" & fromDate & " 00:00:00' "
		sRptTitle = sRptTitle & vbcrlf & "<tr><th>Start Date From: " & fromDate  & "</th><th></th><th><th></th><th></th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	End If

	If toDate <> "" Then
		toDateDisplay = toDate 
		toDate = DateAdd( "d", 1, toDate )
		sWhereClause = sWhereClause & " AND C.startdate < '" & toDate & " 00:00:00' "
		sRptTitle = sRptTitle & vbcrlf & "<tr><th>Start Date To:" & toDateDisplay  & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	End If

	If fromEndDate <> "" Then 
		sWhereClause = sWhereClause & " AND C.enddate >= '" & fromEndDate & " 00:00:00' "
		sRptTitle = sRptTitle & vbcrlf & "<tr><th>End Date From: " & fromEndDate  & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	End If  

	If toEndDate <> "" Then 
		toEndDateDisplay = ""
		toEndDateDisplay = Request("toEndDate")

		toEndDate = DateAdd( "d", 1, Request("toEndDate") )
		sWhereClause = sWhereClause & " AND C.enddate < '" & toEndDate & " 00:00:00' "
		sRptTitle = sRptTitle & vbcrlf & "<tr><th>End Date To: " & toEndDateDisplay & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	End If  

	If iOrderby = clng(1) Then 
		sOrderBy = "classname, activityno"
	ElseIf iOrderby = clng(2) Then 
		sOrderBy = "C.startdate, classname, activityno"
	ElseIf iOrderby = clng(3) Then 
		sOrderBy = "C.enddate, classname, activityno"		
	End If 
	
	Display_Results sWhereClause, sRptTitle, sFrom, sOrderBy


'------------------------------------------------------------------------------------------------------------
' void Display_Results( sWhereClause, sRptTitle, sFrom, sOrderBy )
'------------------------------------------------------------------------------------------------------------
Sub Display_Results( ByVal sWhereClause, ByVal sRptTitle, ByVal sFrom, ByVal sOrderBy )
	Dim sSql, oRs, iOpen, iMeetingCount, dHours, dRevenue, dPayment, dNetIncome, dTotalRevenue
	Dim dTotalPayment, dTotalNetIncome, dTotalHrs, dTotalMeetings, dTotalMin, dTotalMax, dTotalRes
	Dim dResCount, dNonResCount, dTotalNonRes, dTotalEnrollment, dTotalWait, iTotalOpen, iTotalAttendance
	Dim iTotalDropIn, iDropInCount

	dTotalRevenue = CDbl(0.0)
	dTotalPayment = CDbl(0.0)
	dTotalNetIncome = CDbl(0.0)
	dTotalHrs = CDbl(0.0)
	dTotalMeetings = CLng(0)
	dTotalMin = CLng(0)
	dTotalMax = CLng(0)
	dTotalRes = CLng(0)
	dTotalNonRes = CLng(0)
	dTotalEnrollment = CLng(0)
	dTotalWait = CLng(0)
	iTotalOpen = CLng(0)
	iTotalAttendance = CDbl(0.0)
	iTotalDropIn = CLng(0)

	sSql = "SELECT C.classname, C.classid, T.activityno, C.startdate, C.enddate, C.statusid, C.classseasonid, "
	sSql = sSql & " C.locationid, C.supervisorid, isnull(T.instructorid,0) as instructorid, T.min, T.max, T.timeid, "
	sSql = sSql & " T.enrollmentsize, T.waitlistsize, S.seasonname "
	sSql = sSql & " FROM egov_class C, egov_class_time T, egov_class_seasons S " & sFrom
	sSql = sSql & " WHERE C.classid = T.classid AND C.classseasonid = S.classseasonid "
	sSql = sSql & " AND C.orgid = " & Session("orgid") & sWhereClause
	sSql = sSql & " ORDER BY " & sOrderBy

	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF then
		response.Write vbcrlf & "<html>"
		response.write vbcrlf & "<style>  "
		response.write " .moneystyle "
		response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
		response.write vbcrlf & "</style>"
		response.write vbcrlf & "<body><table border=""1"">"
		response.write sRptTitle
		response.write vbcrlf & "<tr><th>ClassName</th><th>Activity No.</th><th>Season</th><th>Start Date</th><th>End Date</th>"
		response.write "<th># Hrs</th><th># Sessions</th><th>Min</th><th>Max</th><th>Res</th><th>Non Res</th><th>Total Enrld</th>"
		response.write "<th>Wait</th><th>Drop In</th><th>Attendance</th><th>Total Revenue</th><th>Instr Payment</th><th>Net Income</th></tr>"
		response.flush

		Do While Not oRs.EOF
			response.write vbcrlf & "<tr>"
			response.write "<td>" & oRs("classname") & "</td>"
			response.write "<td>" & oRs("activityno") & "</td>"
			response.write "<td>" & oRs("seasonname") & "</td>"
			response.write "<td>" & oRs("startdate") & "</td>"
			response.write "<td>" & oRs("enddate") & "</td>"
			dHours = CDbl(0.0)
			iMeetingCount = GetNewActivityMeetingCount( oRs("timeid"), dHours )
			dTotalHrs = dTotalHrs + CDbl(dHours)
			dTotalMeetings = dTotalMeetings + CLng(iMeetingCount)
			response.write "<td>" & dHours & "</td>"
			response.write "<td>" & iMeetingCount & "</td>"
			
			response.write "<td>" & oRs("min") & "</td>"
			If IsNumeric(oRs("min")) Then
				dTotalMin = dTotalMin + CLng(oRs("min"))
			End If 
			response.write "<td>" & oRs("max") & "</td>"
			If IsNumeric(oRs("max")) Then
				dTotalMax = dTotalMax + CLng(oRs("max"))
			End If 
			dResCount = GetResNonResClassCount( oRs("timeid"), "R" )
			dTotalRes = dTotalRes + CLng(dResCount)
			response.write "<td>" & dResCount & "</td>"
			dNonResCount = GetResNonResClassCount( oRs("timeid"), "N" )
			dTotalNonRes = dTotalNonRes + dNonResCount
			response.write "<td>" & dNonResCount & "</td>"
			response.write "<td>" & oRs("enrollmentsize") & "</td>"
			dTotalEnrollment = dTotalEnrollment + CLng(oRs("enrollmentsize"))
			response.write "<td>" & oRs("waitlistsize") & "</td>"
			dTotalWait = dTotalWait + CLng(oRs("waitlistsize"))
			If IsNull(oRs("max")) Then
				iOpen = "N/A"
			Else
				iOpen = CLng(oRs("max")) - CLng(oRs("enrollmentsize"))
				If iOpen < CLng(0) Then 
					iOpen = CLng(0)
				End If 
				iTotalOpen = iTotalOpen + iOpen
			End If 
			'response.write "<td>" & iOpen & "</td>"

			' Display Drop IN Count
			iDropInCount = GetDropInCount( oRs("timeid") )
			iTotalDropIn = iTotalDropIn + iDropInCount
			response.write "<td align=""center"">" & iDropInCount & "</td>"

			response.write "<td>" & CLng(dHours * CDbl(oRs("enrollmentsize"))) & "</td>"
			iTotalAttendance = iTotalAttendance + CLng( dHours * CDbl(oRs("enrollmentsize")))
			getRevenueAndPay oRs("timeid"), dRevenue, dPayment
			dNetIncome = dRevenue - dPayment
			dTotalRevenue = dTotalRevenue + dRevenue
			dTotalPayment = dTotalPayment + dPayment
			dTotalNetIncome = dTotalNetIncome + dNetIncome
			response.write "<td class=""moneystyle"">" & FormatNumber(dRevenue,2) & "</td>"
			response.write "<td class=""moneystyle"">" & FormatNumber(dPayment,2) & "</td>"
			response.write "<td class=""moneystyle"">" & FormatNumber(dNetIncome,2) & "</td>"
			response.write "</tr>"
			response.flush
			oRs.MoveNext
		Loop 
		' Total for all Classes
		response.write vbcrlf & "<tr><td></td><td></td><td></td><td></td><td>Totals:</td>"
		response.write "<td>" & FormatNumber(dTotalHrs,2) & "</td>"
		response.write "<td>" & FormatNumber(dTotalMeetings,0) & "</td>"
		response.write "<td>" & FormatNumber(dTotalMin,0) & "</td>"
		response.write "<td>" & FormatNumber(dTotalMax,0) & "</td>"
		response.write "<td>" & FormatNumber(dTotalRes,0) & "</td>"
		response.write "<td>" & FormatNumber(dTotalNonRes,0) & "</td>"
		response.write "<td>" & FormatNumber(dTotalEnrollment,0) & "</td>"
		response.write "<td>" & FormatNumber(dTotalWait,0) & "</td>"
		'response.write "<td>" & FormatNumber(iTotalOpen,0) & "</td>"
		response.write "<td>" & FormatNumber(iTotalDropIn,0) & "</td>"
		response.write "<td>" & FormatNumber(iTotalAttendance,2) & "</td>"
		response.write "<td class=""moneystyle"">" & FormatNumber(dTotalRevenue,2) & "</td>"
		response.write "<td class=""moneystyle"">" & FormatNumber(dTotalPayment,2) & "</td>"
		response.write "<td class=""moneystyle"">" & FormatNumber(dTotalNetIncome,2) & "</td>"
		response.write "</tr>"
		response.write vbcrlf & "</table></body></html>"
		response.flush
	End If

	oRs.Close
	Set oRs = Nothing 
	
End Sub 


'--------------------------------------------------------------------------------------------------
' string GetLocationName( iLocationid )
'--------------------------------------------------------------------------------------------------
Function GetLocationName( ByVal iLocationid )
	Dim sSql, oRs

	sSql = "SELECT name FROM egov_class_location WHERE locationid = " & iLocationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetLocationName = oRs("name")
	Else
		GetLocationName = ""
	End If 

	oRs.Close 
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetSeasonName( iClassSeasonId )
'--------------------------------------------------------------------------------------------------
Function GetSeasonName( ByVal iClassSeasonId )
	Dim sSql, oRs

	sSql = "SELECT seasonname FROM egov_class_seasons WHERE classseasonid = " & iClassSeasonId
 
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		GetSeasonName = oRs("seasonname")
	Else
		GetSeasonName = ""
	End If

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetInstructorName( iInstrudtorId )
'--------------------------------------------------------------------------------------------------
Function GetInstructorName( ByVal iInstructorId )
	Dim sSql, oRs

	sSql = "SELECT lastname, firstname FROM egov_class_instructor WHERE instructorid = " & iInstructorId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetInstructorName = oRs("firstname") & " " & oRs("lastname")
	Else
		GetInstructorName = ""
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetSupervisorName( iSupervisorId )
'--------------------------------------------------------------------------------------------------
Function GetSupervisorName( ByVal iSupervisorId )
	Dim sSql, oRs

	sSql = "SELECT lastname, firstname FROM users WHERE userid = " & iSupervisorId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetSupervisorName = oRs("firstname") & " " & oRs("lastname")
	Else
		GetSupervisorName = ""
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetCategoryName( iCategoryid )
'--------------------------------------------------------------------------------------------------
Function GetCategoryName( ByVal iCategoryid )
	Dim sSql, oRs

	sSql = "SELECT categorytitle FROM egov_class_categories WHERE categoryid = " & iCategoryid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetCategoryName = oRs("categorytitle") 
	Else
		GetCategoryName = ""
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetResNonResClassCount( iTimeid, sResType )
'--------------------------------------------------------------------------------------------------
Function GetResNonResClassCount( ByVal iTimeid, ByVal sResType )
	Dim sSql, sResMatch, oRs

	If sResType = "R" Then
		sResMatch = " = 'R'"
	Else
		sResMatch = " != 'R'"
	End If 
	
	sSql = "SELECT COUNT(attendeeuserid) AS hits FROM egov_class_list L, egov_users U WHERE L.attendeeuserid = U.userid "
	sSql = sSql & " AND L.status = 'ACTIVE' AND L.classtimeid = " & iTimeid & " AND U.residenttype " & sResMatch

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	GetResNonResClassCount = CLng(oRs("hits"))

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
'  void GetActivityStartAndEndDates iClassid, dStartDate, dEndDate 
'--------------------------------------------------------------------------------------------------
Sub GetActivityStartAndEndDates( ByVal iClassid, ByRef dStartDate, ByRef dEndDate )
	Dim sSql, oRs

	sSql = "SELECT startdate, enddate FROM egov_class WHERE classid = " & iClassid 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		dStartDate = oRs("startdate")
		dEndDate = oRs("enddate")
	Else
		dStartDate = "0/" 
		dEndDate = "0/" 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
'  integer GetNewActivityMeetingCount( iTimeid, dHours )
'--------------------------------------------------------------------------------------------------
Function GetNewActivityMeetingCount( ByVal iTimeid, ByRef dHours )
	Dim sSql, oRs, iMeetingCount

	sSql = "SELECT meetingcount, totalhours FROM egov_class_time WHERE timeid = " & iTimeid 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		dHours = CDbl(oRs("totalhours"))
		iMeetingCount = CLng(oRs("meetingcount"))
	Else
		dHours = CDbl(0.0)
		iMeetingCount = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetNewActivityMeetingCount = iMeetingCount

End Function 


'--------------------------------------------------------------------------------------------------
'  integer GetActivityMeetingCount( iClassid, iTimeid, dHours )
'--------------------------------------------------------------------------------------------------
Function GetActivityMeetingCount( ByVal iClassid, ByVal iTimeid, ByRef dHours )
	Dim dStartDate, dEndDate, dCurrDate, iMonth, iDay, iYear, iMeetingCount

	iMeetingCount = 0
	dHours = CDbl(0.0)

	GetActivityStartAndEndDates iClassid, dStartDate, dEndDate

	If IsNull(dStartDate) Or IsNull(dEndDate) Then
		' one or more dates is missing so cannot create a sheet
		iMeetingCount = 0
	ElseIf IsDate(dStartDate) And IsDate(dEndDate) Then 
		If Day(dStartDate) = Day(dEndDate) And Month(dStartDate) = Month(dEndDate) And Year(dStartDate) = Year(dEndDate) Then 
			' this is a one day event
			iMeetingCount = 1
		Else
			' this class happens over several days
			dCurrDate = dStartDate
			Do While dCurrDate <= dEndDate
				If ClassMeetsThen( iTimeid, WeekDayName(Weekday(dCurrDate)) ) Then
					iMeetingCount = iMeetingCount + 1
					dHours = dHours + GetActivityHoursForDay( iTimeid, WeekDayName(Weekday(dCurrDate)) )
				End If
				dCurrDate = DateAdd("d", 1, dCurrDate )
			Loop 
		End If 
	Else
		' one or more dates is not a date
		iMeetingCount = 0
	End If 

	GetActivityMeetingCount = iMeetingCount

End Function 


'--------------------------------------------------------------------------------------------------
'  boolean ClassMeetsThen( iTimeid, sDayOfWeek )
'--------------------------------------------------------------------------------------------------
Function ClassMeetsThen( ByVal iTimeid, ByVal sDayOfWeek )
	Dim sSql, oDay

	sSql = "SELECT COUNT(timedayid) AS hits FROM egov_class_time_days WHERE timeid = " & iTimeid & " AND " & sDayOfWeek & " = 1" 

	Set oDay = Server.CreateObject("ADODB.Recordset")
	oDay.Open sSql, Application("DSN"), 3, 1

	If Not oDay.EOF Then
		If clng(oDay("hits")) > clng(0) Then
			ClassMeetsThen = True 
		Else
			ClassMeetsThen = False 
		End If 
	Else
		ClassMeetsThen = False 
	End If 

	oDay.Close
	Set oDay = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
'  double GetActivityHoursForDay( iTimeid, sDayOfWeek )
'--------------------------------------------------------------------------------------------------
Function GetActivityHoursForDay( ByVal iTimeid, ByVal sDayOfWeek )
	Dim sSql, oRs, dHours, sAmOrPm, iColonPos, sHour, sMin, sStartDate, sEndDate

	dHours = CDbl(0.0)
	sSql = "SELECT starttime, endtime FROM egov_class_time_days WHERE timeid = " & iTimeid & " AND " & sDayOfWeek & " = 1" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		If oRs("starttime") <> "" And oRs("endtime") <> "" Then 
			' Get the start time in a format that can be used to compute the hours
			sAmOrPm = Right( oRs("starttime"), 2)
			iColonPos = InStr(oRs("starttime"), ":")
			sHour = Left( oRs("starttime"), ( iColonPos-1) )
			sMin = Mid( oRs("starttime"),(iColonPos + 1), ((Len(oRs("starttime"))-2)-iColonPos))
			If UCase(sAmOrPm) = "PM" And clng(sHour) < clng(12) Then 
				sHour = clng(sHour) + clng(12)
			End If
			sStartDate = CDate(Month(now) &"/" & Day(now) & "/" & Year(now) & " " & sHour & ":" & sMin )
			' Get the end time in a format that can be used to compute the hours
			sAmOrPm = Right( oRs("endtime"), 2)
			iColonPos = InStr(oRs("endtime"), ":")
			sHour = Left( oRs("endtime"), ( iColonPos-1) )
			sMin = Mid( oRs("endtime"),(iColonPos + 1), ((Len(oRs("endtime"))-2)-iColonPos))
			If UCase(sAmOrPm) = "PM" And clng(sHour) < clng(12) Then 
				sHour = clng(sHour) + clng(12)
			End If
			sEndDate = CDate(Month(now) &"/" & Day(now) & "/" & Year(now) & " " & sHour & ":" & sMin )
			dHours = dHours + abs(CDbl(DateDiff("s", sStartDate, sEndDate) / (60 * 60)))
		End If
		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

	GetActivityHoursForDay = CDbl(FormatNumber(dHours,2,,,0))

End Function 


'--------------------------------------------------------------------------------------------------
'  void getRevenueAndPay iClassTimeid, dRevenue, dPayment 
'--------------------------------------------------------------------------------------------------
Sub getRevenueAndPay( ByVal iClassTimeid, ByRef dRevenue, ByRef dPayment )
	Dim sSql, oRs
	
	dRevenue = CDbl(0.0)
	dPayment = CDbl(0.0)
	sSql = "SELECT classtimeid, SUM(amount) AS revenue, SUM(instructorpay) AS instructorpay FROM egov_activity_revenue_details "
	sSql = sSql & " WHERE classtimeid = " & iClassTimeid & " GROUP BY classtimeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		dRevenue = CDbl(oRs("revenue"))
		dPayment = CDbl(oRs("instructorpay"))
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetDropInCount( iTimeid )
'--------------------------------------------------------------------------------------------------
Function GetDropInCount( ByVal iTimeid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(attendeeuserid) AS hits FROM egov_class_list L, egov_users U WHERE L.attendeeuserid = U.userid "
	sSql = sSql & " AND L.status = 'DROPIN' AND L.classtimeid = " & iTimeid 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetDropInCount = CLng(oRs("hits"))
	Else
		GetDropInCount = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


%>

<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../classes/class_global_functions.asp" //-->
