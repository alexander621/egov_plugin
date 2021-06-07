<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: attendance_sheet.asp
' AUTHOR: Steve Loar
' CREATED: 07/25/2007
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   07/25/2007	Steve Loar - INITIAL VERSION
' 1.1	01/31/2008	Steve Loar - Catching condition of no classid or timeid. Sending to roster list.
' 2.0	05/09/2011	Steve Loar - Cleaned up SELECT queries to be SQL Server 2008 Compatible
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iClassId, iTimeId

'Check to see if the feature is offline
if isFeatureOffline("activities") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

Set oClassDivOrg = New classOrganization

classCount = 0

'if request("classid") <> "" AND request("timeid") <> "" then
'   iClassId = request("classid")
'   iTimeId = request("timeid")
'else

 if request("classid") = "" AND request("timeid") = "" then
    response.redirect "roster_list.asp"
 end if

 if request("classid") <> "" then
    iClassId = request("classid")
 end if

 if request("timeid") <> "" then
    iTimeId = request("timeid")
 end if
%>
<html>
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8"/>

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />

</head>
<body>


<div id="idControls" class="noprint">
	<input type="button" class="button" onclick="javascript:window.print();" value="Print" />&nbsp;
	<input type="button" class="button" value="<< Back" onclick="javascript:history.back()" />
</div>

<div id="content">

<!--BEGIN: CLASS ROSTER-->
<%
'LOOP THRU EACH CLASS CHECKED AND DISPLAY CLASS INFORMATION AND ROSTER
If request("classid") <> "" And request("timeid") <> "" Then 
	classCount = classCount + 1
	If classCount > 1 Then 
		response.write "<div class=""class_start"">"
	End If 

	'CLASS INFORMATION
	response.write "<p>"
	DisplayItem iclassid, itimeid
	response.write "</p>"

	'CLASS LIST
	response.write "<p>"
	DisplayClassEventsRoster iclassid, itimeid
	response.write "</p>"
Else 
	If request("classid") <> "" then
		For Each item In request("classid")
			sSql = "SELECT timeid "
			sSql = sSql & " FROM egov_class_time "
			sSql = sSql & " WHERE classid = " & item
			sSql = sSql & " ORDER BY activityno, timeid "

			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open sSql, Application("DSN"), 0, 1

       		Do While Not oRs.EOF
				classCount = classCount + 1
				If classCount > 1 Then 
					response.write "<div class=""class_start"">"
				End If 

				'CLASS INFORMATION
				response.write "<p>"
				DisplayItem item, oRs("timeid")
				response.write "</p>"

				'CLASS LIST
				response.write "<p>"
				DisplayClassEventsRoster item, oRs("timeid")
				response.write "</p>"

				response.write "<div class=""footerbox"">"
				response.write "<table width=""100%"" cellspacing=""0"" cellpadding=""0"" border=""0"">"
				response.write "<tr>"
				response.write "<td height=""5"" bgcolor=""#93bee1"" style=""border-bottom: solid 1px #000000;"">&nbsp; </td>"
				response.write "</tr>"
				response.write "<tr>"
				response.write "<td valign=""top"" align=""center"">"
				response.write "<font style=""font-size:10px;font-weight:bold;"">Copyright &copy;2004 - "
%>
				<script type="text/javascript">
				<!--
					var theDate=new Date();
					document.write(theDate.getFullYear());
				//-->
				</script>
<%
				response.write ". All Rights Reserved. " & oClassDivOrg.GetOrgDisplayName( "admin footer brand link" ) & "</font><br />&nbsp;</font>"
				response.write "</td>"
				response.write "</tr>"
				response.write "</table>"
				response.write "</div>"

				If classCount > 1 Then 
					response.write "</div>"
				End If 

				oRs.MoveNext
			Loop
			
			oRs.Close
			Set oRs = Nothing 

		Next 
	End If 
End If

%>
<!--END: CLASS ROSTER-->
	<!-- </div> -->
</div>

<!--#include file="../admin_footer.asp"-->  
</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' DisplayClassEventsRoster( iClassId, iTimeId )
'--------------------------------------------------------------------------------------------------
Sub DisplayClassEventsRoster( ByVal iClassId, ByVal iTimeId )
	Dim sSql, oRs, iWaitlistCount, iActiveCount, iAge, iRow, sHeader, sBody

	iWaitlistCount = 0
	iActiveCount   = 0
	iRow           = clng(0)
	sHeader        = ""
	sBody          = ""

	' GET ALL CLASS\EVENTS FOR ORG
	sSql = "SELECT * "
	sSql = sSql & " FROM egov_class_roster "
	sSql = sSql & " WHERE classid = " & iClassId
	sSql = sSql & " AND classtimeid = " & iTimeId 
	sSql = sSql & " AND status = 'ACTIVE' "
	sSql = sSql & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		
		GetClassDates iclassid, itimeid, sHeader, sBody

		' IF REGISTRATION REQUIRED SHOW REGISTERED USERS
		If oRs("optionid") = "1" Then
			response.write "<div class=""shadow"">"
			response.write "<table cellpadding=""2"" cellspacing=""0"" border=""1"" id=""attendancesheet"">"
			
			'HEADER ROW
			response.write "<tr>"
			response.write "<th>&nbsp;</th>"
			response.write "<th>Enrollee Name</th>"
			response.write "<th>Waivers<br />On File</th>"
			response.write      sHeader
			response.write "</tr>"

			' LOOP THRU AND DISPLAY CLASS ROSTER
			Do While Not oRs.EOF
				iRow = iRow + 1
				If iRow Mod 2 = 0 Then
					sClass = " class=""altrow"""
				Else
					sClass = ""
				End If 
				response.write vbcrlf & "<tr" & sClass & ">"
				response.write "<td align=""center"">" & iRow & "</td>"
				response.write "<td nowrap=""nowrap"">" & oRs("firstname") & " " & oRs("lastname") & "</td>"

				response.write "<td align=""center"">"
				If oRs("waiveronfile") Then
					response.write "yes"
				Else
					response.write "&nbsp;"
				End If 
				response.write "</td>"

				response.write sBody & "</tr>"
				oRs.MoveNext
			Loop 

		Else
			'NON-REGISTERED USERS SHOW PAID USERS
			
			' DRAW TABLE WITH CLASSES LISTED
			response.write vbcrlf & "<b>List Effective: " & Now() & "<br />"
			response.write vbcrlf & "<b>Total Participants: </b>" & fnIsNull(oRs("enrollmentsize"),0) & " - (Min: " & fnIsNull(oRs("min"),"n/a") & ", Max: " & fnIsNull(oRs("max"),"n/a") & ") <br />"
			response.write vbcrlf & "<b>Total Payees: </b>" & oRs.RecordCount & "<br /><br />"

			response.write vbcrlf & "<div class=""shadow"">"
			response.write vbcrlf & "<table cellpadding=""2"" cellspacing=""0"" border=""1"" id=""attendancesheet"">"
			
			' HEADER ROW
			response.write vbcrlf & "<tr><th>&nbsp;</th><th>Attendee Name</th>" & sHeader & "</tr>"

			' LOOP THRU AND DISPLAY CLASS ROSTER
			Do While Not oRs.EOF
				iRow = iRow + 1
				response.write vbcrlf & "<tr>"
				response.write "<td>" & iRow & "</td>"
				response.write "<td>" & oRs("userlname") & ", " & oRs("userfname") & "</td>"
				response.write sBody & "</tr>"
				oRs.MoveNext
			Loop 

		End If

		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>"

	Else
		' NO CLASS\EVENTS WERE FOUND
		response.write "<font style=""font-size:10px;"" color=""red""><strong>No purchases/registrations have been made for this activity.</strong></font>"
	
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
'  GetClassDates( iClassid, iTimeid, sHeader, sBody )
'--------------------------------------------------------------------------------------------------
Sub GetClassDates( ByVal iClassid, ByVal iTimeid, ByRef sHeader, ByRef sBody )
	Dim dStartDate, dEndDate, dCurrDate, iMonth, iDay, iYear

	GetStartAndEndDates iClassid, dStartDate, dEndDate

	If IsNull(dStartDate) Or IsNull(dEndDate) Then
		' one or more dates is missing so cannot create a sheet
		sHeader = "<th>&nbsp;</th>"
		sBody = "<td>&nbsp;</td>"
	ElseIf IsDate(dStartDate) And IsDate(dEndDate) Then 
		If Day(dStartDate) = Day(dEndDate) And Month(dStartDate) = Month(dEndDate) And Year(dStartDate) = Year(dEndDate) Then 
			' this is a one day event
			sHeader = "<th>" & Month(dStartDate) & "/" & Day(dStartDate) & "</th>"
			sBody = "<td>&nbsp;</td>"
		Else
			' this class happens over several days
			dCurrDate = dStartDate
			Do While dCurrDate <= dEndDate
				If ClassMeetsThen( iTimeid, WeekDayName(Weekday(dCurrDate)) ) Then
					sHeader = sHeader & "<th>" & Month(dCurrDate) & "/" & Day(dCurrDate) & "</th>"
					sBody = sBody & "<td>&nbsp;</td>"
				End If
				dCurrDate = DateAdd("d", 1, dCurrDate )
			Loop 
		End If 
	Else
		' one or more dates is not a date
		sHeader = "<th>&nbsp;</th>"
		sBody = "<td>&nbsp;</td>"
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
'  ClassMeetsThen( iTimeid, sDayOfWeek )
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
'  GetStartAndEndDates( iClassid, iTimeid, dStartDate, dEndDate )
'--------------------------------------------------------------------------------------------------
Sub GetStartAndEndDates( ByVal iClassid, ByRef dStartDate, ByRef dEndDate )
	Dim sSql, oRs

	sSql = "SELECT startdate, enddate FROM egov_class WHERE classid = " & iclassid 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

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
'  DisplayItem iClassid, iTimeid
'--------------------------------------------------------------------------------------------------
Sub DisplayItem( ByVal iClassid, ByVal iTimeid )
	Dim sSql, oRs

	arrDetails = Array("startdate","enddate","alternatedate","minage","maxage")
	arrDetailLabels = Array("Start Date","End Date","Make Up Date","Minimum Age","Maximum Age")

	'GET SELECTED FACILITY INFORMATION
	sSql = "SELECT classname, classseasonid, alternatedate, minage, maxage, locationid, "
	sSql = sSql & "ISNULL(egov_class.startdate,0) AS startdate, IsNull(egov_class.enddate,0) AS enddate, "
	sSql = sSql & "IsNull(egov_class.imgurl,'EMPTY') AS imgurl, (firstname + ' ' + lastname) AS Instructor "
	sSql = sSql & "FROM egov_class "
	sSql = sSql & "LEFT JOIN egov_class_time ON egov_class.classid = egov_class_time.classid "
	sSql = sSql & "LEFT JOIN egov_class_instructor ON egov_class_time.instructorid = egov_class_instructor.instructorid "
	sSql = sSql & "WHERE egov_class.classid = " & iClassid
	sSql = sSql & " AND egov_class_time.timeid = " & iTimeid
	sSql = sSql & " ORDER BY noenddate desc, startdate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	'DISPLAY ITEM INFORMATION
	If Not oRs.EOF Then 

		'WRITE TITLE
		response.write "<h2>" & Session("sOrgName") & " Attendance Sheet</h2>"
		response.write "<h3>" & oRs("classname")  & " ( " & GetActivityNo( iTimeid ) & " ) </h3>"

		'DISPLAY ITEM DETAILS
		response.write "<div>"
		response.write "<fieldset style=""border:0;"">"
		response.write "<table style=""align=left; margin: 5px;width:450px;"">"

		'Show the season
		response.write "<tr>"
		response.write "<td class=""classdetaillabel"">Season: </td>"
		response.write "<td class=""classdetailvalue"">" & GetSeasonName( oRs("classseasonid") ) & "</td>"
		response.write "</tr>"

		'Show the location
		response.write "<tr>"
		response.write "<td class=""classdetaillabel"">Location: </td>"
		response.write "<td class=""classdetailvalue"">" & GetLocationName( oRs("locationid") ) & "</td>"
		response.write "</tr>"

		'DISPLAY DETAILS VALUE PAIR
		For d = 0 To UBound(arrDetails)
			If Trim(oRs(arrDetails(d))) <> "" And Not IsNull(oRs(arrDetails(d))) Then 

				'IF DATE THEN FORMAT
				If IsDate(oRs(arrDetails(d))) Then 
					'FORMAT DATE
					sValue = FormatDateTime(oRs(arrDetails(d)),1)
				Else 
					'DISPLAY STORED VALUE UNFORMATTED
					sValue = oRs(arrDetails(d))
				End If 

				response.write "<tr>"
				response.write "<td class=""classdetaillabel"">" & arrDetailLabels(d) & ": </td>"
				response.write "<td class=""classdetailvalue"">" & sValue & "</td>"
				response.write "</tr>"
			End If 
		Next 

		'DISPLAY INSTRUCTOR
		If Trim(oRs("Instructor")) <> "" And Not IsNull(oRs("Instructor")) Then 
			response.write "<tr>"
			response.write "<td class=""classdetaillabel"">Instructor: </td>"
			response.write "<td>" & oRs("Instructor") & "</td>"
			response.write "</tr>"
		End If 

		'Display Waiver Links
		response.write "<tr>"
		response.write "<td class=""classdetaillabel"">Waivers: </td>"
		response.write "<td>"
		ShowClassWaiverNames iClassid 
		response.write "</td>"
		response.write "</tr>"
		response.write "</table>"

		DisplayClassActivities iClassid, iTimeid, False   ' In class_global_functions.asp

		response.write "</fieldset>"
		response.write "</div>"

	End If 


	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
'  DISPLAYCLASSEVENTS( iOrgId )
'--------------------------------------------------------------------------------------------------
Sub DisplayClassEvents( ByVal iOrgId )
	Dim sSql, oRs

	' GET ALL CLASS\EVENTS FOR ORG
	sSql = "SELECT * FROM egov_roster_list2 where orgid = '" & iOrgId & "' and parentclassid is null order by classname"
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		DisplayTimes oRs("classid"), oRs("classname")

		If oRs("isparent") Then
			DisplayChildClassEvents iorgid, oRs("classid")
		End If

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' DisplayChildClassEvents iOrgId, iparentid
'--------------------------------------------------------------------------------------------------
Sub DisplayChildClassEvents( ByVal iOrgId, ByVal iparentid )
	Dim sSql, oRs

	' GET ALL CLASS\EVENTS FOR ORG
	sSql = "SELECT classid, classname FROM egov_roster_list2  WHERE orgid = " & iOrgId & " AND parentclassid = " & iparentid & " ORDER BY classname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		DisplayTimes oRs("classid"), oRs("classname" )
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' string fnGetPercentFull( sMax, sCurrent )
'--------------------------------------------------------------------------------------------------
Function fnGetPercentFull( ByVal sMax, ByVal sCurrent )

	If IsNumeric(sMax) AND IsNumeric(sCurrent) Then
		 fnGetPercentFull = FormatNumber(clng(sCurrent) / clng(sMAX) * 100,0)  
	Else
		 fnGetPercentFull = "n/a"
	End If

End Function


'--------------------------------------------------------------------------------------------------
' void DisplayTimes iClassId, sClassName
'--------------------------------------------------------------------------------------------------
Sub DisplayTimes( ByVal iClassId, ByVal sClassName )
	Dim sSql, oRs

	sSql = "SELECT  egov_class_time.starttime, egov_class_time.endtime, egov_class_time.min, egov_class_time.max, timeid "
	sSql = sSql & "FROM egov_class_time WHERE egov_class_time.classid = " & iClassId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	Do While Not oRs.EOF 
		response.write "<option value=""" &iclassid & "," & oRs("timeid")& """>" & sClassName & " --- (" & oRs("starttime") & " - " & oRs("endtime") & " " & fnGetTimeDaysofWeek(iclassid) & ")</option>"
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION FNGETTIMEDAYSOFWEEK(ICLASSID)
'--------------------------------------------------------------------------------------------------
Function fnGetTimeDaysofWeek( ByVal iClassId )
	Dim sSql, oRs
	
	sReturnValue = ""

	' GET THE DAY OF THE WEEK VALUES FOR THE SPECIFIED
	sSql = "SELECT dayofweek FROM egov_class_dayofweek WHERE classid = " & iClassId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	' LOOP THRU AVAILABLE DAYS OF THE WEEK
	Do While Not oRs.EOF 
		sReturnValue = sReturnValue &  WeekDayName(oRs("dayofweek"),true) & " "
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing

	' RETURN DAYS OF THE WEEK
	fnGetTimeDaysofWeek = Trim(sReturnValue)

End Function



%>
