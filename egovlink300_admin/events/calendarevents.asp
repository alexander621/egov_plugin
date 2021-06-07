<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="events_global_functions.asp" //-->
<%
'Check to see if the feature is offline
 if isFeatureOffline("internal_calendars") = "Y" OR isFeatureOffline("custom_calendars") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../" ' Override of value from common.asp

'Check to see if this is a Custom Calendar
' if trim(request("cal")) <> "" then
'    lcl_calendarfeature = trim(request("cal"))
' else
'    lcl_calendarfeature = getFirstCalendarInList(session("orgid"))
' end if

' lcl_calendarfeature_url  = "&cal=" & lcl_calendarfeature
' lcl_calendarfeature_name = " [" & getFeatureName(lcl_calendarfeature) & "]"

'Check to see if this is a Custom Calendar
 if trim(request("cal")) <> "" then
    if not isnumeric(trim(request("cal"))) then
       response.redirect sLevel & "permissiondenied.asp"
    else
       lcl_calendarfeatureid = CLng(trim(request("cal")))
    end if
 else
    lcl_calendarfeatureid = getFirstCalendarInList(session("orgid"))
 end if

 if lcl_calendarfeatureid <> "" then
    lcl_calendarfeature = getFeatureByID(session("orgid"), lcl_calendarfeatureid)

    if orghasfeature("internal_calendars") and userhaspermission(session("userid"), "internal_calendars") then
       lcl_calendarfeature_url  = "&cal=" & lcl_calendarfeatureid
       lcl_calendarfeature_name = " [" & getFeatureName(lcl_calendarfeature) & "]"
    else
       response.redirect sLevel & "permissiondenied.asp"
    end if
 else
    lcl_calendarfeatureid    = ""
    lcl_calendarfeature      = ""
    lcl_calendarfeature_url  = ""
    lcl_calendarfeature_name = ""
 end if
%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<title><%=langBSEventsCalendar%><%=lcl_calendar_name%></title>
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="eventstyles.css" />

	<script src="../scripts/selectAll.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>

	<style type="text/css">
	<!--
		body {scrollbar-base-color:#6699cc; scrollbar-highlight-color:#ffffff; scrollbar-arrow-color:#99ccff; font-family:Verdana,Tahoma,Arial; font-size:11px;}
		.cal {border-left:1px solid #93bee1; border-top:1px solid #93bee1; border-right:1px solid #93bee1;}
		.cal th {border-right:1px solid #93bee1; border-bottom:1px solid #93bee1; font-family:Tahoma,Arial; font-size:11px; color:#336699; text-align:left;}
		.cal td {border-bottom:1px solid #93bee1; font-family:Tahoma,Arial; font-size:11px;}
		select {font-family:Arial,Tahoma,Verdana; font-size:13px;}
	//-->
	</style>
</head>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 
<body bgcolor="#ffffff">
<div id="content">
	 <div id="centercontent">
  <%
Dim dDate

If IsDate(Request.QueryString("date")) Then
	dDate = CDate(Request.QueryString("date"))
Else
	If IsDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year")) Then
		dDate = CDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year"))
	Else
		dDate = Date()
	End If
End If

Dim oCmd, oRst, sEvents

lcl_bgcolor = "#eeeeee"

'Retrieve the details about the event(s) for the date selected.
sSql = "SELECT e.EventID, e.EventDate, e.EventDuration, t.TZAbbreviation, e.Subject, e.Message, e.CategoryID, c.Color "
sSql = sSql & " FROM Events e LEFT OUTER JOIN EventCategories c ON e.CategoryID = c.CategoryID, TimeZones t  "
sSql = sSql & " WHERE t.TimeZoneID = e.EventTimeZoneID "
sSql = sSql & " AND e.OrgID = " & session("orgid")
sSql = sSql & " AND DateDiff(dd, e.EventDate, '" & dDate & "') = 0 "
sSql = sSql & " AND DateDiff(mm, e.EventDate, '" & dDate & "') = 0 "
sSql = sSql & " AND DateDiff(yy, e.EventDate, '" & dDate & "') = 0"

if lcl_calendarfeature <> "" then
	sSql = sSql & " AND e.calendarfeature = '" & lcl_calendarfeature & "' "
else
	sSql = sSql & " AND (e.calendarfeature IS NULL OR e.calendarfeature <> '') "
end if

set oRst = Server.CreateObject("ADODB.Recordset")
oRst.Open sSql, Application("DSN"), 3, 1

if Not oRst.eof then
	do while Not oRst.eof
		sEvents = sEvents & "<tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
		sEvents = sEvents & "    <td width=""100"" valign=""top"" nowrap=""nowrap"">" & oRst("EventDate") & " " & oRst("TZAbbreviation") & "</td>" & vbcrlf
		sEvents = sEvents & "    <td><strong>" & oRst("Subject") & "</strong><br />" & oRst("Message") & "</td>" & vbcrlf
		sEvents = sEvents & "</tr>" & vbcrlf

		lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

		oRst.movenext
	loop
end if
set oRst = nothing
%>

		<h3>Calendar Event<%=lcl_calendarfeature_name%></h3>

		<p>
			<img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="calendar.asp<%=replace(lcl_calendarfeature_url,"&","?")%>"><%=langBackToCalendar%></a><br />
		</p>

		<div class="calendartitle"><%=langEvents%>: <%= FormatDateTime(dDate, vbLongDate) %></div>
			
		<table border="0" cellpadding="4" cellspacing="0" class="tablelist" width="100%" id="eventsdetails">
			<tr align="left">
				<th><%=langDateTime%></th>
				<th><%=langEvent%></th>
			</tr>
			<%= sEvents %>
		</table>

		</div>
	</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>