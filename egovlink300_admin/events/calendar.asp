<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="events_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: calendar.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the calendar for internal calendars
'
' MODIFICATION HISTORY
' 1.0 08/20/08  David Boyer - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("internal_calendars") = "Y" OR isFeatureOffline("custom_calendars") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../" ' Override of value from common.asp

'Check to see if any Custom Calendars exist
 lcl_hasCustomCalendars = checkForCustomCalendars(session("orgid"))

'Check to see if this is a Custom Calendar
' if trim(request("cal")) <> "" then
'    lcl_calendarfeature = trim(request("cal"))
' else
'    if trim(request("calendarfeature")) <> "" then
'       lcl_calendarfeature = trim(request("calendarfeature"))
'    else
'       lcl_calendarfeature = getFirstCalendarInList(session("orgid"))
'    end if
' end if

' lcl_calendarfeature_url  = "&cal=" & lcl_calendarfeature

' if lcl_calendarfeature <> "" then
'    lcl_calendarfeature_name = " [" & getFeatureName(lcl_calendarfeature) & "]"
' else
'    lcl_calendarfeature_name = ""
' end if

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

   'Allow the user to access the Internal Calendars if any/all of the following:
     '1. The user has the "edit events" permission assigned
     '2. The user has a specific Custom Calendar feature assigned: [session("calendarfeature") <> ""]
   'MODIFIED (change requested by Peter on 05/07/2010)
     '- Now only have to check for the "View Internal Calendar" feature (internal_calendars) to see ALL internal calendars
     '  whether user has permission to edit the internal calendar(s) or not.
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

'Allow the user to access the Internal Calendars if any/all of the following:
 '1. The user has the "edit events" permission assigned
 '2. The user has a specific Custom Calendar feature assigned: [session("calendarfeature") <> ""]
'MODIFIED (change requested by Peter on 05/07/2010)
 '- Now only have to check for the "View Internal Calendar" feature (internal_calendars) to see ALL internal calendars
 '  whether user has permission to edit the internal calendar(s) or not.

' if orghasfeature("internal_calendars") and userhaspermission(session("userid"), "internal_calendars") then
'    if lcl_calendarfeatureid <> "" then
'       session("calendarfeature") = trim(request("cal"))
'       lcl_calendarfeature_url    = "&cal=" & session("calendarfeature")
'       lcl_calendarfeature_name   = " [" & GetFeatureName(session("calendarfeature")) & "]"
'    end if
' else
'    session("calendarfeature") = ""
'    response.redirect sLevel & "permissiondenied.asp"
' end if

 'if trim(request("cal")) <> "" then
 '   if OrgHasFeature(trim(request("cal"))) AND UserHasPermission(session("userid"), trim(request("cal"))) then
 '      session("calendarfeature") = trim(request("cal"))
 '      'lcl_calendarfeature = trim(request("cal"))
 '      lcl_calendarfeature_url  = "&cal=" & session("calendarfeature")
 '      lcl_calendarfeature_name = " [" & GetFeatureName(session("calendarfeature")) & "]"
 '   else
 '     	response.redirect sLevel & "permissiondenied.asp"
 '   end if
 'else
 '   if NOT UserHasPermission( Session("UserId"), "edit events" ) then
 '	     response.redirect sLevel & "permissiondenied.asp"
 '   end if

 '   session("calendarfeature") = ""
 'end if

 Function GetDaysInMonth(iMonth, iYear)
	  Dim dTemp
	  dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
	  GetDaysInMonth = Day(dTemp)
 End Function

 Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth)
	  Dim dTemp
	  dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth)
	  GetWeekdayMonthStartsOn = WeekDay(dTemp)
 End Function

 Function SubtractOneMonth(dDate)
	  SubtractOneMonth = DateAdd("m", -1, dDate)
 End Function

 Function AddOneMonth(dDate)
	  AddOneMonth = DateAdd("m", 1, dDate)
 End Function

 Dim dDate     ' Date we're displaying calendar for
 Dim iDIM      ' Days In Month
 Dim iDOW      ' Day Of Week that month starts on
 Dim iCurrent  ' Variable we use to hold current day of month as we write table
 Dim iPosition ' Variable we use to hold current position in table

'Set variables for values passed in
 if isDate(request("date")) then
    lcl_date = CDate(request("date"))
 else
    lcl_date = ""
 end if

 If IsNumeric(request("month")) Then 
   	lcl_month = CLng(request("month"))
 Else
   	response.redirect "calendar.asp"
 End If 

 If IsNumeric(request("day")) Then 
   	lcl_day   = CLng(request("day"))
 Else
   	response.redirect "calendar.asp"
 End If

 If IsNumeric(request("year")) Then 
   	lcl_year  = CLng(request("year"))
 Else
   	response.redirect "calendar.asp"
 End If

'Get selected date
If IsDate(lcl_date) Then
	  dDate = CDate(lcl_date)
Else
	  If IsDate(lcl_month & "-1-" & lcl_year) Then
		    dDate = CDate(lcl_month & "-1-" & lcl_year)
	  Else
		    dDate = Date()
		   'The annoyingly bad solution for those of you running IIS3
		    If Len(lcl_month) <> 0 Or Len(lcl_day) <> 0 Or Len(lcl_year) <> 0 Or Len(lcl_date) <> 0 Then
			      lcl_message = "The date you picked was not a valid date.  The calendar was set to today's date.<br /><br />"
		    End If
	  End If
End If

 bFilter = 0
 If IsNumeric(request("Category")) Then 
   	iCategory = CLng(Request("Category"))
 Else
 	  iCategory = CLng(0)
 End If

 if NOT iCategory = "" then
    if iCategory > CLng(0) then
      	bFilter = 1
    else
       bFilter = 0
    end if
 end if

'Now we've got the date.  Now get Days in the choosen month and the day of the week it starts on.
iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
iDOW = GetWeekdayMonthStartsOn(dDate)

'Check to see if the org has:
'1. the "Calendar Request" option "turned-on"
'2. what form to use for the Calendar Request if #1 is "turned-on"

'These options can be found/set within "Organization Features => Properties".
 checkForCalendarRequest blnCalRequest,iCalForm
%>

<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
	
	<title><%=langBSEvents%><%=lcl_calendarfeature_name%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="eventstyles.css" />

	<script src="../scripts/selectAll.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>

	<style type="text/css">
	<!--
		body {scrollbar-base-color:#6699cc; scrollbar-highlight-color:#ffffff; scrollbar-arrow-color:#99ccff;}
		.cal {border-left:1px solid #93bee1; border-top:1px solid #93bee1;}
		.cal td {border-right:1px solid #93bee1; border-bottom:1px solid #93bee1; font-family:Tahoma,Arial; font-size:11px;}
	//-->
	</style>

	<script language="javascript">
	<!--
		function goToCalendar(p_cal) 
		{
			if(p_cal!="" || p_cal!=undefined) 
			{
				location.href='calendar.asp?cal='+p_cal;
			}
			else
			{
				location.href='calendar.asp';
			}
		}

		function SubmitByMonth() 
		{
			document.frmDate.day.value = "01";
			document.frmDate.submit();
		}

		function PrintPreview() 
		{
			window.print();
		}
	//-->
	</script>
</head>

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<body topmargin="0" leftmargin="0" bottommargin="0" rightmargin="0" marginwidth="0" marginheight="0">
  
<div id="content">
	 <div id="centercontent">

    <form name="frmDate" id="frmDate" action="calendar.asp" method="post">
      <input type="hidden" name="day" value="<%= Day(Now) %>" />
    <p>
    <table border="0" cellspacing="0" cellpadding="0" width="90%">
      <tr valign="top">
          <td>
              <p><h3>Calendar of Events<%=lcl_calendarfeature_name%></h3></p>
          </td>
          <td align="right">
            <p>
            <%
              if lcl_hasCustomCalendars = "Y" then
                  response.write "<strong>Show Calendar: </strong>" & vbcrlf
                  response.write "<select name=""cal"" onChange=""goToCalendar(this.value)"">" & vbcrlf
'                  response.write "  <option value="""">Community Calendar</option>" & vbcrlf
                  displayCustomCalendarOptions session("orgid"), lcl_calendarfeatureid
                  response.write "</select>" & vbcrlf
              else
                  response.write "&nbsp;" & vbcrlf
              end if
            %>
            </p>
          </td>
      </tr>
      <tr>
          <td colspan="2">
           <% if lcl_hasCustomCalendars = "Y" then %>
              <div id="itemshow" align="center">
                <strong>
                <%
                  if blnCalRequest then
                     response.write "<a href=""../action_line/action.asp?actionid=" & iCalForm & lcl_calendarfeature_url & """>REQUEST to add Calendar Item</a> | " & vbcrlf
                  end if
                %>
              		<a href="#listView">View Calendar by Event List</a> |
              		<a href="searchevents.asp<%=replace(lcl_calendarfeature_url,"&","?")%>">SEARCH for Calendar Items</a> |
          		    <a href="#" onclick="javascript:PrintPreview();">Print the Calendar</a>
               	</strong>
              </div>
              <%
                if OrgHasDisplay( session("orgid"), "calendar notice" ) then
                   response.write "<p id=""calendarnotice"">" & vbcrlf
                   GetOrgDisplay session("orgid"), "calendar notice"
                   response.write "</p>" & vbcrlf
                end if

              end if
              %>
          </td>
      </tr>
    </table>
    </p>

  <table border="0" cellpadding="3" cellspacing="0" bgcolor="#ffffff" class="cal" width="100%" height="100%" id="calendarview">
    <tr height="2%">
      <td bgcolor="#336699" align="center" colspan="7">
        <table border="1" cellspacing="0" cellpadding="3" width="100%" height="100%" style="background-color: #336699; border-style: none; left: 0px; top: 0px;">
          <tr>
            <td height="30" width="60%" align="left" style="border:0px;" nowrap="nowrap">
                <span class="noprint">
                  &nbsp;<a href="calendar.asp?date=<%= SubtractOneMonth(dDate) %><%=lcl_calendarfeature_url%>"><img src="../images/arrow_back.gif" align="absmiddle" border="0" /></a>
                  &nbsp;<a href="calendar.asp?date=<%= SubtractOneMonth(dDate) %><%=lcl_calendarfeature_url%>"><font family="Tahoma,Arial" color="#ffffff" size="1"><%=langPreviousMonth%></font></a>
                </span>
                &nbsp;&nbsp;
                <select id="month" name="month" onChange="SubmitByMonth();">
                  <option value="1"><%=langMonth01%></option>
                  <option value="2"><%=langMonth02%></option>
                  <option value="3"><%=langMonth03%></option>
                  <option value="4"><%=langMonth04%></option>
                  <option value="5"><%=langMonth05%></option>
                  <option value="6"><%=langMonth06%></option>
                  <option value="7"><%=langMonth07%></option>
                  <option value="8"><%=langMonth08%></option>
                  <option value="9"><%=langMonth09%></option>
                  <option value="10"><%=langMonth10%></option>
                  <option value="11"><%=langMonth11%></option>
                  <option value="12"><%=langMonth12%></option>
                </select>
                <select id="year" name="year" onChange="document.getElementById('frmDate').submit();">
                		<%
                    for x = (Year(dDate) - 5) to (Year(dDate) + 5)
                        response.write "  <option value=""" & x & """>" & x & "</option>" & vbcrlf
                    next
                  %>
                </select>
                <script language="javascript">
                  document.getElementById("month").selectedIndex = <%= Month(dDate)-1 %>;
                  document.getElementById("year").value = <%= Year(dDate) %>;
                </script>
                <span class="noprint">
                  &nbsp;&nbsp;&nbsp;
                  <a href="calendar.asp?date=<%= AddOneMonth(dDate) %><%=lcl_calendarfeature_url%>"><font family="Tahoma,Arial" color="#ffffff" size="1"><%=langNextMonth%></font></a>&nbsp;
                  <a href="calendar.asp?date=<%= AddOneMonth(dDate) %><%=lcl_calendarfeature_url%>"><img src="../images/arrow_forward.gif" align="absmiddle" border="0"></a>&nbsp;
                  <img src="../images/spacer.gif" width="1" height="30" border="0" align="absmiddle" />
                </span>
            </td>
            <td align="right" width="40%" style="border:0px;" nowrap="nowrap">
                <div class="noprint">
                  <font color="#ffffff">View:</font>
                  <select name="Category" class="time" onchange="document.getElementById('frmDate').submit();">
                  		<option value="0">All Categories</option>
                    <% getEventCategoryOptions session("orgid"), lcl_calendarfeature, iCategory %>
                  </select>
               	</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#93bee1" align="center" style="font-size: 8pt; color: #003366; height: 8px; font-weight: bold;">
      <td width="14.25%"><%=langDay1%></td>
      <td width="14.25%"><%=langDay2%></td>
      <td width="14.25%"><%=langDay3%></td>
      <td width="14.25%"><%=langDay4%></td>
      <td width="14.25%"><%=langDay5%></td>
      <td width="14.25%"><%=langDay6%></td>
      <td width="14.25%"><%=langDay7%></td>
    </tr>
    <%
    ' Write spacer cells at beginning of first row if month doesn't start on a Sunday.
    If iDOW <> 1 Then
      response.write vbtab & "<tr>" & vbcrlf
      iPosition = 1
      Do While iPosition < iDOW
        response.write vbtab & vbtab & "<td>&nbsp;</td>" & vbcrlf
        iPosition = iPosition + 1
      Loop
    End If

    ' Write days of month in proper day slots
    iCurrent = 1
    iPosition = iDOW
    Do While iCurrent <= iDIM
      ' If we're at the begginning of a row then write TR
      If iPosition = 1 Then
        response.write vbtab & "<tr>" & vbcrlf
      End If

      ' If the day we're writing is the selected day then highlight it somehow.
      If iCurrent = Day(dDate) And Month(dDate) = Month(Now()) And Year(dDate) = Year(Now()) Then
        response.write vbtab & vbtab & "<td id=""day_" & iCurrent & """ style=""border:2px solid #ff0000; valign=""top""><font size=""1"">" & iCurrent & "</font></td>" & vbcrlf
      Else
        response.write vbtab & vbtab & "<td id=""day_" & iCurrent & """ valign=""top""><font size=""1"">" & iCurrent & "</font></td>" & vbcrlf
      End If
      
      ' If we're at the endof a row then write /TR
      If iPosition = 7 Then
        response.write vbtab & "</tr>" & vbcrlf
        iPosition = 0
      End If
      
      ' Increment variables
      iCurrent = iCurrent + 1
      iPosition = iPosition + 1
    Loop

    ' Write spacer cells at end of last row if month doesn't end on a Saturday.
    If iPosition <> 1 Then
      Do While iPosition <= 7
        response.write vbtab & vbtab & "<td>&nbsp;</td>" & vbcrlf
        iPosition = iPosition + 1
      Loop
      response.write vbtab & "</tr>" & vbcrlf
    End If


   'Determine what days we have events on and mark in calendar --------------------------------

   'Setup the query to pull all of the events for the month
    lcl_sql = "SELECT e.eventdate, e.subject, c.color "
    lcl_sql = lcl_sql & " FROM events as e "
    lcl_sql = lcl_sql & " LEFT JOIN eventcategories as c ON e.categoryid = c.categoryid "
    lcl_sql = lcl_sql & " WHERE e.orgid = " & session("orgid")
    lcl_sql = lcl_sql & " AND datediff(mm, e.eventdate, '" & dDate & "') = 0 "
    lcl_sql = lcl_sql & " AND datediff(yy, e.eventdate, '" & dDate & "') = 0 "

    if lcl_calendarfeature <> "" then
       lcl_sql = lcl_sql & " AND UPPER(e.calendarfeature) = '" & UCASE(lcl_calendarfeature) & "' "
    else
       'lcl_sql = lcl_sql & " AND (e.calendarfeature = '' OR e.calendarfeature IS NULL)"
       lcl_sql = lcl_sql & " AND (e.calendarfeature <> '' OR e.calendarfeature IS NOT NULL)"
    end if

    lcl_sql = lcl_sql & " AND '" & lcl_hasCustomCalendars & "' = 'Y' "

   'Setup the ORDER BY
    lcl_orderby = " ORDER BY e.eventdate "

   'Build the query
    sSQL = lcl_sql & lcl_orderby

    set oRst = Server.CreateObject("ADODB.Recordset")
    oRst.Open sSQL, Application("DSN"), 3, 1

    sScript = ""

    If Not oRst.EOF Then
      sScript = "<script language=""javascript"">"
      Do While Not oRst.EOF
        iDay = Day(oRst("EventDate"))

                   'Determine if the events returned for the month are filtered or not.
                   'FILTERED
                    if bFilter then
                       sSQL2 = lcl_sql & " AND e.categoryid = " & iCategory
                       sSQL2 = sSQL2 & lcl_orderby
                   'NOT FILTERED
                    else
                       sSQL2 = lcl_sql & lcl_orderby
                    end if

                    set oRst2 = Server.CreateObject("ADODB.Recordset")
                    oRst2.Open sSQL2, Application("DSN"), 3, 1

                  		perday  = 0
                  		sEvents = ""

                  		if NOT oRst2.eof then
                       do while NOT oRst2.eof
                          if Day(oRst2("EventDate")) = iDay then
                         				perday = perday + 1

                        					if perday > 4 then
                          						sEvents = sEvents & "<li><a href='calendarevents.asp?date="& Month(dDate) & "-" & iDay & "-" & Year(dDate) & "'>More...</a></li>"
                          						exit do
                         				else
                          						truncMessage = oRst2("Subject")
                         						'dcajacob 6/15/05 removed at request of Peter,John
                         						'If Len(truncMessage) > 22 Then
                             						'truncMessage=Left(truncMessage,19) & "..."
                         						'End If
                          						sEvents = sEvents & "<li><span style='cursor:pointer; cursor:hand;' onclick='document.location.href=\""calendarevents.asp?date="& Month(dDate) & "-" & iDay & "-" & Year(dDate) & "\""'><font style='font-size: 8pt; font-family:Tahoma,Arial; text-transform:uppercase; color:" & oRst2("Color") & "'>" & escDblQuote(truncMessage) & "</font></span></li>"
                             end if
                      				end if

                       			oRst2.movenext
                       loop
                    end if

                  		oRst2.close
                  	 set oRst2 = nothing

                  		sTemp = "document.getElementById('day_" & iDay & "').innerHTML = ""<strong>" & iDay & "</strong><br />" & sEvents & """;" & vbcrlf
                  		'sScript = sScript & sTemp & "document.all.day_" & iDay & ".style.backgroundColor = '#99ccff';" &  vbCrLf
                   	sScript = sScript & sTemp & vbcrlf
        oRst.MoveNext
      Loop
      sScript = sScript & "</script>"
    End If
    Set oRst = Nothing

    response.write sScript
    %>

  </table>
<!--</div>-->

</form>

          <a name="listView"></a>
<%
          sEvents = ""

         'Setup the query to pull all of the details for the events for a month
          lcl_sql2 = "SELECT e.eventid, e.eventdate, e.eventduration, t.TZabbreviation, e.subject, e.message, e.categoryid, c.color "
          lcl_sql2 = lcl_sql2 & " FROM events as e "
          lcl_sql2 = lcl_sql2 &   " LEFT JOIN timezones as t ON t.timezoneid = e.eventtimezoneid "
          lcl_sql2 = lcl_sql2 &   " LEFT JOIN eventcategories as c ON e.categoryid = c.categoryid "
          lcl_sql2 = lcl_sql2 & " WHERE e.orgid = " & session("orgid")
          lcl_sql2 = lcl_sql2 & " AND datediff(mm, e.eventdate, '" & dDate & "') = 0 "
          lcl_sql2 = lcl_sql2 & " AND datediff(yy, e.eventdate, '" & dDate & "') = 0 "

          if lcl_calendarfeature <> "" then
             lcl_sql2 = lcl_sql2 & " AND UPPER(e.calendarfeature) = '" & UCASE(lcl_calendarfeature) & "' "
          else
             lcl_sql2 = lcl_sql2 & " AND (e.calendarfeature = '' OR e.calendarfeature IS NULL)"
          end if

          lcl_sql2 = lcl_sql2 & " AND '" & lcl_hasCustomCalendars & "' = 'Y' "

         'Setup the order by
          lcl_orderby2 = " ORDER BY e.eventdate "

          if bFilter then
             sSQL2 = lcl_sql2 & " AND e.categoryid = " & iCategory
             sSQL2 = sSQL2 & lcl_orderby2
          else
             sSQL2 = lcl_sql2 & lcl_orderby2
          end if

          set oRst = Server.CreateObject("ADODB.Recordset")
          oRst.Open sSQL2, Application("DSN"), 3, 1

          if Not oRst.EOF then
             lcl_bgcolor = "#eeeeee"
             Do while Not oRst.EOF
               	'if oRst("EventDuration") > 0 then
                '	  dEnd = DateAdd("n",oRst("EventDuration"),oRst("EventDate"))

                '	  if DateDiff("d",dEnd,oRst("EventDate")) = 0 then
                ' 	    dEnd = FormatDateTime(dEnd,vbLongTime)
                '   end if

                '	  dEnd = " - " & dEnd
               	'else
                '	  dEnd = ""
                'end if

				If oRst("EventDuration") > 0 Then
					If CLng(oRst("EventDuration")) = CLng(1440) Then
						dEnd = ""
					Else
						dEnd = DateAdd("n",oRst("EventDuration"),oRst("EventDate"))

						If DateDiff("d",dEnd,oRst("EventDate")) = 0 Then 
							dEnd = FormatDateTime(dEnd,vbLongTime)
						End If 

						dEnd = " - " & dEnd
					End If 
				Else 
					dEnd = ""
				End If 

				if oRst("CategoryID") <> 0 then
					sCategory = "(" & getCategoryName(oRst("categoryid")) & ") "
				else
					sCategory = ""
				end if

				'Format the event date/time (remove the seconds from the date(s))
				lcl_event_date = oRst("eventdate")

				'This displays the "12:00:00 AM" IF a duration exists
				If Left(FormatDateTime(oRst("eventdate"),vbLongTime),11) = "12:00:00 AM" And oRst("eventduration") > 0 Then 
					If CLng(oRst("EventDuration")) <> CLng(1440) Then
						lcl_event_date = oRst("eventdate") & " " & FormatDateTime(oRst("eventdate"),vbLongTime)
					End If 
				End If 

				If oRst("eventduration") > 0 Then 
					If CLng(oRst("EventDuration")) <> CLng(1440) Then
						If Left(FormatDateTime(Replace(dEnd," - ", ""),vbLongTime),11) = "12:00:00 AM" then
							lcl_end_time = Replace(dEnd," - ", "")
							lcl_end_time = lcl_end_time & " " & FormatDateTime(lcl_end_time,vbLongTime)
							lcl_end_time = " - " & lcl_end_time
							dEnd         = lcl_end_time
						End If 
					End If 
				End If 

				formatEventDateTime lcl_event_date, dEnd, sDate1, sDate2

				if oRst("CategoryID") <> 0 then
					sCategory = "(" & getCategoryName(oRst("categoryid")) & ") "
				else
					sCategory = ""
				end if

             		'-------------------------------------------------------------
              	'Used to trim seconds from dates displayed on calendar pages.
             		'9/9/2005 Vincent Evans
             		'Start trim code
             		'-------------------------------------------------------------
              		'sDate1 = cStr(oRst("EventDate"))
              		'sDate2 = cStr(dEnd)

              		'iTrimDate1 = clng(InStrRev(sDate1,":"))
              		'iTrimDate2 = clng(InStrRev(sDate2,":"))

              	'Retrieves AM/PM, trims final :00 and builds string
              		'if iTrimDate1 > 0 then
                '			sTemp  = Right(sDate1, 2)
                '			sDate1 = Left(sDate1,iTrimDate1 - 1) & " " & sTemp
                '			sTemp  = ""
              		'end if

              		'if iTrimDate2 > 0 then
                '			sTemp  = Right(sDate2, 2)
                ' 		sDate2 = Left(sDate2,iTrimDate2 - 1) & " " & sTemp
                '			sTemp  = ""
              		'end if

             		'-------------------------------------------------------------
             		'End trim code
              	'-------------------------------------------------------------

                'sEvents = sEvents & vbcrlf & "<tr><td width=""25%"" valign=top>" & sDate1 & " " & sDate2 & " " & oRst("TZAbbreviation") & "</td><td><i><font color=""" & oRst("Color") & """>" & sCategory & "</i><strong>" & oRst("Subject") & "</font></strong><br />" & oRst("Message") & "</td></tr>"
              	sEvents = sEvents & "<tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
                sEvents = sEvents & "    <td width=""25%"" valign=""top"">" & sDate1 & " " & sDate2 & "</td>" & vbcrlf
                sEvents = sEvents & "    <td>" & vbcrlf
                sEvents = sEvents & "        <font style=""font-size: 8pt; color: " & oRst("Color") & ";""><i>" & sCategory & "</i><strong>" & oRst("Subject") & "</strong></font>" & vbcrlf
                sEvents = sEvents & "        <br />" & oRst("Message") & vbcrlf
                sEvents = sEvents & "    </td>" & vbcrlf
                sEvents = sEvents & "</tr>" & vbcrlf

                lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

                oRst.movenext
             Loop 
          end if

         	oRst.close
          set oRst = nothing
        %>
          <div align="center" class="calendartitle">Calendar Event List</div>
          <p>
          <table id="eventlist" border="0" cellpadding="4" cellspacing="0" width="70%" align="center" class="tablelist">
            <tr>
                <th><%=langDateTime%></th>
                <th><%=langEvent%></th>
            </tr>
<%
			If Len(sEvents) > 1 Then 
				response.write sEvents
			Else 
				response.write "<tr><td colspan=""2""><p><strong>There are no events currently scheduled for this day.</strong></p></td></tr>" & vbcrlf
			End If 
%>
          </table>
          </p>

  </div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
function escDblQuote( strDB )
 	if VarType( strDB ) <> vbString then
	   	escDblQuote = strDB 
 	else
	   	escDblQuote = Replace( strDB, Chr(34), "\" & Chr(34) )
 	end if
end function

'------------------------------------------------------------------------------
function checkForCalendarRequest(ByRef blnCalRequest, ByRef iCalForm)
  blnCalRequest = ""
  iCalForm      = ""

  sSQL = "SELECT OrgRequestCalOn, OrgRequestCalForm "
  sSQL = sSQL & " FROM organizations "
  sSQL = sSQL &      " INNER JOIN TimeZones ON Organizations.OrgTimeZoneID = TimeZones.TimeZoneID "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")

 	set oOrgInfo = Server.CreateObject("ADODB.Recordset")
 	oOrgInfo.Open sSQL, Application("DSN"), 3, 1
	
 	if NOT oOrgInfo.eof then
   		blnCalRequest = oOrgInfo("OrgRequestCalOn")
     iCalForm      = oOrgInfo("OrgRequestCalForm")
  end if

  oOrgInfo.close
  set oOrgInfo = nothing

end function
%>
