<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="events_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: calendar.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Main Calendar display.
'
' MODIFICATION HISTORY
'	1.?	11/30/07	Steve Loar - Changed to set day to the first when month is picked. Handles Feb EOM problem.
' 1.2 08/08/08 David Boyer - Added Custom Calendar
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()
session("RedirectLang") = "Return to Calendar"

'Check to see if this is a custom calendar
 if trim(request("cal")) <> "" then
    lcl_calendarfeature      = trim(request("cal"))
    lcl_calendarfeature_url  = "&cal=" & lcl_calendarfeature
    lcl_calendarfeature_name = " [" & getFeatureName(lcl_calendarfeature) & "]"
 else
    lcl_calendarfeature      = ""
    lcl_calendarfeature_url  = ""
    lcl_calendarfeature_name = ""
 end if

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
Dim dToday    ' Today

dToday = Date()

'Set variables for values passed in
 if isDate(request("date")) then
    lcl_date = request("date")
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
 	  dDate = lcl_date
 Else
 	  If IsDate(lcl_month & "-" & lcl_day & "-" & lcl_year) Then
 		    dDate = lcl_month & "-" & lcl_day & "-" & lcl_year
 	  Else
 		    dDate = Date()
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

'Now we have got the date.  Now get Days in the choosen month and the day of the week it starts on.
 iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
 iDOW = GetWeekdayMonthStartsOn(dDate)
%>

<html>
<head>
	<title>E-Gov Services - <%=sOrgName%></title>
 <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />
	<link rel="stylesheet" type="text/css" href="calendarprint.css" media="print" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script type="text/javascript" src="https://s7.addthis.com/js/200/addthis_widget.js"></script>
	<script type="text/javascript">var addthis_pub="cschappacher";</script>

	<script language="javascript">
	<!--
		function SubmitByMonth()
		{
			document.frmDate.day.value = "01";
			document.frmDate.submit();
		}

		function openWin2(url, name) 
		{
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
		}

		if(document.all && !document.getElementById) 
		{
			document.getElementById = function(id) 
			{
				 return document.all[id];
			}
		}

		function PrintPreview()
		{
			window.print();
		}

	//-->
	</script>

</head>

<!--#Include file="../include_top.asp"-->

<!--BODY CONTENT-->
<table border="0" cellspacing="0" cellpadding="0" width="90%">
  <tr>
      <td>
          <p>
          <table border="0" cellspacing="0" cellpadding="0" width="90%">
            <tr valign="top">
                <td>
                    <font class="pagetitle">Calendar of Events<%=lcl_calendarfeature_name%></font>
                    <% checkForRSSFeed iorgid, "", "COMMUNITYCALENDAR", sEgovWebsiteURL %>
                    <br />
                   	<%	RegisteredUserDisplay( "../" ) %>
                </td>
                <td align="right">
                    <table border="0" cellspacing="0" cellpadding="2">
                      <tr valign="top">
                          <td><% displayYahooBuzzButton iorgid %></td>
                          <td><% displayAddThisButton iorgid %></td>
                      </tr>
                    </table>
                </tr>
            </tr>
            <tr valign="top">
                <td align="right" colspan="2">
                <%
                  if checkForCustomCalendars(iorgid) = "Y" then
                     response.write "<strong>Show Calendar: </strong>" & vbcrlf
                     response.write "<select name=""calendarfeature"" onChange=""location.href='calendar.asp?cal='+this.value"">" & vbcrlf
                     response.write "  <option value="""">Community Calendar</option>" & vbcrlf
                     displayCustomCalendarOptions iorgid, lcl_calendarfeature
                     response.write "</select>" & vbcrlf
                  end if
                %>
                </td>
            </tr>
          </table>
          </p>
          <p style="padding-left:20px">
          <div id="itemshow" align="center">
            <strong>
            <%
              if blnCalRequest then
                 response.write "<a href=""../action.asp?actionid=" & iCalForm & lcl_calendarfeature_url & """>REQUEST to add Calendar Item</a> | " & vbcrlf
              end if
            %>
          		<a href="#listView">View Calendar by Event List</a> |
          		<a href="searchevents.asp<%=replace(lcl_calendarfeature_url,"&","?")%>">SEARCH for Calendar Items</a> |
          		<a href="#" onclick="javascript:PrintPreview();">Print the Calendar</a>
           	</strong>
          </div>
          </p>
          <%
            if OrgHasDisplay( iOrgId, "calendar notice" ) then
               response.write "<p id=""calendarnotice"">" & vbcrlf
               GetOrgDisplay iOrgId, "calendar notice"
               response.write "</p>" & vbcrlf
            end if
          %>
          <form id="frmDate" name="frmDate" action="calendar.asp" method="get">
            <input type="hidden" name="day" value="<%=Day(Now)%>" />
            <input type="hidden" name="cal" value="<%=request("cal")%>" />

          <table id="calendar" border="0" cellpadding="2" cellspacing="0" width="90%" height="90%">
            <tr height="2%">
                <th align="center" colspan="7">
                    <table id="calendarheader" width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                          <td height="30" width="15%" style="border:0px;" nowrap="nowrap">
                            		<span class="noprint">
                             			&nbsp;<a href="calendar.asp?date=<%=SubtractOneMonth(dDate) & lcl_calendarfeature_url%>"><img src="../images/arrow_back.gif" align="absmiddle" border="0" /></a>&nbsp;
                                      <a href="calendar.asp?date=<%=SubtractOneMonth(dDate) & lcl_calendarfeature_url%>"><%=langPreviousMonth%></a>
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
                                <a href="calendar.asp?date=<%=AddOneMonth(dDate) & lcl_calendarfeature_url%>"><%=langNextMonth%></a>&nbsp;
                                <a href="calendar.asp?date=<%=AddOneMonth(dDate) & lcl_calendarfeature_url%>"><img src="../images/arrow_forward.gif" align="absmiddle" border="0"></a>&nbsp;
                                <img src="../images/spacer.gif" width="1" height="30" border="0" align="absmiddle" />
                			           </span>
                      			 </td>
                          <td align="right" width="50%" nowrap="nowrap">
                              <div class="noprint">
                                View:
                                <select name="Category" class="time" onchange="document.getElementById('frmDate').submit();">
                  		              <option value="0">All Categories</option>
                                  <% getEventCategoryOptions iorgid, lcl_calendarfeature, iCategory %>
                                </select>
               			            </div>
                          </td>
                      </tr>
                    </table>
                </th>
            </tr>
            <tr id="calendardayrow" height="10">
                <td width="13%" height="10">Sun</td>
                <td width="13%" height="10">Mon</td>
                <td width="13%" height="10">Tues</td>
                <td width="13%" height="10">Wed</td>
                <td width="13%" height="10">Thurs</td>
                <td width="13%" height="10">Fri</td>
                <td width="13%" height="10">Sat</td>
            </tr>
            <%
             'Write spacer cells at beginning of first row if month doesn't start on a Sunday.
              If iDOW <> 1 Then
                 response.write "<tr>" & vbcrlf
                 iPosition = 1
                 Do While iPosition < iDOW
                    response.write "<td>&nbsp;</td>" & vbcrlf
                    iPosition = iPosition + 1
                 Loop
              End If

             'Write days of month in proper day slots
              iCurrent  = 1
              iPosition = iDOW
              Do While iCurrent <= iDIM

                'If we're at the begginning of a row then write TR
                 If iPosition = 1 Then
                    response.write "<tr>" & vbcrlf
                 End If

                'If the day we're writing is the selected day then highlight it somehow.
                 if iCurrent = Day(Date()) AND Month(dDate) = Month(Date()) AND Year(dDate) = Year(Date()) then
                 	  'If Day(dToday) = iCurrent And Month(dToday) = Month(dDate) And Year(dToday) = Year(dDate) Then
                    response.write "    <td height=""55"" id=""day_" & iCurrent & """ valign=""top"">" & vbcrlf
                    response.write "        <strong>" & iCurrent & " - Today</strong>" & vbcrlf
                    response.write "    </td>" & vbcrlf
                 else
                    response.write "    <td height=""55"" id=""day_" & iCurrent & """ valign=""top"">" & vbcrlf
                    response.write "        <strong>" & iCurrent & "</strong>" & vbcrlf
                    response.write "    </td>" & vbcrlf
                 end if

                'If we're at the endof a row then write /TR
                 if iPosition = 7 then
                    response.write "</tr>" & vbcrlf
                    iPosition = 0
                 end if

                'Increment variables
                 iCurrent  = iCurrent  + 1
                 iPosition = iPosition + 1
              Loop

             'Write spacer cells at end of last row if month doesn't end on a Saturday.
              if iPosition <> 1 then
                 Do While iPosition <= 7
                    response.write "    <td>&nbsp;</td>" & vbcrlf
                    iPosition = iPosition + 1
                 Loop
                 response.write "</tr>" & vbcrlf
              end if

             'Determine what days we have events on and mark in calendar --------------------------------

             'Setup the query to pull all of the events for the month
              lcl_sql = "SELECT e.eventdate, e.subject, c.color "
              lcl_sql = lcl_sql & " FROM events as e "
              lcl_sql = lcl_sql &      " LEFT JOIN eventcategories as c ON e.categoryid = c.categoryid "
              lcl_sql = lcl_sql & " WHERE e.orgid = " & iorgid
              lcl_sql = lcl_sql & " AND datediff(mm, e.eventdate, '" & dDate & "') = 0 "
              lcl_sql = lcl_sql & " AND datediff(yy, e.eventdate, '" & dDate & "') = 0 "

              if lcl_calendarfeature <> "" then
                 lcl_sql = lcl_sql & " AND UPPER(e.calendarfeature) = '" & UCASE(lcl_calendarfeature) & "' "
              else
                 lcl_sql = lcl_sql & " AND (e.calendarfeature = '' OR e.calendarfeature IS NULL)"
              end if

             'Setup the ORDER BY
              lcl_orderby = " ORDER BY e.eventdate "

             'Build the query
              sSQL = lcl_sql & lcl_orderby

              set oRst = Server.CreateObject("ADODB.Recordset")
              oRst.Open sSQL, Application("DSN"), 3, 1

              sScript = ""

              if NOT oRst.eof then
                 sScript = "<script language=""javascript"">" & vbcrlf

                 while NOT oRst.eof

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
                          						sEvents = sEvents & "<li><a href='calendarevents.asp?date="& Month(dDate) & "-" & iDay & "-" & Year(dDate) & lcl_calendarfeature_url & "'>More...</a></li>"
                          						exit do
                         				else
                          						truncMessage = oRst2("Subject")
                         						'dcajacob 6/15/05 removed at request of Peter,John
                         						'If Len(truncMessage) > 22 Then
                             						'truncMessage=Left(truncMessage,19) & "..."
                         						'End If
                          						sEvents = sEvents & "<li><span style='cursor:pointer; cursor:hand;' onclick='document.location.href=\""calendarevents.asp?date="& Month(dDate) & "-" & iDay & "-" & Year(dDate) & lcl_calendarfeature_url & "\""'><font size=1 style='text-transform:uppercase' color='" & oRst2("Color") & "'>" & escDblQuote(truncMessage) & "</font></span></li>"
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

                  		oRst.movenext
	                wend

                	sScript = sScript & "</script>" & vbcrlf

              end if

             	oRst.close
              set oRst = nothing

              response.write sScript
            %>

          </table>
          </form>

          <a name="listView"></a>
        <%
          sEvents = ""

         'Setup the query to pull all of the details for the events for a month
          lcl_sql2 = "SELECT e.eventid, e.eventdate, e.eventduration, t.TZabbreviation, e.subject, e.message, e.categoryid, c.color "
          lcl_sql2 = lcl_sql2 & " FROM events as e "
          lcl_sql2 = lcl_sql2 &   " LEFT JOIN timezones as t ON t.timezoneid = e.eventtimezoneid "
          lcl_sql2 = lcl_sql2 &   " LEFT JOIN eventcategories as c ON e.categoryid = c.categoryid "
          lcl_sql2 = lcl_sql2 & " WHERE e.orgid = " & iorgid
          lcl_sql2 = lcl_sql2 & " AND datediff(mm, e.eventdate, '" & dDate & "') = 0 "
          lcl_sql2 = lcl_sql2 & " AND datediff(yy, e.eventdate, '" & dDate & "') = 0 "

          if lcl_calendarfeature <> "" then
             lcl_sql2 = lcl_sql2 & " AND UPPER(e.calendarfeature) = '" & UCASE(lcl_calendarfeature) & "' "
          else
             lcl_sql2 = lcl_sql2 & " AND (e.calendarfeature = '' OR e.calendarfeature IS NULL)"
          end if

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

          if NOT oRst.eof then
             while NOT oRst.eof
               	if oRst("EventDuration") > 0 then
                	  dEnd = DateAdd("n",oRst("EventDuration"),oRst("EventDate"))

                	  if DateDiff("d",dEnd,oRst("EventDate")) = 0 then
                 	    dEnd = FormatDateTime(dEnd,vbLongTime)
                   end if

                	  dEnd = " - " & dEnd
               	else
                	  dEnd = ""
                end if

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
              		sDate1 = cStr(oRst("EventDate"))
              		sDate2 = cStr(dEnd)

              		iTrimDate1 = clng(InStrRev(sDate1,":"))
              		iTrimDate2 = clng(InStrRev(sDate2,":"))

              	'Retrieves AM/PM, trims final :00 and builds string
              		if iTrimDate1 > 0 then
                			sTemp  = Right(sDate1, 2)
                			sDate1 = Left(sDate1,iTrimDate1 - 1) & " " & sTemp
                			sTemp  = ""
              		end if

              		if iTrimDate2 > 0 then
                			sTemp  = Right(sDate2, 2)
                 		sDate2 = Left(sDate2,iTrimDate2 - 1) & " " & sTemp
                			sTemp  = ""
              		end if

             		'-------------------------------------------------------------
             		'End trim code
              	'-------------------------------------------------------------

                'sEvents = sEvents & vbcrlf & "<tr><td width=""25%"" valign=top>" & sDate1 & " " & sDate2 & " " & oRst("TZAbbreviation") & "</td><td><i><font color=""" & oRst("Color") & """>" & sCategory & "</i><strong>" & oRst("Subject") & "</font></strong><br />" & oRst("Message") & "</td></tr>"
              		sEvents = sEvents & "<tr>" & vbcrlf
                sEvents = sEvents & "    <td width=""25%"" valign=""top"">" & sDate1 & " " & sDate2 & "</td>" & vbcrlf
                sEvents = sEvents & "    <td><i><font color=""" & oRst("Color") & """>" & sCategory & "</i><strong>" & oRst("Subject") & "</strong></font><br />" & oRst("Message") & "</td>" & vbcrlf
                sEvents = sEvents & "</tr>" & vbcrlf

                oRst.movenext
             wend
          end if

         	oRst.close
          set oRst = nothing
        %>
          <p>
          <!-- formerly the table was class="cal" -->
          <table id="eventlist" border="0" cellpadding="4" cellspacing="0" width="70%" align="center">
            <tr>
                <td colspan="2" align="center"><strong>Calendar Event List</strong></td>
            </tr>
            <tr>
                <th><%=langDateTime%></th>
                <th><%=langEvent%></th>
            </tr>
          <%
            if len(sEvents) > 1 then
            			response.write sEvents
           	else
            			response.write "<tr>"  & vbcrlf
               response.write "    <td colspan=""2""><p><strong>There are no events currently scheduled for this day.</strong></p></td>" & vbcrlf
               response.write "</tr>" & vbcrlf
          		end if
         	%>
          </table>
          </p>

          <p>&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</p>

<!--#Include file="../include_bottom.asp"-->

<%
'------------------------------------------------------------------------------
' BEGIN: VISITOR TRACKING
'------------------------------------------------------------------------------
	iSectionID     = 5
	sDocumentTitle = "MAIN"
	sURL           = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
	datDate        = Date()
	datDateTime    = Now()
	sVisitorIP     = request.servervariables("REMOTE_ADDR")
	Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
'------------------------------------------------------------------------------
' END: VISITOR TRACKING
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' Function escDblQuote( strDB )
'------------------------------------------------------------------------------
Function escDblQuote( strDB )
 	If VarType( strDB ) <> vbString Then 
	   	escDblQuote = strDB 
 	Else 
	   	escDblQuote = Replace( strDB, Chr(34), "\" & Chr(34) )
 	End If 
End Function

'-------------------------------------------------------------------------------------------------
function GetFeatureName( sFeature )
	Dim sSQL, oFeature

	sSQL = "SELECT isnull(FO.featurename,F.featurename) as featurename "
	sSQL = sSQL & " FROM egov_organizations_to_features FO, egov_organization_features F "
	sSQL = sSQL & " WHERE FO.featureid = F.featureid "
 sSQL = sSQL & " AND FO.orgid = " & iorgid
 sSQL = sSQL & " AND feature = '" & sFeature & "'" 

	set oFeature = Server.CreateObject("ADODB.Recordset")
	oFeature.Open sSQL, Application("DSN"), 3, 1

	If Not oFeature.EOF Then
		  GetFeatureName = oFeature("featurename")
	Else
		  GetFeatureName = ""
	End If 

	oFeature.close
	set oFeature = nothing 
end function
%>
