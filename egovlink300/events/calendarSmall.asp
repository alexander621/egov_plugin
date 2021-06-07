
<%
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

' Get selected date
If IsDate(Request.QueryString("date")) Then
	dDate = CDate(Request.QueryString("date"))
Else
	If IsDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year")) Then
		dDate = CDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year"))
	Else
		dDate = Date()
		' The annoyingly bad solution for those of you running IIS3
		If Len(Request.QueryString("month")) <> 0 Or Len(Request.QueryString("day")) <> 0 Or Len(Request.QueryString("year")) <> 0 Or Len(Request.QueryString("date")) <> 0 Then
			Response.Write "The date you picked was not a valid date.  The calendar was set to today's date.<BR><BR>"
		End If
	End If
End If

'Now we've got the date.  Now get Days in the choosen month and the day of the week it starts on.
iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
iDOW = GetWeekdayMonthStartsOn(dDate)
%>

<style type="text/css">
  <!--
    body {scrollbar-base-color:#6699cc; scrollbar-highlight-color:#ffffff; scrollbar-arrow-color:#99ccff;}
    .cal {border-left:1px solid #1c4aab; border-top:1px solid #1c4aab;background-color: #ffffff;font-size:11px;}
    .cal td {border-right:1px solid #1c4aab; border-bottom:1px solid #1c4aab; font-family:Tahoma,Arial; font-size:11px;}
    select {font-family:Arial,Tahoma,Verdana; font-size:13px;}
  //-->
  </style>


<div style="cursor:hand;" onclick="document.location.href='events/calendar.asp';" title="View Calendar" >
  <form name="frmDate" action="calendar.asp" method="get">
  <input type="hidden" name="day" value="<%= Day(Now) %>">
  
  <table border="0" cellpadding="2" cellspacing="0" bgcolor="#ffffff" class="cal" width="100%" height="100%">
    <tr>
	<td ALIGN="center" BGCOLOR="#1c4aab" style="font-weight:bold; font-family:Tahoma,Arial; font-size:11px; color:#99CCFF;" colspan=7><%= MonthName(Month(Now())) & " " & Year(Now())%></td>
    </tr>
    <!--<tr>
      <td ALIGN="center" BGCOLOR=#1c4aab width="13%"><FONT COLOR=#ffffff size=1>S</FONT></td>
      <td ALIGN="center" BGCOLOR=#1c4aab width="13%"><FONT COLOR=#ffffff size=1>M</FONT></td>
      <td ALIGN="center" BGCOLOR=#1c4aab width="13%"><FONT COLOR=#ffffff size=1>T</FONT></td>
      <td ALIGN="center" BGCOLOR=#1c4aab width="13%"><FONT COLOR=#ffffff size=1>W</FONT></td>
      <td ALIGN="center" BGCOLOR=#1c4aab width="13%"><FONT COLOR=#ffffff size=1>TH</FONT></td>
      <td ALIGN="center" BGCOLOR=#1c4aab width="13%"><FONT COLOR=#ffffff size=1>F</FONT></td>
      <td ALIGN="center" BGCOLOR=#1c4aab width="13%"><FONT COLOR=#ffffff size=1>S</FONT></td>
    </TR>-->
    
    <%
    ' Write spacer cells at beginning of first row if month doesn't start on a Sunday.
    If iDOW <> 1 Then
      Response.Write vbTab & "<tr>" & vbCrLf
      iPosition = 1
      Do While iPosition < iDOW
        Response.Write vbTab & vbTab & "<td>&nbsp;</td>" & vbCrLf
        iPosition = iPosition + 1
      Loop
    End If

    ' Write days of month in proper day slots
    iCurrent = 1
    iPosition = iDOW
    Do While iCurrent <= iDIM
      ' If we're at the begginning of a row then write TR
      If iPosition = 1 Then
        Response.Write vbTab & "<tr>" & vbCrLf
      End If

      ' If the day we're writing is the selected day then highlight it somehow.
      If iCurrent = Day(dDate) And Month(dDate) = Month(Now()) And Year(dDate) = Year(Now()) Then
        Response.Write vbTab & vbTab & "<td bgcolor=#1c4aab height=""15%"" id=""day_" & iCurrent & """ valign=""top""><font color=white>" & iCurrent & "</font></td>" & vbCrLf
      Else
        Response.Write vbTab & vbTab & "<td height=""15%"" id=""day_" & iCurrent & """ valign=""top"">" & iCurrent & "</td>" & vbCrLf
      End If
      
      ' If we're at the endof a row then write /TR
      If iPosition = 7 Then
        Response.Write vbTab & "</tr>" & vbCrLf
        iPosition = 0
      End If
      
      ' Increment variables
      iCurrent = iCurrent + 1
      iPosition = iPosition + 1
    Loop

    ' Write spacer cells at end of last row if month doesn't end on a Saturday.
    If iPosition <> 1 Then
      Do While iPosition <= 7
        Response.Write vbTab & vbTab & "<td>&nbsp;</td>" & vbCrLf
        iPosition = iPosition + 1
      Loop
      Response.Write vbTab & "</tr>" & vbCrLf
    End If


    '------------------------------------------------------ determine what days we have events on and mark in calendar
   ' Dim oCmd, oRst, sScript, iDay

    Set oCmd = Server.CreateObject("ADODB.Command")
    With oCmd
      ''**.ActiveConnection = Application("DSN")
      .ActiveConnection = Application("DSN")
      .CommandText = "ListMonthEvents"
      .CommandType = adCmdStoredProc
      ''**.Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
      .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, iorgid)
      .Parameters.Append oCmd.CreateParameter("Date", adDateTime, adParamInput, 4, dDate)
    End With
                
    Set oRst = Server.CreateObject("ADODB.Recordset")
    With oRst
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockReadOnly
      .Open oCmd
    End With
    Set oCmd = Nothing

    sScript = ""

    If Not oRst.EOF Then
      sScript = "<script language=""Javascript"">"
      Do While Not oRst.EOF
        iDay = Day(oRst("EventDate"))
      
  	  Set oCmd2 = Server.CreateObject("ADODB.Command")
	    With oCmd2
	      .ActiveConnection = Application("DSN")
	      .CommandText = "ListMonthEvents"
	      .CommandType = adCmdStoredProc
	      .Parameters.Append oCmd2.CreateParameter("OrgID", adInteger, adParamInput, 4, iorgid)
	      .Parameters.Append oCmd2.CreateParameter("Date", adDateTime, adParamInput, 4, dDate)
	    End With

          Set oRst2 = Server.CreateObject("ADODB.Recordset")
	    With oRst2
	      .CursorLocation = adUseClient
	      .CursorType = adOpenStatic
	      .LockType = adLockReadOnly
	      .Open oCmd2
	    End With
	 Set oCmd2 = Nothing

	perday = 0
	Do While Not oRst2.EOF
		if Day(oRst2("EventDate")) = iDay then
			perday = perday + 1
		end if

		oRst2.MoveNext
	Loop


	
	        'sScript = sScript & "document.all.day_" & iDay & ".innerHTML = ""<span style='width:100%; height:100%; cursor:hand;' onclick='document.location.href=\""events/calendar.asp\""'>" & iDay & "<br><font size=1>" & perday & " Events</font></span>"";" & vbCrLf
	        'sScript = sScript & "document.all.day_" & iDay & ".innerHTML = ""<span style='width:100%; height:100%; cursor:hand;' onclick='document.location.href=\""events/calendar.asp?date="& Month(dDate) & "-" & iDay & "-" & Year(dDate) & "\""'>" & iDay & "<br><font size=1>" & perday & " Events</font></span>"";" & vbCrLf

	        sScript = sScript & "document.all.day_" & iDay & ".innerHTML = ""<span style='width:100%; height:100%; cursor:hand;' onclick='document.location.href=\""events/calendar.asp\""'>" & iDay & "</span>"";" & vbCrLf
	        'sScript = sScript & "document.all.day_" & iDay & ".innerHTML = ""<span style='width:100%; height:100%; cursor:hand;' onclick='document.location.href=\""events/calendar.asp?date="& Month(dDate) & "-" & iDay & "-" & Year(dDate) & "\""'>" & iDay & "</span>"";" & vbCrLf


 sScript = sScript & "document.all.day_" & iDay & ".style.backgroundColor = '#0099ff';" &  vbCrLf
        oRst.MoveNext
      Loop
      sScript = sScript & "</script>"
    End If
    Set oRst = Nothing

    Response.Write sScript
    %>
  </table>
  </form>
  </div>

