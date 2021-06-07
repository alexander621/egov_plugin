<!-- #include file="../includes/common.asp" //-->
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
If IsDate(Request("date")) Then
	dDate = CDate(Request("date"))
Else
	If IsDate(Request("month") & "/01/" & Request("year")) Then
		dDate = CDate(Request("month") & "/01/" & Request("year"))
	Else
		dDate = Date()
		' The annoyingly bad solution for those of you running IIS3
		'If Len(Request("month")) <> 0 Or Len(Request("day")) <> 0 Or Len(Request("year")) <> 0 Or Len(Request("date")) <> 0 Then
		'	Response.Write "The date you picked was not a valid date.  The calendar was set to today's date.<BR><BR>"
		'End If
	End If
End If

'Now we've got the date.  Now get Days in the choosen month and the day of the week it starts on.
iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
iDOW = GetWeekdayMonthStartsOn(dDate)
%>

<html>
<head>
  <title><%=langBSEventsCalendar%></title>

  <link rel="stylesheet" type="text/css" href="../classes/classes.css" />

  <style type="text/css">
  <!--
		body {scrollbar-base-color:#6699cc; scrollbar-highlight-color:#ffffff; scrollbar-arrow-color:#99ccff;}
		.cal {border-left:1px solid #93bee1; border-top:1px solid #93bee1;}
		.cal td {border-right:1px solid #93bee1; border-bottom:1px solid #93bee1; font-family:Tahoma,Arial; font-size:11px;}
		select {font-family:Arial,Tahoma,Verdana; font-size:13px;}
		td#today { background-color:#FEF76E; font-weight: bold; }

		td.calendardate:hover 
		{
			background-color: #898989 !important;
			color: #fff !important;
			cursor: pointer !important;
		}
  //-->
  </style>

	<script language="Javascript">
	<!--

		function doDateSelect( d ) 
		{
			//ClassForm
			//window.opener.document.<%=request("updateform")%>.<%=request("updatefield")%>.value = d;
			window.opener.document.getElementById( '<%=request("updatefield")%>' ).value = d;
			window.close();
		}

	//-->
	</script>

</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
  <form name="frmDate" action="calendarpicker.asp" method="post">
  <input type="hidden" name="day" value="1" />
  <input type="hidden" name="updatefield" value="<%=request("updatefield")%>" />
  <input type="hidden" name="updateform" value="<%=request("updateform")%>" />
  
  <table border="0" cellpadding="3" cellspacing="0" bgcolor="#ffffff" class="cal" width="100%" height="100%">
    <tr height="30">
      <td bgcolor="#336699" align="center" colspan="7">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15%" style="border:0px;" nowrap>&nbsp;<img src="../images/arrow_back.gif" align="absmiddle">&nbsp;<a href="calendarpicker.asp?date=<%= SubtractOneMonth(dDate) %>&updatefield=<%=request("updatefield")%>&updateform=<%=request("updateform")%>"><font face="Arial" color="#ffffff" size="1"><%=langPreviousMonth%></font></a></td>
            <td width="70%" align="center" style="border:0px;">
              <select name="month" onchange="document.frmDate.submit();">
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
              <select name="year" onchange="document.frmDate.submit();">
				<% For x = (Year(dDate) - 5) to (Year(dDate) + 5)
	      				response.write "<option value=""" & x & """>" & x & "</option>"
	      			next%>
              </select>
              <script language="Javascript">
                document.frmDate.month.selectedIndex = <%= Month(dDate)-1 %>;
                document.frmDate.year.value = <%= Year(dDate) %>;
              </script>
            </td>
            <td width="15%" align="right" style="border:0px;" nowrap><a href="calendarpicker.asp?date=<%= AddOneMonth(dDate) %>&updatefield=<%=request("updatefield")%>&updateform=<%=request("updateform")%>"><font face="Arial" color="#ffffff" size="1"><%=langNextMonth%></font></a>&nbsp;<img src="../images/arrow_forward.gif" align="absmiddle">&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr height="30">
		<td align="center" bgcolor="#93bee1" width="13%"><font color="#003366"><b><%=langDay1%></b></font></td>
		<td align="center" bgcolor="#93bee1" width="13%"><font color="#003366"><b><%=langDay2%></b></font></td>
		<td align="center" bgcolor="#93bee1" width="13%"><font color="#003366"><b><%=langDay3%></b></font></td>
		<td align="center" bgcolor="#93bee1" width="13%"><font color="#003366"><b><%=langDay4%></b></font></td>
		<td align="center" bgcolor="#93bee1" width="13%"><font color="#003366"><b><%=langDay5%></b></font></td>
		<td align="center" bgcolor="#93bee1" width="13%"><font color="#003366"><b><%=langDay6%></b></font></td>
		<td align="center" bgcolor="#93bee1" width="13%"><font color="#003366"><b><%=langDay7%></b></font></td>
    </tr>
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
      'Response.Write vbTab & vbTab & "<td id=""day_" & iCurrent & """ valign=""top""><font size=""2"">" & iCurrent & "</font></td>" & vbCrLf
      'Response.Write vbTab & vbTab & "<td><span class=""pickerday"" style='width:100%; height:100%; cursor:pointer;' onclick=""javascript:doDateSelect('" & Month(dDate) & "/" & iCurrent & "/" & Year(dDate) & "');""><font size=2>" & iCurrent & "</font></span></td>" & vbCrLf

	  Response.Write vbTab & vbTab & "<td"
	  If DateDiff("d", Date(), Month(dDate) & "/" & iCurrent & "/" & Year(dDate)) = 0 Then
		response.write " id=""today"" title=""today"""
	  End If 
	  response.write " class=""calendardate"" onclick=""doDateSelect('" & Month(dDate) & "/" & iCurrent & "/" & Year(dDate) & "')"">"
	  response.write "<span class=""pickerday"" style='width:100%; height:100%; cursor:pointer;' "
	  response.write "onclick=""doDateSelect('" & Month(dDate) & "/" & iCurrent & "/" & Year(dDate) & "')""><font size=""2"">" & iCurrent & "</font></span></td>" & vbCrLf


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
    %>
  </table>
  </form>
</body>
</html>