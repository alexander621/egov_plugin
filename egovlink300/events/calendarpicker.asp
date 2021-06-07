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
	elseIf IsDate(Request.QueryString("month") & "-1-" & Request.QueryString("year")) Then
		dDate = CDate(Request.QueryString("month") & "-1-" & Request.QueryString("year"))
	Else
		dDate = Date()
		' The annoyingly bad solution for those of you running IIS3
		If Len(Request.QueryString("month")) <> 0 Or Len(Request.QueryString("day")) <> 0 Or Len(Request.QueryString("year")) <> 0 Or Len(Request.QueryString("date")) <> 0 Then
			Response.Write "The date you picked was not a valid date.  The calendar was set to today's date.<BR><BR>"
		End If
	End If
End If

if month(date()) = month(dDate) and year(date()) = year(dDate) then
	dDate = date()
end if

'Now we've got the date.  Now get Days in the choosen month and the day of the week it starts on.
iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
iDOW = GetWeekdayMonthStartsOn(dDate)
%>

<html>
<head>
  <title>Choose Date</title>
  <style type="text/css">
  <!--
    body {scrollbar-base-color:#6699cc; scrollbar-highlight-color:#ffffff; scrollbar-arrow-color:#99ccff;}
    .cal {border-left:1px solid #93bee1; border-top:1px solid #93bee1;}
    .cal td {border-right:1px solid #93bee1; border-bottom:1px solid #93bee1; font-family:Tahoma,Arial; font-size:11px;}
    select {font-family:Arial,Tahoma,Verdana; font-size:13px;}
  //-->
  </style>
 
<SCRIPT TYPE="text/javascript">
<!--
window.focus();
//-->
</SCRIPT>

  
  
  <script language="Javascript">
  <!--
    function doDateSelect( d ) {
<%if request.querystring("n") > 1 then  'NOTE.. A BUNCH of lines (40?) of this Javascript are NOT sent to the browser, 
										'if n=1 (that is, if the user can only choose ONE date, rather than
										'a date range!)  Confusing kludge, sorry.
										%>

		var p;
		
		if (document.frmDate.range.value == "between") {
			p= document.frmDate.DateRange.value.indexOf("[end]");
			if (document.frmDate.DateRange.value.substring(0,1) == "["  || p==-1) {
				document.frmDate.DateRange.value = d + " and [end]"
			} else {
				document.frmDate.DateRange.value = document.frmDate.DateRange.value.substring(0,p)  + d
			}
		} else {
			document.frmDate.DateRange.value = d;
		}
	}
  //-->
  </script>
  <script language="Javascript">
  <!--
    function SendBackDate( d ) {

	  var s;
	  var p;
	  if (d.indexOf("[") != -1) {
		s=""
	  } else {
	    p= d.indexOf(" and ")
		if (p == -1) {
		  s=d
		} else {
		  s=d.substring(0,p) + "-" + d.substring(p+5)
		}
	  
		  if (document.frmDate.range.value == "since") {
			s=s + "-"
		  }
		  if (document.frmDate.range.value == "before") {
			s="-" + s
		  }
	  }
	  d=s;
<%end if								' END OF Java Script LINE SKIPPING... (note half of 2 functions are skipped, 
										' so that one doesnt even exist anymore by name!)
%>
	  window.opener.document.all.<%=request.querystring("r")%>.value = d;
      window.close();
    }
  //-->
  </script>
  
<script language="Javascript">
<!--
function SetDateRange() {
		if (document.frmDate.range.value == "between") {
			if (document.frmDate.DateRange.value.indexOf("[")!=-1) {
				document.frmDate.DateRange.value="[start] and [end]";
			} else {
				document.frmDate.DateRange.value=document.frmDate.DateRange.value + " and [end]";
			}
		} else {
			if (document.frmDate.DateRange.value.indexOf("[d")!=-1 || document.frmDate.DateRange.value.indexOf("[s")!=-1) {
				document.frmDate.DateRange.value="[date]"
			} else {
				if (document.frmDate.DateRange.value.indexOf(" ")!=-1) {
					document.frmDate.DateRange.value=document.frmDate.DateRange.value.substr(0,document.frmDate.DateRange.value.indexOf(" "))
				}
			}					
		}
    }
//-->
</script>

<script language="Javascript">
<!--
function prevMonth() {
	if (document.all.month.selectedIndex <1) {
		document.all.year.value=parseInt(document.all.year.value)-1;
		document.all.month.selectedIndex= 11;
	} else {
			document.all.month.selectedIndex=document.all.month.selectedIndex-1;
	}
	document.all.frmDate.submit();
    }
//-->
</script>


<script language="Javascript">
<!--
function nextMonth() {
	if (document.all.month.selectedIndex >10) {
		document.all.year.value=parseInt(document.all.year.value)+1;
		document.all.month.selectedIndex= 0;
	} else {
			document.all.month.selectedIndex=document.all.month.selectedIndex+1;
	}
	document.all.frmDate.submit();
    }
//-->
</script>

</head>

<body topmargin="0" leftmargin="0" bottommargin="0" rightmargin="0" marginwidth="0" marginheight="0">
  <form name="frmDate" action="calendarpicker.asp" method="get">
  <input type="hidden" name="day" value="<%= Day(Now) %>">
  
  <table border="0" cellpadding="3" cellspacing="0" bgcolor="#ffffff" class="cal" width="100%" height="100%">
    <tr height="30">
      <td bgcolor="#336699" align="center" colspan="7">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15%" style="border:0px;" nowrap>&nbsp;<A  href="javascript:void prevMonth();"><img src="../images/arrow_back.gif" align="absmiddle" border="0"></A>&nbsp;<A  href="javascript:void prevMonth();"><FONT face="Arial" COLOR=#ffffff SIZE="1"><u>Previous Month</u></FONT></A></td>
            <td width="70%" align="center" style="border:0px;">
              <select name="month" onchange="document.all.frmDate.submit();">
                <option value=1>January</option>
                <option value=2>February</option>
                <option value=3>March</option>
                <option value=4>April</option>
                <option value=5>May</option>
                <option value=6>June</option>
                <option value=7>July</option>
                <option value=8>August</option>
                <option value=9>September</option>
                <option value=10>October</option>
                <option value=11>November</option>
                <option value=12>December</option>
              </select>
              <select name="year" onchange="document.all.frmDate.submit();">
	      <% For x = (Year(dDate) - 5) to (Year(dDate) + 5)
	      		response.write "<option value=" & x & ">" & x
	      	 next%>
              </select>
              <script language="Javascript">
                document.all.month.selectedIndex = <%= Month(dDate)-1 %>;
                document.all.year.value = <%= Year(dDate) %>;
              </script>
            </td>
            <td width="15%" align="right" style="border:0px;" nowrap><A href="javascript:void nextMonth();"><FONT face="Arial" COLOR=#ffffff SIZE="1"><u>Next Month</u></FONT></A>&nbsp;<A href="javascript:void nextMonth();"><img src="../images/arrow_forward.gif" align="absmiddle" border="0"></A>&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr height="30">
      <td ALIGN="center" BGCOLOR=#93bee1 width="13%"><FONT COLOR=#003366><B>Sun</B></FONT></td>
      <td ALIGN="center" BGCOLOR=#93bee1 width="13%"><FONT COLOR=#003366><B>Mon</B></FONT></td>
      <td ALIGN="center" BGCOLOR=#93bee1 width="13%"><FONT COLOR=#003366><B>Tue</B></FONT></td>
      <td ALIGN="center" BGCOLOR=#93bee1 width="13%"><FONT COLOR=#003366><B>Wed</B></FONT></td>
      <td ALIGN="center" BGCOLOR=#93bee1 width="13%"><FONT COLOR=#003366><B>Thu</B></FONT></td>
      <td ALIGN="center" BGCOLOR=#93bee1 width="13%"><FONT COLOR=#003366><B>Fri</B></FONT></td>
      <td ALIGN="center" BGCOLOR=#93bee1 width="13%"><FONT COLOR=#003366><B>Sat</B></FONT></td>
    </TR>
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
      if  Month(dDate) = Month(Date()) and  iCurrent = day(Date()) and Year(dDate) = year(date()) then
      	Response.Write vbTab & vbTab & "<td bgcolor=yellow><span style='width:100%; height:100%; cursor:hand;' onclick=""javascript:doDateSelect('" & Month(dDate) & "/" & iCurrent & "/" & Year(dDate) & "');""><font size=2>" & iCurrent & "</font></span></td>" & vbCrLf
      else
      	Response.Write vbTab & vbTab & "<td><span style='width:100%; height:100%; cursor:hand;' onclick=""javascript:doDateSelect('" & Month(dDate) & "/" & iCurrent & "/" & Year(dDate) & "');""><font size=2>" & iCurrent & "</font></span></td>" & vbCrLf
      end if


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
	<%if request.querystring("n") > 1 then %>
	<tr > 
	  <td colspan=5>
		  <select name="range" onchange="SetDateRange();">
			  <option value=on <%if request.querystring("range")="on" then response.write "Selected"%>>On</option>
			  <option value=since <%if request.querystring("range")="since" then response.write "Selected"%>>Since</option>
			  <option value=before <%if request.querystring("range")="before" then response.write "Selected"%>>Before</option>
			  <option value=between <%if request.querystring("range")="between" then response.write "Selected"%>>Between</option>
		  </select>
		<input type=text name="DateRange" style="background-color:#ffffff; border:0px solid #ffffff; width:150px;" value='<%
		  if request.querystring("DateRange")<>"" then response.write request.querystring("DateRange") else response.write "[date]"		  
		  %>'>
	  </td>
	  <td colspan=2>
		Choose date range and press <a href="javascript:void SendBackDate(document.frmDate.DateRange.value);">Close</a>
	  </td>
	</tr>
	<%end if %>
  </table>
  <input type=hidden name="rand" value="<%
	  randomize
	  response.write rnd()
	%>">
  <input type=hidden name="r" value="<%=request.querystring("r")%>">
  <input type=hidden name="n" value="<%=request.querystring("n")%>">
  </form>
</body>
</html>
