<style>
ul {list-style-type: none;}

/* Month header */
.month {
  padding: 70px 25px;
  width: 100%;
  background: #1abc9c;
  text-align: center;
}

/* Month list */
.month ul {
  margin: 0;
  padding: 0;
}

.month ul li {
  color: white;
  font-size: 20px;
  text-transform: uppercase;
  letter-spacing: 3px;
}

/* Previous button inside month header */
.month .prev {
  float: left;
  padding-top: 10px;
}

/* Next button */
.month .next {
  float: right;
  padding-top: 10px;
  padding-right:45px;
}

.month .prev a,
.month .next a
{
	color: white;
	text-decoration:none;
}




.calendar td
{
	border: 1px solid black;
  position: relative;
  width:14%;
}
.calendar td:after{
    content:'';
    display:block;
    margin-top:100%;
}
td .content {
    position:absolute;
    top:0;
    bottom:0;
    left:0;
    right:0;
    padding:10px;
}
</style>


<%
'What Month and year is it?
dtDate = request.querystring("date")
if dtDate = "" then
	intMonth = Month(date)
	intYear = Year(date)
else
	intMonth = Month(dtDate)
	intYear = Year(dtDate)
end if

dtFirstOfMonth = intMonth & "/1/" & intYear

dtPrevMonth = DateAdd("m",-1, dtFirstOfMonth)
dtNextMonth = DateAdd("m",1, dtFirstOfMonth)

'How many days in month?
intMaxDays = Day(DateAdd("d",-1,DateAdd("m",1,dtFirstOfMonth)))

'How many blank days?
intBlankDays = WeekDay(dtFirstOfMonth) - 1

'Get Bookings for this month
response.write session("orgid")



%>

<div class="month">
  <ul>
    <li class="prev"><a href="?date=<%=dtPrevMonth%>">&#10094;</a></li>
    <li class="next"><a href="?date=<%=dtNextMonth%>">&#10095;</a></li>
    <li><%=MonthName(intMonth)%><br><span style="font-size:18px"><%=intYear%></span></li>
  </ul>
</div>

<table class="calendar" style="width:100%;border-collapse: collapse;">
	<tr>
		<th>Su</th>
		<th>Mo</th>
		<th>Tu</th>
		<th>We</th>
		<th>Th</th>
		<th>Fr</th>
		<th>Sa</th>
	</tr>
	<tr>
<%
	For intX = 1 to intBlankDays
		response.write "<td>&nbsp;</td>"
	next

	intRowDays = intBlankDays + 1

	for intX = 1 to intMaxDays
		strActiveOpen = ""
		strActiveClose = ""
		if date = CDate(intMonth & "/" & intX & "/" & intYear) then
			strActiveOpen = "<span class=""active"">"
			strActiveClose = "</span>"
		end if
		response.write "<td><div class=""content"">" & strActiveOpen & intX & strActiveClose
		if intX < 8 then response.write " Partially Available"
		response.write "</div></td>"

		intRowDays = intRowDays + 1
		if intRowDays = 8 then
			intRowDays = 1
			response.write "</tr>" & vbcrlf
			response.write "<tr>" & vbcrlf
		end if
	next
%>
	</tr>
</table>
