<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="rentalcommonfunctions.asp" //-->
<html>
<head>
  	<meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

	<%If iorgid = 7 Then %>
		<title><%=sOrgName%></title>
	<%Else%>
		<title>E-Gov Services <%=sOrgName%></title>
	<%End If%>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script language="Javascript" src="../scripts/modules.js"></script>
	<script language="Javascript" src="../scripts/easyform.js"></script>

	<style>
		body>table {width: 100%;}
		ul {list-style-type: none;}
		.calview 
		{
			width:50%;
			margin: 0 auto;
		}
		.iframeformat .calview
		{
			width:100%;
		}
		
		/* Month header */
		.month {
  			padding: 70px 0;
  			width: 100%;
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
  		padding-left:45px;
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
			font-size:30px !important;
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
		.indent20
		{
			padding:0 20px 0 20px;
		}
		.footerbox.month
		{
			display:block !important;
		}


		/* The Modal (background) */
		.modal {
  			display: none; /* Hidden by default */
  			position: fixed; /* Stay in place */
  			z-index: 1; /* Sit on top */
  			left: 0;
  			top: 0;
  			width: 100%; /* Full width */
  			height: 100%; /* Full height */
  			overflow: auto; /* Enable scroll if needed */
  			background-color: rgb(0,0,0); /* Fallback color */
  			background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
		}

		/* Modal Content/Box */
		.modal-content {
  			background-color: #fefefe;
  			margin: 15% auto; /* 15% from the top and centered */
  			padding: 20px;
  			border: 1px solid #888;
  			width: 352px; /* Could be more or less, depending on screen size */
		}
		.iframeformat .modal-content
		{
			margin: 50% auto !important;
		}
		
		/* The Close Button */
		.close {
  			color: #aaa;
  			float: right;
  			font-size: 28px;
  			font-weight: bold;
		}
		
		.close:hover,
		.close:focus {
  			color: black;
  			text-decoration: none;
  			cursor: pointer;
		}
		table.schedule td
		{
			vertical-align:top;
		}
		table.schedule td.starttime,
		table.schedule td.endtime
		{
			width:15%;
			white-space: nowrap;
		}
		table.schedule td.dash
		{
			width:5%
		}
		table.calendar td:hover
		{
  			cursor: pointer;
		}

		.notavailable
		{
			color:white;
			background-color: #666;
		}
		.available
		{
			background-color: lightgreen;
		}
		.partiallyavailable
		{
			background-color: orange;
		}
	</style>

</head>

<!--#Include file="../include_top.asp"-->

<%
iCitizenUserId = request.cookies("userid")
strResidentType = GetUserResidentType( iCitizenUserId )
'If they are not one of these (R, N), we have to figure which they are
If strResidentType <> "R" And strResidentType <> "N" Then 
	'This leaves E and B - See if they are a resident, also
	strResidentType = GetResidentTypeByAddress( iCitizenUserId, iOrgId )
End If 



intRentalID = request.querystring("rid")
if intRentalID = "" or not isnumeric(intRentalID) then response.redirect "rentalcategories.asp"



'What Month and year is it?
dtDate = request.querystring("date")
if dtDate = "" then
	intMonth = Month(date)
	intYear = Year(date)
	dtDate = date
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

'Get Rental Information
sSQL = "SELECT r.rentalid, l.name as locationname, r.rentalname, ISNULL(U.businessnumber,'') AS businessnumber, IsNull(residentrentalperiod, 12) as residentrentalperiod, IsNull(nonresidentrentalperiod, 12) as nonresidentrentalperiod,publiccanreserve " _
	& " FROM egov_rentals r " _
	& " INNER JOIN egov_class_location l ON l.locationid = r.locationid " _
	& " LEFT JOIN Users u ON u.userid = r.supervisoruserid " _
	& " WHERE r.isdeactivated = 0 and r.orgid = " & iorgid & " and r.rentalid = '" & intRentalID & "' "
Set oRs = Server.CreateObject("ADODB.RecordSet")
oRs.Open sSQL, Application("DSN"), 3, 1
if oRs.EOF then response.redirect "rentalcategories.asp"

strRentalName = oRs("locationname") & ": " & oRs("rentalname")
strPhoneNumber = trim(oRs("businessnumber"))
if strPhoneNumber = "" or isnull(strPhoneNumber) then 
	strPhoneNumber = "by phone"
else
	strPhoneNumber = "at " & strPhoneNumber
end if

If strResidentType = "R" Then
	intMonthsOut = oRs("residentrentalperiod")
else
	intMonthsOut = oRs("nonresidentrentalperiod")
end if

blnPublicCanReserve = oRs("publiccanreserve")

oRs.Close
Set oRs = Nothing


'Get Bookings for this month
sSQL = "SELECT D.reservationstarttime as resstart, D.billingendtime as resend " _
	& " FROM egov_rentalreservationdates D " _
	& " INNER JOIN egov_rentalreservations rr ON rr.reservationid = D.reservationid " _
	& " INNER JOIN egov_rentalreservationstatuses DS ON D.statusid = DS.reservationstatusid " _
	& " INNER JOIN egov_rentalreservationstatuses RS ON rr.reservationstatusid = RS.reservationstatusid " _
	& " WHERE  d.orgid = " & iorgid & " " _
	& " AND D.rentalid = " & intRentalID & " " _
	& " AND D.reservationstarttime BETWEEN '" & dtFirstOfMonth & "' AND '" & dtNextMonth & "'  " _
	& " AND DS.iscancelled = 0  " _
	& " AND RS.iscancelled = 0  " _
	& " AND  rr.isonhold = 0  " _
	& " ORDER BY reservationstarttime "
Set oRes = Server.CreateObject("ADODB.RecordSet")
oRes.Open sSQL, Application("DSN"), 3, 1



sSQL  = "SELECT C.date,  " _
	& " DateAdd(n,openingminute,DATEADD(hh,openinghour + CASE WHEN openingampm = 'PM' and openinghour <> 12 THEN 12 ELSE 0 END,c.date)) as opens,  " _
	& " DateAdd(n,closingminute,DATEADD(hh,closinghour + CASE WHEN closingampm = 'PM' and closinghour <> 12 THEN 12 ELSE 0 END,c.date)) as closes,  " _
	& " rd.isopen, rd.isavailabletopublic, ISNULL(rd.postbuffer,0) as postbuffer, pbt.ishours as postbufferishours, pbt.isminutes as postbufferisminutes, " _
	& " rd.minimumrental, mt.isminutes as minisminutes, mt.ishours as minishours, mt.isallday as minallday, " _
	& " DateAdd(n,lateststartminute,DATEADD(hh,lateststarthour + CASE lateststartampm WHEN 'PM' THEN 12 ELSE 0 END,c.date)) as lateststart " _
	& " FROM CalendarDays c " _
	& " LEFT JOIN egov_rentaldays rd ON DATEPART(dw,c.date) = rd.dayofweek  " _
	& " AND rd.isoffseason = ISNULL((SELECT ISNULL(hasoffseason,0) as hasoffseason FROM egov_rentals WHERE orgid = rd.orgid AND rentalid = rd.rentalid AND c.date BETWEEN  " _
	& " CONVERT(varchar(2),ISNULL(offseasonstartmonth, 1)) + '/' + CONVERT(varchar(2),ISNULL(offseasonstartday,1)) + '/' + CASE WHEN ISNULL(offseasonendmonth,1) < ISNULL(offseasonstartmonth,1) and Month(c.date) < ISNULL(offseasonendmonth,1) THEN CONVERT(varchar(4),YEAR(c.date)-1) ELSE CONVERT(varchar(4),YEAR(c.date)) END  " _
	& " AND CONVERT(varchar(2),ISNULL(offseasonendmonth, 1)) + '/' + CONVERT(varchar(2),ISNULL(offseasonendday,1)) + '/' + CASE WHEN ISNULL(offseasonendmonth,1) < ISNULL(offseasonstartmonth,1) and Month(c.date) < ISNULL(offseasonendmonth,1) THEN CONVERT(varchar(4),YEAR(c.date)) ELSE CONVERT(varchar(4),YEAR(c.date)+ offseasonendyear) END) ,0) " _
	& " LEFT JOIN egov_rentaltimetypes pbt ON pbt.timetypeid = rd.postbuffertimetypeid " _
	& " LEFT JOIN egov_rentaltimetypes mt ON mt.timetypeid = rd.minimumrentaltimetypeid " _
	& " WHERE c.date between '" & dtFirstOfMonth & "' AND '" & DateAdd("d",-1, dtNextMonth) & "' " _
	& " AND rd.orgid = " & iorgid & " AND rd.rentalid = " & intRentalID & " " _
	& " ORDER BY c.date "
'response.write sSQL
Set oDays = Server.CreateObject("ADODB.RecordSet")
oDays.Open sSQL, Application("DSN"), 3, 1
Dim arrDays(31,2)
Dim arrBlocks()
ReDim arrBlocks(oDays.RecordCount,0)
intBlocks = 0
intArrayMax = 0
Do While Not oDays.EOF
	'Build array for each day showing blocks of available and unavailable
	strAvail = "Available"
	intLoopBlocks = 0

	dtResEnd = oDays("opens")
	blnFoundHours = false
	do while not oRes.EOF
		if DateDiff("n", oDays("date"), oRes("ResStart")) > 0 AND DateDiff("n", oRes("ResStart"), DateAdd("d",1,oDays("date"))) > 0 then
			blnFoundHours = true
			'if oDays("date") = "5/1/2020" then response.write "<h1>HERE</h1>"
			if DateDiff("n",dtResEnd, oRes("ResStart")) > 0 then
				'Add Open Block TO Array
				'response.write oDays("date") & " - " & oDays("opens") & " to " & oRes("ResStart") & " OPEN 1<br />" & vbcrlf
				intLoopBlocks = intLoopBlocks + 1

				strOpen = "OPEN"
				if not oDays("isavailabletopublic") or not blnPublicCanReserve then strOpen = "CALL"
				if DateDiff("d",oDays("Date"),Date) > 0 then strOpen = "PAST"
				if not oDays("isopen") then strOpen = "CLOSED"
				if DateDiff("d",DateAdd("m",intMonthsOut,Date), oDays("Date")) > 0 then
					strOpen = "NONRESLIMIT"
					if strResidentType = "R" then strOpen = "RESLIMIT"
				end if
				strDiffType = "h"
				if oDays("minisminutes") then strDiffType = "n"
				if DateDiff(strDiffType,dtResEnd, oRes("ResStart")) < oDays("minimumrental") then strOpen = "UNAVAILABLE"
		
				AddToArray Day(oDays("Date")), intLoopBlocks, FormatTime(dtResEnd) & "|" & FormatTime(oRes("ResStart")) & "|" & strOpen

				if strOpen <> "UNAVAILABLE" then strAvail = "Partially Available"
			end if

			dtResEnd = oRes("ResEnd")
			if oDays("postbufferishours") then dtResEnd = DateAdd("h",oDays("postbuffer"), dtResEnd)
			if oDays("postbufferisminutes") then dtResEnd = DateAdd("n",oDays("postbuffer"), dtResEnd)

			'Add Reservation To Array
			'response.write oDays("date") & " - " & oRes("ResStart") & " to " & dtResEnd & " RESERVED<br />" & vbcrlf
			intLoopBlocks = intLoopBlocks + 1
			AddToArray Day(oDays("Date")), intLoopBlocks, FormatTime(oRes("ResStart")) & "|" & FormatTime(dtResEnd) & "|RESERVED"

			if strAvail = "Available" then strAvail = "Not Available"
		end if
		oRes.MoveNext
	loop
	'response.write "<h1>" & oDays("date") & "-" & blnFoundHours & "</h1>"
	if oRes.RecordCount > 0 then oRes.MoveFirst

	dtCloses = oDays("closes")
	if DateDiff("n",oDays("opens"),dtCloses) < 0 then dtCloses = DateAdd("d",1,dtCloses)

	if DateDiff("n",dtResEnd,dtCloses) > 0 Then 
		blnFoundHours = true
		'Add Open Block to Array
		'Response.Write oDays("date") & " - " & dtResEnd & " to " & oDays("closes") & " OPEN 2<br />" & vbcrlf
		intLoopBlocks = intLoopBlocks + 1
		strOpen = "OPEN"
		if not oDays("isavailabletopublic") or not blnPublicCanReserve then strOpen = "CALL"
		if DateDiff("d",oDays("Date"),Date) > 0 then strOpen = "PAST"
		if not oDays("isopen") or DateDiff("m",dtResEnd,oDays("lateststart")) < 0 then strOpen = "CLOSED"
		if DateDiff("d",DateAdd("m",intMonthsOut,Date), oDays("Date")) > 0 then
			strOpen = "NONRESLIMIT"
			if strResidentType = "R" then strOpen = "RESLIMIT"
		end if
		strDiffType = "h"
		if oDays("minisminutes") then strDiffType = "n"
		if DateDiff(strDiffType,dtResEnd, dtCloses) < oDays("minimumrental") then strOpen = "UNAVAILABLE"
		AddToArray Day(oDays("Date")), intLoopBlocks, FormatTime(dtResEnd) & "|" & FormatTime(dtCloses) & "|" & strOpen

		if strAvail <> "Available" and strOpen <> "UNAVAILABLE" then strAvail = "Partially Available"
	end if

	if DateDiff("d",oDays("date"),date) > 0 or DateDiff("d",DateAdd("m",intMonthsOut,Date), oDays("Date")) > 0  or (not oDays("isopen") and strAvail <> "Partially Available") then strAvail = "Not Available"

	if not blnFoundHours then 
		intLoopBlocks = intLoopBlocks + 1
		AddToArray Day(oDays("Date")), intLoopBlocks, "12:00 AM|12:00 AM|CLOSED"
		'Response.Write oDays("date") & " - " & dtResEnd & " to " & oDays("closes") & " OPEN 2<br />" & vbcrlf
	end if

	'Build Array Of Date Attributes
	arrDays(day(oDays("Date")),0) = strAvail
	arrDays(day(oDays("Date")),1) = oDays("isOpen")
	arrDays(day(oDays("Date")),2) = oDays("isavailabletopublic")
	
	
	oDays.MoveNext
loop

oDays.Close
oRes.Close
Set oDays = Nothing
Set oRes = Nothing

response.write "<script>"
response.write "var arrBlocks = [];" & vbcrlf
for x = 1 to UBOUND(arrBlocks,1)
	response.write "arrBlocks.push(["
	for y = 1 to UBOUND(arrBlocks,2)
		'response.write x & "," & y & " " & arrBlocks(x,y) & "<br />"
		'response.write "arrBlocks.push([" & x & ", '" & arrBlocks(x,y) & "']);"
		if y > 1 then response.write ","
		response.write "'" & arrBlocks(x,y) & "'"
	next
	response.write "]);" & vbcrlf
next
response.write "</script>"


%>
<div class="calview">
<%


sSQL = "SELECT r.rentalid, l.name as locationname, r.rentalname, ISNULL(U.businessnumber,'') AS businessnumber, residentrentalperiod, nonresidentrentalperiod " _
	& " FROM egov_rentals r " _
	& " INNER JOIN egov_class_location l ON l.locationid = r.locationid " _
	& " LEFT JOIN Users u ON u.userid = r.supervisoruserid " _
	& " WHERE r.isdeactivated = 0 and r.publiccanview = 1 and r.orgid = " & iorgid & " " _
	& " ORDER BY l.name, r.rentalname "
Set oRs = Server.CreateObject("ADODB.RecordSet")
oRs.Open sSQL, Application("DSN"), 3, 1
intCategoryID = 0
If not oRs.EOF then
	response.write "<form name=""rentalSel"">"
	response.write "<input type=""hidden"" name=""date"" value=""" & dtDate & """>"
	response.write "Rental Location: <select name=""rid"" onChange=""document.rentalSel.submit()""><option>Select...</option>"
	strCat = ""
	Do While Not oRs.EOF
		if strCat <> oRs("locationname") then
			if strCat <> "" then response.write "</optgroup>" & vbcrlf
			strCat = oRs("locationname")
			response.write "<optgroup label=""" & strCat & """>" & vbcrlf
		end if
		strSelected = ""
		if cint(intRentalID) = cint(oRs("rentalid")) then 
			strSelected = " selected"

		end if
		response.write "<option value=""" & oRs("rentalid") & """ " & strSelected & ">" & oRs("rentalname") & "</option>" & vbcrlf
		oRs.MoveNext
	loop
	response.write "</optgroup></select>"
	response.write "</form> <br />"
end if
oRs.Close
Set oRs = Nothing


response.write "<form name=""monthSel"">"
response.write "<input type=""hidden"" name=""rid"" value=""" & intRentalID & """>"
response.write "Jump To: <select name=""date"" onChange=""document.monthSel.submit()"">"
for intLoopMonth = 0 to intMonthsOut
	dtLoopMonth = DateAdd("m",intLoopMonth,Date)
	strSelected = ""
	if Month(dtDate) = Month(dtLoopMonth) and Year(dtDate) = Year(dtLoopMonth) then strSelected = " selected"
	response.write "<option value=""" & dtLoopMonth & """ " & strSelected & ">" & MonthName(month(dtLoopMonth)) & " " & year(dtLoopMonth) & "</option>" & vbcrlf
next
response.write "</select>"
response.write "</form> <br />"
%>


<div class="month footerbox">
  <ul>
    <li class="prev"><a href="?rid=<%=intRentalID%>&date=<%=dtPrevMonth%>">&#10094;</a></li>
    <li class="next"><a href="?rid=<%=intRentalID%>&date=<%=dtNextMonth%>">&#10095;</a></li>
    <li><%=strRentalName%><br /><%=MonthName(intMonth)%><br /><span style="font-size:18px"><%=intYear%></span></li>
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
		response.write "<td onClick=""showDay(" & intX-1 & ");"" class=" & replace(lcase(arrDays(intX,0))," ","") & "><div class=""content"">" & strActiveOpen & intX & strActiveClose
		response.write " <span class=""calDesc"">" & arrDays(intX,0) & "</span>"
		'if intX < 8 then response.write " Partially Available"
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
</div>
<!--SPACING CODE-->
<p><bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->
<!-- The Modal -->
<div id="myModal" class="modal">

  <!-- Modal content -->
  <div class="modal-content">
    <span class="close">&times;</span>
    <h3>Schedule for <%=MonthName(intMonth)%> <span id="modalday"></span></h3>
    <p id="modalcontent">Some text in the Modal..</p>
  </div>

</div>
<script>
	// Get the modal
	var modal = document.getElementById("myModal");

	// Get the button that opens the modal
	var btn = document.getElementById("myBtn");

	// Get the <span> element that closes the modal
	var span = document.getElementsByClassName("close")[0];
	
	// When the user clicks on the button, open the modal
	function showDay(x) {
  		modal.style.display = "block";

		//Now we need to populate the modal
		var strSchedule = '<table class="schedule">';
		for(y = 0; y < arrBlocks[x].length; y++) { 
			var arrSchedule = (arrBlocks[x][y]).split("|");
			if (arrBlocks[x][y] != "")
			{
				strSchedule += "<tr><td class=\"starttime\">" + arrSchedule[0] + "</td><td class=\"dash\">-</td><td class=\"endtime\">" + arrSchedule[1] + "</td><td>" + blockStatus(arrSchedule[2],x+1) + "</td></tr>";
			}
		}
		document.getElementById("modalcontent").innerHTML = strSchedule + "</table>";
		document.getElementById("modalday").innerHTML = x+1;

	}


	// When the user clicks on <span> (x), close the modal
	span.onclick = function() {
  		modal.style.display = "none";
	}

	function blockStatus(x, day)
	{
		switch(x) {
			case "CALL": return "Contact us <%=strPhoneNumber%> to inquire about reservations on this date."; break;
			case "CLOSED": return "Sorry, it's Closed"; break;
			case "PAST": return "Past blocks cannot be reserved"; break;
			case "OPEN": return "<a href=\"javascript:selectDate('<%=month(dtFirstOfMonth)%>/" + day + "/<%=year(dtFirstOfMonth)%>');\">Click to Reserve</a>"; break;
			case "NONRESLIMIT": return "This location limits non-residents to reservations that are <%=intMonthsOut%> months out."; break;
			case "RESLIMIT": return "This location limits residents to reservations that are <%=intMonthsOut%> months out."; break;
			default: return x;
		}
	}

	function selectDate(x)
	{
		document.selectForm.selecteddate.value = x;
		document.selectForm.submit();
	}

	// When the user clicks anywhere outside of the modal, close it
	window.onclick = function(event) {
  		if (event.target == modal) {
    			modal.style.display = "none";
  		}
	}
</script>
<form name="selectForm" method="post" action="rentalcontrol.asp">
	<input type="hidden" id="cid" name="cid" value="0" />
	<input type="hidden" id="rid" name="rid" value="<%=intRentalID%>" />
	<input type="hidden" id="src" name="src" value="dp" />
	<input type="hidden" id="viewtype" name="viewtype" value="viewsingledate" />
	<input type="hidden" id="selecteddate" name="selecteddate" value="" />
	<input type="hidden" id="selectedrid" name="selectedrid" value="<%=intRentalID%>" />
	<input type="hidden" id="selectdate" name="selectdate" value="" />
	<input type="hidden" id="startdate" name="startdate" value="" />
	<input type="hidden" id="enddate" name="enddate" value="" />
	<input type="hidden" id="wanteddows" name="wanteddows" value="" />
</form>
<!--#Include file="../include_bottom.asp"-->  

<%
Sub AddToArray(intDay, intIndex, strValue)
	if intArrayMax < intIndex then
		intArrayMax = intIndex
		ReDim Preserve arrBlocks(UBOUND(arrBlocks,1), intArrayMax)
	end if
	
	'response.write intDay & " " & intIndex & "<br />"
	'response.flush
	arrBlocks(intDay, intIndex) = strValue 

End Sub

Function FormatTime(dtDateTime)
	strAMPM = " AM"
	strHour = Hour(dtDateTime)
	if strHour > 11 then strAMPM = " PM"
	if strHour > 12 then strHour = strHour - 12
	FormatTime = strHour & ":" & Right("00" & Minute(dtDateTime),2) & strAMPM
End Function
%>
