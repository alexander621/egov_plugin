<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<%
' GET ORG ID FROM QUERYSTRING
iorgid = request("orgid")
If iorgid = "" Then
	iorgid = 0
End If

' GET DATE RANGE
Select Case request("date")
Case "1"
	' Today
	datBeginDate = DateAdd("d",-1,Date()) 
	datEndDate = Date()
	sRange = "Last 24 Hours"
Case "2"
	' LAST 2 WEEKS
	datBeginDate = DateAdd("d",-14,Date()) 
	datEndDate = Date()
	sRange = "Last 2 Weeks"
Case "3"
	' LAST THIRTY DAYS
	datBeginDate = DateAdd("d",-30,Date()) 
	datEndDate = Date()
	sRange = "Last 30 Days"
Case "4"
	' LAST 90 DAYS
	datBeginDate = DateAdd("d",-90,Date()) 
	datEndDate = Date()
	sRange = "Last 90 Days"
Case "5"
	' LAST YEAR
	datBeginDate = DateAdd("yyyy",-1,Date()) 
	datEndDate = Date()
	sRange = "Last Year"
Case Else
	' The past month's - Changed 10/16/2009, Steve Loar
	datEndDate = DateAdd("d", -1, CDate(Month(Date()) & "/1/" & Year(Date())))
	datBeginDate = DateAdd("m", -1, CDate(Month(Date()) & "/1/" & Year(Date())))
	sRange = "Last Month"
End Select 

' ADD TIME TO RANGES
datBeginDate = datBeginDate & " 12:00:00 AM"
datEndDate = datEndDate & " 11:59:59 PM"

%>


<html>
<head>
	<title>E-GovLink Web Statistics</title>

	<link href="../global.css" rel="stylesheet" type="text/css" />
</head>
<body>

	<!-- BEGIN: DISPLAYING STATS-->
	<div style="padding:20 px;">

		<h3 class="webstattitle"><%=fnGetName(iorgid)%> E-Govlink Web Statistics: <font class="sectiontitle"><%=sRange%> (<%=datBeginDate & " to " & datEndDate%>)</font><br />
		<font style="size:10px;"><a class="webstat" href="webstats.asp?orgid=<%=iorgid%>&date=1">Last 24 Hours</a> | <a class="webstat" href="webstats.asp?orgid=<%=iorgid%>&date=2">Last 2 Weeks</a> | <a class="webstat" href="webstats.asp?orgid=<%=iorgid%>&date=3">Last 30 Days</a> | <a class="webstat" href="webstats.asp?orgid=<%=iorgid%>&date=4">Last 90 Days</a> | <a class="webstat" href="webstats.asp?orgid=<%=iorgid%>&date=5">Last Year</a></font>
		<hr align="left" style="color:#000000;size:1 px; width:75%;" ></h3>

		<h3 class="webstattitle">E-Govlink Web Site Total Statistics<hr align="left" class="webstat"></h3>
		<p>
			<%

			' TOTAL STATS
			subDisplayTotalStats iorgid, "", datBeginDate, datEndDate

			%>
		</p>

		<h3 class="webstattitle">E-Govlink Web Site Section Detail Statistics<hr align="left" class="webstat"></h3>
		<p>
			<%
			' ACTION LINE
			subDisplayStats iorgid, 22, 2, "Action Request Form Hit Statistics", datBeginDate, datEndDate
			%>
		</p>

		<p>
			<%
			' PAYMENTS
			subDisplayStats iorgid, 33, 3, "Payment Form Hit Statistics", datBeginDate, datEndDate
			%>
		</p>

		<p>
			<%
			' CALENDAR
			subDisplayStats iorgid, 55, 5, "Calendar Hit Statistics", datBeginDate, datEndDate
			%>
		</p>

		<p>
			<%
			' DOCUMENTS
			subDisplayStats iorgid, 44, 4, "Document Hit Statistics", datBeginDate, datEndDate
			%>
		</p>

	</div>
	<!-- END: DISPLAYING STATS-->


	<!-- BEGIN: DISPLAY FOOTER-->
	<center>
		<hr  style="color:#000000;size:1 px; width:90%;" >
		<div class="copyright_text">Copyright &copy;2004-<%=Year(Date())%>. <i>electronic commerce</i> link, inc. dba <a href="http://www.egovlink.com" target="_NEW">egovlink</a>.</div>
	</center>
	<!-- END: DISPLAY FOOTER-->

</body>
</html>

<%
'------------------------------------------------------------------------------------------------------------
' FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' void SUBDISPLAYSTATS IORGID, ISECTIONID, ISECTIONHOMEID, SSECTIONTITLE, DATSTARTDATE, DATENDDATE
'------------------------------------------------------------------------------------------------------------
Sub subDisplayStats( ByVal iorgid, ByVal isectionid, ByVal isectionhomeid, ByVal sSectionTitle, ByVal datStartDate, ByVal datEndDate )
	Dim oRs, sSql

	' GET HOME HITS
	sSql = "SELECT ISNULL(COUNT(*),0) AS TotalHomeHits FROM visitor_log WHERE  ((orgid = " & iorgid & ") AND (sectionid = " & isectionhomeid & ")) AND ((visitdate >='" & datStartDate & "' AND visitdate <= '" & datEndDate & "'))"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iHomeHits = oRs("TotalHomeHits")
	Else
		iHomeHits = 0
	End If

	oRs.Close
	Set oRs = Nothing

	' GET DETAIL HITS FOR TOP 10
	sSql = "SELECT TOP 10 COUNT(documenttitle) AS NumHits, documenttitle,(SELECT ISNULL(COUNT(*),0) FROM visitor_log "
	sSql = sSql & " WHERE (orgid = " & iorgid & ") AND (sectionid = " & isectionid & ")  AND ((visitdate >='" & datStartDate & "' "
	sSql = sSql & " AND visitdate <= '" & datEndDate & "'))) AS TotalHits FROM visitor_log WHERE (orgid = " & iorgid & ") "
	sSql = sSql & " AND (sectionid = " & isectionid & ") AND (visitdate >='" & datStartDate & "' AND visitdate <= '" & datEndDate
	sSql = sSql & "') GROUP BY documenttitle ORDER BY NumHits DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		
		' DETAIL HITS
		response.write "<font class=""sectiontitle"">" & sSectionTitle & "</font>"
		response.write "<div class=""webstat"">"
		response.write "<table cellspacing=""0"" cellpadding=""2"" width=""100%"">"
		response.write "<tr><td class=""statlabel"">Total Number of Hits: </td><td> " & CLng(oRs("TotalHits")) + CLng(iHomeHits) & " </td></tr>"
		response.write "<tr><td class=""statlabel"">Home Page Number of Hits: </td><td> " & CLng(iHomeHits) & " </td></tr>"
		response.write "<tr><td class=""statlabel"" colspan=""2""><br /> Top 10 Individual Page Hits  </td></tr>"

		Do while Not oRs.EOF 
			response.write "<tr><td class=""statdetail"" align=""left"">" & oRs("documenttitle") & "</td><td>" & oRs("NumHits")  & "</td></tr>"
			oRs.MoveNext
		Loop

		response.write "</table>"
		response.write "</div>"

'		oRs.Close

	Else
		
		' NO DETAIL HITS
		response.write "<font class=""sectiontitle"">" & sSectionTitle & "</font>"
		response.write "<div class=""webstat"">"
		response.write "<table cellspacing=""0"" cellpadding=""2"" width=""100%"">"
		response.write "<tr><td class=""statlabel"">Total Number of Hits: </td><td> " & CLng(iHomeHits) & " </td></tr>"
		response.write "<tr><td class=""statlabel"">Home Page Number of Hits: </td><td> " & CLng(iHomeHits) & " </td></tr>"
		response.write "<tr><td class=""statlabel"" colspan=""2"" ><br /> Top 10 Individual Page Hits  </td></tr>"
		response.write "<tr><td class=""statdetail"" align=""left"" colspan=""2"">No Hits</td></tr>"
		response.write "</table>"
		response.write "</div>"

	End If

	oRs.Close 
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------------------------------------
' void SUBDISPLAYTOTALSTATS IORGID, SSECTIONTITLE, DATSTARTDATE, DATENDDATE
'------------------------------------------------------------------------------------------------------------
Sub subDisplayTotalStats( ByVal iorgid, ByVal sSectionTitle, ByVal datStartDate, ByVal datEndDate )
	Dim sSql, oRs

	' GET HOME HITS
	sSql = "SELECT ISNULL(COUNT(*),0) AS TotalHomeHits FROM visitor_log WHERE  (orgid = " & iorgid & ") AND (visitdate >='" & datStartDate & "' AND visitdate <= '" & datEndDate & "')"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	

	If NOT oRs.EOF Then
		iHomeHits = oRs("TotalHomeHits")
	Else
		iHomeHits = 0
	End If

	oRs.Close
	Set oRs = Nothing

	' GET DETAIL HITS FOR TOP 10
	sSql = "SELECT TOP 10 COUNT(documenttitle) AS NumHits, documenttitle = Case sectionid WHEN '1' THEN 'E-GovLink Home Page' WHEN '2' THEN 'Action Request Home Page'  WHEN '3' THEN 'Payment Home Page'  WHEN '5' THEN 'Calendar Home Page'  WHEN '4' THEN 'Document Home Page' ELSE documenttitle END,(SELECT ISNULL(COUNT(*),0) FROM visitor_log WHERE  (orgid = " & iorgid & ")  AND ((visitdate >='" & datStartDate & "' AND visitdate <= '" & datEndDate & "'))) AS TotalHits FROM visitor_log WHERE (orgid = " & iorgid & ") AND ((visitdate >='" & datStartDate & "' AND visitdate <= '" & datEndDate & "')) GROUP BY documenttitle,sectionid ORDER BY NumHits DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If NOT oRs.EOF Then
		
		' DETAIL HITS
		response.write "<font class=""sectiontitle"">" & sSectionTitle & "</font>"
		response.write "<div class=""webstat"">"
		response.write "<table cellspacing=""0"" cellpadding=""2"" width=""100%"">"
		response.write "<tr><td class=""statlabel"">Total Number of Hits: </td><td> " & CLng(iHomeHits) & " </td></tr>"
		'response.write "<tr><td class=statlabel >Main Page Number of Hits: </td><td> " & CLng(iHomeHits) & " </td></tr>"
		response.write "<tr><td class=""statlabel"" colspan=""2"">  <br> Top 10 Individual Page Hits  </td></tr>"

		Do while NOT oRs.EOF 
			response.write "<tr><td class=""statdetail"" align=""left"">" & oRs("documenttitle") & "</td><td>" & oRs("NumHits")  & "</td></tr>"
			oRs.MoveNext
		Loop

		response.write "</table>"
		response.write "</div>"

	End If

	oRs.Close
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------------------------------------
' string FNGETNAME( IORGID )
'------------------------------------------------------------------------------------------------------------
Function fnGetName( ByVal iorgid )
	Dim sSql, oRs
	
	sReturnValue = "UNKNOWN"

	' GET ORG NAME
	sSql = "SELECT orgname FROM Organizations WHERE orgid = " & iorgid
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If NOT oRs.EOF Then
		sReturnValue = oRs("orgname")
	End If

	fnGetName = sReturnValue

	oRs.Close
	Set oRs = Nothing 

End Function


%>
