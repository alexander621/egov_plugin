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
	' LAST THIRTY DAYS
	datBeginDate = DateAdd("d",-30,Date()) 
	datEndDate = Date()
	sRange = "Last 30 Days"
End Select 

' ADD TIME TO RANGES
datBeginDate = datBeginDate & " 12:00:00 AM"
datEndDate = datEndDate & " 11:59:59 PM"
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> E-GovLink Web Statistics  </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<link href="../global.css" rel="stylesheet" type="text/css">
</HEAD>



<BODY>


<!-- BEGIN: DISPLAYING STATS-->
<div style="padding:20 px;">

<h3 class=webstattitle ><%=fnGetName(iorgid)%> E-Govlink Web Statistics:  <font class=sectiontitle><%=sRange%> (<%=datBeginDate & " to " & datEndDate%>)</font><br>
<font style="size:10px;"><a class=webstat href="webstats.asp?orgid=<%=iorgid%>&date=1">Last 24 Hours</a> | <a class=webstat href="webstats.asp?orgid=<%=iorgid%>&date=2">Last 2 Weeks</a> | <a class=webstat href="webstats.asp?orgid=<%=iorgid%>&date=3">Last 30 Days</a> | <a class=webstat href="webstats.asp?orgid=<%=iorgid%>&date=4">Last 90 Days</a> | <a class=webstat href="webstats.asp?orgid=<%=iorgid%>&date=5">Last Year</a></font>
<hr align=left style="color:#000000;size:1 px; width:75%;" ></h3>

<h3 class=webstattitle >E-Govlink Web Site Total Statistics<hr align=left  class=webstat></h3>
<p>
<%
' TOTAL STATS
Call subDisplayTotalStats(iorgid,"",datBeginDate,datEndDate)
%>
</P>

<h3 class=webstattitle >E-Govlink Web Site Section Detail Statistics<hr align=left class=webstat ></h3>
<p>
<%
' ACTION LINE
Call subDisplayStats(iorgid,22,2,"Action Request Form Hit Statistics",datBeginDate,datEndDate)
%>
</P>

<p>
<%
' PAYMENTS
Call subDisplayStats(iorgid,33,3,"Payment Form Hit Statistics",datBeginDate,datEndDate)
%>
</P>

<p>
<%
' CALENDAR
Call subDisplayStats(iorgid,55,5,"Calendar Hit Statistics",datBeginDate,datEndDate)
%>
</P>

<p>
<%
' DOCUMENTS
Call subDisplayStats(iorgid,44,4,"Document Hit Statistics",datBeginDate,datEndDate)
%>
</P>

</div>
<!-- END: DISPLAYING STATS-->


<!-- BEGIN: DISPLAY FOOTER-->
<center>
<hr  style="color:#000000;size:1 px; width:90%;" >
<div class=copyright_text>Copyright &copy;2004-2005. <i>electronic commerce</i> link, inc. dba <a href="http://www.egovlink.com" target="_NEW">egovlink</a>.</div>
</center>
<!-- END: DISPLAY FOOTER-->


</BODY>
</HTML>


<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYSTATS(IORGID,ISECTIONID,ISECTIONHOMEID,SSECTIONTITLE,DATSTARTDATE,DATENDDATE)
'------------------------------------------------------------------------------------------------------------
Sub subDisplayStats(iorgid,isectionid,isectionhomeid,sSectionTitle,datStartDate,datEndDate)

	' GET HOME HITS
	sSQL = "SELECT ISNULL(COUNT(*),0) AS TotalHomeHits FROM visitor_log WHERE  ((orgid = " & iorgid & ") AND (sectionid = " & isectionhomeid & ")) AND ((visitdate >='" & datStartDate & "' AND visitdate <= '" & datEndDate & "'))"
	
	Set oStats = Server.CreateObject("ADODB.Recordset")
	oStats.Open sSQL, Application("DSN") , 3, 1
	

	If NOT oStats.EOF Then
		iHomeHits = oStats("TotalHomeHits")
	Else
		iHomeHits = 0
	End If

	' GET DETAIL HITS FOR TOP 10
	sSQL = "SELECT TOP 10 COUNT(documenttitle) AS NumHits, documenttitle,(SELECT ISNULL(COUNT(*),0) FROM visitor_log WHERE  (orgid = " & iorgid & ") AND (sectionid = " & isectionid & ")  AND ((visitdate >='" & datStartDate & "' AND visitdate <= '" & datEndDate & "'))) AS TotalHits FROM visitor_log WHERE (orgid = " & iorgid & ") AND (sectionid = " & isectionid & ") AND (visitdate >='" & datStartDate & "' AND visitdate <= '" & datEndDate & "') GROUP BY documenttitle ORDER BY NumHits DESC"

	Set oStats = Server.CreateObject("ADODB.Recordset")
	oStats.Open sSQL, Application("DSN") , 3, 1

	If NOT oStats.EOF Then
		
		' DETAIL HITS
		response.write "<font class=sectiontitle>" & sSectionTitle & "</font>"
		response.write "<div class=webstat>"
		response.write "<table cellspacing=0 cellpadding=2 width=100% >"
		response.write "<tr><td class=statlabel >Total Number of Hits: </td><td> " & clng(oStats("TotalHits")) + clng(iHomeHits) & " </td></tr>"
		response.write "<tr><td class=statlabel >Home Page Number of Hits: </td><td> " & clng(iHomeHits) & " </td></tr>"
		response.write "<tr><td  class=statlabel colspan=2 >  <br> Top 10 Individual Page Hits  </td></tr>"

		Do while NOT oStats.EOF 
			response.write "<tr><td class=statdetail align=left>" & oStats("documenttitle") & "</td><td>" & oStats("NumHits")  & "</td></tr>"
			oStats.MoveNext
		Loop

		response.write "</table>"
		response.write "</div>"

		oStats.Close

	Else
		
		' NO DETAIL HITS
		response.write "<font class=sectiontitle>" & sSectionTitle & "</font>"
		response.write "<div class=webstat>"
		response.write "<table cellspacing=0 cellpadding=2 width=100% >"
		response.write "<tr><td class=statlabel >Total Number of Hits: </td><td> " & clng(iHomeHits) & " </td></tr>"
		response.write "<tr><td class=statlabel >Home Page Number of Hits: </td><td> " & clng(iHomeHits) & " </td></tr>"
		response.write "<tr><td  class=statlabel colspan=2 >  <br> Top 5 Individual Page Hits  </td></tr>"
		response.write "<tr><td class=statdetail align=left colspan=2 >No Hits</td></tr>"
		response.write "</table>"
		response.write "</div>"

	End If

	
	Set oStats = Nothing 

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYTOTALSTATS(IORGID,SSECTIONTITLE,DATSTARTDATE,DATENDDATE)
'------------------------------------------------------------------------------------------------------------
Sub subDisplayTotalStats(iorgid,sSectionTitle,datStartDate,datEndDate)

	' GET HOME HITS
	sSQL = "SELECT ISNULL(COUNT(*),0) AS TotalHomeHits FROM visitor_log WHERE  (orgid = " & iorgid & ") AND (visitdate >='" & datStartDate & "' AND visitdate <= '" & datEndDate & "')"

	Set oStats = Server.CreateObject("ADODB.Recordset")
	oStats.Open sSQL, Application("DSN") , 3, 1
	

	If NOT oStats.EOF Then
		iHomeHits = oStats("TotalHomeHits")
	Else
		iHomeHits = 0
	End If

	' GET DETAIL HITS FOR TOP 10
	sSQL = "SELECT TOP 10 COUNT(documenttitle) AS NumHits, documenttitle = Case sectionid WHEN '1' THEN 'E-GovLink Home Page' WHEN '2' THEN 'Action Request Home Page'  WHEN '3' THEN 'Payment Home Page'  WHEN '5' THEN 'Calendar Home Page'  WHEN '4' THEN 'Document Home Page' ELSE documenttitle END,(SELECT ISNULL(COUNT(*),0) FROM visitor_log WHERE  (orgid = " & iorgid & ")  AND ((visitdate >='" & datStartDate & "' AND visitdate <= '" & datEndDate & "'))) AS TotalHits FROM visitor_log WHERE (orgid = " & iorgid & ") AND ((visitdate >='" & datStartDate & "' AND visitdate <= '" & datEndDate & "')) GROUP BY documenttitle,sectionid ORDER BY NumHits DESC"

	Set oStats = Server.CreateObject("ADODB.Recordset")
	oStats.Open sSQL, Application("DSN") , 3, 1

	If NOT oStats.EOF Then
		
		' DETAIL HITS
		response.write "<font class=sectiontitle>" & sSectionTitle & "</font>"
		response.write "<div class=webstat>"
		response.write "<table cellspacing=0 cellpadding=2 width=100% >"
		response.write "<tr><td class=statlabel >Total Number of Hits: </td><td> " & clng(iHomeHits) & " </td></tr>"
		'response.write "<tr><td class=statlabel >Main Page Number of Hits: </td><td> " & clng(iHomeHits) & " </td></tr>"
		response.write "<tr><td  class=statlabel colspan=2 >  <br> Top 10 Individual Page Hits  </td></tr>"

		Do while NOT oStats.EOF 
			response.write "<tr><td class=statdetail align=left>" & oStats("documenttitle") & "</td><td>" & oStats("NumHits")  & "</td></tr>"
			oStats.MoveNext
		Loop

		response.write "</table>"
		response.write "</div>"

		oStats.Close

	End If

	
	Set oStats = Nothing 

End Sub


'------------------------------------------------------------------------------------------------------------
' FUNCTION FNGETNAME(IORGID)
'------------------------------------------------------------------------------------------------------------
Function fnGetName(iorgid)
	
	sReturnValue = "UNKNOWN"

	' GET ORG NAME
	sSQL = "SELECT orgname FROM Organizations WHERE orgid=" & iorgid
	
	Set oOrg = Server.CreateObject("ADODB.Recordset")
	oOrg.Open sSQL, Application("DSN") , 3, 1

	If NOT oOrg.EOF Then
		sReturnValue = oOrg("orgname")
	End If

	fnGetName = sReturnValue

End Function
%>
