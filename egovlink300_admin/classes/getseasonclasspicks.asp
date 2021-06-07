<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getseasonclasspicks.asp
' AUTHOR: Steve Loar
' CREATED: 2/15/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets the classes for a selected season. It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   2/15/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iClassSeasonId, sSql, oRs, sResults

iClassSeasonId = CLng(request("classseasonid"))

sSql = "SELECT T.timeid, C.classname, T.activityno, C.classseasonid "
sSql = sSql & "FROM egov_class_time T, egov_class C, egov_class_status S "
sSql = sSql & "WHERE C.classid = T.classid AND C.statusid = S.statusid "
sSql = sSql & "AND C.orgid = " & SESSION("orgid")
sSql = sSql & " AND C.classseasonid = " & iClassSeasonId
sSql = sSql & " AND S.iscancelled = 0 AND T.iscanceled = 0 "
sSql = sSql & "ORDER BY C.classname, T.activityno"
'sResults = sSql

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

If Not oRs.EOF Then 
	sResults = "<select id='classtimeid' name='classtimeid'>"
	Do While Not oRs.EOF
		sResults = sResults & "<option value='" & oRs("timeid") & "'"
		sResults = sResults & ">" & oRs("classname") & " - " & oRs("activityno") & "</option>"
		oRs.MoveNext
	Loop
	sResults = sResults &  "</select>"
Else
	sResults = "<input type='hidden' id='classtimeid' name='classtimeid' value='0' />No Classes Found"
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults

%>