<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getseasonclasses.asp
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

iClassSeasonId = request("classseasonid")

sSQL = "SELECT C.classid, C.classname "
sSql = sSql & " FROM egov_class C, egov_class_status S, egov_registration_option RO "
sSql = sSql & " WHERE C.statusid = S.statusid AND S.statusname = 'ACTIVE' AND C.classseasonid = " & iClassSeasonId
sSql = sSql & " AND RO.optionid = C.optionid AND RO.canpurchase = 1 "
sSql = sSql & " AND C.orgid = " & SESSION("orgid") & " ORDER BY C.classname"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 0, 1

If Not oRs.EOF Then 
	sResults = "<select id='earlyregistrationclassid' name='earlyregistrationclassid' multiple='multiple' size='20'>"
	Do While Not oRs.EOF
		sResults = sResults & "<option value='" & oRs("classid") & "'"
		sResults = sResults & ">" & oRs("classname") & "</option>"
		oRs.MoveNext
	Loop
	sResults = sResults &  "</select>"
Else
	sResults = "<input type='hidden' id='earlyregistrationclassid' name='earlyregistrationclassid' value='0' />No Classes Found"
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults

%>