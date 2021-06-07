<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: checkinpsecteddateneed.asp
' AUTHOR: Steve Loar
' CREATED: 07/24/08
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets whether an inspection status requires an inspected date. It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   07/24/08	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, sResults, iInspectionStatusId

iInspectionStatusId = CLng(request("inspectionstatusid"))

sSql = "SELECT needsinspecteddate FROM egov_inspectionstatuses WHERE orgid = " & session("orgid")
sSql = sSql & " AND inspectionstatusid = " & iInspectionStatusId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRs.EOF Then 
	If oRs("needsinspecteddate") Then
		sResults = "NEEDED"
	Else 
		sResults = "NOT NEEDED"
	End If 
Else
	sResults = "NOT NEEDED"
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults



%>
