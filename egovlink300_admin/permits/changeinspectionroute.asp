<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: changeinspectionroute.asp
' AUTHOR: Steve Loar
' CREATED: 08/11/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sets the route order of an inspection. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   08/11/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitInspectionId, iRouteOrder

iPermitInspectionId = CLng(request("permitinspectionid"))
iRouteOrder = CLng(request("routeorder"))


sSql = "UPDATE egov_permitinspections SET routeorder = " & iRouteOrder & " WHERE permitinspectionid = " & iPermitInspectionId
RunSQL sSql

response.write "UPDATED"


%>