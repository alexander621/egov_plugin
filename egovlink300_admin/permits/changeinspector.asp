<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: changeinspector.asp
' AUTHOR: Steve Loar
' CREATED: 08/11/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sets the inspector of an inspection. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   08/11/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitInspectionId, iInspectorUserId

iPermitInspectionId = CLng(request("permitinspectionid"))
iInspectorUserId = CLng(request("inspectoruserid"))


sSql = "UPDATE egov_permitinspections SET inspectoruserid = " & iInspectorUserId & " WHERE permitinspectionid = " & iPermitInspectionId
RunSQL sSql

response.write "UPDATED"


%>