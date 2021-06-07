<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: removepermitinspection.asp
' AUTHOR: Steve Loar
' CREATED: 07/10/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Removes an inspection from a permit
'
' MODIFICATION HISTORY
' 1.0   07/10/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitInspectionId, sSql

iPermitInspectionId = CLng(request("permitinspectionid"))

' Remove from the permit inspection table
sSql = "DELETE FROM egov_permitinspections WHERE permitinspectionid = " & iPermitInspectionId
RunSQL sSql

response.write "Success"

%>
