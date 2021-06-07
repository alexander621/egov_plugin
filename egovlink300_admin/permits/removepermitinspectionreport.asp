<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: removepermitinspectionreport.asp
' AUTHOR: Terry Foster
' CREATED: 12/16/2019
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Removes an inspection report from a permit
'
' MODIFICATION HISTORY
' 1.0   12/16/2019	Terry Foster - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitInspectionReportId, sSql

iPermitInspectionReportId = CLng(request("permitinspectionReportid"))

' Remove from the permit inspection table
sSql = "DELETE FROM egov_permitinspectionreports WHERE permitinspectionreportid = " & iPermitInspectionReportId
RunSQL sSql

response.write "Success"

%>
