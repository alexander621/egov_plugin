<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: expirationdateupdate.asp
' AUTHOR: Steve Loar
' CREATED: 05/21/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates expiration dates. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   05/21/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sExpirationDate, sSql

iPermitId = CLng(request("permitid"))
sExpirationDate = request("expirationdate")

sSql = "UPDATE egov_permits SET expirationdate = '" & sExpirationDate & "', isexpired = 0 WHERE permitid = " & iPermitId 
sSql = sSql & " AND orgid = " & session("orgid")
RunSQL sSql

response.write "Success"


%>