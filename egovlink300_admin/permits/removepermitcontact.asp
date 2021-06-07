<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: removepermitcontact.asp
' AUTHOR: Steve Loar
' CREATED: 06/06/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Removes contractors from permits. Called via AJAX from permitedit.asp
'
' MODIFICATION HISTORY
' 1.0   06/06/2008	Steve Loar - INITIAL VERSION
' 1.1	06/12/2008	Steve Loar - Changed to move into prior contacts not delete them
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitContactId, sSql

iPermitContactId = CLng(request("permitcontactid"))

' Remove from the permit contact table
'sSql = "DELETE FROM egov_permitcontacts WHERE permitcontactid = " & iPermitContactId
sSql = "UPDATE egov_permitcontacts SET ispriorcontact = 1 WHERE permitcontactid = " & iPermitContactId
RunSQL sSql

' Remove from the permit contact licenses table
'sSql = "DELETE FROM egov_permitcontacts_licenses WHERE permitcontactid = " & iPermitContactId
'RunSQL sSql

response.write "Success"

%>