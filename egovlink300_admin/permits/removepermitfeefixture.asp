<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: removepermitfeefixture.asp
' AUTHOR: Steve Loar
' CREATED: 05/13/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Removes fees from permits
'
' MODIFICATION HISTORY
' 1.0   05/13/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFixtureId, sSql

iPermitFixtureId = CLng(request("permitfixtureid"))

' Remove from the fixtures table
sSql = "DELETE FROM egov_permitfixtures WHERE permitfixtureid = " & iPermitFixtureId
RunSQL sSql

' Remove from the fixtures step table
sSql = "DELETE FROM egov_permitfixturestepfees WHERE permitfixtureid = " & iPermitFixtureId
RunSQL sSql


response.write "Success"

%>
