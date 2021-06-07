<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: changefixturetypeorder.asp
' AUTHOR: Steve Loar
' CREATED: 12/04/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sets the display order of a fixture type. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   12/04/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFixtureTypeid, iRouteOrder

iPermitFixtureTypeid = CLng(request("permitfixturetypeid"))
iDisplayOrder = CLng(request("displayorder"))


sSql = "UPDATE egov_permitfixturetypes SET displayorder = " & iDisplayOrder & " WHERE permitfixturetypeid = " & iPermitFixtureTypeid
RunSQL sSql

response.write "UPDATED"


%>