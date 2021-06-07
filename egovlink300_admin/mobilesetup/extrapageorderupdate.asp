<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: extrapageorderupdate.asp
' AUTHOR: Steve Loar
' CREATED: 08/26/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sets the order of an extra mobile page. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   08/26/2011	Steve Loar - INITIAL VERSION
' 1.1  11/19/13  Terry Foster - CLng Bug Fix
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPageId, iDisplayOrder, sSql

if isnumeric(request("pageid")) then

	iPageId = CLng(request("pageid"))

	iDisplayOrder = CLng(request("displayorder"))

	sSql = "UPDATE egov_extramobilepages SET displayorder = " & iDisplayOrder & " WHERE pageid = " & iPageId
	sSql = sSql & " AND orgid = " & session("orgid")

	RunSQLStatement sSql

	response.write "UPDATED"
else
	response.write "FAILED"
end if

%>

