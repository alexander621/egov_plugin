<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: extrapagedelete.asp
' AUTHOR: Steve Loar
' CREATED: 08/25/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This script allows the deleting of extra mobile pages
'
' MODIFICATION HISTORY
' 1.0   08/25/2011   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, iPageId

iPageId = CLng(request("pageid"))


sSql = "DELETE FROM egov_extramobilepages WHERE orgid = " & session("orgid") & " AND pageid = " & iPageId


RunSQLStatement sSql


' Take them back To the extra pages list
response.redirect "extrapagelist.asp?s=d"


%>