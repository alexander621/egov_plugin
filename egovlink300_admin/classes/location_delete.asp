<!-- #include file="../includes/common.asp" //-->
<%
Dim sSql, iLocationId

iLocationId = CLng(request("iLocationId"))

sSql = "DELETE FROM egov_class_Location WHERE Locationid = " &  iLocationId & ""
'response.write sSQL
RunSQLStatement sSql 

sSql = "UPDATE egov_class SET locationid = NULL WHERE locationid = " & iLocationId
RunSQLStatement sSql 


' REDIRECT TO location list page
response.redirect "location_mgmt.asp"

%>
