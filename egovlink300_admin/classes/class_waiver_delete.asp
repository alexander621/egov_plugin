<!-- #include file="../includes/common.asp" //-->

<%
Dim iWaiverId

iWaiverId = CLng(request("iWaiverId"))
	
sSql = "DELETE FROM egov_class_waivers WHERE waiverid = " &  iWaiverId 
'	response.write sSQL
	
RunSQLStatement sSql

' REDIRECT TO facility waivers page
response.redirect "class_waivers.asp"


%>
