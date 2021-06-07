<!--#include file="../../egovlink300_global/includes/inc_rye.asp"-->
<%
if request.querystring("date") = FindNextRyeBusinessDay(request.querystring("date")) then
	response.write "YES"
else
	response.write "NO"
end if
%>
