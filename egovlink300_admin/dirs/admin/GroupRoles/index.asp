<link href="../../../global.css" rel="stylesheet" type="text/css">

<!--#include file="forum.asp"-->
<CENTER><U><FONT SIZE="4" COLOR="blue"><%=langAdminTitle%></FONT></U><br>
<a href='javascript:self.close();'><FONT SIZE="2" COLOR=""><%=langAdminCloseWindow%></FONT></a></font>
<br><br>
<%
thisname=request.servervariables("script_name")
response.write "<a href="&thisname&"?iOfaction=" & ActNewPost &"&groupid="&request.querystring("groupid")&">"&langNewRecord&"</a>"
response.write "&nbsp;&nbsp;<a href="&thisname&"?iOfaction=" & ActDisplayRecords&">"&langRecordList&"</a><br>"

	%>
<%
SHowForum
%>

