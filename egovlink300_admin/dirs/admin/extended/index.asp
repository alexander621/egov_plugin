<link href="../../../global.css" rel="stylesheet" type="text/css">
<%
const langAdminPropertyAdded= "Property added"
const langAdminUserExtendedTitle= "Manage User Extended properties"
const langAdminCloseWindow	= "Close Window"
const langAdminNewProperty	= "New property"
const langAdminAllProperty	= "All properties"
const langDelete="Delete"
fields_description_extended	= array("Extended ID","User ID","Property","Value","Added Time")
fields_description_committee=array("GroupID","Organizatioin ID","GroupName","Description","Added Time")
%>
<body <%  if request.querystring("onload")<>1 then response.write "onload=""javasript:opener.location.reload(true);"""%>>
<!--#include file="forum.asp"-->
<CENTER><U><FONT SIZE="4" COLOR="blue"><%=langAdminUserExtendedTitle%></FONT></U><br>
<a href='javascript:self.close();'><FONT SIZE="2" COLOR=""><%=langAdminCloseWindow%></FONT></a></font>
<br><br>
<%
thisname=request.servervariables("script_name")
response.write "<a href="&thisname&"?iOfaction=" & ActNewPost &"&userid="&request.querystring("userid")&">"&langAdminNewProperty&"</a>"
response.write "&nbsp;&nbsp;<a href="&thisname&"?iOfaction=" & ActDisplayRecords &"&userid="&request.querystring("userid")& ">"&langAdminAllProperty&"</a><br>"

%>
<%
SHowForum
%>
</body>
