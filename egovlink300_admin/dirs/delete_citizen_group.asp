<!--#include file='header.asp'-->


<%
dim conn,strSQL,thisname,currentpage,pagesize,totalpages,delete,id
pagesize=20 
totalpages=1

thisname=request.servervariables("script_name")
if not isempty(request.querystring("currentpage")) then
	CurrentPage=clng(request.querystring("currentpage"))
else
	currentpage=1
end if
%>


  <table border="0" cellpadding="10" cellspacing="0" width="100%">

    <tr>
      <td valign="top" width='151'>
		 <center> <img src='../images/icon_directory.jpg'></center>
	 <br>
	       <!--#include file='quicklink.asp'-->   
      </td>
      <td colspan="2" valign="top">

<%
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
for each delete in request.form("delete")
	id=clng(delete)
	strSQL = "delete from citizengroups where groupid="&id
	conn.execute(strSQL)
next

if request.form("delete").count=0 then response.write langNoDelete
response.write "<br><a href='display_citizen_groups.asp'>Back to Group List</a>"
response.redirect("display_citizen_groups.asp")
%>

</td>
</tr>	
</table>


<!--#include file='footer.asp'-->
