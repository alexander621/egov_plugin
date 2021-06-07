<!--#include file='header.asp'-->
<% 
dim conn,cmd,resultid,strSuccess
if not HasPermission("CanRegisterRole") then
response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleRegisterCommittee)
 end if %> 
  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_directory.jpg"></td>
      <td><font size="+1"><b><%=langRegisterRoleTitle%></b></font>
	  <br>
	   <div id="goback" name="goback">
	  <img src='../images/arrow_back.gif' align='absmiddle'><a href='javascript:history.go(-1)'><%=langGoBack%></a>
	  </div>
	  </td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
        <!--#include file='quicklink.asp'-->      
      </td>
      <td colspan="2" valign="top">
        
<%
set conn = Server.CreateObject("ADODB.Connection")
set cmd=Server.CreateObject("ADODB.Command")
conn.Open Application("DSN")
cmd.ActiveConnection=Application("DSN")
cmd.commandtext="NewRole"
cmd.commandtype=&H0004
cmd.Parameters.Refresh
With request
cmd.parameters(1)=left(.form("rolename"),30)
cmd.parameters(2)=left(.form("roledescription"),150)
cmd.parameters(3)=Session("OrgID")
cmd.execute
end with
ResultID=cmd.parameters(0)
conn.close
set conn=nothing
set cmd=nothing
'response.write "<br>ResultID="&ResultID
select case ResultID
case -100
response.write "<br><li>"&langInsertDatabaseError&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>"&langGoBack&"</a>"
case 2
response.write "<br><li>"&langDuplicateRolename&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>"&langGoBack&"</a>"
case 1
response.write "<br><li>"&langInsertRole
strSuccess="<br><img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<A HREF=display_roles.asp>"&langBackToRoleDisplay&"</A>"
response.write "<script>document.all.goback.innerHTML="""&strSuccess&"""</script>"
end select
%>
<!--#include file='footer.asp'-->

