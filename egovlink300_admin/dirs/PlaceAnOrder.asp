<!--#include file='header.asp'-->
<% 
dim conn,cmd,resultid,strSuccess
 %> 
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
cmd.commandtext="RegularMoveToQueueForSendEmail"
cmd.commandtype=&H0004
cmd.Parameters.Refresh
cmd.parameters(1)=30
cmd.execute

cmd.Parameters.Refresh
cmd.parameters(1)=7
cmd.execute

cmd.Parameters.Refresh
cmd.parameters(1)=1
cmd.execute

cmd.Parameters.Refresh
cmd.parameters(1)=0
cmd.execute

conn.close
set conn=nothing
set cmd=nothing
response.write "<br>You will receive your subscription in just 1 minutes, check your email account!! "

%>
<!--#include file='footer.asp'-->

