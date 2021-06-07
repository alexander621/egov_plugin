<% if not HasPermission("CanRegisterUser") then
response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleRegisterContact)
 end if %> 
<!--#include file='header.asp'-->
  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_directory.jpg"></td>
      <td><font size="+1"><b><%=langRegisterContactTitle%></b></font><br>
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
Set cmd.ActiveConnection=conn
cmd.commandtext="NewContactUser"
cmd.commandtype=&H0004
cmd.Parameters.Refresh
With request
cmd.parameters(1)=.form("orgid")
cmd.parameters(2)=.form("firstname")
cmd.parameters(3)=.form("middleinitial")

cmd.parameters(4)=.form("lastname")
cmd.parameters(5)=.form("nickname")
cmd.parameters(6)=.form("jobtitle")
cmd.parameters(7)=.form("department")
cmd.parameters(8)=.form("homeaddress")

cmd.parameters(9)=.form("businessaddress")
cmd.parameters(10)=.form("homenumber")
cmd.parameters(11)=.form("businessnumber")
cmd.parameters(12)=.form("mobilenumber")
cmd.parameters(13)=.form("pagenumber")

cmd.parameters(14)=.form("faxnumber")
cmd.parameters(15)=.form("email")
cmd.parameters(16)=.form("email2")
cmd.parameters(17)=.form("webpage")
cmd.parameters(18)=.form("birthday")
cmd.execute
end with
ResultID=cmd.parameters(0)
conn.close
set conn=nothing
set cmd=nothing

select case ResultID
case -100
response.write "<br><li>"&langInsertDatabaseError&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>"&langGoBack&"</a>"
case -3
response.write "<br><li>"&langInsertNormalUser1&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>"&langGoBack&"</a>"
case -2
response.write "<br><li>"&langInsertNormalUser2&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>"&langGoBack&"</a>"
case 2
response.write "<br><li>"&langInsertNormalUser5&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>"&langGoBack&"</a>"
case 1
response.write "<br><li>"&langInsertNormalUser6&"</li>"
strSuccess="<br><img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<A HREF='display_contact.asp'>"&langBackToContactDisplay&"</A>"
response.write "<script>document.all.goback.innerHTML="""&strSuccess&"""</script>"
end select
%>
<!--#include file='footer.asp'-->

