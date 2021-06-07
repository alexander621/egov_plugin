<!--#include file='header.asp'-->
<% 
dim conn,cmd,resultid,strSuccess
if not HasPermission("CanRegisterContact") then
response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleAdmin)
 end if %>   
  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_directory.jpg"></td>
      <td><font size="+1"><b>Registration: New User</b></font><br>
	  <div id="goback" name="goback">
	  <img src='../images/arrow_back.gif' align='absmiddle'><a href='display_citizen.asp'><%=langGoBack%></a>
	  </div>
	  
	  </td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
        <!--#include file='quicklink_citizen.asp'-->      
      </td>
      <td colspan="2" valign="top">
        
<%
set conn = Server.CreateObject("ADODB.Connection")
set cmd=Server.CreateObject("ADODB.Command")
conn.Open Application("DSN")
Set cmd.ActiveConnection=conn
cmd.commandtext="NewCitizen"
cmd.commandtype=&H0004
cmd.Parameters.Refresh
With request
cmd.parameters(1)= session("orgid")
cmd.parameters(2)=.form("Password")
cmd.parameters(3)=.form("First Name")
cmd.parameters(4)=.form("Last Name")
cmd.parameters(5)=.form("Business Name")
cmd.parameters(6)=.form("Address Line 1")
cmd.parameters(7)=.form("Address Line 2")
cmd.parameters(8)=.form("Home Phone")
cmd.parameters(9)=.form("Work Phone")
cmd.parameters(10)=.form("City")
cmd.parameters(11)=.form("State")
cmd.parameters(12)=.form("Zip")
cmd.parameters(13)=.form("Country")
cmd.parameters(14)=.form("Fax Number")
cmd.parameters(15)=.form("Email")
cmd.parameters(16)=1
cmd.execute
end with
ResultID=cmd.parameters(0)
conn.close
set conn=nothing
set cmd=nothing


' DEBUG CODE: 
'For each item in Request.Form
	'response.write item & " = " & request(item) & "<BR>"
'Next
'response.write "<br>ResultID="&ResultID


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
	case -1
		response.write "<br><li>"&langInsertNormalUser3&"</li>"
		'response.write "<br><a href='javascript:history.go(-1)'>"&langGoBack&"</a>"
	case 0
		response.write "<br><li>No Email address was entered.</li>"
		'response.write "<br><a href='javascript:history.go(-1)'>"&langGoBack&"</a>"
	case -4
		response.write "<br><li>"&langInsertNormalUser5&"</li>"
		'response.write "<br><a href='javascript:history.go(-1)'>"&langGoBack&"</a>"
	case else
		if ResultID>0 then
		response.write "<br><li>"&langInsertNormalUser6&"</li>"
		strSuccess="<br><img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='display_citizen.asp'>Registration: New User</a>"
		response.write "<script>document.all.goback.innerHTML="""&strSuccess&"""</script>"

		'response.write "<br><br><a HREF=javascript:openWin2('admin/extended/index.asp?onload=1&iOfaction=4&UserID="&ResultID&"','_blank')>"&langInsertNormalUser7&"</a> "
		end if
end select
%>


<!--#include file='footer.asp'-->

<script language=javascript>
	function openWin2(url, name) {
  popupWin = window.open(url, name,
"resizable,width=500,height=450");
}
</script>
