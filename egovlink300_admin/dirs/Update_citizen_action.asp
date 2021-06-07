<!--#include file='header.asp'-->
<% 
dim pagesize, totalpages,RA,totalrecords,groupname,thisname,currentpage,conn,rs,groupmode,strSQL,CName,AdditonURL,numstartid,numendid,i,deleteurl,EventOrNot,Str_Bgcolor,username,password,str_image,editurl,FullName,l_length,l_name,b_update,j,fld,cmd,ResultID,strSuccess

if not HasPermission("CanEditUser") and session("userid")<>clng(request.form("userid")) then
response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleEditUser)
 end if %>

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_directory.jpg"></td>
      <td><font size="+1"><b><%=langUpdateUserAccount%></b></font><br>
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
cmd.commandtext="Updatecitizen"
cmd.commandtype=&H0004
cmd.Parameters.Refresh
With request
cmd.parameters(1)=.form("userid")
cmd.parameters(2)=.form("orgid")
cmd.parameters(3)=.form("userpassword")
cmd.parameters(4)=.form("userfname")
cmd.parameters(5)=.form("userlname")
cmd.parameters(6)=.form("useraddress")
cmd.parameters(7)=.form("useraddress2")
cmd.parameters(8)=.form("userhomephone")
cmd.parameters(9)=.form("userworkphone")
cmd.parameters(10)=.form("userfax")
cmd.parameters(11)=.form("useremail")
cmd.parameters(12)=.form("userbusinessname")
cmd.parameters(13)=.form("email_o")
cmd.execute
end with
ResultID=cmd.parameters(0)
conn.close
set conn=nothing
set cmd=nothing

'response.write "<br>ResultID="&ResultID
select case ResultID
case -100
response.write "<br><li>"&langErrorDatabase&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>Go Back</a>"
case -4
response.write "<br><li>"&langExpiredSession&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>Go Back</a>"
case -3
response.write "<br><li>"&langNoFirstName&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>Go Back</a>"
case -2
response.write "<br><li>"&langNoLastName&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>Go Back</a>"
case -1
response.write "<br><li>"&langNoPassword&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>Go Back</a>"
case 0
response.write "<br><li>"&langNoUserName&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>Go Back</a>"
case 2
response.write "<br><li>"&langUserNameIsTaken&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>Go Back</a>"
case 1
response.write "<br><li>"&langSucessUpdate&"</li><br>"

strSuccess="<br><img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='display_member.asp'>"&langBackToUserDisplay&"</a>"
response.write "<script>document.all.goback.innerHTML="""&strSuccess&"""</script>"

end select
%>
<!--#include file='footer.asp'-->

