<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../../egovlink300_global/includes/inc_passencryption.asp" //-->
<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: insert_normaluser.asp
' AUTHOR: ????
' CREATED: ??/??/????
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module adds users to the admin side of the application.
'
' MODIFICATION HISTORY
' 1.0 ??/??/??  ???? - INITIAL VERSION
' 1.1	02/13/07	 Steve Loar - Added locationid
' 1.2 04/01/09  David Boyer - Added imagefilename
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim iUserId, sInsertResults, sUserName
 sLevel = "../"  'Override of value from common.asp

 If Not userhaspermission(session("userid"),"add users") Then 
 	  response.redirect sLevel & "permissiondenied.asp"
 End If 

 sUserName      = " "
 iUserId        = 0
 sInsertResults = InsertNewUser( iUserId, sUserName )

'Check for user permissions
 lcl_userhaspermission_user_permission = userhaspermission(session("userid"),"user permission")
%>
<html>
<head>
	<title><%=langBSCommittees%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script src="../scripts/selectAll.js"></script>

<script language="javascript">
<!--
function openWin2(url, name) {
  popupWin = window.open(url, name, "resizable,width=500,height=450");
}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="menu">
  <tr>
      <td background="../images/back_main.jpg">
          <% 'DrawTabs tabCommittees,2 %>

       			<% ShowHeader sLevel %>
       			<!--#Include file="../menu/menu.asp"--> 
      </td>
  </tr>
</table>
<!-- #include file="dir_constants.asp"-->
<div id="content">
  <div id="centercontent">
<table border="0" cellpadding="10" cellspacing="0" width="100%">
  <tr>
      <td>
          <font size="+1"><strong>User: <%=sUserName%></strong></font><br />
          <% displayButtons iUserID %>
   	  </td>
      <td>&nbsp;</td>
  </tr>
  <tr>
      <td colspan="2" valign="top">
          <%	response.write sInsertResults %>
     	</td>
  </tr>
</table>
  </div>
</div>
<p>

<!--#Include file="../admin_footer.asp"-->  

<!--#include file="footer.asp"-->

<%

'------------------------------------------------------------------------------
Function InsertNewUser( ByRef iUserId, ByRef sUserName )
	Dim conn, cmd, resultid, strSuccess

	sUserName = request("firstname") & " " & request("lastname")

	Set conn = Server.CreateObject("ADODB.Connection")
	Set cmd  = Server.CreateObject("ADODB.Command")
	conn.Open Application("DSN")
	Set cmd.ActiveConnection=conn
	cmd.commandtext = "NewNormalUser"
	cmd.commandtype = &H0004
	cmd.Parameters.Refresh

	newpassword = createHashedPassword(request.form("password"))

	With request
		cmd.parameters(1)=session("orgid")
		cmd.parameters(2)=.form("username")
		cmd.parameters(3)=newpassword
		cmd.parameters(4)=.form("firstname")
		cmd.parameters(5)=.form("middleinitial")

		cmd.parameters(6)=.form("lastname")
		cmd.parameters(7)=.form("nickname")
		cmd.parameters(8)=.form("jobtitle")
		'cmd.parameters(9)=.form("department")
		cmd.parameters(9)=""
		cmd.parameters(10)=.form("homeaddress")

		cmd.parameters(11)=.form("businessaddress")
		cmd.parameters(12)=.form("homenumber")
		cmd.parameters(13)=.form("businessnumber")
		cmd.parameters(14)=.form("mobilenumber")
		cmd.parameters(15)=.form("pagernumber")

		cmd.parameters(16)=.form("faxnumber")
		cmd.parameters(17)=.form("email")
		cmd.parameters(18)=.form("email2")
		cmd.parameters(19)=.form("webpage")
		cmd.parameters(20)=.form("birthday")
		cmd.parameters(21)=.form("companyname")
		cmd.parameters(22)=.form("locationid")
		cmd.parameters(23)=.form("staff_dir_display")
		cmd.parameters(24)=.form("imagefilename")
		cmd.execute
	End with

	ResultID = cmd.parameters(0)
	conn.close
	Set conn = Nothing 
	Set cmd  = Nothing 

'Now update the staff directory - organizational groups which was the department column.
 lcl_org_group_ids = request("department")

'If the field has been entered then set up a loop so that we can update the assignment table
 if lcl_org_group_ids <> "" then
    sSQLg = "SELECT distinct org_group_id "
    sSQLg = sSQLg & " FROM egov_staff_directory_groups "
    sSQLg = sSQLg & " WHERE org_group_id IN (" & lcl_org_group_ids & ") "
    sSQLg = sSQLg & " ORDER BY org_group_id "

  	 set rsg = Server.CreateObject("ADODB.Recordset")
   	rsg.Open sSQLg, Application("DSN"), 0, 1

    if not rsg.eof then
       while not rsg.eof
          sSQLi = "INSERT INTO egov_staff_directory_usergroups (userid, org_group_id) VALUES ("
          sSQLi = sSQLi & ResultID & ", "
          sSQLi = sSQLi & rsg("org_group_id") & ") "

        	 set rsi = Server.CreateObject("ADODB.Recordset")
         	rsi.Open sSQLi, Application("DSN"), 0, 1

          rsg.movenext
       wend
    end if
 end if

	select case ResultID
 		case -100
	  		InsertNewUser = "<br /><li>" & langInsertDatabaseError & "</li>"
 		case -3
	  		InsertNewUser = "<br /><li>" & langInsertNormalUser1   & "</li>"
 		case -2
	  		InsertNewUser = "<br /><li>" & langInsertNormalUser2   & "</li>"
 		case -1
	  		InsertNewUser = "<br /><li>" & langInsertNormalUser3   & "</li>"
 		case 0
	  		InsertNewUser = "<br /><li>" & langInsertNormalUser4   & "</li>"
 		case -4
	  		InsertNewUser = "<br /><li>" & langInsertNormalUser5   & "</li>"
 		case else
	  		if ResultID > 0 Then
			    	iUserId = ResultID
    				InsertNewUser = "<br /><li>" & langInsertNormalUser6 & "</li>"
   		end if
	end select	
End Function  


'------------------------------------------------------------------------------
sub displayButtons( ByVal p_userid )

  response.write "<input type=""button"" name=""backButton"" id=""backButton"" value=""Back to User List"" class=""button"" onclick=""location.href='display_member.asp'"" />" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "<input type=""button"" name=""updateButton"" id=""updateButton"" value=""Save Changes"" class=""button"" onclick=""location.href='update_user.asp?userid=" & p_userid & "'"" />" & vbcrlf
  response.write "<input type=""button"" name=""addAnotherButton"" id=""addAnotherButton"" value=""Add Another User"" class=""button"" onclick=""location.href='register_normaluser.asp'"" />" & vbcrlf

  if lcl_userhaspermission_user_permission then
     response.write "<input type=""button"" name=""userPermissionsButton"" id=""userPermissionsButton"" value=""User Permissions"" class=""button"" onclick=""location.href='../security/edit_user_security.asp?iuserid=" & p_userid & "'"" />" & vbcrlf
  end if

end Sub


%>
