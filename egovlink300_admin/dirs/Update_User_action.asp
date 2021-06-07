<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../../egovlink300_global/includes/inc_passencryption.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: update_user_action.asp
' AUTHOR: ????
' CREATED: ??/??/????
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  ??/??/??  ????? ????? - INITIAL VERSION
' 1.1	 02/13/07	 Steve Loar  - Added locationid
' 1.2	 02/21/07  Steve Loar  - Added class supervisor flag
' 1.3  12/18/07  David Boyer	- Added staff directory options
' 1.4  03/31/09  David Boyer - Added image field
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
sLevel = "../"  'Override of value from common.asp

' You can edit if you have edit permision or the user is editing themselves
if NOT UserHasPermission(Session("UserId"),"edit users") then
   if NOT UserHasPermission(session("userid"),"edit staff directory") then
     	if Session("UserID") <> clng(Trim(request("userid"))) then
		       response.redirect sLevel & "permissiondenied.asp"
     	end if
   end if
end if

'Check for org features
 lcl_orghasfeature_class_supervisors = orghasfeature("class supervisors")

'Check for user permissions
 lcl_userhaspermission_edit_users = userhaspermission(session("userid"),"edit users")
%>
<html>
<head>
  <title><%=langBSCommittees%></title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

  <script src="../scripts/selectAll.js"></script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="menu">
  <tr>
      <td background="../images/back_main.jpg">
          <%  'DrawTabs tabCommittees,2  %>

        		<% ShowHeader sLevel %>
       			<!--#Include file="../menu/menu.asp"--> 
      </td>
  </tr>
</table>
<!-- #include file="dir_constants.asp"-->
<div id="content">
 	<div id="centercontent">
<% 
dim pagesize, totalpages,RA,totalrecords,groupname,thisname,currentpage,conn,rs,groupmode,strSQL,CName,AdditonURL
Dim numstartid,numendid,i,deleteurl,EventOrNot,Str_Bgcolor,username,password,str_image,editurl,FullName,l_length
Dim l_name,b_update,j,fld,cmd,ResultID,strSuccess

' You can edit if you have edit permision or the user is editing themselves
'If Not UserHasPermission( Session("UserId"), "edit users" ) Then
'	If Session("UserID") <> clng(Trim(request("userid"))) Then 
'		response.redirect sLevel & "permissiondenied.asp"
'	End If 
'End If 

'if not HasPermission("CanEditUser") and session("userid")<>clng(request.form("userid")) then
'	response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleEditUser)
'end if
%>
<table border="0" cellpadding="10" cellspacing="0" width="100%">
  <tr>
      <td width="151" align="center">&nbsp;<!--<img src="../images/icon_directory.jpg">--></td>
      <td><font size="+1"><b><%=langUpdateUserAccount%></b></font><br />
       	  <div id="goback" name="goback">
         	<img src="../images/arrow_back.gif" align="absmiddle"><a href="javascript:history.go(-1)"><%=langGoBack%></a>
          </div>
   	  </td>
      <td width="200">&nbsp;</td>
  </tr>
  <tr>
      <td valign="top">
          <!--#include file="quicklink.asp"-->
      </td>
      <td colspan="2" valign="top">
<%
		set conn = Server.CreateObject("ADODB.Connection")
		set cmd  = Server.CreateObject("ADODB.Command")
		conn.Open Application("DSN")

		Set cmd.ActiveConnection=conn
		'if request.form("orgid") = "5" then
		cmd.commandtext = "UpdateUser_optpass"
		'else
		'cmd.commandtext = "UpdateUser"
		'end if
		cmd.commandtype = &H0004
		cmd.Parameters.Refresh

		
		newpassword = createHashedPassword(request.form("password"))

		With request
			cmd.parameters(1)=.form("userid")
			cmd.parameters(2)=.form("orgid")
			cmd.parameters(3)=.form("username")
			cmd.parameters(4)=newpassword
			cmd.parameters(5)=.form("firstname")
			cmd.parameters(6)=.form("middleinitial")

			cmd.parameters(7)=.form("lastname")
			cmd.parameters(8)=.form("nickname")
			cmd.parameters(9)=.form("jobtitle")
   			cmd.parameters(10)=""
			cmd.parameters(11)=.form("homeaddress")
			cmd.parameters(12)=.form("businessaddress")
			cmd.parameters(13)=.form("homenumber")
			cmd.parameters(14)=.form("businessnumber")
			cmd.parameters(15)=.form("mobilenumber")
			cmd.parameters(16)=.form("pagernumber")
			cmd.parameters(17)=.form("faxnumber")
			cmd.parameters(18)=.form("email")
			cmd.parameters(19)=.form("email2")
			cmd.parameters(20)=.form("webpage")
			cmd.parameters(21)=.form("birthday")
			cmd.parameters(22)=.form("companyname")
			cmd.parameters(23)=.form("username_o")
			cmd.parameters(24)=.form("locationid")
   cmd.parameters(25)=0 'primarygroupid
   cmd.parameters(26)=.form("staff_dir_display")
   cmd.parameters(27)=.form("imagefilename")
			cmd.execute
		end with

	'	ResultID = cmd.parameters(0)
		ResultID = 1
		'conn.close

		set conn = Nothing 
		set cmd  = Nothing 

 'Now update the staff directory - organizational groups which was the department column.
  lcl_org_group_ids = request("department")

 'If the field has been entered then set up a loop so that we can update the assignment table
  if lcl_org_group_ids <> "" then
     sSQLg = "SELECT org_group_id "
     sSQLg = sSQLg & " FROM egov_staff_directory_groups "
     sSQLg = sSQLg & " WHERE org_group_id IN (" & lcl_org_group_ids & ") "
     sSQLg = sSQLg & " ORDER BY org_group_id "

   	 set rsg = Server.CreateObject("ADODB.Recordset")
    	rsg.Open sSQLg, Application("DSN"), 0, 1

     if not rsg.eof then
       'If the user has selected value(s) then clear the existing values and re-insert them.
        sSQLd = "DELETE FROM egov_staff_directory_usergroups WHERE userid = " & request("userid")

      	 set rsd = Server.CreateObject("ADODB.Recordset")
       	rsd.Open sSQLd, Application("DSN"), 0, 1

        while not rsg.eof
           sSQLi = "INSERT INTO egov_staff_directory_usergroups (userid, org_group_id) VALUES ("
           sSQLi = sSQLi & request("userid") & ", "
           sSQLi = sSQLi & rsg("org_group_id") & ") "

         	 set rsi = Server.CreateObject("ADODB.Recordset")
          	rsi.Open sSQLi, Application("DSN"), 0, 1

           rsg.movenext
        wend
     end if
  else
     sSQLd = "DELETE FROM egov_staff_directory_usergroups WHERE userid = " & request("userid")

   	 set rsd = Server.CreateObject("ADODB.Recordset")
    	rsd.Open sSQLd, Application("DSN"), 0, 1
  end if


		if lcl_orghasfeature_class_supervisors then
  			SetClassSupervisor request("userid"), request("isclasssupervisor")
  end if

	'response.write "<br />ResultID="&ResultID
		select case ResultID
  		case -100
 		     	response.write "<br /><li>"&langErrorDatabase&"</li>"
     			'response.write "<br /><a href='javascript:history.go(-1)'>Go Back</a>"
  		case -4
		      	response.write "<br /><li>"&langExpiredSession&"</li>"
     			'response.write "<br /><a href='javascript:history.go(-1)'>Go Back</a>"
		  case -3
  	    		response.write "<br /><li>"&langNoFirstName&"</li>"
   		  	'response.write "<br /><a href='javascript:history.go(-1)'>Go Back</a>"
  		case -2
		      	response.write "<br /><li>"&langNoLastName&"</li>"
     			'response.write "<br /><a href='javascript:history.go(-1)'>Go Back</a>"
		  case -1
  	    		response.write "<br /><li>"&langNoPassword&"</li>"
   		  	'response.write "<br /><a href='javascript:history.go(-1)'>Go Back</a>"
  		case 0
			      response.write "<br /><li>"&langNoUserName&"</li>"
    	 		'response.write "<br /><a href='javascript:history.go(-1)'>Go Back</a>"
  		case 2
     	 		response.write "<br /><li>"&langUserNameIsTaken&"</li>"
     			'response.write "<br /><a href='javascript:history.go(-1)'>Go Back</a>"
  		case 1
'		      	response.write "<br /><li>"&langSucessUpdate&"</li><br />"
         		response.redirect("update_user.asp?userid=" & request("userid") & "&success=SU&sc_lastname=" & request("sc_lastname") & "&sc_firstname=" & request("sc_firstname") & "&sc_orderby=" & request("sc_orderby") & "&groupid=" & request("groupid"))

      			if lcl_userhaspermission_edit_users then
        				strSuccess="<br /><img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='display_member.asp'>"&langBackToUserDisplay&"</a>"
      			else
        				strSuccess="<br /><img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='javascript:history.go(-1)'>" & langGoBack & "</a>"
       			end if
      			response.write "<script>document.all.goback.innerHTML="""&strSuccess&"""</script>"
		end select
%>
      </td>
  </tr>
</table>
  </div>
</div>
<!--#Include file="../admin_footer.asp"-->  
</body>
</html>
<%
'------------------------------------------------------------------------------
Sub SetClassSupervisor( iUserId, sIsClassSupervisor )
	Dim oCmd, sNewFlag

	if sIsClassSupervisor = "on" then
  		sNewFlag = 1
	else
  		sNewFlag = 0
	end if

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = "Update users set isclasssupervisor = " & sNewFlag & " where userid = " & iUserId
	oCmd.Execute
	Set oCmd = Nothing
End Sub 
%>
