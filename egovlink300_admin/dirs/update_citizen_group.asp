<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->

<% 
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "groups" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>

<html>
<head>
	<title><%=langBSCommittees%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script src="../scripts/selectAll.js"></script>

	<script language="Javascript">
	  <!--

		function UpdateFamily( sUserId )
		{
			location.href='../dirs/family_members.asp?userid=' + sUserId;
		}
	//-->
	</script>

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
<!------------------------------------- -->

  
<table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center">&nbsp;<!--<img src="../images/icon_directory.jpg">--></td>
      <td><font size="+1"><b>Registration: Update Group Information</b></font>
	  <br><br><img src='../images/arrow_back.gif' align='absmiddle'> <a href="display_citizen_groups.asp">Back to Citizen Groups</a>
	  <br></td>
    </tr>
    <tr>
      <td valign="top">
        <!--#include file='quicklink_citizen.asp'-->      
      </td>
      <td colspan="2" valign="top">

<%
dim thisname,conn,rs,strSQL,rolename,title,groupname,rs2,strWhoCanView,strWhoCanEdit,strRoleNameList,groupdescription,newgroupname,newgroupdescription
thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 
rs.CursorType = 3

if request.form("start")="Y" then
	'==========================================================
	strSQL = "select groupname,groupdescription from citizengroups where groupid="&clng(trim(request.form("groupid")))
	rs.Open strSQL
	groupname=rs("groupname")
	groupdescription=rs("groupdescription")
	newgroupname=replace(trim(request.form("groupname")),"'","''")
	newgroupdescription=replace(trim(request.form("groupdescription")),"'","''")
	grouptype=replace(trim(request.form("grouptype")),"'","''")
	rs.close

	if trim(groupname)<>trim(newgroupname) then
		Title="<li>Directory name: <B>"&groupname&"</B> is sucessfully updated to <B>"&newgroupname&"</B></li>"
	end If

	if trim(groupdescription)<>trim(newgroupdescription) then
		Title=Title+"<li>Directory description is sucessfully updated!!</li>"
	else
		'if trim(groupname)=trim(newgroupname) then Title="<li>No changes have been made!!</li>"
		response.write " <LI> Changes have been saved. </LI> "
	end if

	strSQL = "update citizengroups set groupname='"&newgroupname&"',  groupdescription='"&newgroupdescription&"', grouptype='"&grouptype&"'  where groupid="&clng(trim(request.form("groupid")))
	conn.execute(strSQL)
	response.write "<br>"&Title
	'===========================================================
else 
	'===========================================================
	'-- check is the group id is entered or not ---------
	if trim(request.querystring("groupid"))="" then
		response.write "<br>No GroupID is entered, end program here"
		response.end
	else
		strSQL = "select *  from citizengroups where groupid="&clng(trim(request.querystring("groupid")))
	end if
	'-----------------
	rs.Open strSQL
	'-----------------
	if rs.recordcount=0 then
		response.write "<br>Cannot find the Directory name in database"
		response.write "<a href='javascript:> Go back</a>"
		response.end
	end if
	'---------------------------
	groupname=rs("groupname")
	Title="Update Directory:"&groupname
	'if not HasPermission("CanEditRoles") and not HasPermission("CanEdit"&groupname) then
	'response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleUpdateCommittee)
	'end if 
	'----------------
	set rs2 = Server.CreateObject("ADODB.Recordset")
	set rs2.ActiveConnection = conn
	rs2.CursorLocation = 3 
	rs2.CursorType = 3
	strSQL="exec getgroupViewlist "&clng(trim(request.querystring("groupid")))&","&Session("OrgID")
	rs2.open strSQL
	strWhoCanView="Viewable by:<B>&nbsp;"&rs2("GroupNameViewList")&"</B>"
	strWhoCanView=replace(strWhoCanView,",",", ")
	rs2.close
	strSQL="exec getgroupEditlist "&clng(trim(request.querystring("groupid")))&","&Session("OrgID")
	rs2.open strSQL
	strWhoCanEdit="Editable by:<B>&nbsp;"&rs2("GroupNameEditList")&"</B>"
	strWhoCanEdit=replace(strWhoCanEdit,",",", ")
	rs2.close
	strSQL="exec GetGroupRoleList "&clng(trim(request.querystring("groupid")))&","&Session("OrgID")
	rs2.open strSQL
	strRoleNameList="Existing Roles:<B>&nbsp;"&rs2("RoleNameList")&"</B>"
	strRoleNameList=replace(strRoleNameList,",",", ")
	rs2.close 
	'------------------------
	set rs2=nothing
	%>
	<FORM METHOD=POST ACTION="update_citizen_group.asp" name="UpdateCommittee">
	<div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.UpdateCommittee.submit();" onclick="return CheckCommitteeField();"><%=langUpdate%></a></div>


	<table border="0" width="100%"   class='tablelist' cellpadding='5' cellspacing='0'>
	 <tr><th colspan="2" width="100" align=left><%=langUpdate%>&nbsp;<%=langCommittee%></th><tr>  
		<tr>
		<td width="10%" valign="top"><%=langGroup%>:</td>
		<td width="80%"><input type=text name="GroupName" value="<%=rs("groupname")%>" size=50 maxlength=50></td>
		</tr>	          
		<tr>
		<td width="10%" valign="top"><%=langDescription%>:</td>
		<td width="80%"><textarea rows="2" cols="50" name="GroupDescription"><%=rs("GroupDescription")%></textarea></td>  
		</tr> 
		
	</table>
	<input type=hidden name="GroupID" value="<%=rs("groupid")%>">
	<input type=hidden name="start" value="Y">

	<div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.UpdateCommittee.submit();" onclick="return CheckCommitteeField();"><%=langUpdate%></a></div>
	</FORM>

	<!-- 
	<table border="0" width="100%"  class='tablelist' cellpadding='5' cellspacing='0'>
	<tr><th align=left colspan=2>&nbsp;&nbsp;Edit Directory Security and Roles</th></tr>
	<tr>
	<td width=110><img src="../images/newpermission.gif" border=0 align="absmiddle">&nbsp;<A HREF="javascript:doGroupsAccess()""><%=langEdit%> Security</A></td>
	<td><%=strWhoCanView%><br><%=strWhoCanEdit%></td>
	</tr>  
	<% 'if HasPermission("CanEditRoles") then %>
	<tr><td><%	response.write	"<img src='../images/newrole.gif' width='16' height='16' align='absmiddle'>&nbsp;&nbsp;"
		response.write	"<a  href=""javascript:openWin2('ManageCommitteeRoles.asp?groupid="&request.querystring("groupid")&"','_blank')"">"&langEdit&" "&langrole&"</a>" %>
	</td><td colspan=1>	<%=strRoleNameList%></td></tr>  
	<%'end if%>
	</table>-->

	<%
	'=================================================
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if %>


</td></tr></table>
 </div>
 </div>

<!--#Include file="../admin_footer.asp"-->  

<!--#include file='footer.asp'-->


<script language="JavaScript">
<!--
	function CheckCommitteeField()
	{
		if (document.UpdateCommittee.GroupName.value == "")
		{
			alert("Group name is required");
			document.UpdateCommittee.GroupName.focus();
		return false;				
		}					
		return true;
	}

    function doGroupsAccess() 
	{
      x = (screen.width-450)/2;
      y = (screen.height-400)/2;
      win = window.open("ManageCommitteeAccess2.asp?groupid=<%=Request("groupid")%>", "disc_members", "width=450,height=350,status=0,menubar=0,scrollbars=1,toolbar=0,left="+x+",top="+y+",z-lock=yes");
      win.focus();
    }

	function openWin2(url, name) 
	{
		popupWin = window.open(url, name,"resizable,width=380,height=300");
	}
//-->
</script>
