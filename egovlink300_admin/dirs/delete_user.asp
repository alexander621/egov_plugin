<!-- #include file="../includes/common.asp" //-->
<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: delete_user.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2005 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes an admin user from the system.
'
' MODIFICATION HISTORY
' 1.0	??/??/????	???? - INITIAL VERSION
' 2.0	10/14/2011	Steve Loar - Changed from a hard delete to flagged as deleted
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iUserId, sSQL, sSQL2, sAdminName

If Trim(request("userid")) = "" Then 
	response.redirect("display_member.asp")
Else
	iUserId = CLng(request("userid"))
End If 

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit users" ) Then
  	response.redirect sLevel & "permissiondenied.asp"
End If 

sAdminName = GetAdminName( iUserId )	' in common.asp

' Flag them as deleted
sSql = "UPDATE users SET isdeleted = 1, deleteddate = GETDATE(), deletedbyuserid = " & session("UserID") & " WHERE userid = " & iUserId
sSql = sSql & " AND orgid = " & session("orgid")
'response.write sSql & "<br /><br />"

RunSQLStatement sSql	' in common.asp

'Delete all of the egov_staff_directory_usergroups assignments
sSQL  = "DELETE FROM egov_staff_directory_usergroups WHERE userid = " & iUserID
sSQL2 = "DELETE FROM usersgroups WHERE userid = " & iUserID
'response.write sSql & "<br /><br />"

RunSQLStatement sSQL	' in common.asp
RunSQLStatement sSQL2	' in common.asp
%>
<html>
<head>
	<title><%=langBSCommittees%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" class="menu">
		<tr>
			<td background="../images/back_main.jpg">
				<% ShowHeader sLevel %>
				<!--#Include file="../menu/menu.asp"-->
			</td>
		</tr>
	</table>

	<div id="content">
 		<div id="centercontent">

		<table border="0" cellpadding="10" cellspacing="0" width="100%">
			<tr>
				<td><font size="+1"><b>User Deletion Completed</b></font><br /><br /><br />
					<div id="goback" name="goback">
						<input type="button" value="<< Return To User List" class="button" onclick="location.href='display_member.asp';" />
					</div>
				</td>
				<td width="200">&nbsp;</td>
			</tr>
			<tr>
				<td>
					<p><font size="+1"><strong><%=sAdminName%> has been successfully deleted.</strong></font></p><br /><br />
				</td>
			</tr>
		</table>

		</div>
	</div>

	<!--#include file="../admin_footer.asp"-->  

</body>
</html>

