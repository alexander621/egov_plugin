<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: edit_user_security.asp
' AUTHOR: Steve Loar
' CREATED: 09/20/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Displays the screen for selecting features that a user gets permission to access.
'
' MODIFICATION HISTORY
' 1.0	09/28/2006	Steve Loar - Initial version completed
' 1.1	08/11/2009	David Boyer - Added screen messages.
' 1.2	03/03/2011	Steve Loar - Modified to show message from copy security script
' 1.3	04/17/2013	Steve Loar - Changed user drop down to only show active admin users
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iUserID, bIsRootAdmin, sShowDetails, sLoadMsg, lcl_success, lcl_msg

sLevel = "../"  'Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "user permission", sLevel	' In common.asp

If request("iUserId") = "" Then 
	iUserID = Session("UserID")
Else 
	iUserID = request("iUserId")
End If 

bIsRootAdmin = UserIsRootAdmin(session("userid"))

'Check for a screen message
lcl_onload = ""
lcl_success = request("success")

If lcl_success <> "" Then 
	lcl_msg = setupScreenMsg(lcl_success)
	lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
End If
 
If request("s") = "u" Then
	lcl_onload = "displayScreenMsg('User Permissions Were Successfully Copied.');"
End If

%>
<html>
<head>
	 <title>E-GovLink Administration Console {User Permissions}</title>

	 <link rel="stylesheet" type="text/css" href="../global.css" />
	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="security.css" />

	<script language="javascript">
	<!--

	function ShowDetails()
	{
		if (UserForm.showdetails.value == '')
		{
			UserForm.showdetails.value = 'checkit';
		}
		else
		{
			UserForm.showdetails.value = '';
		}
		alert( UserForm.iUserID.value );
		UserForm.submit();
	}

	function CopyUser( iUserID )
	{
		location.href = 'copy_user_security.asp?userid=' + iUserID;
	}

	function displayScreenMsg(iMsg) 
	{
		if(iMsg!="") 
		{
			document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
			window.setTimeout("clearScreenMsg()", (10 * 1000));
		}
	}

	function clearScreenMsg() 
	{
		document.getElementById("screenMsg").innerHTML = "";
	}

	function submitForm()
	{
		document.UserForm.submit();
	}

	//-->
	</script>

</head>
<body onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

		<table border="0" cellspacing="0" cellpadding="0" width="1000px">
			<tr valign="top">
				<td>
					<font size="+1"><strong>User Permissions</strong></font><br /><br />

					<form name="UserForm" method="post" action="edit_user_security.asp">
						<input type="hidden" name="showdetails" id="showdetails" value="<%=sShowDetails%>" />

						<label for="userid">User Name:</label> <% ShowAdminUserPicks session("orgid"), iUserID, bIsRootAdmin %>
					</form>
				</td>
				<td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
			</tr>
		</table>

		<form name="SecurityForm" method="post" action="updatesecurity.asp">
			<input type="hidden" name="iUserId" id="iUserId" value="<%=iUserId%>" />
			<input type="hidden" name="bIsRootAdmin" id="bIsRootAdmin" value="<%=bIsRootAdmin%>" />
<%
			displayButtons

			displayFeatureList session("orgid"), iUserID, bIsRootAdmin

			displayButtons
%>
		</form>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%
'------------------------------------------------------------------------------
Sub ShowAdminUserPicks( ByVal iOrgID, ByVal iUserID, ByVal bIsRootAdmin )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname FROM users "
	sSql = sSql & "WHERE orgid = " & iOrgID & " AND isdeleted = 0"

	If Not bIsRootAdmin Then 
		sSql = sSql & " AND (isrootadmin IS NULL OR isrootadmin = 0) "
	End If 

	sSql = sSql & " ORDER BY lastname, firstname"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oRs.EOF Then 
		response.write "<select id=""userid"" name=""iUserID"" onchange=""submitForm();"">" & vbcrlf

		Do While Not oRs.EOF
			If clng(iUserID) = clng(oRs("userid")) Then 
				lcl_selected_user = " selected=""selected"""
			Else 
				lcl_selected_user = ""
			End If 

			response.write "<option value=""" & oRs("userid") & """" & lcl_selected_user & ">" & oRs("lastname") & ", " & oRs("firstname") & "</option>" & vbcrlf

			oRs.MoveNext
		Loop 
		response.write "</select>" & vbcrlf
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
Sub displayFeatureList( ByVal iOrgID, ByVal iUserID, ByVal bIsRootAdmin )
	Dim sChecked, sSql, oRs

	sChecked = ""

	'Get the features that the organization has
	sSql = "SELECT F.featureid, P.permission, F.featurename, F.haspermissionlevels "
	sSql = sSql & " FROM egov_organization_features F, egov_feature_permissions P, "
	sSql = sSql & " egov_features_to_permissions FP, egov_organizations_to_features FO "
	sSql = sSql & " WHERE F.parentfeatureid = 0 "
	sSql = sSql & " AND F.haspermissions = 1 "
	sSql = sSql & " AND FP.permissionid = P.permissionid "
	sSql = sSql & " AND FP.featureid = F.featureID "
	sSql = sSql & " AND FO.featureid = F.featureid "
	sSql = sSql & " AND FO.orgid = " & iOrgID

	If Not bIsRootAdmin Then 
		sSql = sSql & " AND rootadminrequired = 0 "
	End If 

	sSql = sSql & " ORDER BY F.admindisplayorder, F.featurename"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oRs.EOF Then 

		response.write "<div class=""shadow"">" & vbcrlf
		response.write "<table id=""securitydisplay"" cellspacing=""0"" cellpadding=""5"" border=""0"">" & vbcrlf
		response.write "  <tr>" & vbcrlf
		response.write "      <th>Feature</th>" & vbcrlf
		response.write "      <th>Permission</th>" & vbcrlf
		response.write "      <th>Permission Level</th>" & vbcrlf
		response.write "  </tr>" & vbcrlf

		Do While Not oRs.EOF
			response.write "<tr>" & vbcrlf
			response.write "<td class=""featureparent"">" & oRs("FeatureName") & "</td>" & vbcrlf

			'View Display Check
			If lCase(oRs("permission")) = "permission" Then 
				sChecked = GetPermissionSetValue(UserHasPermissionToFeature(oRs("featureid"),iUserID,"permission"))

				response.write "<td class=""featureparent"" align=""center""><input type=""checkbox"" name=""viewPermission"" value=""" & oRs("featureid") & """" & sChecked & " /></td>" & vbcrlf

				sChecked = ""
			Else 
				response.write "<td class=""featureparent"" align=""center"">&nbsp;</td>" & vbcrlf
			End If 

			response.write "<td class=""featureparent"" align=""center"">&nbsp;</td>" & vbcrlf
			response.write "</tr>" & vbcrlf

			GetChildrenFeatureList iOrgID, oRs("Featureid"), iUserID, bIsRootAdmin

			oRs.MoveNext
		Loop 

		response.write "</table>" & vbcrlf
		response.write "</div>" & vbcrlf

	End If 

	oRs.Close
	set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
Sub GetChildrenFeatureList( ByVal iOrgID, ByVal iParentID, ByVal iUserID, ByVal bIsRootAdmin )
	Dim sClass, sSql

	'Get the child features that the organization has
	sSql = "SELECT F.featureid, P.permission, F.featurename, F.haspermissionlevels "
	sSql = sSql & "FROM egov_organization_features F, egov_feature_permissions P, "
	sSql = sSql & "egov_features_to_permissions FP, egov_organizations_to_features FO "
	sSql = sSql & "WHERE F.haspermissions = 1 "
	sSql = sSql & "AND parentfeatureid = " & iParentID
	sSql = sSql & " AND FP.permissionid = P.permissionid "
	sSql = sSql & "AND FP.featureid = F.featureID "
	sSql = sSql & "AND FO.featureid = F.featureid "
	sSql = sSql & "AND FO.orgid = " & iOrgID
	'response.write sSQL

	If Not bIsRootAdmin Then 
		sSql = sSql & " AND rootadminrequired = 0 "
	End If 

	sSql = sSql & " ORDER BY securitydisplayorder, F.featurename"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oRs.EOF Then 
		sClass = "oddrow"

		Do While Not oRs.EOF
			response.write "<tr>" & vbcrlf
			response.write "<td class=""" & sClass & """>&nbsp;&nbsp;&nbsp;&nbsp; &bull; " & oRs("FeatureName") & "</td>" & vbcrlf

			'View Display Check
			If lCase(oRs("permission")) = "permission" Then 
				sChecked = GetPermissionSetValue(UserHasPermissionToFeature(oRs("featureid"), iUserID, "permission"))

				response.write "<td class=""" & sClass & """ align=""center""><input type=""checkbox"" name=""viewPermission"" value=""" & oRs("featureid") & """" & sChecked & " /></td>" & vbcrlf

				sChecked = ""
			Else 
				response.write "<td class=""" & sClass & """ align=""center"">&nbsp;</td>" & vbcrlf
			End If   


			'ODA level CHECK
			If oRs("haspermissionlevels") Then 
				response.write "<td class=""" & sClass & """ align=""center"">" & vbcrlf
				DisplayPermissionLevels GetUserPermissionLevel(oRs("featureid"), iuserid, "permission"), "edit_oda" & oRs("featureid")
				response.write "</td>" & vbcrlf
			Else 
				response.write "<td class=""" & sClass & """ align=""center"">&nbsp;</td>" & vbcrlf
			End If 

			response.write "</tr>" & vbcrlf

			oRs.MoveNext

			If sClass = "oddrow" Then 
				sClass = "evenrow"
			Else 
				sClass = "oddrow"
			End If 
		Loop 
	End If 

End Sub 


'------------------------------------------------------------------------------
Sub DisplayPermissionLevels( ByVal iSelected, ByVal sName )
	Dim oRs, sSql

	sSql = "SELECT * FROM egov_feature_permission_levels ORDER BY permissionlevelid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oRs.EOF Then 
		response.write "<select name=""" & sName & """>" & vbcrlf

		Do While Not oRs.EOF
			If Clng(iSelected) = oRs("permissionlevelid") Then 
				sSelected = " selected=""selected"""
			Else 
				sSelected = ""
			End If 
			response.write "<option value=""" & oRs("permissionlevelid") & """" & sSelected & ">" & oRs("permissionlevel") & "</option>" & vbcrlf
			oRs.MoveNext
		Loop 
		response.write "</select>" & vbcrlf
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
Function UserHasPermissionToFeature( ByVal iFeatureID, ByVal iUserID, ByVal sPermission )
	Dim sSql, oRs, iPermissionID

	iPermissionID = GetPermissionId( sPermission )

	UserHasPermissionToFeature = False

	sSql = "SELECT permissionid "
	sSql = sSql & " FROM egov_users_to_features "
	sSql = sSql & " WHERE userid = " & iUserID
	sSql = sSql & " AND featureid = " & iFeatureID
	sSql = sSql & " AND permissionid = " & iPermissionID 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oRs.EOF Then 
		UserHasPermissionToFeature = True
	End If 

End Function 


'------------------------------------------------------------------------------
Function GetPermissionId( ByVal sPermission )
	Dim sSql, oRs

	sSql = "SELECT permissionid "
	sSql = sSql & " FROM egov_feature_permissions "
	sSql = sSql & " WHERE permission = '" & sPermission & "' "

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oRs.EOF Then 
		GetPermissionId = oRs("permissionid")
	End If 

	oRs.Close
	set oRs = Nothing 

End Function


'------------------------------------------------------------------------------
Function GetPermissionSetValue( ByVal blnValue )
	
	sReturnValue = ""

	If blnValue Then 
		sReturnValue = " checked=""checked"" "
	End If 

	GetPermissionSetValue = sReturnValue

End Function


'------------------------------------------------------------------------------
Function GetUserPermissionLevel( ByVal iFeatureID, ByVal iUserID, ByVal sPermission )
	Dim sSql, oRs

	GetUserPermissionLevel = 0

	sSql = "SELECT isnull(F.permissionlevelid,0) as permissionlevelid "
	sSql = sSql & " FROM egov_users_to_features F, egov_feature_permissions P "
	sSql = sSql & " WHERE F.userid = " & iUserID
	sSql = sSql & " AND F.featureid = " & iFeatureID 
	sSql = sSql & " AND F.permissionid = P.permissionid "
	sSql = sSql & " AND P.permission = '" & sPermission & "' "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oRs.EOF Then 
		GetUserPermissionLevel = oRs("permissionlevelid")
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'------------------------------------------------------------------------------
Sub displayButtons()

	response.write "<p>" & vbcrlf
	response.write "<input class=""button"" type=""button"" value=""Copy Permissions"" onClick=""javascript: CopyUser(" & iUserID & ");"" />" & vbcrlf
	response.write "<input class=""button"" type=""submit"" value=""Save Changes"" />" & vbcrlf
	response.write "</p>" & vbcrlf

End Sub


'------------------------------------------------------------------------------
Function setupScreenMsg( ByVal iSuccess )
	Dim lcl_return

	lcl_return = ""

	If iSuccess <> "" then
		iSuccess = UCase(iSuccess)

		If iSuccess = "SU" Then 
			lcl_return = "Successfully Updated..."
		ElseIf iSuccess = "SA" Then 
			lcl_return = "Successfully Created..."
		ElseIf iSuccess = "SR" Then 
			lcl_return = "Successfully Reordered..."
		ElseIf iSuccess = "SD" Then 
			lcl_return = "Successfully Deleted..."
		End If 
	End If 

	setupScreenMsg = lcl_return

End Function


%>
