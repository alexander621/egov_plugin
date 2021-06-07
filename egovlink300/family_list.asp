<!DOCTYPE html>
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: family_list.asp
' AUTHOR: Steve Loar
' CREATED: 12/27/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This script lists family members.
'
' MODIFICATION HISTORY
' 1.0   12/27/2006	Steve Loar - INITIAL VERSION
' 1.2	07/25/2008	Steve Loar - Changed to use deleted flag
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

%>
<html lang="en">
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<meta charset="UTF-8">
<%
If iorgid = 7 Then 
%>
	<title><%=sOrgName%></title>
<%
Else
%>
	<title>E-Gov Services <%=sOrgName%></title>
<%
End If
%>


<link rel="stylesheet" href="css/styles.css" />
<link rel="stylesheet" href="global.css" />
<link rel="stylesheet" href="css/style_<%=iorgid%>.css" />

<script src="scripts/jquery-1.7.2.min.js"></script>

<script src="scripts/modules.js"></script>
<script src="scripts/easyform.js"></script>

<script>
<!--

	function goBack( returnToURL )
	{
		if (returnToURL != "")
		{
			location.href=returnToURL;
		}
		else
		{
			history.go(-1);
		}
	}

	function deleteFamilyMember( userId, name, deletedById ) 
	{
		var msg = "Are you sure you want to delete " + name + "?"
		if (confirm(msg))
		{
			//alert(deletedById);
			location.href='delete_family_member.asp?iUserId=' + userId + '&deletedbyid=' + deletedById;
		}
	}

	function editFamilyMember( userId )
	{
		location.href='manage_family_member.asp?u=' + userId;
	}

	function createNewFamilyMember()
	{
		location.href="manage_family_member.asp?u=0";
	}

  //-->
 </script>

</head>

<!--#Include file="include_top.asp"-->
<!-- #include file="./class/classFamily.asp" //-->

<% 
Dim iUserId, iOrgId, sMessage, iFamilyCount, iFamilyId, oFamily

Set oFamily  = New classFamily

If request.cookies("userid") <> "" Then 
	iUserId = CLng(request.cookies("userid"))
Else
	If request("userid") <> "" Then 
		iUserId = CLng(request("userid"))
	Else 
		response.redirect "manage_account.asp"
	End if
End If 

'response.write iUserId & "<br /><br />"
iFamilyId    = oFamily.GetFamilyId( iUserId )
'response.write session("GetFamilyIdSql") & "<br /><br />"
sMessage     = ""
iFamilyCount = 0
Set oFamily  = Nothing 

%>

<!--BEGIN PAGE CONTENT-->
<%	RegisteredUserDisplay( "" ) %>

<div id="content">
<%	If Session("RedirectPage") <> "" Then %>
		<br /><br /><a href="<%=Session("RedirectPage")%>"><img src="images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=Session("RedirectLang")%></a><br /><br />
<%	Else %>
		<br /><br /><a href="javascript:goBack('manage_account.asp')"><img src="images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Return to Manage Account</a><br /><br />
<%	End If %>

<p>
	<!-- <a href="manage_family_member.asp?u=0">Create New Family Member</a> -->
	<input type="button" class="button" value="Create New Family Member" onclick="createNewFamilyMember();" />
</p>
	<div class="reserveformtitle">Family of <%=GetUserName( iUserId )%></div>
	<div class="reserveforminputarea">

		<table border="1" cellpadding="5" cellspacing="0" width="100%">
			<tr><th>Name</th><th>Relation</th><th>Birthdate</th><th>Actions</th></tr>
<% 
			' Get the family members of the user
			iFamilyCount = GetFamilyMembers( iFamilyId, iUserId )
%>
		</table>
	</div>
</div>
<!--END: PAGE CONTENT-->


<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="include_bottom.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' integer GetFamilyMembers( iFamilyId, iUserId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyMembers( ByVal iFamilyId, ByVal iUserId )
	Dim sSql, iCount, sPreselected, oRs

	iCount = 0
	sSql = "SELECT userid, userfname, userlname, birthdate, relationship, selftag "
	sSql = sSql & " FROM egov_users U, egov_familymember_relationships F WHERE U.relationshipid = F.relationshipid"
	sSql = sSql & " AND U.isdeleted = 0 AND familyid = " & iFamilyId & " ORDER BY birthdate ASC"

	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	Do While Not oRs.EOF 
		iCount = iCount + 1
		response.write vbcrlf & "<tr><td>" & oRs("userfname") & " " & oRs("userlname") & "</td>"
		If CLng(iUserId) = CLng(oRs("userid")) Then 
			' What they call themselves
			response.write vbcrlf & "<td align=""center"">" & oRs("selftag") & "</td>"
		Else 
			' what they call others
			response.write vbcrlf & "<td align=""center"">" & oRs("relationship") & "</td>"
		End If 
		If IsNull(oRs("birthdate")) Then 
			response.write vbcrlf & "<td>&nbsp;</td>"
		Else 
			response.write vbcrlf & "<td align=""center"">" & oRs("birthdate") & "</td>"
		End If 
		response.write "<td align=""center""><input type=""button"" class=""button"" value=""Edit"" onClick=""editFamilyMember(" & oRs("userid") & ");"" />"
		If CLng(iUserId) <> CLng(oRs("userid")) Then
			' do not want you to be able to delete yourself
			response.write " &nbsp; <input type=""button"" class=""button"" value=""Delete"" onClick=""deleteFamilyMember( " & oRs("userid") & ",'" & oRs("userfname") & " " & oRs("userlname") & "', " & iUserId & " );"" /> "
		End If 
		response.write "</td></tr>"
		oRs.MoveNext
	Loop 
		
	oRs.Close
	Set oRs = Nothing

	GetFamilyMembers = iCount

End Function  


'--------------------------------------------------------------------------------------------------
' string GetUserName( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetUserName( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname "
	sSql = sSql & "FROM egov_users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	GetUserName = Trim(oRs("userfname") & " " & oRs("userlname"))

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetFamilyLastName( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyLastName( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT userlname FROM egov_users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	GetFamilyLastName = oRs("userlname")

	oRs.Close
	Set oRs = Nothing

End Function 



%>
