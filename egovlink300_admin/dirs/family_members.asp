<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: family_members.asp
' AUTHOR: Steve Loar
' CREATED: 02/14/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This script allows the ading and updating of family member information
'
' MODIFICATION HISTORY
' 1.0   01/14/06	Steve Loar - INITIAL VERSION
' 1.1	10/10/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit citizens" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 
%>



<html>
<head>

<title>E-Gov Family Members</title>

<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
<link rel="stylesheet" type="text/css" href="../global.css" />
<link rel="stylesheet" type="text/css" href="family_members.css" />

<script language="Javascript">
<!--
	function GoBack(ReturnToURL)
	{
		if (ReturnToURL != "")
		{
			location.href=ReturnToURL;
		}
		else
		{
			history.go(-1);
		}
	}

	function Validate(inForm)
	{
		var rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
		var Ok = rege.test(inForm.birthdate.value);

		if (inForm.firstname.value == "")
		{
			alert("Please input a first name.");
			inForm.firstname.focus();
			return;
		}
		
		if (inForm.lastname.value == "")
		{
			alert("Please input a last name.");
			inForm.lastname.focus();
			return;
		}

		if (inForm.relation.value == 'Child')
		{
			if (inForm.birthdate.value == "")
			{
				alert("Please input a birth date for the child.");
				inForm.birthdate.focus();
				return;
			}
			if (! Ok)
			{
				alert("Birth date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
				inForm.birthdate.focus();
				return;
			}
		}
		else
		{
			if (inForm.birthdate.value != "")
			{
				if (! Ok)
				{
				alert("Birth date should be in the format of MM/DD/YYYY.  \nPlease enter it again, or leave it blank.");
				inForm.birthdate.focus();
				return;
				}
			}
		}

		inForm.submit();

	}

	function ConfirmDelete(sFirstName,sLastName,iFamilyMemberId, iUserId) 
	{
		var msg = "Do you wish to delete " + sFirstName + " " + sLastName + "?"
		if (confirm(msg))
		{
			location.href='family_member_delete.asp?iFamilyMemberId=' + iFamilyMemberId + '&iUserId=' + iUserId;
		}
	}

	function SetFocus()
	{
		document.addFamily.firstname.focus();
	}

  //-->
 </script>


</head>

<body onload="javascript:SetFocus();">

<%'DrawTabs tabRegistration,2 %>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<% Dim iUserId, sMessage, iFamilyCount

	iUserId = request("userid")
	sMessage = ""
	iFamilyCount = 0

%>


<!--BEGIN PAGE CONTENT-->
<!--<font class=datetagline>Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%> </font><br /><br />-->

<div id="content">
	<div id="centercontent">

	<p><strong>Family Members of <%=GetUserName(iUserId)%></strong></p>
	<a href="javascript:GoBack('<%=Session("RedirectPage")%>')"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=Session("RedirectLang")%></a><br /><br />

	<div class="shadow">
		<table border="0" cellpadding="6" cellspacing="0" class="tableadmin">
			<tr><th>First Name</th><th>Last Name</th><th>Relation</th><th>Birthdate</th><th>Action</th></tr>
			<% 		' GetPurchaserInfo(iUserId)
				iFamilyCount = GetFamilyMembers(iUserId)
			%>
			<tr>
				<form method="post" name="addFamily" action="family_members_update.asp">
					<input type="hidden" name="iuserid" value="<%=iUserId%>" />
					<input type="hidden" name="familymemberid" value="0" />
				<td align="center">
					<input type="text" name="firstname" value="" size="20" maxlength="50" />
				</td>
				<td align="center">
					<input type="text" name="lastname" value="<%=GetFamilyLastName(iUserId)%>" size="20" maxlength="50" />
				</td>
				<td align="center">
					<select name="relation" size="1">
						<option value="Spouse">Spouse</option>
						<option value="Child">Child</option>
						<option value="Sitter">Sitter</option>
						<option value="Parent">Parent</option>
						<option value="Yourself">Yourself</option>
					</select>
				</td>
				<td align="center"><input type="text" name="birthdate" value="" size="10" maxlength="10" /></td>
				<td align="center"><input type="button" class="button" name="add" value="Add" onclick="javascript:Validate(document.addFamily);" /></td>
				</form>
			</tr>
		</table>
	</div>
	
	</div>
</div>
<!--END: PAGE CONTENT-->


<!--SPACING CODE-->
<!--<p><bR>&nbsp;<bR>&nbsp;</p>-->
<!--SPACING CODE-->


<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function GetFamilyMembers( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyMembers( iUserId )
	Dim sSQL, iCount, sPreselected

	iCount = 0
	sSQL = "Select familymemberid, firstname, lastname, birthdate, relationship FROM egov_familymembers WHERE belongstouserid = " & iUserId & " order by birthdate asc"

	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open sSQL, Application("DSN"), 0, 1

	Do While Not oUser.eof 
		iCount = iCount + 1
		response.write vbcrlf & "<form method=""post"" name=""FamilyMember" & iCount & """ action=""family_members_update.asp"">"
		response.write vbcrlf & "<input type=""hidden"" name=""iuserid"" value=""" & iUserId & """ />"
		response.write vbcrlf & "<input type=""hidden"" name=""familymemberid"" value=""" & oUser("familymemberid") & """ />"

		If iCount Mod 2 = 1 Then
			response.write vbcrlf & "<tr class=""alt_row"">"
		Else
			response.write vbcrlf & "<tr>"
		End If
		
		response.write vbcrlf & "<td align=""center""><input type=""text"" name=""firstname"" value=""" & oUser("firstname") & """ size=""20"" maxlength=""50"" /></td>"
		response.write vbcrlf & "<td align=""center""><input type=""text"" name=""lastname"" value=""" & oUser("lastname") & """ size=""20"" maxlength=""50"" /></td>"
				
		response.write vbcrlf & "<td align=""center"">"
		response.write vbcrlf & "<select name=""relation"" size=""1"">" 
		response.write vbcrlf & "<option value=""Spouse"""
		If oUser("relationship") = "Spouse" Then
			response.write " selected=""selected"" "
		End If 
		response.write ">Spouse</option>"

		response.write vbcrlf & "<option value=""Child"""
		If oUser("relationship") = "Child" Then
			response.write " selected=""selected"" "
		End If 
		response.write ">Child</option>"

		response.write vbcrlf & "<option value=""Sitter"""
		If oUser("relationship") = "Sitter" Then
			response.write " selected=""selected"" "
		End If 
		response.write ">Sitter</option>"

		response.write vbcrlf & "<option value=""Parent"""
		If oUser("relationship") = "Parent" Then
			response.write " selected=""selected"" "
		End If 
		response.write ">Parent</option>"
		
		response.write vbcrlf & "<option value=""Yourself"""
		If oUser("relationship") = "Yourself" Then
			response.write " selected=""selected"" "
		End If 
		response.write ">Yourself</option>"
		
		response.write vbcrlf & "</select></td>"

		response.write vbcrlf & "<td align=""center""><input type=""text"" name=""birthdate"" value=""" & oUser("birthdate") & """ size=""10"" maxlength=""10"" /></td>"
		response.write vbcrlf & "<td align=""center"">"
		response.write vbcrlf & "<input type=""button"" name=""update"" value=""Update"" onclick=""javascript:Validate(document.FamilyMember" & iCount & ");"" />"
		response.write vbcrlf & "<input type=""button"" name=""delete"" value=""Delete"" onclick=""javascript:ConfirmDelete('" & JavascriptSafe(oUser("firstname")) & "', '" & JavascriptSafe(oUser("lastname")) & "'," & oUser("familymemberid") & "," & iUserId & ");"" />"
		response.write vbcrlf & "</td></tr></form>"
		oUser.movenext
	Loop 
		
	oUser.close
	Set oUser = Nothing
	GetFamilyMembers = iCount
End Function  


'--------------------------------------------------------------------------------------------------
' Function GetUserName(iUserId)
'--------------------------------------------------------------------------------------------------
Function GetUserName(iUserId)
	Dim sSql 

	sSql = "Select userfname, userlname from egov_users where userid = " & iUserId 
	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open sSQL, Application("DSN"), 0, 1

	GetUserName = oUser("userfname") & " " & oUser("userlname")

	oUser.close
	Set oUser = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function JavascriptSafe( strDB )
'--------------------------------------------------------------------------------------------------
Function JavascriptSafe( strDB )
  If Not VarType( strDB ) = vbString Then JavascriptSafe = strDB : Exit Function
  JavascriptSafe = Replace( strDB, "'", "\'" )
End Function


'--------------------------------------------------------------------------------------------------
' Function GetFamilyLastName( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyLastName( iUserId )
	Dim sSql, oUser

	sSql = "Select userlname from egov_users where userid = " & iUserId 

	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open sSQL, Application("DSN"), 0, 1

	GetFamilyLastName = oUser("userlname")

	oUser.close
	Set oUser = Nothing
End Function 



%>