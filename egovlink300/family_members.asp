<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
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
' 1.0   01/14/2006	Steve Loar - INITIAL VERSION
' 1.1	12/26/2006	Steve Loar - Redirect to family structure page
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
%>
<html>
<head>

	<META NAME="ROBOTS" CONTENT="NOINDEX, NOFOLLOW">

	<%If iorgid = 7 Then %>
		<title><%=sOrgName%></title>
	<%Else%>
		<title>E-Gov Services <%=sOrgName%></title>
	<%End If
	Dim sReturnTo, sBackLang

	If session("ManageURL") <> "" Then
		sReturnTo = session("ManageURL")
		sBackLang = Session("ManageLang")
	Else
		If Session("RedirectPage") <> "" Then
			sReturnTo = Session("RedirectPage")
			sBackLang = Session("RedirectLang")
		Else
			sBackLang = "Back"
		End If 
	End If 

	%>

	<link rel="stylesheet" href="css/styles.css" type="text/css" />
	<link rel="stylesheet" href="global.css" type="text/css" />
	<link rel="stylesheet" href="css/style_<%=iorgid%>.css" type="text/css" />

	<script language="Javascript" src="scripts/modules.js"></script>
	<script language="Javascript" src="scripts/easyform.js"></script>

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
					alert("Please input a birth date for this child.");
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

	  //-->
	 </script>

</head>

<!--#Include file="include_top.asp"-->

<% Dim iUserId, iOrgId, sMessage, iFamilyCount

	iUserId = CLng(-1)

	iUserId = request.cookies("userid")
	If IsNumeric(iUserId) Then
		iUserId = CLng(iUserId)
	Else 
		iUserId = CLng(-1)
	End If 

	If CLng(iUserId) = CLng(-1) Then 
		response.redirect sEgovWebsiteURL
	End If 

	sMessage = ""
	iFamilyCount = 0

%>

<!--BEGIN PAGE CONTENT-->
<font class="datetagline">Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%> </font><br /><br />

<div id="content">

<br /><br /><a href="javascript:GoBack('<%=sReturnTo%>')"><img src="images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=sBackLang%></a><br /><br />

	<div class="reserveformtitle">Family Members of <%=GetUserName(iUserId)%></div>
	<div class="reserveforminputarea">

			<table border="1" cellpadding="5" cellspacing="0" width="100%">
			<tr><th>First Name</th><th>Last Name</th><th>Relation</th><th>Birthdate</th><th>Action</th></tr>
			<% 
				' Get the family members of the user
				iFamilyCount = GetFamilyMembers( iUserId )
			%>


			<form method="post" name="addFamily" action="family_members_update.asp">
			<input type="hidden" name="iuserid" value="<%=iUserId%>" />
			<input type="hidden" name="familymemberid" value="0" />
			<tr>
			<td align="center"><input type="text" name="firstname" value="" size="20" maxlength="50" /></td>
			<td align="center"><input type="text" name="lastname" value="<%=GetFamilyLastName(iUserId)%>" size="20" maxlength="50" /></td>
			<td align="center"><select name="relation" size="1">
					<option value="Spouse">Spouse</option>
					<option value="Child">Child</option>
					<option value="Sitter">Sitter</option>
					<option value="Parent">Parent</option>
					<option value="Yourself">Yourself</option>
				</select>
			</td>
			<td align="center"><input type="text" name="birthdate" value="" size="10" maxlength="10" /></td>
			<td align="center"><input type="button" name="add" value="Add" onclick="javascript:Validate(document.addFamily);" /></td></tr>
			</form>
			</table>
	</div>
</div>
<!--END: PAGE CONTENT-->


<!--SPACING CODE-->
<p><bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="include_bottom.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function GetFamilyMembers( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyMembers( ByVal iUserId )
	Dim sSql, iCount, sPreselected, oRs

	iCount = 0
	sSql = "SELECT familymemberid, firstname, lastname, birthdate, relationship FROM egov_familymembers "
	sSql = sSql & "WHERE belongstoUserid = " & iUserId & " ORDER BY birthdate ASC"
	session("familysql") = sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	'session("familysql") = ""

	Do While Not oRs.EOF 
		iCount = iCount + 1
		response.write vbcrlf & "<form method=""post"" name=""FamilyMember" & iCount & """ action=""family_members_update.asp"">"
		response.write vbcrlf & "<input type=""hidden"" name=""iuserid"" value=""" & iUserId & """ />"
		response.write vbcrlf & "<input type=""hidden"" name=""familymemberid"" value=""" & oRs("familymemberid") & """ />"

		response.write vbcrlf & "<tr><td align=""center""><input type=""text"" name=""firstname"" value=""" & oRs("firstname") & """ size=""20"" maxlength=""50"" /></td>"
		response.write vbcrlf & "<td align=""center""><input type=""text"" name=""lastname"" value=""" & oRs("lastname") & """ size=""20"" maxlength=""50"" /></td>"
				
		response.write vbcrlf & "<td align=""center"">"
		response.write vbcrlf & "<select name=""relation"" size=""1"">" 
		response.write vbcrlf & "<option value=""Spouse"""
		If oRs("relationship") = "Spouse" Then
			response.write " selected=""selected"" "
		End If 
		response.write ">Spouse</option>"

		response.write vbcrlf & "<option value=""Child"""
		If oRs("relationship") = "Child" Then
			response.write " selected=""selected"" "
		End If 
		response.write ">Child</option>"

		response.write vbcrlf & "<option value=""Sitter"""
		If oRs("relationship") = "Sitter" Then
			response.write " selected=""selected"" "
		End If 
		response.write ">Sitter</option>"

		response.write vbcrlf & "<option value=""Parent"""
		If oRs("relationship") = "Parent" Then
			response.write " selected=""selected"" "
		End If 
		response.write ">Parent</option>"
		
		response.write vbcrlf & "<option value=""Yourself"""
		If oRs("relationship") = "Yourself" Then
			response.write " selected=""selected"" "
		End If 
		response.write ">Yourself</option>"
		
		response.write vbcrlf & "</select></td>"

		response.write vbcrlf & "<td align=""center""><input type=""text"" name=""birthdate"" value=""" & oRs("birthdate") & """ size=""10"" maxlength=""10"" /></td>"
		response.write vbcrlf & "<td align=""center"">"
		response.write vbcrlf & "<input type=""button"" name=""update"" value=""Update"" onclick=""javascript:Validate(document.FamilyMember" & iCount & ");"" />"
		response.write vbcrlf & "<input type=""button"" name=""delete"" value=""Delete"" onclick=""javascript:ConfirmDelete('" & oRs("firstname") & "', '" & oRs("lastname") & "'," & oRs("familymemberid") & "," & iUserId & ");"" />"
		response.write vbcrlf & "</td></tr></form>"
		oRs.movenext
	Loop 
		
	oRs.Close
	Set oRs = Nothing

	session("familysql") = sSql

	GetFamilyMembers = iCount

End Function  


'--------------------------------------------------------------------------------------------------
' Function GetUserName( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetUserName( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT userfname, userlname FROM egov_users WHERE userid = " & iUserId & " AND orgid = " & iOrgId
	session("familysql") = sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	'session("familysql") = ""

	If Not oRs.EOF Then 
		GetUserName = oRs("userfname") & " " & oRs("userlname")
	Else
		GetUserName = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetFamilyLastName( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyLastName( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT userlname FROM egov_users WHERE userid = " & iUserId & " AND orgid = " & iOrgId
	session("familysql") = sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	'session("familysql") = ""

	If Not oRs.EOF Then 
		GetFamilyLastName = oRs("userlname")
	Else
		GetFamilyLastName = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 



%>