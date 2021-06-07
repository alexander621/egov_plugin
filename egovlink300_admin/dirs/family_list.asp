<!DOCTYPE html>
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
' 1.2	1/5/2009	Steve Loar - Added page availability and user rights call
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sLoadMsg, hasCitizenAccounts

sLevel = "../" ' Override of value from common.asp

hasCitizenAccounts = OrgHasFeature("citizen accounts")

' Check the page availability and user access rights in one call
PageDisplayCheck "edit citizens", sLevel	' In common.asp

If request("status") <> "" Then
	If request("status") = "deletesuccess" Then
		sLoadMsg = "displayScreenMsg('The Family Member Was Successfully Deleted');"
	End If
	If request("status") = "deletefail" Then
		sLoadMsg = "displayScreenMsg('The Family Member Could Not be Deleted Due To An Account Balance');"
	End If 
End If 

%>
<html lang="en">
<head>
	<meta charset="UTF-8">

	<title>E-Gov Family Members</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="family_members.css" />
	
	<script src="../scripts/jquery-1.7.2.min.js"></script>

	<script>
	<!--
		GoBack = function(ReturnToURL) {
			if (ReturnToURL != "")
			{
				location.href=ReturnToURL;
			}
			else
			{
				history.go(-1);
			}
		};


		ConfirmDelete = function(sFirstName,sLastName,iFamilyMemberId, iUserId) {
			var msg = "Do you wish to delete " + sFirstName + " " + sLastName + "?"
			<% If OrgHasFeature("citizen accounts") Then %>
				msg += "\n\nIf a family member has an account balance, they will not be deleted.";
			<% End If %>

			if (confirm(msg))
			{
				location.href='family_member_delete.asp?iFamilyMemberId=' + iFamilyMemberId + '&iUserId=' + iUserId;
			}
		};

		DeleteFamilyMember = function(iUserId, sName, iReturn) {
			var msg = "Do you wish to delete " + sName + "?"
			<% If OrgHasFeature("citizen accounts") Then %>
				//msg += "\n\nIf a family member has an account balance, they will not be deleted.";
			<% End If %>
			
			if (confirm(msg))
			{
				location.href='delete_family_member.asp?iUserId=' + iUserId + '&iReturn=' + iReturn;
			}
		};

		EditFamilyMember = function( iUserId, iReturnTo ) {
			location.href='manage_family_member.asp?u=' + iUserId + '&iReturn=' + iReturnTo;
		}

		EditCitizen = function( iUserId ) {
			location.href='update_citizen.asp?userid=' + iUserId;
		};
		
		newFamilyMember = function( iUserId ) {
			location.href='manage_family_member.asp?u=0&iReturn=' + iUserId;
		}
		
		displayScreenMsg = function( iMsg ) 
		{
			if( iMsg != "" ) 
			{
				$("#screenMsg").html( "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;" );
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		};

		clearScreenMsg = function() 
		{
			$("#screenMsg").html("");
		};
		
		SetUpPage = function()
		{
			<%=sLoadMsg%>
		};

	  //-->
	 </script>

</head>
<body onload="SetUpPage();">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<% Dim iUserId, sMessage, iFamilyCount, iFamilyId

	iUserId = request("userid")
	iFamilyId = GetFamilyId( iUserId )
	sMessage = ""
	iFamilyCount = 0

%>


<!--BEGIN PAGE CONTENT-->
<!--<font class=datetagline>Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%> </font><br /><br />-->

<div id="content">
	<div id="centercontent">

		<p id="title">
			Family Members of <%=GetUserName(iUserId)%>
		</p>
		
		<p>
			<span id="screenMsg"></span>
			<input type="button" class="button" value="<< <%=Session("RedirectLang")%>" onclick="javascript:GoBack('<%=Session("RedirectPage")%>')" />
		</p>
		
		<p>
			<input type="button" class="button" value="Add a New Family Member" onclick="newFamilyMember( <%=iUserId%> )" />
		</p>

		<table border="0" cellpadding="6" cellspacing="0" class="tableadmin" id="family_table">
			<tr><th>Name</th><th align="center">Relation</th><th align="center">Age</th>
				<% If hasCitizenAccounts Then
					response.write "<th align=""center"">Account<br />Balance</th>"
				End If  %>
				<th align="center" colspan="2">Action</th></tr>
			<% 		
				iFamilyCount = GetFamilyMembers( iFamilyId, iUserId, hasCitizenAccounts )
			%>
		</table>
	
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
' Function GetFamilyMembers( iFamilyId, iUserId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyMembers( ByVal iFamilyId, ByVal iUserId, ByVal hasCitizenAccounts )
	Dim sSql, iCount, sPreselected, oRs, sBgcolor

	iCount = 0
	sSql = "SELECT userid, userfname, userlname, birthdate, relationship, selftag, headofhousehold, ISNULL(accountbalance, 00000.0000) AS accountbalance "
	sSql = sSql & " FROM egov_users U, egov_familymember_relationships F WHERE U.relationshipid = F.relationshipid"
	sSql = sSql & " AND isdeleted = 0 AND familyid = " & iFamilyId & " ORDER BY birthdate ASC, userfname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	Do While Not oRs.eof 
		iCount = iCount + 1
		If iCount Mod 2 = 0 Then
			sBgcolor = " bgcolor=""#eeeeee"" "
		Else
			sBgcolor = ""
		End If 
		response.write vbcrlf & "<tr" & sBgcolor & "><td class=""name"">" & oRs("userfname") & " " & oRs("userlname") & "</td>"
		'If CLng(iUserId) = CLng(oRs("userid")) Then 
		If oRs("headofhousehold") Then 
			' What they call themselves
			response.write vbcrlf & "<td>Head of Household</td>"
		Else 
			' what they call others
			response.write vbcrlf & "<td>" & oRs("relationship") & "</td>"
		End If 
		
		' Display Age
		response.write vbcrlf & "<td>" 
		If Trim(oRs("birthdate")) = "" Or IsNull(oRs("birthdate")) Then
			response.write "Adult"
		Else
			response.write GetCitizenAge( oRs("birthdate") )
		End If 
		'response.write oRs("birthdate") 
		response.write "</td>"
		
		accountBalance = CDbl(oRs("accountbalance"))
		
		' display the account balance'
		If hasCitizenAccounts Then 
			response.write vbcrlf & "<td class=""accountbalance""><a href=""citizen_account_history.asp?u=" & oRs("userid") & """>" 
			response.write FormatCurrency(accountBalance, 2)
			response.write "</a></td>"
		End If 

		If CLng(iUserId) <> CLng(oRs("userid")) Then
			response.write "<td class=""actionbtn1""><input type=""button"" class=""button"" value=""Edit"" onClick=""EditFamilyMember(" & oRs("userid") & "," & iUserId & ");"" /></td>"
			' do not want you to be able to delete the registered citizen from here
			response.write "<td class=""actionbtn2"">"
			If accountBalance = CDbl(0) Then 
				response.write " <input type=""button"" class=""button"" value=""Delete"" onClick=""DeleteFamilyMember(" & oRs("userid") & ",'" & oRs("userfname") & " " & oRs("userlname") & "'," & iUserId & ");"" /></td>"
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"
		Else
			response.write "<td class=""actionbtn1""><input type=""button"" class=""button"" value=""Edit"" onClick=""EditCitizen(" & oRs("userid") & ");"" /></td><td>&nbsp;</td>"
		End If 
		response.write "</tr>"
		oRs.movenext
	Loop 
		
	oRs.close
	Set oRs = Nothing
	
	GetFamilyMembers = iCount
	
End Function  


'--------------------------------------------------------------------------------------------------
' Function GetUserName(iUserId)
'--------------------------------------------------------------------------------------------------
Function GetUserName( ByVal iUserId )
	Dim sSql, oRs

	sSql = "Select userfname, userlname from egov_users where userid = " & iUserId 
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetUserName = oRs("userfname") & " " & oRs("userlname")
	Else
		GetUserName = ""
	End If 

	oRs.close
	Set oRs = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' Function JavascriptSafe( strDB )
'--------------------------------------------------------------------------------------------------
Function JavascriptSafe( ByVal strDB )

	If Not VarType( strDB ) = vbString Then JavascriptSafe = strDB : Exit Function
	JavascriptSafe = Replace( strDB, "'", "\'" )
	
End Function


'--------------------------------------------------------------------------------------------------
' Function GetFamilyLastName( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyLastName( ByVal iUserId )
	Dim sSql, oRs

	sSql = "Select userlname from egov_users where userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	GetFamilyLastName = oRs("userlname")

	oRs.close
	Set oRs = Nothing
	
End Function 



%>