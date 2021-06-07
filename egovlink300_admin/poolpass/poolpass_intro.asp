<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../class/classMembership.asp" -->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: poolpass_intro.asp
' AUTHOR: Steve Loar
' CREATED: 03/14/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This allows admins to edit the intro text on the poolpass purchase public initial page
'
' MODIFICATION HISTORY
' 1.0   03/14/06   Steve Loar - Initial code to edit intro text
' 2.0   07/11/2006 Steve Loar - Changed to be membership 
' 2.1	10/05/06	Steve Loar - Changed header and Nav
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMembershipId, oMembership

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "membership intro" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

Set oMembership = New classMembership
 if request("sMembershipType") <> "" then
    lcl_membership_type = request("sMembershipType")
 else
    'lcl_membership_type = "pool"
    lcl_membership_type = oMembership.GetFirstMembershipType()
 end if
' oMembership.SetMembershipId( "pool" )
 oMembership.SetMembershipId(lcl_membership_type)
 response.write GetMembership

If request("sIntroText") <> "" Then 
	oMembership.SaveMembershipIntro( request("sIntroText") )
End If 

If request("public_purchase") <> "" Then 
	If clng(request("public_purchase")) = 1 Then 
		oMembership.SetPublicDisplay( 1 )
	Else
		oMembership.SetPublicDisplay( 0 )
	End If 
End If

%>

<html>
<head>
	<title>E-Gov Membership Management</title>

	<link rel="stylesheet" type="text/css" href="../global.css">
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="./poolpass.css">

<script language="Javascript">
  <!--

  	function ValidateForm()
	{
		//alert(document.MemberForm.sIntroText.value);
		// Check the description
		if (document.MemberForm.sIntroText.value == "")
		{
			alert("Please provide some introductory text.");
			document.MemberForm.sIntroText.focus();
			return;
		}

		document.MemberForm.submit();	
	}


  //-->
 </script>
</head>

<body>
 
<%'DrawTabs tabRecreation,1%>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<p>
		<font size="+1"><strong>Membership Introductory Text</strong></font> 
	</p>
  	Membership Type: <select name="iMembershipId" id="iMembershipId" onChange="window.location.href = 'poolpass_intro.asp?smembershiptype=' + this.value">
  	<% showMembershipTypePicks oMembership.MembershipId %>
  	</select><br /><br />

	<div class="shadow"><table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
		<tr><th colspan="3" align="left"><%=Session("sOrgName")%> Membership Introductory Text</th></tr>
		<tr>
			<td>&nbsp;</td>
<!--			<td>&nbsp;
		<form name="nameform" method="post" action="poolpass_intro.asp">
			<strong>Memberships:</strong> 
				<select name="iMembershipId" onchange="javascript:document.nameform.submit();">
					<% '=oMembership.ShowMembershipPicks() %>
				</select>
			</form>
			</td>
-->
			<td>
			<form name="checkform" method="post" action="poolpass_intro.asp">
				<input type="hidden" name="iMembershipId" value="<%=oMembership.MembershipId%>" />
				<input type="hidden" name="sMembershiptype" value="<%=lcl_membership_type%>" />
				<%'response.write "[" & request("publicpurchase") & "]"
				'response.write "[" & request("iMembershipId") & "]"
				%>
				<% sDisplayText = oMembership.ShowPublicDisplayCheck() %>
					<input type="hidden" name="public_purchase" value="<%If sDisplayText <> "" Then 
						response.write "0"   ' This is what they will change to, not what they are
					  Else 
					    response.write "1"
					  End If %>" />
				<input type="checkbox" name="publicpurchase" <%=sDisplayText%> onclick="javascript:document.checkform.submit();" /> Public Can Purchase Memberships
			</form>
			</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<form name="MemberForm" method="post" action="poolpass_intro.asp">
			<td>&nbsp;</td>
			<td>
				<input type="hidden" name="iMembershipId" value="<%=oMembership.MembershipId%>" />
				<input type="hidden" name="sMembershiptype" value="<%=lcl_membership_type%>" />
				<strong>Introductory Text (include HTML tags for formatting)</strong><br />
				<textarea name="sIntroText" style="width:600px;height:350px;"><% =oMembership.ShowMembershipIntro() %></textarea>
			</td>
			<td width="30%" valign="top"><input type="submit" class="button" name="submit" value="Save Text Changes" onclick="ValidateForm();" /></td>
			</form>
		</tr>
	</table>
	</div>
</div>
<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  

</body>

</html>

<%
Set oMembership = Nothing 
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' Function ShowPoolPassIntro(  )
'------------------------------------------------------------------------------------------------------------
Sub ShowPoolPassIntro(  )
	Dim sSQL

	sSQL = "Select OrgPoolPassIntro FROM organizations WHERE orgid = " & Session("OrgID")

	Set oSeason = Server.CreateObject("ADODB.Recordset")
	oSeason.Open sSQL, Application("DSN") , 3, 1

	If Not oSeason.eof Then 
		response.write Trim(oSeason("OrgPoolPassIntro"))
	End If
		
	oSeason.close
	Set oSeason = Nothing

End Sub  


'------------------------------------------------------------------------------------------------------------
' Function SaveIntro(sIntroText)
'------------------------------------------------------------------------------------------------------------
Sub SaveIntro(sIntroText)
	Dim sSql, oCmd
	
	sIntroText = DBsafe(sIntroText)

	sSql = "Update organizations set OrgPoolPassIntro = '" & sIntroText & "' where orgid = " & Session("OrgID")
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

End Sub 


'------------------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function


'------------------------------------------------------------------------------------------------------------
' Function ShowMembershipPicks(iMembershipId, iOrgId)
'------------------------------------------------------------------------------------------------------------
Function ShowMembershipPicks(iMembershipId, iOrgId)
	Dim sSQL, oMembers

	' Get the memberships
	sSQL = "Select membershipid, membershipdesc FROM egov_memberships WHERE orgid = " & iOrgId & " order by membershipdesc"
	ShowMembershipPicks = ""

	Set oMembers = Server.CreateObject("ADODB.Recordset")
	oMembers.Open sSQL, Application("DSN"), 3, 1
	
	Do While not oMembers.eof 
		ShowMembershipPicks = ShowMembershipPicks & vbcrlf & "<option value=""" & oMembers("membershipid") & """ "
		If clng(iMembershipId) = clng(oMembers("membershipid"))  Then
			ShowMembershipPicks = ShowMembershipPicks & " selected=""selected"" "
		End If 
		ShowMembershipPicks = ShowMembershipPicks & ">" & oMembers("membershipdesc") & "</option>"
		oMembers.movenext
	Loop 

	oMembers.close
	Set oMembers = Nothing

End Function 


'------------------------------------------------------------------------------------------------------------
' Function GetInitialMembershipId( iOrgID )
'------------------------------------------------------------------------------------------------------------
Function GetInitialMembershipId( iOrgID )
	Dim sSql, oMember

	sSQL = "Select MIN(membershipid) as membershipid FROM egov_memberships WHERE orgid = " & Session("OrgID") 
	
	Set oMember = Server.CreateObject("ADODB.Recordset")
	oMember.Open sSQL, Application("DSN"), 3, 1
	
	If IsNull(oMember("membershipid")) Then
		GetInitialMembershipId = 0
	Else
		GetInitialMembershipId = oMember("membershipid")
	End If 
	
	oMember.close
	Set oMember = Nothing
End Function 

sub showMembershipTypePicks(p_membership_type)

  sSQL = "SELECT membershipid, membership, membershipdesc "
  sSQL = sSQL & " FROM egov_memberships "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " ORDER BY membershipdesc "

 	set rs = Server.CreateObject("ADODB.Recordset")
	 rs.Open sSQL, Application("DSN"), 3, 1

  if not rs.eof then
     while not rs.eof
	     selected = ""
	     if p_membership_type = rs("membershipid") then selected = " selected"

        response.write "  <option value=""" & rs("membership") & """ " & selected & ">" & rs("membershipdesc") & "</option>" & vbcrlf
        rs.movenext
     wend
  end if

  set rs = nothing

end sub

%>


