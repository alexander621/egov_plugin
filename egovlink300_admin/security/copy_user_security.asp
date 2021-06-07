<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->

<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: copy_user_security.ASP
' AUTHOR: Steve Loar
' CREATED: 10/02/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   10/02/2006	Steve Loar - INITIAL VERSION
' 1.1	03/03/2011	Steve Loar - Upgraded to jQuery to work in FireFox
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim bIsRootAdmin, sShowDetails, iUserID

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "user permission" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

If request("userid") = "" Then 
	response.redirect "edit_user_security.asp"
End If 

iOrgid = session("orgid")
iUserID = CLng(request("userid"))

bIsRootAdmin = UserIsRootAdmin( Session("UserID") )


%>

<html>

	<head>
		<title>E-GovLink Administration Console - Copy User Permissions</title>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
		<link rel="stylesheet" type="text/css" href="security.css" />

		<script type="text/javascript" src="https://code.jquery.com/jquery-1.5.min.js"></script>

		<script language="Javascript">
		<!--

		function CopyUser()
		{
			if ( $("#fromuserid").val() == $("#touserid").val() )
			{
				alert( 'You have selected the same user for both the source and target.\nPlease change one of your selections and try again.' );
			}
			else
			{
				document.UserForm.submit();
			}
		}

		//-->
		</script>
</head>
<body>

<% ShowHeader sLevel %>

<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
 <div id="content">
	<div id="centercontent">
		<input type="button" class="button" value="<< Back" onclick="javascript:history.back()" />
		<br /><br />

		<h3>Copy User Permissions</h3>

		<form name="UserForm" method="post" action="copysecurity.asp">
			<p>
				<label for="fromuserid">From User: </label><select name="iFromUserID" id="fromuserid">
	<%
				' Get the From user drop down 
				ShowOrgUserPicks iOrgId, iUserID, bIsRootAdmin
	%>
				</select>
				
				<label for="touserid">To User: </label><select name="iToUserID" id="touserid">
	<%
				' Get the From user drop down 
				ShowOrgUserPicks iOrgId, iUserID, bIsRootAdmin
	%>
				</select>
			</p>

			<p id="securitycopybutton">
				<input type="button" class="button" value="Copy Permissions" onClick="CopyUser();" />
			</p>
			
		</form>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%

'--------------------------------------------------------------------------------------------------
' void ShowOrgUserPicks iOrgId, iUserID, bIsRootAdmin 
'--------------------------------------------------------------------------------------------------
Sub ShowOrgUserPicks( ByVal iOrgId, ByVal iUserID, ByVal bIsRootAdmin )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname FROM users WHERE orgid = " & iOrgId & " AND isdeleted = 0"
	If Not bIsRootAdmin Then
		sSql = sSql & " AND (isrootadmin IS NULL OR isrootadmin = 0) "
	End If 
	sSql = sSql & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("userid") & """"
		If clng(iUserID) = clng(oRs("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write " > " & oRs("lastname") & ", " & oRs("firstname") & "</option>"
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing 
	
End Sub 


%>


