<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: refund_policy_edit.asp
' AUTHOR: Steve Loar
' CREATED: 7/11/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the editing of the refund policy
'
' MODIFICATION HISTORY
' 1.0   7/11/07		Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMessageDisplayId, sMessage, iDisplayOrgId, sDisplayName, sOrgname, sDescription, bUsesDisplayName, sDisplayField

sLevel = "../" ' Override of value from common.asp

If Not UserIsRootAdmin( session("UserID") ) Then
	response.redirect "../default.asp"
End If 

iDisplayOrgId = CLng(request("orgid"))

' These functions are all in common.asp
If request("displayid") <> "" Then
	iMessageDisplayId = CLng(request("displayid"))
Else
	iMessageDisplayId = GetInitialEditDisplayId( )
End If 
sDisplayName = GetDisplayName( iMessageDisplayId )
sDescription = GetDisplayDescription ( iMessageDisplayId, bUsesDisplayName ) 
If bUsesDisplayName Then
	sDisplayField = "displayname"
Else
	sDisplayField = "displaydescription"
End If 
sMessage = GetOrgDisplayWithId( iDisplayOrgId, iMessageDisplayId, bUsesDisplayName )


%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="admin.css" />

	<script language="Javascript">
	<!--

		function Validate()
		{
			var rege;
			var Ok;

			// check the message
			if (document.formReceipt.message.value == "")
			{
				alert('Please enter text for the display.');
				document.formReceipt.message.focus();
				return;
			}
			//alert("OK");
			document.formReceipt.submit();
		}

		function DeletePolicy()
		{
			if (confirm('Delete the display?'))
			{
				location.href='org_display_delete.asp?orgid=<%=iDisplayOrgId%>&displayid=<%=iMessageDisplayId%>';
			}
		}

		function ShowDisplay()
		{
			document.pickForm.submit();
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
	
		<!--BEGIN: PAGE TITLE-->
		<p>
			<font size="+1"><strong><%=GetOrgName( iDisplayOrgId )%> Organization Displays</strong></font>
		</p>
		<!--END: PAGE TITLE-->


		<!--BEGIN: FUNCTION LINKS-->
		<div id="functionlinks">
			<input type="button" class="button" value="<< Return to Feature Selection" onclick="location.href='featureselection.asp?orgid=<%=iDisplayOrgId%>'"; />&nbsp;&nbsp;
			<input type="button" class="button" value="Update" onclick="Validate()"; />&nbsp;&nbsp;
			<input type="button" class="button" value="Delete" onclick="DeletePolicy()"; />
		</div>
		<!--END: FUNCTION LINKS-->
		
		<div id="pickform">
			<form name="pickForm" method="post" action="org_display_edit.asp">
				<input type="hidden" name="orgid" value="<%=iDisplayOrgId%>" />
				<%ShowDisplayPicks iMessageDisplayId %>
			</form>
		</div>

		<!--BEGIN: EDIT FORM-->
		<form name="formReceipt" action="org_display_update.asp" method="post">
			<input type="hidden" name="displayid" value="<%=iMessageDisplayId%>" />
			<input type="hidden" name="orgid" value="<%=iDisplayOrgId%>" />
			<input type="hidden" name="displayfield" value="<%=sDisplayField%>" />
			<div class="shadow">
				<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
					<tr><th><%=sDisplayName%></th></tr>
					<tr>
						<td>
							<table border="0" cellpadding="10" cellspacing="0">
								<tr>
									<td>Description: &nbsp; <%= sDescription%></td>
								</tr>
								<tr>
									<td>
										<textarea id="refundpolicy" name="message" cols="50" rows="10"><%=sMessage%></textarea>
										<br />* Use Simple HTML for formatting
									</td>
								</tr>
								<tr>
									<td>
										Input text to override the default display text. Delete this text to use the default display text again.
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</div>
		</form>
		<!--END: EDIT FORM-->

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>



<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function GetInitialEditDisplayId( )
'--------------------------------------------------------------------------------------------------
Function GetInitialEditDisplayId( )
	Dim sSql, oDisplay

	sSql = "SELECT TOP 1 displayid FROM egov_organization_displays WHERE admincanedit = 1 ORDER BY displayname"

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open  sSQL, Application("DSN"), 3, 1
	
	If Not oDisplay.EOF Then
		GetInitialEditDisplayId = CLng(oDisplay("displayid"))
	Else
		GetInitialEditDisplayId = CLng(0)
	End If
	
	oDisplay.close 
	Set oDisplay = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' ShowDisplayPicks iDisplayId
'--------------------------------------------------------------------------------------------------
Sub ShowDisplayPicks( ByVal iDisplayId )
	Dim sSql, oRs
	
	sSql = "SELECT displayid, displayname FROM egov_organization_displays WHERE admincanedit = 1 ORDER BY displayname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""displayid"" onChange=""ShowDisplay();"">"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("displayid") & """"
			If CLng(oRs("displayid")) = CLng(iDisplayId) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("displayname") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetDisplayDescription ( iMessageDisplayId ) 
'--------------------------------------------------------------------------------------------------
Function GetDisplayDescription ( ByVal iMessageDisplayId, ByRef bUsesDisplayName ) 
	Dim sSql, oDisplay, sDescription
	
	sSql = "SELECT displaydescription, usesdisplayname FROM egov_organization_displays WHERE displayid = " & iMessageDisplayId 

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open  sSQL, Application("DSN"), 3, 1

	If Not oDisplay.EOF Then 
		sDescription = oDisplay("displaydescription")
		If oDisplay("usesdisplayname") Then 
			bUsesDisplayName = True 
		Else
			bUsesDisplayName = False 
		End If 
	Else
		sDescription = "No Description Found."
		bUsesDisplayName = False 
	End If 

	oDisplay.close
	Set oDisplay = Nothing 

	GetDisplayDescription = sDescription

End Function



%>