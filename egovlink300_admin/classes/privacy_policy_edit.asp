<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: privacy_policy_edit.asp
' AUTHOR: Steve Loar
' CREATED: 7/12/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the editing of the Privacy policy
'
' MODIFICATION HISTORY
' 1.0   7/12/07		Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMessageDisplayId, sMessage

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit privacypolicy" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

' These functions are all in common.asp
iMessageDisplayId = GetDisplayId( "privacy policy" )
sMessage = GetOrgDisplay( Session("orgid"), "privacy policy" )

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />

	<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>

	<script language="Javascript">
	<!--

		function Validate()
		{
			var rege;
			var Ok;

			// check the message
			if (document.formReceipt.privacypolicy.value == "")
			{
				alert('Please enter a privacy Policy.');
				document.formReceipt.privacypolicy.focus();
				return;
			}
			//alert("OK");
			document.formReceipt.submit();
		}

		function DeletePolicy()
		{
			if (confirm('Delete the Privacy Policy?'))
			{
				location.href='privacy_policy_delete.asp?messagedisplayid=<%=iMessageDisplayId%>';
			}
		}

	//-->
	</script>

</head>

<body onload="setMaxLength();">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
		<!--BEGIN: PAGE TITLE-->
		<p>
			<font size="+1"><strong>Privacy Policy</strong></font>
		</p>
		<!--END: PAGE TITLE-->


		<!--BEGIN: FUNCTION LINKS-->
		<div id="functionlinks">
				<a href="javascript:Validate();"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;Update</a>&nbsp;&nbsp;
				<a href="javascript:DeletePolicy();"><img src="../images/small_delete.gif" align="absmiddle" border="0">&nbsp;Delete</a>&nbsp;&nbsp;
		</div>
		<!--END: FUNCTION LINKS-->


		<!--BEGIN: EDIT FORM-->
		<form name="formReceipt" action="privacy_policy_update.asp" method="post">
			<input type="hidden" name="messagedisplayid" value="<%=iMessageDisplayId%>" />
			<div class="shadow">
				<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
					<tr><th>Privacy Policy</th></tr>
					<tr>
						<td>
							<table border="0" cellpadding="10" cellspacing="0">
								<tr>
									<td class="privacypolicy">
										<textarea id="privacypolicy" name="privacypolicy" maxlength="9000" wrap="soft"><%=sMessage%></textarea>
									</td>
								</tr>
								<tr>
									<td class="privacypolicy">
										* Use Simple HTML for formatting
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



%>