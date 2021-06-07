<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rec_message_edit.asp
' AUTHOR: Steve Loar
' CREATED: 7/6/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the editing of the message that is at the top of the rec activities public pages
'
' MODIFICATION HISTORY
' 1.0   7/6/07		Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMessageDisplayId, sMessage

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit recmessage" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

' These functions are all in common.asp
iMessageDisplayId = GetDisplayId( "classdetailsnotice" )
sMessage = GetOrgDisplay( Session("orgid"), "classdetailsnotice" )

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />

	<script language="Javascript">
	<!--

		function Validate()
		{
			var rege;
			var Ok;

			// check the message
			if (document.formReceipt.message.value == "")
			{
				alert('Please enter a page message.');
				document.formReceipt.message.focus();
				return;
			}
			//alert("OK");
			document.formReceipt.submit();
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
			<font size="+1"><strong>Recreation Page Message</strong></font>
		</p>
		<!--END: PAGE TITLE-->


		<!--BEGIN: FUNCTION LINKS-->
		<div id="functionlinks">
				<a href="javascript:Validate();"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;Update</a>&nbsp;&nbsp;
		</div>
		<!--END: FUNCTION LINKS-->


		<!--BEGIN: EDIT FORM-->
		<form name="formReceipt" action="rec_message_update.asp" method="post">
			<input type="hidden" name="messagedisplayid" value="<%=iMessageDisplayId%>" />
			<div class="shadow">
				<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
					<tr><th>Recreation Page Message</th></tr>
					<tr>
						<td>
							<table border="0" cellpadding="30" cellspacing="0">
								<tr>
									<td>
										<textarea id="recpagemessage" name="message"><%=sMessage%></textarea>
										<br />* Use Simple HTML for formatting
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