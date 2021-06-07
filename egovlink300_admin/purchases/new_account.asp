<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: new_account.asp
' AUTHOR: Steve Loar
' CREATED: 2/15/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is where accounts are added
'
' MODIFICATION HISTORY
' 1.0   2/15/07   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iAccountId, sAccountName, sAccountNumber

sLevel = "../" ' Override of value from common.asp

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<style>
		#content * { font-family:'Open Sans', sans-serif !important; font-size:14px; }
		.ui-button, .ui-button:hover, .ui-button:active, .ui-button:focus { font-size:14px;background-color: #2C5F93 !important; color:white !important; border:inherit !important; margin-bottom:5px; padding: .4em 1em !important; }
		.ui-button:disabled {
			background-color: #ccc;
			color:#999;
			margin-bottom:5px;
			cursor: not-allowed;
		}
		div.glaccounts_shadow, .shadow {background:none !important;}
	</style>

	<script language="Javascript" src="tablesort.js"></script>

	<script language="Javascript">
	<!--

		function Validate() 
		{
			if (document.frmAccount.accountname.value == '')
			{
				alert("Account Name cannot be blank.");
				document.frmAccount.accountname.focus();
				return;
			}
			if (document.frmAccount.accountnumber.value == '')
			{
				alert("Account Number cannot be blank.");
				document.frmAccount.accountnumber.focus();
				return;
			}
			document.frmAccount.submit();
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
			<font size="+1"><strong>New GL Account</strong></font><br />
		</p>
		<!--END: PAGE TITLE-->


		<!--BEGIN: FUNCTION LINKS-->
		<div id="functionlinks">
				<a href="gl_account_mgmt.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Back</a><br /><br />
				<input type="button" value="Create" onClick="Validate();" class="button ui-button ui-widget ui-corner-all"><br /><br />
		</div>
		<!--END: FUNCTION LINKS-->


		<!--BEGIN: EDIT FORM-->
		<form name="frmAccount" action="save_accounts.asp" method="post">
		<input type="hidden" name="action" value="new" />
		<div class="shadow">
			<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
				<tr>
					<th>Account Information</th>
				</tr>
				<tr>
					<td>
						<table>
							<tr>
								<td align="right">Account Number: </td>
								<td><input type="text" name="accountnumber" value="" size="20" maxlength="20" /></td>
							</tr>
							<tr>
								<td align="right">Account Name: </td>
								<td><input type="text" name="accountname" value="" size="50" maxlength="50" /></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<input type="button" value="Create" onClick="Validate();" class="button ui-button ui-widget ui-corner-all">
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

