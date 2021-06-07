<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: header_edit.asp
' AUTHOR: Steve Loar
' CREATED: 4/11/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the editing of receipt headers and footers
'
' MODIFICATION HISTORY
' 1.0   2/1/07		Steve Loar - INITIAL VERSION
' 1.1	4/24/2007	Steve Loar - Added Refund Footer
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iHeaderDisplayId, iFooterDisplayId, sHeader, sFooter, sRefundFooter, iRefundFooterId

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit receipt headers" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

' These functions are all in common.asp
iHeaderDisplayId = GetDisplayId( "receipt header" )
sHeader = GetOrgDisplay( Session("orgid"), "receipt header" )
iFooterDisplayId = GetDisplayId( "receipt footer" )
sFooter = GetOrgDisplay( Session("orgid"), "receipt footer" )
iRefundFooterId = GetDisplayId( "refund footer" )
sRefundFooter = GetOrgDisplay( Session("orgid"), "refund footer" )

%>

<html>
<head>
 	<title>E-Gov Administration Console</title>

	 <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

 	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />

	<script language="Javascript">
	<!--

		function Validate()
		{
			var rege;
			var Ok;

			// check the header
			if (document.formReceipt.header.value == "")
			{
				alert('Please enter a header.');
				document.formReceipt.header.focus();
				return;
			}
			// check the footer
			if (document.formReceipt.footer.value == "")
			{
				alert('Please enter a footer.');
				document.formReceipt.footer.focus();
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
			<font size="+1"><strong>Receipt Header and Footer</strong></font>
		</p>
		<!--END: PAGE TITLE-->


		<!--BEGIN: FUNCTION LINKS-->
		<div id="functionlinks">
				<a href="javascript:Validate();"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;Update</a>&nbsp;&nbsp;
		</div>
		<!--END: FUNCTION LINKS-->


		<!--BEGIN: EDIT FORM-->
		<form name="formReceipt" action="header_update.asp" method="post">
			<input type="hidden" name="headerdisplayid" value="<%=iHeaderDisplayId%>" />
			<input type="hidden" name="footerdisplayid" value="<%=iFooterDisplayId%>" />
			<input type="hidden" name="refundfooterid" value="<%=iRefundFooterId%>" />
			<div class="shadow">
				<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
					<tr><th>Receipt Header and Footer</th></tr>
					<tr>
						<td>
							<table border="0" cellpadding="3" cellspacing="0">
								<tr>
									<td class="firstheadercell" align="right">Header: </td><td><textarea class="headertextarea" name="header"><%=sHeader%></textarea></td>
								</tr>
								<tr>
									<td class="firstheadercell" align="right">Footer: </td><td><textarea class="footertextarea" name="footer"><%=sFooter%></textarea></td>
								</tr>
								<tr>
									<td class="firstheadercell" align="right">Refund Footer: </td><td><textarea class="headertextarea" name="refundfooter"><%=sRefundFooter%></textarea></td>
								</tr>
								<tr><td>&nbsp;</td><td>* Use Simple HTML for formatting</td></tr>
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