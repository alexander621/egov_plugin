<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: header_edit.asp
' AUTHOR: Steve Loar
' CREATED: 5/16/2007
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the editing of permit headers and footers
'
' MODIFICATION HISTORY
' 1.0   05/15/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iInvoiceHeaderDisplayId, sInvoiceHeader, sLogoURL, iLogoURLDisplayId

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "edit permit headers", sLevel	' In common.asp

' These functions are all in common.asp
iInvoiceHeaderDisplayId = GetDisplayId( "invoice header" )
sInvoiceHeader = GetOrgDisplay( Session("orgid"), "invoice header" )

iLogoURLDisplayId = GetDisplayId( "invoice url" )
sLogoURL = GetOrgDisplay( Session("orgid"), "invoice url" )

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>

	<script language="Javascript">
	<!--

		function Validate()
		{
			var rege;
			var Ok;

			// check the header
			if (document.formReceipt.invoiceheader.value == "")
			{
				alert('Please enter an invoice header.');
				document.formReceipt.invoiceheader.focus();
				return;
			}

			//alert("OK");
			document.formReceipt.submit();
		}

		function doPicker(sFormField) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			eval('window.open("imagepicker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function insertAtURL (textEl, text) 
		{
			if (textEl.createTextRange && textEl.caretPos) 
			{
				var caretPos = textEl.caretPos;
				caretPos.text =
			caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
			text + ' ' : text;
			}
			else
				textEl.value  = text;
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
			<font size="+1"><strong>Permit Invoice Header</strong></font>
		</p>
		<!--END: PAGE TITLE-->


		<!--BEGIN: FUNCTION LINKS-->
		<div id="functionlinks">
				<input type="button" class="button ui-button ui-widget ui-corner-all" value="Save Changes" onclick="Validate();" />
		</div>
		<!--END: FUNCTION LINKS-->


		<!--BEGIN: EDIT FORM-->
		<form name="formReceipt" action="header_update.asp" method="post">
			<input type="hidden" name="headerdisplayid" value="<%=iInvoiceHeaderDisplayId%>" />
			<input type="hidden" name="logodisplayid" value="<%=iLogoURLDisplayId%>" />

			<div class="shadow">
				<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
					<tr>
						<td align="right" nowrap="nowrap">
							Logo: 
						</td>
						<td valign="middle">
							<input type="text" id="logourl" name="logourl" value="<%=sLogoURL%>" size="50" maxlength="250" /> &nbsp; &nbsp; 
							<input type="button" class="button ui-button ui-widget ui-corner-all" value="Select Logo" onclick="doPicker('formReceipt.logourl');" /> &nbsp; &nbsp;
<%							If sLogoURL <> "" Then  %>
								<img src="<%=sLogoURL%>" alt="" border="0" align="absmiddle" />
<%							End If  %>
						</td>
					</tr>
					<tr>
						<td align="right" nowrap="nowrap">Invoice Header Text: </td>
						<td><textarea class="headertextarea" name="invoiceheader" maxlength="200" wrap="soft"><%=sInvoiceHeader%></textarea>
						</td>
					</tr>
					<tr><td>&nbsp;</td><td>* Use Simple HTML for formatting</td>
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
