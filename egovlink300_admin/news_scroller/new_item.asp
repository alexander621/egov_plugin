<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: new_item.asp
' AUTHOR: Steve Loar
' CREATED: 10/31/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module adds News Scroller Items.
'
' MODIFICATION HISTORY
' 1.0   09/11/06	Steve Loar - Initial Version.
' 1.1	12/04/07	Steve Loar - Added Pub start and end dates
' 1.2	02/25/2008	Steve Loar - Added textarea limit JavaScript and increased size to 400 from 200
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit scroller" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>

<html>
<head>
<title>E-GovLink News Scroller Item</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>

	<script language="Javascript">
	<!--

		function doCalendar(sField) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmNewItem", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		 function storeCaret (textEl) 
		 {
		   if (textEl.createTextRange) 
			 textEl.caretPos = document.selection.createRange().duplicate();
		 }

		 function insertAtCaret (textEl, text) 
		 {
		   if (textEl.createTextRange && textEl.caretPos) {
			 var caretPos = textEl.caretPos;
			 caretPos.text =
			   caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
				 text + ' ' : text;
		   }
		   else
			 textEl.value  = text;
		 }

		function doSitePicker(sFormField) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			eval('window.open("../sitelinker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=470,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function doPicker(sFormField) {
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  eval('window.open("../picker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function fnCheckSubject()
		{
			if (document.NewEvent.Subject.value != '') {
				return true;
			}
			else
			{
				return false;
			}
		}

		function ValidateForm() 
		{
			var rege;
			var OK;
			
			// validate the item title
			if (document.frmNewItem.itemtitle.value == "")
			{
				alert("Please enter a title for this News Item.");
				document.frmNewItem.itemtitle.focus();
				return;
			}

			// validate the item date
			if (document.frmNewItem.itemdate.value == "")
			{
				alert("Please enter a date for this News Item.");
				document.frmNewItem.itemdate.focus();
				return;
			}
			else
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.frmNewItem.itemdate.value);
				if (! Ok)
				{
					alert("The date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.frmNewItem.itemdate.focus();
					return;
				}
			}

			// Validate the message text
			if (document.frmNewItem.itemtext.value == "")
			{
				alert("Please enter a message for this News Item.");
				document.frmNewItem.itemtext.focus();
				return;
			}
			else 
			{
				if (document.frmNewItem.itemtext.value.length > document.frmNewItem.itemtext.getAttribute('maxlength'))
				{
					alert("The maxium length for the Message is " + document.frmNewItem.itemtext.getAttribute('maxlength') + " characters.\nPlease correct this and try saving again.");
					document.frmNewItem.itemtext.focus();
					return;
				}
			}

			// check the publication start date
			if (document.frmNewItem.publicationstart.value != "")
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.frmNewItem.publicationstart.value);
				if (! Ok)
				{
					alert("The Publication Start date should be in the format of MM/DD/YYYY, or be blank.  \nPlease enter it again.");
					document.frmNewItem.publicationstart.focus();
					return;
				}
			}
			// check the publication end date
			if (document.frmNewItem.publicationend.value != "")
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.frmNewItem.publicationend.value);
				if (! Ok)
				{
					alert("The Publication End date should be in the format of MM/DD/YYYY, or be blank.  \nPlease enter it again.");
					document.frmNewItem.publicationend.focus();
					return;
				}
			}

			// all is ok, so submit and save 
			document.frmNewItem.submit();
		}

	//-->
	</script>

</head>

<body onload="setMaxLength();">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="content">
	<div id="centercontent">

	<p>
		<font size="+1"><strong>Create News Item</strong></font><br /><br />

		<a href="list_items.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
	</p>

		<!--BEGIN: FAQ -->
	<form name="frmNewItem" action="save_item.asp" method="post">
		<input type="hidden" name="newsitemid" value="0" />
		<p>
			<div class="shadow">
			<table id="newsitemedit" border="0" cellpadding="6" cellspacing="0">
				<tr><td>
					<strong>Title:</strong> &nbsp; <input type="text" name="itemtitle" value="" maxlength="100" size="100" />
				</td></tr>
				<tr><td>
					<strong>Date:</strong> (MM/DD/YYYY) &nbsp; <input type="text" name="itemdate" value="" maxlength="10" size="15" />
					&nbsp; <img class="calendarimg" src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('itemdate');" />
				</td></tr>
				<tr><td><strong>Message:</strong> &nbsp;* You May Use Simple HTML for formatting</td></tr>
				<tr><td>
					<textarea name="itemtext" id="itemtext" rows="5" cols="100" maxlength="400" wrap="soft"></textarea>
				</td></tr>
				<tr><td>
					<strong>Link URL:</strong> &nbsp; <input type="text" name="itemlinkurl" value="" maxlength="500" size="50" />
					&nbsp; <input type="button" value="Find a Link" class="button" onClick="doSitePicker('frmNewItem.itemlinkurl');" />
				</td></tr>
				<tr><td><strong>Publication Start: </strong> &nbsp; <input type="text" name="publicationstart" value="" />
				&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publicationstart');" /></span>
				</td></tr>
				<tr><td><strong>Publication End: </strong> &nbsp; <input type="text" name="publicationend" value="" />
				&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publicationend');" /></span>
				</td></tr>
			</table>
			</div>
		</p>
		<p>
			<input type="button" class="button" value="Create News Item" onclick="ValidateForm();" />
		</p>
	</form>
		<!--END: FAQ -->

	</div>
</div>
	
	<!--#Include file="../admin_footer.asp"--> 

</body>
</html>


<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------



%>
