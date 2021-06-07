<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: new_faq.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module adds Frequently Asked Questions (FAQ).
'
' MODIFICATION HISTORY
' 1.?   09/11/06   Steve Loar - Changes for categories.
' 1.2	10/09/06	Steve Loar - Security, Header and Nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "manage faq" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>

<html>
<head>
<title> E-GovLink Faq Management </title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>

	<script language="Javascript">
	<!--
		function doCalendar(sField) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=NewEvent", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
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

		function doPicker(sFormField) 
		{
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

		function validate() 
		{
			var rege;
			var OK;

			// check the publication start date
			if (document.NewEvent.publicationstart.value != "")
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.NewEvent.publicationstart.value);
				if (! Ok)
				{
					alert("The Publication Start date should be in the format of MM/DD/YYYY, or be blank.  \nPlease enter it again.");
					document.NewEvent.publicationstart.focus();
					return;
				}
			}
			// check the publication end date
			if (document.NewEvent.publicationend.value != "")
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.NewEvent.publicationend.value);
				if (! Ok)
				{
					alert("The Publication End date should be in the format of MM/DD/YYYY, or be blank.  \nPlease enter it again.");
					document.NewEvent.publicationend.focus();
					return;
				}
			}
			document.NewEvent.submit();
		}

	//-->
	</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="setMaxLength();">
  <%'DrawTabs tabActionline,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 


  <table border="0" cellpadding="6" cellspacing="0" class="start" width="100%">
    <tr>
      <td valign="top">

	<div style="margin-top:20px; margin-left:20px;">

	<p><font class=label>FAQ - Add Faq </font><br /><br />
		<!--<small>[<a class=edit href="copy_form.asp?task=copyme&iformid=<%=iFormID%>&iorgid=<%=iorgid%>">Copy This Form</a>]</small> -->
		<!--<small>[<a class=edit href="../action_line/edit_form.asp?task=name&control=<%=iFormID%>&iorgid=<%=iorgid%>">Manage This Form</a>]</small> -->
		<!--<small>[<a class=edit href="list_faq.asp">Return to FAQ List</a>]</small>
		<hr size="1" width="600px;" style="text-align:left; color:#000000;">-->

		<a href="list_faq.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
	</p>

	<div class="group">

		<div class="orgadminboxf">

		<!--BEGIN: FAQ -->
			<P>
			<form name="NewEvent" action="save_faq.asp" method="post">
				<p>
					<table border="0" cellpadding="6" cellspacing="0">
<%
						' Display FAQ category picks
						ShowCategories
%>
						<tr><td><b>Question:</b></td></tr>
						<tr><td><textarea class="formtextarea" name="FaqQ" maxlength="700"></textarea></td></tr>
						<tr><td><b>Answer:</b></td></tr>
						<tr><td><textarea class="formtextareaBig" name="FaqA" maxlength="4000" onselect="storeCaret(this);" onclick="storeCaret(this);" onkeyup="storeCaret(this);" ONDBLCLICK="storeCaret(this);"></textarea> 
						&nbsp; <input type="button" class="actionbtn" value="Add Link" onClick="doPicker('NewEvent.FaqA');" /></td></tr>
						<tr><td><strong>Publication Start: </strong> &nbsp; <input type="text" name="publicationstart" value="" size="10" maxlength="10" />
						&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publicationstart');" /></span>
						</td></tr>
						<tr><td><strong>Publication End: </strong> &nbsp; <input type="text" name="publicationend" value="" size="10" maxlength="10" />
						&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publicationend');" /></span>
						</td></tr>
				</table>
				</p>
				<p>
					<input type="button" class="actionbtn" value="ADD FAQ" onclick="validate();" />
				</p>
			</form>
		<!--END: FAQ -->

		</div>
	</div>

	<!--include file="bottom_include.asp"-->
	</div>
      </td>
       
    </tr>
  </table>
</body>
</html>


<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' Sub ShowCategories( )
'------------------------------------------------------------------------------------------------------------
Sub ShowCategories( )
	Dim sSql, oFAQCats

	sSql = "Select FAQCategoryName, FAQCategoryId, isnull(internalonly,0) AS internalonly FROM faq_categories where orgid = " & session("orgid") & " Order by displayorder"

	Set oFAQCats = Server.CreateObject("ADODB.Recordset")
	oFAQCats.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly
	
	response.write "<tr><td><strong>Category:</strong> &nbsp; <select name=""FAQCategoryId"">"
	response.write vbcrlf & vbtab & "<option value=""0"">None</option>"
	Do While Not oFAQCats.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oFAQCats("FAQCategoryId") & """"
		response.write ">" & oFAQCats("FAQCategoryName")
		If oFAQCats("internalonly") Then
			response.write " (Internal)"
		End If 
		response.write "</option>"
		oFAQCats.MoveNext
	Loop 
	response.write vbcrlf & "</select></td></tr>"

	oFAQCats.close
	Set oFAQCats = Nothing 

End Sub 


%>