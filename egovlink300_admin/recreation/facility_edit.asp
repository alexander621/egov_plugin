<!-- #include file="../includes/common.asp" //-->
<!--#Include file="facility_functions.asp"-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CLIENT_TEMPLATE_PAGE.ASP
' AUTHOR: Steve Loar
' CREATED: 01/24/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/24/06   Steve Loar - Code added
' 1.1	10/06/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFacilityId
Dim sFacilityName, chkisviewable, chkisreservable
Dim sFound

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "edit facilities", sLevel	' In common.asp

'If Not UserHasPermission( Session("UserId"), "edit facilities" ) Then
'	response.redirect sLevel & "permissiondenied.asp"
'End If 

sFound = ""

If request("facilityid") = "" Then
	response.redirect( "facility_management.asp" )
Else 
	iFacilityId = request("facilityid")
End If

' GET FACILITY INFORMATION
SetFacilityInformation(iFacilityID)
'sFacilityName = GetFacilityName(iFacilityId)

Dim oFacilities
Dim iRowCount
Dim x
%>

<html>
<head>
	<title>E-Gov Facility Edit</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="facility.css" />

<script language="Javascript">
  <!--
  	function VerifyName(passForm)
	{
		if (passForm.sFacilityName.value == "") 
		{
			alert("Please enter a name for the facility.");
			passForm.sName.focus();
			return;
		}

		 passForm.submit();
	}

	function doPicker(sFormField) 
	{
      w = (screen.width - 350)/2;
      h = (screen.height - 350)/2;
      eval('window.open("imagepicker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
    }

     function storeCaret (textEl) {
       if (textEl.createTextRange)
         textEl.caretPos = document.selection.createRange().duplicate();
     }

     function insertAtURL (textEl, text) {
       if (textEl.createTextRange && textEl.caretPos) {
         var caretPos = textEl.caretPos;
         caretPos.text =
           caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
             text + ' ' : text;
       }
       else
         textEl.value  = text;
     }


	function CheckElement(passForm)
	{
		var msg;
		if (passForm.content.value == "") 
		{
			if (passForm.sequence.value > 4)
			{
				msg = "Please enter a URL for the image before saving.";
			}
			else
			{
				msg = "Please enter some text before saving";
			}
			alert(msg);
			passForm.content.focus();
			return;
		}
		passForm.submit();
	}

	function ClearElement(passForm)
	{
		var msg;
		if (passForm.sequence.value > 4)
		{
			 msg = "Do you wish to clear Image " + (passForm.sequence.value - 4) + "?";
		}
		else
		{
			msg = "Do you wish to clear Text " + passForm.sequence.value + "?";
		}
		
		if (confirm(msg))
		{
			passForm.content.value = "";
			passForm.alt_tag.value = "";
			passForm.submit();
		}

	}

  //-->
 </script>
   <script language="Javascript">
  <!--
    function openpreview() {
      w = (screen.width - 350)/2;
      h = (screen.height - 350)/2;
      eval('window.open("http://www.egovlink.com/eclink/admin/recreation/facility_template_preview.asp?ifacilityid=<%=request("facilityid")%>", "_preview", "width=750,height=550,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
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
	<font size="+1"><strong>Recreation: Facility Edit - 
	<% If iFacilityId <> "0" Then %>
		<%=sFacilityName%>
	<% 
		Else
			response.write "New Facility"
		End If %></strong></font><br />

	<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
	</p>

	<div id="functionlinks">
		<a href="javascript:history.go(-1)"><img src="../images/cancel.gif" align="absmiddle" border="0">&nbsp;Cancel</a>&nbsp;&nbsp;
		<% If iFacilityId <> "0" Then %>
			<a href="javascript:openpreview();"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;Preview</a>&nbsp;&nbsp;
		<% End If %>
	</div>

	<div class="shadow">
	<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
		<tr>
			<th>Facility name</th><th>Category</th><th>Display Template</th><th>Is Viewable?</th><th>Is Reservable?</th><th>&nbsp;</th>
		</tr>
		<tr>
			<td>
				<form name="nameform" method="post" action="facility_name_save.asp" >
					<input type="hidden" name="iFacilityId" value="<%=iFacilityId%>" />
					<input type="text" name="sFacilityName" size="50" maxlength="50" value="<%=sFacilityName%>" />
			</td>
			<td>
				<% 
					iCurrentCategory = GetFacilityCategory( ifacilityid ) 
					ShowCategoryPicks iCurrentCategory 
				%>
			</td>
			<td>
				<% subDrawSelectTemplate( ifacilityid ) %>
			</td>
			<td>
				<input <%If chkisviewable = true Then response.write " checked=""checked"" " End If%> type="checkbox" name="chkisviewable" />
			</td>
			<td>
				<input <%If chkisreservable = true Then response.write " checked=""checked"" " End If%> type="checkbox" name="chkisreservable" />
			</td>
			<td>
				<a href="javascript:VerifyName(document.nameform);">Save</a>
				</form>
			</td>
		</tr>
<% 	If iFacilityId = "0" Then %>
		<tr><td colspan="2">To create a new facility, please start by creating a name for it.</td></tr>
<%	End If %>
	</table>
	</div><br /><br />

<%
	If iFacilityId <> "0" Then %>
	<div class="shadow">
	<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
		<tr>
			<th colspan="3">Display Elements</th>
		</tr>

<%		For iRowCount = 1 To 8

			If iRowCount Mod 2 = 0 Then
				response.write "<tr class=""alt_row"">"
			Else
				response.write "<tr>"
			End If

			If iRowCount < 5 Then %>
				<td valign="top" colspan="2">
					<form name="element<%=iRowCount%>" method="post" action="element_save.asp">
					<input type="hidden" name="iFacilityId" value="<%=iFacilityId%>" />
					<input type="hidden" name="sequence" value="<%=iRowCount%>" />
					<input type="hidden" name="alt_tag" value="" />
					<input type="hidden" name="element_type" value="textarea" />
<%					' Get a text row
					response.write "<strong>Text " & iRowCount & "</strong><br />"
					response.write GetTextArea(iFacilityId,iRowCount, sFound)
%>
				</td>
<%			Else %>
				<td valign="top">
					<form name="element<%=iRowCount%>" method="post" action="element_save.asp">
					<input type="hidden" name="iFacilityId" value="<%=iFacilityId%>" />
					<input type="hidden" name="sequence" value="<%=iRowCount%>" />
					<input type="hidden" name="element_type" value="img" />
<%					' get an image row
					response.write "<strong>Image " & (iRowCount - 4) & "</strong><br />"
					response.write GetImageInfo(iFacilityId, iRowCount, sFound)
%>
				</td>
<%
			End If 
%>
				<td>
					<a href="javascript:CheckElement(document.element<%=iRowCount%>);">Save</a>
					<% If sFound = "yes" Then %>
						<a href="javascript:ClearElement(document.element<%=iRowCount%>);">Clear</a>
					<% End If %>
					</form>
				</td>
			</tr>
<%
		Next 
%>
	</table>
	</div>

<%	End If %>

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


