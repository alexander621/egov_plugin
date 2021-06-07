<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: WAIVER_MGMT.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 03/21/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   05/17/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sName, sDescription, sType, sURL, blnRequired, sBody, iWaiverID, sTitle, sLinkText, sChecked

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "waivers" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

' GET waiver ID
If request("waiverid") = "" OR Not isnumeric(request("waiverid")) OR request("waiverid") = "0" Then
	' CREATE NEW waiver
	iWaiverID = 0
	sTitle = "Add New Waiver"
	sLinkText = "Create Waiver"
Else
	' EDIT EXISTING waiver
	iWaiverID = CLng( request("waiverid") )
	sTitle = "Edit Waiver"
	sLinkText = "Save Changes"
End If

blnHasWP = hasWordPress()
sHomeWebsiteURL = getOrganization_WP_URL(session("orgid"), "OrgPublicWebsiteURL")

' GET waiver INFORMATION
GetwaiverInfo iWaiverID, sName, sType, sDescription, sBody, sUrl, blnRequired

If blnRequired Then
	sChecked = " checked=""checked"" "
End If
%>

<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="classes.css" />
	<link rel="stylesheet" href="../recreation/facility.css" />

	<script src="tablesort.js"></script>

  	<script src="//code.jquery.com/jquery-1.12.4.js"></script>
   	<script src="//code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
	<!--#include file="../includes/wp-image-picker.asp"-->

	<script>
	<!--

		function doPicker(sFormField) 
		{
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  eval('window.open("documentpicker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function openWin2(url, name) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 550)/2;
			popupWin = eval('window.open(url, name,"resizable,width=820,height=600,left=' + 80 + ',top=' + h + '")');
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

			$('#documentURLpic').html('<a href="' + text + '" target="_newwindow">View Document</a>&nbsp;&nbsp;');
		}
//-->
	</script>

</head>

<body>
 
<%'DrawTabs tabRecreation,1%>
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content" style="width:auto;">
	<div id="centercontent">
	
<!--BEGIN: PAGE TITLE-->
<p>
	<font size="+1"><strong>Recreation: <%=sTitle%></strong></font><br />
	<a href="class_waivers.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
</p>
<!--END: PAGE TITLE-->


<!--BEGIN: FUNCTION LINKS-->
<div id="functionlinks">
		<a href="class_waivers.asp"><img src="../images/cancel.gif" align="absmiddle" border="0">&nbsp;Cancel</a>&nbsp;&nbsp;
		<a href="javascript:document.frmwaiver.submit()"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;<%=sLinkText%></a>&nbsp;&nbsp;
</div>
<!--END: FUNCTION LINKS-->


<!--BEGIN: EDIT FORM-->
<form name="frmwaiver" accept-charset="UTF-8" action="class_waiver_save.asp" method="post">
<input type="hidden" name="iwaiverid" value="<%=iWaiverID%>" />

<div class="shadow">
	<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
		<tr>
			<th>Waiver Information</th>
		</tr>
		<tr>
			<td>
				<table>

					<tr>
						<td class="label">Type:</td>
						<td>
							<select name="sType">
								<option value="LINK" <%if UCase(sType) = "LINK" then response.write " selected=""selected"" "%>>LINK</option>
								<option value="TERM" <%if UCase(sType) = "TERM" then response.write " selected=""selected"" "%>>TERM</option>
							</select>
						</td>
					</tr>
					<tr>
						<td class="label">Name:</td><td><input class="waiver" type="text" name="sName" maxlength="50" value="<%=sName%>" /></td>
					</tr>
					<tr>
						<td class="label">Description:</td><td><textarea name="sDescription" maxlength="4000"><%=sDescription%></textarea></td>
					</tr>
					<tr>
						<td class="label">Body:</td><td><textarea name="sBody" maxlength="4000"><%=sBody%></textarea></td>
					</tr>
					<tr>
						<td class="label">URL:</td>
						<td>
							<input class="waiver" type="<% if blnHasWP then %>hidden<%else%>text<%end if %>" name="sURL" id="documentURL" maxlength="1024" style="width:100%" value="<%=sURL%>" />
							<span id="documentURLpic">
							<% if sURL <> "" then%>
								<a href="<%=sURL%>" target="_newwindow">View Document</a>&nbsp;&nbsp;
							<% end if %>
							</span>
							<% if blnHasWP then %>
								<input type="button" class="button" value="Pick" onclick="showModal('Pick File', 65, 80, 'documentURL');" />
							<%else %>
								<input type="button" class="button" value="Browse..." onclick="javascript:doPicker('frmwaiver.sURL');" />
					 			&nbsp; &nbsp; 
					 			<input type="button" class="button" name="upload" value="Upload" onclick="openWin2('../docs/default.asp','_blank')" />
							<% end if %>
						</td>
					</tr>
					<tr>
						<td class="label">Applies to all Event\Classes?:</td><Td> <input type="checkbox" <%=sChecked%> name="bRequired" /></td>
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
' GETWAIVERINFO iWaiverID, sName, sType, sDescription, sBody, sUrl, blnRequired
'--------------------------------------------------------------------------------------------------
Sub GetwaiverInfo( ByVal iWaiverID, ByRef sName, ByRef sType, ByRef sDescription, ByRef sBody, ByRef sUrl, ByRef blnRequired )
	dim sSql, oRs

	sSql = "SELECT * FROM egov_class_waivers WHERE waiverid = " & iWaiverID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sName = oRs("waivername")
		sType = oRs("waivertype")
		sDescription = oRs("waiverdescription")
		sBody = oRs("waiverbody")
		sUrl = oRs("waiverurl")
		blnRequired = oRs("isrequired")
	End If

	oRs.close
	Set oRs = nothing

End Sub


%>


