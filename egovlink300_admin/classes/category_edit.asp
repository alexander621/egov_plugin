<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CATEGORY_MGMT.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 03/21/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/17/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.0   04/26/06   TERRY FOSTER - MADE FUNCTIONAL
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "categories" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 


' INITIALIZE VARIABLES
Dim sCategoryTitle, sSubTitle, sDescription, sURL, blnRoot
Dim icategoryID, sAltTag, sScreenMsg

' GET category ID
If request("categoryid") = "" OR NOT isnumeric(request("categoryid")) OR request("categoryid") = 0 Then
	' CREATE NEW category
	icategoryID = 0
	sTitle = "Add New Category"
	sLinkText = "Create Category"
Else
	' EDIT EXISTING category
	icategoryID = request("categoryid")
	sTitle = "Edit Category"
	sLinkText = "Save Changes"
End If

blnHasWP = hasWordPress()
sHomeWebsiteURL = getOrganization_WP_URL(session("orgid"), "OrgPublicWebsiteURL")

' GET category INFORMATION
GetcategoryInfo icategoryID 

If blnRoot Then
	sChecked= " checked=""checked"" "
End If

If request("msg") = "1" Then
	' Created
	sScreenMsg = "Category Created."
ElseIf request("msg") = "2" Then
	' Updated
	sScreenMsg = "Changes Saved."
Else
	sScreenMsg = ""
End If


%>


<html lang="en">
<head>
 	<title>E-Gov Administration Console</title>

	 <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

 	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" href="../global.css" />
 	<link rel="stylesheet" href="classes.css" />
 	<link rel="stylesheet" href="../recreation/facility.css" />


  	<script src="//code.jquery.com/jquery-1.12.4.js"></script>
   	<script src="//code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
	<!--#include file="../includes/wp-image-picker.asp"-->
	<script src="tablesort.js"></script>
	
	<script>
	<!--
	
	function displayScreenMsg(iMsg) 
	{
		if(iMsg!="") 
		{
			$("#screenMsg").html( "*** " + iMsg + " ***" );
			window.setTimeout("clearScreenMsg()", (10 * 1000));
		}
	}

	function clearScreenMsg() 
	{
		$("#screenMsg").html( "" );
	}
	
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

		$("#" + textEl.name + "pic").attr("src",text);
	}
	
	<% If sScreenMsg <> "" Then %>
	$( document ).ready(function() {
		displayScreenMsg( "<%= sScreenMsg %>" );
	});
	<% End If %>
	
//-->
	</script>

</head>
<body>
 
	<%'DrawTabs tabRecreation,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 


	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		
	<!--BEGIN: PAGE TITLE-->
	<p>
		<font size="+1"><strong>Recreation: <%=sTitle%></strong></font><br />
		<span id="screenMsg"></span><br />
		<input type="button" class="button" value="<< Back" onclick="location.href='category_mgmt.asp';" />
	</p>
	<!--END: PAGE TITLE-->


	<!--BEGIN: FUNCTION LINKS-->
	<div id="functionlinks">
			<input type="button" class="button" value="Cancel" onclick="location.href='category_mgmt.asp';" />&nbsp;&nbsp;
			<input type="button" class="button" value="<%=sLinkText%>" onclick="javascript:document.frmcategory.submit();" />&nbsp;&nbsp;
	</div>
	<!--END: FUNCTION LINKS-->


	<!--BEGIN: EDIT FORM-->
	<form name="frmcategory" action="category_save.asp" method="post">
	<input type="hidden" name="icategoryid" value="<%=icategoryID%>" />


	<table cellpadding="5" cellspacing="0" border="0" class="tableadmin" id="categorydetails">
		<tr>
			<th>Category Information</th>
		</tr>
		<tr>
			<td>
				<table>
					<tr>
						<td>Image:</td><td><input class="waiver imageurl" type="<%if blnHasWP then%>hidden<%else%>text<%end if%>" id="sURL" name="sURL" maxlength="1024" value="<%=sURL%>" style="display:block;">
							<img src="<%=sUrl%>" id="sURLpic" align="middle" width="180" height="180"  onerror="this.src = '../images/placeholder.png';" />
							<% if blnHasWP then %>
								&nbsp;&nbsp;<input type="button" class="button" value="Change" onclick="showModal('Pick Image',65,80,'sURL');" /-->
							<% else %>
							&nbsp; &nbsp; <input type="button" class="button" value="Browse..." onclick="javascript:doPicker('frmcategory.sURL');" />
							&nbsp; &nbsp; <input type="button" class="button" name="upload" value="Upload" onclick="openWin2('../docs/default.asp','_blank')" />
							<% end if %>
						</td>
					</tr>
					<tr>
						<td>Image Alt Tag:</td><td><input class=waiver type=text name=sAltTag maxlength=255 value="<%=sAltTag%>" /></td>
					</tr>
					<tr>
						<td>Title:</td><td><input class=waiver type=text name=sTitle maxlength=255 value="<%=sCategoryTitle%>" /></td>
					</tr>
					<tr>
						<td>SubTitle:</td><td><input class=waiver type=text name=sSubTitle maxlength=512 value="<%=sSubTitle%>" /></td>
					</tr>
					<tr>
						<td>Description:</td><td><textarea name=sDescription><%=sDescription%></textarea></td>
					</tr>
					<!--<tr>
						<td>Is Root?:</td><td><input <%=sChecked%> name=bRoot Type=checkbox></td>
					</tr>-->
				</table>
			</td>
		</tr>
	</table>


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
' Sub GETcategoryINFO(IcategoryID)
'--------------------------------------------------------------------------------------------------
Sub GetcategoryInfo( ByVal icategoryID)
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_class_categories WHERE categoryid = " & icategoryID 
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If NOT oRs.EOF Then
		sCategoryTitle = oRs("categorytitle")
		sSubTitle = oRs("categorysubtitle")
		sDescription = oRs("categorydescription")
		sURL = oRs("imgurl")
		blnRoot = oRs("isroot")
		sAltTag = oRs("imgalttag")
	End If

	oRs.close
	Set oRs = Nothing 

End Sub


%>


