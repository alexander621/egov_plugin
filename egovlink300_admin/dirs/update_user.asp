<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: update_user.asp
' AUTHOR: ????
' CREATED: ??/??/????
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0 ??/??/?? ????? ????? - INITIAL VERSION
' 1.1	02/13/2007	Steve Loar - Added locationid
' 1.2	02/21/2007	Steve Loar - Added class supervisor flag
' 1.3 ??/??/07 David Boyer - Added Email check to not allow the removal of an email if user is assigned to a form.
' 1.4 12/18/07 David Boyer - Reformatted the code for easier maintenance and added a few new features
                            '1. Now tracks the search criteria and page number from the main user list
                            '2. Now displays a "successful" message when the record has been updated.
' 1.5 12/18/07 David Boyer - Added Staff Directory options
' 1.6	01/17/2008	Steve Loar - Put class supervisor back as a checkbox and added permit inspector checkbox.
' 1.7 03/30/09 David Boyer - Added User Pic input field.
' 1.8 07/31/09 David Boyer - Added Delegate.
' 1.9	08/17/2009	Steve Loar - Added Rental supervisor pick for Menlo Park Rentals project
' 2.0	10/14/2011	Steve Loar - Changed to not fetch those users who have been flagged as deleted.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
sLevel = "../"  'Override of value from common.asp

if trim(request.querystring("userid")) = "" then
	response.redirect sLevel & "default.asp"
end if

'You can edit if you have edit permision or the user is editing themselves
If Not userhaspermission(session("userid"),"edit users" ) Then 
	lcl_edit_type = "ALL"
	if Not userhaspermission(session("userid"),"edit staff directory") then
		lcl_edit_type = "STAFF"
		if session("userid") <> clng(Trim(request("userid"))) then
			response.redirect sLevel & "permissiondenied.asp"
		else
			lcl_edit_type = "ALL"
		end if
	end if
end if

lcl_hidden = "hidden"   'Show/Hide hidden fields.  TEXT = Show, HIDDEN = Hide

'Retrieve the page parameters from the main list screen
lcl_sc_firstname = request("sc_firstname")
lcl_sc_lastname  = request("sc_lastname")
lcl_sc_orderby   = request("sc_orderby")
lcl_group_id     = request("groupid")

'Check for org features
lcl_orghasfeature_admin_locations         = orghasfeature("admin locations")
lcl_orghasfeature_class_supervisors       = orghasfeature("class supervisors")
lcl_orghasfeature_rental_supervisors      = orghasfeature("rental supervisors")
lcl_orghasfeature_permit_inspection_types = orghasfeature("permit inspection types")
lcl_orghasfeature_permit_review_types     = orghasfeature("permit review types")
lcl_orghasfeature_action_line             = orghasfeature("action line")
lcl_orghasfeature_edit_permits		  = orghasfeature("edit permits")

'Check for user permissions
lcl_userhaspermission_edit_users                 = userhaspermission(session("userid"),"edit users")
lcl_userhaspermission_user_permission            = userhaspermission(session("userid"),"user permission")
lcl_userhaspermission_edit_organizational_groups = userhaspermission(session("userid"),"edit_organizational_groups")
lcl_userhaspermission_staff_directory            = userhaspermission(session("userid"),"staff_directory")
lcl_userhaspermission_action_line                = userhaspermission(session("userid"),"action line")

'Set up the page variables
lcl_userid             = ""
lcl_orgid              = ""
lcl_username           = ""
lcl_password           = ""
lcl_isloggedin         = ""
lcl_enabled            = ""
lcl_pagesize           = ""
lcl_firstname          = ""
lcl_middleinitial      = ""
lcl_lastname           = ""
lcl_nickname           = ""
lcl_companyname        = ""
lcl_jobtitle           = ""
lcl_department         = ""
lcl_homeaddress        = ""
lcl_businessaddress    = ""
lcl_locationid         = ""
lcl_homenumber         = ""
lcl_businessnumber     = ""
lcl_mobilenumber       = ""
lcl_pagernumber        = ""
lcl_faxnumber          = ""
lcl_email              = ""
lcl_email2             = ""
lcl_webpage            = ""
lcl_birthday           = ""
lcl_isclasssupervisor  = ""
lcl_isrentalsupervisor = ""
lcl_staff_dir_display  = ""
lcl_ispermitinspector  = ""
lcl_ispermitreviewer   = ""
lcl_imagefilename      = ""
lcl_delegate_id        = ""

'Get the user information
sSql = "SELECT UserID,"
sSql = sSql & "OrgID,"
sSql = sSql & "Username,"
sSql = sSql & "Password,"
sSql = sSql & "IsLoggedIn,"
sSql = sSql & "Enabled,"
sSql = sSql & "PageSize,"
sSql = sSql & "FirstName,"
sSql = sSql & "MiddleInitial,"
sSql = sSql & "LastName,"
sSql = sSql & "Nickname,"
sSql = sSql & "CompanyName,"
sSql = sSql & "JobTitle,"
sSql = sSql & "Department,"
sSql = sSql & "HomeAddress,"
sSql = sSql & "BusinessAddress,"
sSql = sSql & "isnull(locationid,0) as locationid,"
sSql = sSql & "HomeNumber,"
sSql = sSql & "BusinessNumber,"
sSql = sSql & "MobileNumber,"
sSql = sSql & "PagerNumber,"
sSql = sSql & "FaxNumber,"
sSql = sSql & "Email,"
sSql = sSql & "Email2,"
sSql = sSql & "WebPage,"
sSql = sSql & "Birthday,"
sSql = sSql & "isclasssupervisor,"
sSql = sSql & "isrentalsupervisor,"
sSql = sSql & "staff_dir_display,"
sSql = sSql & "ispermitinspector,"
sSql = sSql & "ispermitreviewer,"
sSql = sSql & "imagefilename,"
sSql = sSql & "delegateid "
sSql = sSql & " FROM users u "
sSql = sSql & " WHERE userid = " & clng(trim(request.querystring("userid")))
sSql = sSql & " AND isdeleted = 0 AND orgid = " & session("orgid")

set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sSql, Application("DSN"), 1, 3

if not rs1.eof then
	lcl_userid             = rs1("userid")
	lcl_orgid              = rs1("orgid")
	lcl_username           = rs1("username")
	lcl_password           = rs1("password")
	lcl_isloggedin         = rs1("isloggedin")
	lcl_enabled            = rs1("enabled")
	lcl_pagesize           = rs1("pagesize")
	lcl_firstname          = trim(rs1("firstname"))
	lcl_middleinitial      = trim(rs1("middleinitial"))
	lcl_lastname           = trim(rs1("lastname"))
	lcl_nickname           = trim(rs1("nickname"))
	lcl_companyname        = trim(rs1("companyname"))
	lcl_jobtitle           = trim(rs1("jobtitle"))
	lcl_department         = trim(rs1("department"))
	lcl_homeaddress        = trim(rs1("homeaddress"))
	lcl_businessaddress    = trim(rs1("businessaddress"))
	lcl_locationid         = trim(rs1("locationid"))
	lcl_homenumber         = trim(rs1("homenumber"))
	lcl_businessnumber     = trim(rs1("businessnumber"))
	lcl_mobilenumber       = trim(rs1("mobilenumber"))
	lcl_pagernumber        = trim(rs1("pagernumber"))
	lcl_faxnumber          = trim(rs1("faxnumber"))
	lcl_email              = trim(rs1("email"))
	lcl_email2             = trim(rs1("email2"))
	lcl_webpage            = trim(rs1("webpage"))
	lcl_birthday           = trim(rs1("birthday"))
	lcl_isclasssupervisor  = rs1("isclasssupervisor")
	lcl_isrentalsupervisor = rs1("isrentalsupervisor")
	lcl_staff_dir_display  = rs1("staff_dir_display")
	lcl_ispermitinspector  = rs1("ispermitinspector")
	lcl_ispermitreviewer   = rs1("ispermitreviewer")
	lcl_imagefilename      = rs1("imagefilename")
	lcl_delegateid         = rs1("delegateid")
end if

rs1.close
set rs1 = nothing

'Setup the return button.
lcl_return_label   = ""
lcl_return_onclick = ""

if lcl_userhaspermission_edit_users then
	lcl_return_label   = langBackToUserDisplay
	lcl_return_onclick = "location.href='display_member.asp?sc_firstname=" & lcl_sc_firstname & "&sc_lastname=" & lcl_sc_lastname & "&groupid=" & lcl_group_id & "';"
else
	lcl_return_label   = langGoback
	lcl_return_onclick = "javascript:history.go(-1);"
end if

'Check for groups COMMENTED OUT 20180608, NOT SURE WHY THIS IS HERE
'sGroupList = displayGroups(request.querystring("userid"))

'Check for a screen message
lcl_success = request("success")
lcl_onload  = ""

if lcl_success <> "" then
	lcl_success = UCASE(lcl_success)
	lcl_msg     = setupScreenMsg(lcl_success)
	lcl_onload  = "displayScreenMsg('" & lcl_msg & "');"
end if

%>
<html>
<head>
	<title>E-Gov Administration Console {Update User}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script language="javascript" src="../scripts/selectAll.js"></script>
	<script language="javascript" src="../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="javascript">
	<!--

	function ChangeSupervisor( iUserId ) 
	{
		doAjax('setclasssupervisor.asp', 'userid=' + iUserId, '', 'get', '0');
	}

	function ChangeRentalSupervisor( iUserId ) 
	{
		doAjax('setrentalsupervisor.asp', 'userid=' + iUserId, '', 'get', '0');
	}

	function ChangePermitInspector( iUserId ) 
	{
		doAjax('setpermitinspector.asp', 'userid=' + iUserId, '', 'get', '0');
	}

	function ChangePermitReviewer( iUserId ) 
	{
		doAjax('setpermitreviewer.asp', 'userid=' + iUserId, '', 'get', '0');
	}

	function assignDelegate() 
	{
		lcl_userid     = document.getElementById("userid").value;
		lcl_delegateid = document.getElementById("delegateid").value;

		//Build the parameter string
		var sParameter = 'orgid='       + encodeURIComponent("<%=session("orgid")%>");
		sParameter    += '&userid='     + encodeURIComponent(lcl_userid);
		sParameter    += '&delegateid=' + encodeURIComponent(lcl_delegateid);
		sParameter    += '&isAjaxRoutine=Y';

		doAjax('savedelegate.asp', sParameter, 'displayScreenMsg', 'post', '0');
	}

	function confirmDelete() 
	{
		if (confirm("Are you sure you want to delete this user?")) 
		{ 
			location.href='delete_user.asp?userid=<%=CLng(trim(request.querystring("userid")))%>';
		}
	}

<%
	'Setup the valid image types
	lcl_imgTypesDisplay = "BMP,GIF,JPG,JPEG,PNG,TIF"

	lcl_imgTypes = ""
	lcl_imgTypes = lcl_imgTypes & "(lcl_ext==""BMP"")"
	lcl_imgTypes = lcl_imgTypes & "||(lcl_ext==""GIF"")"
	lcl_imgTypes = lcl_imgTypes & "||(lcl_ext==""JPG"")"
	lcl_imgTypes = lcl_imgTypes & "||(lcl_ext==""JPEG"")"
	lcl_imgTypes = lcl_imgTypes & "||(lcl_ext==""PNG"")"
	lcl_imgTypes = lcl_imgTypes & "||(lcl_ext==""TIF"")"
%>


	function validate() 
	{
		var lcl_focus       = "";
		var lcl_false_count = 0;

		//Image
		if (document.getElementById("imagefilename").value!="") 
		{
			lcl_imagefilename = document.getElementById("imagefilename").value.toUpperCase();
			lcl_ext_start_pos = lcl_imagefilename.indexOf(".");
			lcl_ext           = lcl_imagefilename.substr(lcl_ext_start_pos+1,lcl_imagefilename.length);

			if(<%=lcl_imgTypes%>) 
			{
				clearMsg("findImageButton");
			}
			else
			{
				inlineMsg(document.getElementById("findImageButton").id,'<strong>Invalid Value: </strong> The image file extension is not valid. Valid file extensions:<br /><strong><%=lcl_imgTypesDisplay%></strong>',10,'findImageButton');
				lcl_false_count = lcl_false_count + 1;
				lcl_focus       = "imagefilename";
			}
		}
		else
		{
			clearMsg("findImageButton");
		}

		//Password
		document.UpdateUser.password.value = removeSpaces(document.UpdateUser.password.value);

/*
		if (document.UpdateUser.password.value == "") 
		{
			inlineMsg(document.getElementById("password").id,'<strong>Required Field Missing: </strong>Password',10,'password');
			lcl_false_count = lcl_false_count + 1;
			lcl_focus       = "password";
		}
		else
		{
			clearMsg("password");
		}
		*/

		//UserName
		document.UpdateUser.username.value = removeSpaces(document.UpdateUser.username.value);

		if (document.UpdateUser.username.value == "") 
		{
			inlineMsg(document.getElementById("username").id,'<strong>Required Field Missing: </strong>User Name',10,'username');
			lcl_false_count = lcl_false_count + 1;
			lcl_focus       = "username";
		}
		else
		{
			clearMsg("username");
		}

		if(lcl_false_count > 0) 
		{
			document.getElementById(lcl_focus).focus();
			return false;
		}
<%
		'If a user has been assigned to a form then do not allow the email to be removed
		sSql = "SELECT COUNT(*) AS total_count "
		sSql = sSql & "FROM egov_action_request_forms "
		sSql = sSql & "WHERE (assigned_userID = " & clng(request("userid"))
		sSql = sSql & " OR assigned_userID2 = "  & clng(request("userid"))
		sSql = sSql & " OR assigned_userID3 = "  & clng(request("userid")) & ")"

		set oValidate = Server.CreateObject("ADODB.Recordset")
		oValidate.Open sSql, Application("DSN"), 0, 1

		if oValidate("total_count") > 0 then
			lcl_exists = "Y"
		else
			lcl_exists = "N"
		end if

		oValidate.Close
		set oValidate = nothing
%>

		if(("<%=lcl_exists%>"=="Y")&&(document.UpdateUser.email.value=="")) 
		{
			alert('The email address cannot be removed as the user is currently assigned to be notified when an Action Request is submitted');
			document.UpdateUser.email.value = document.UpdateUser.email_original.value;
			document.UpdateUser.email.focus();
			return;
		}
		else
		{
			document.UpdateUser.submit();
		}
	}

	function doPicker(sFormField)
	{
		//w = (screen.width - 350)/2;
		//h = (screen.height - 350)/2;
		w = 600;
		h = 400;
		l = (screen.AvailWidth/2)-(w/2);
		t = (screen.AvailHeight/2)-(h/2);

		pickerURL  = "../picker_new/default.asp";
		pickerURL += "?name=" + sFormField;
		pickerURL += "&folderStart=unpublished_documents";
		pickerURL += "&returnAsHTMLLink=N";
		pickerURL += "&displayDocuments=Y";
		pickerURL += "&returnOnlyFileName=Y";

		eval('window.open("' + pickerURL + '", "_picker", "width=' + w + ',height=' + h + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
	}

	function storeCaret (textEl) 
	{
		if (textEl.createTextRange)
			textEl.caretPos = document.selection.createRange().duplicate();
	}

	function insertAtCaret (textEl, text) 
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

	function displayScreenMsg(iMsg) 
	{
		if(iMsg!="") 
		{
			document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
			window.setTimeout("clearScreenMsg()", (10 * 1000));
		}
	}

	function clearScreenMsg() 
	{
		document.getElementById("screenMsg").innerHTML = "";
	}

//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="menu">
  <tr>
      <td background="../images/back_main.jpg">
		<% ShowHeader sLevel %>
       	<!--#Include file="../menu/menu.asp"--> 
      </td>
  </tr>
</table>

<!-- #include file="dir_constants.asp"-->

<div id="content">
 	<div id="centercontent">

<table border="0" cellpadding="10" cellspacing="0" style="width:900px;">
  <tr>
      <td>
          <table border="0" cellspacing="0" cellpadding="0" width="100%">
            <tr valign="top">
                <td>
                    <font size="+1"><strong><%=langUpdateUserAccount%></strong></font><br /><br /><br />
                    <input type="button" name="returnButton" id="returnButton" value="<< <%=lcl_return_label%>" class="button" onclick="<%=lcl_return_onclick%>" />
                </td>
                <td>
<% if lcl_orghasfeature_action_line AND lcl_userhaspermission_action_line then %>
                    <fieldset>
                      <legend>Select a Delegate&nbsp;</legend>
                      <select name="delegateid" id="delegateid" style="margin-top:10px;" onchange="assignDelegate();">
                        <option value=""></option>
                        <% DrawAdminUsersNew lcl_delegateid,"Y" %>
                      </select>
                      </legend>
                    </fieldset>
<%
   else
      response.write "<input type=""hidden"" name=""delegateid"" id=""delegateid"" value=""" & lcl_delegateid & """ />" & vbcrlf
   end if
%>
                </td>
            </tr>
          </table>
      </td>
  </tr>
  <tr>
      <td valign="top">
          <form method="post" name="UpdateUser" action="update_user_action.asp">
            <input type="<%=lcl_hidden%>" name="userid" id="userid" value="<%=lcl_userid%>" size="5" maxlength="4" />
            <input type="<%=lcl_hidden%>" name="orgid" value="<%=lcl_orgid%>" size="5" maxlength="4" />
            <input type="<%=lcl_hidden%>" name="isloggedin" value="<%=lcl_isloggedin%>" size="5" maxlength="4" />
            <input type="<%=lcl_hidden%>" name="enabled" value="<%=lcl_enabled%>" size="5" maxlength="4" />
            <input type="<%=lcl_hidden%>" name="pagesize" value="<%=lcl_pagesize%>" size="5" maxlength="4" />
            <input type="<%=lcl_hidden%>" name="sc_firstname" value="<%=lcl_sc_firstname%>" size="15" maxlength="25" />
            <input type="<%=lcl_hidden%>" name="sc_lastname" value="<%=lcl_sc_lastname%>" size="15" maxlength="25" />
            <input type="<%=lcl_hidden%>" name="sc_orderby" value="<%=lcl_sc_orderby%>" size="15" maxlength="25" />
            <input type="<%=lcl_hidden%>" name="groupid" value="<%=lcl_group_id%>" size="10" maxlength="10" />

          <table border="0" width="800" height="220" class="tablelist" cellpadding="3" cellspacing="0">
            <caption>
              <table border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr>
                    <td align="left" style="font-size:10px;">
                        <% displayButtons %>
                    </td>
                    <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
                </tr>
              </table>
            </caption>
            <tr>
               	<th width="175" height="23" align="left"><%=langProperty%></th>
               	<th width="525" height="23" align="left"><%=langValue%></th>
            </tr>
            <tr>
                <td>User Name: </td>
                <td><input type="text" name="username" id="username" value="<%=lcl_username%>" size="15" maxlength="32" onchange="clearMsg('username');" /></td>
            </tr>
            <tr>
                <td>Password: </td>
                <td><input type="password" name="password" id="password" placeholder="Enter a New Password Only" value="" size="25" maxlength="16" onchange="clearMsg('password');" /></td>
            </tr>
            <tr>
                <td>First Name: </td>
                <td><input type="text" name="firstname" value="<%=lcl_firstname%>" size="25" maxlength="25" /></td>
            </tr>
            <tr>
                <td>Middle Initial: </td>
                <td><input type="text" name="middleinitial" value="<%=lcl_middleinitial%>" size="1" maxlength="1" /></td>
            </tr>
            <tr>
                <td>Last Name: </td>
                <td><input type="text" name="lastname" value="<%=lcl_lastname%>" size="25" maxlength="25" /></td>
            </tr>
            <tr>
                <td>Nickname: </td>
                <td><input type="text" name="nickname" value="<%=lcl_nickname%>" size="25" maxlength="25" /></td>
            </tr>
            <tr>
                <td>Company Name: </td>
                <td><input type="text" name="companyname" value="<%=lcl_companyname%>" size="50" maxlength="50" /></td>
            </tr>
            <tr>
                <td>Job Title: </td>
                <td><input type="text" name="jobtitle" value="<%=lcl_jobtitle%>" size="50" maxlength="50" /></td>
            </tr>
<%
  if lcl_userhaspermission_edit_organizational_groups AND lcl_userhaspermission_staff_directory then
     response.write "            <tr valign=""top"">" & vbcrlf
     response.write "                <td>Organizational Group: </td>" & vbcrlf
     response.write "                <td>" & vbcrlf
     response.write "                    <select name=""department"" size=""10"" MULTIPLE>" & vbcrlf
     response.write "                      <option value=""""></option>" & vbcrlf

     display_organizational_groups "", 0, lcl_userid

     response.write "                    </select>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  else
     response.write "            <input type=""" & lcl_hidden & """ name=""department"" value="""" size=""4"" maxlength=""10"" />" & vbcrlf
  end if
%>
            <tr valign="top">
                <td>Home Address: </td>
                <td><textarea rows="2" cols="50" name="homeaddress"><%=lcl_homeaddress%></textarea></td>
            </tr>
            <tr valign="top">
                <td>Business Address: </td>
                <td><textarea rows="2" cols="50" name="businessaddress"><%=lcl_businessaddress%></textarea></td>
            </tr>
<%
		  if lcl_orghasfeature_admin_locations then
			 response.write "            <tr>" & vbcrlf
			 response.write "                <td>Location: </td>" & vbcrlf

			 ShowLocations lcl_locationid

			 response.write "            </tr>" & vbcrlf
		  else
			 response.write "            <tr style=""display: none"">" & vbcrlf
			 response.write "                <td colspan=""3""><input type=""" & lcl_hidden & """ name=""locationid"" value=""" & lcl_locationid & """ size=""5"" maxlength=""4"" /></td>" & vbcrlf
			 response.write "            </tr>" & vbcrlf
		  end if
%>
            <tr>
                <td>Home Phone: </td>
                <td><input type="text" name="homenumber" value="<%=lcl_homenumber%>" size="20" maxlength="20" /></td>
            </tr>
            <tr>
                <td>Business Phone: </td>
                <td><input type="text" name="businessnumber" value="<%=lcl_businessnumber%>" size="20" maxlength="20" /></td>
            </tr>
            <tr>
                <td>Mobile Phone: </td>
                <td><input type="text" name="mobilenumber" value="<%=lcl_mobilenumber%>" size="20" maxlength="20" /></td>
            </tr>
            <tr>
                <td>Pager Number: </td>
                <td><input type="text" name="pagernumber" value="<%=lcl_pagernumber%>" size="20" maxlength="20" /></td>
            </tr>
            <tr>
                <td>Fax Number: </td>
                <td><input type="text" name="faxnumber" value="<%=lcl_faxnumber%>" size="20" maxlength="20" /></td>
            </tr>
            <tr>
                <td>Email: </td>
                <td><input type="text" name="email" value="<%=lcl_email%>" size="50" maxlength="50" /></td>
            </tr>
            <tr>
                <td>Email (alternate): </td>
                <td><input type="text" name="email2" value="<%=lcl_email2%>" size="50" maxlength="50" /></td>
            </tr>
            <tr valign="top">
                <td>Web Page: </td>
                <td><textarea rows="2" cols="50" name="webpage"><%=lcl_webpage%></textarea></td>
            </tr>
            <tr>
                <td>Birthday: </td>
                <td><input type="text" name="birthday" value="<%=lcl_birthday%>" size="10" maxlength="10" /></td>
            </tr>
<%
 'Display on Staff Directory
  if lcl_userhaspermission_edit_organizational_groups AND lcl_userhaspermission_staff_directory then
     if lcl_staff_dir_display = "Y" then
        lcl_selected_yes = " selected=""selected"""
        lcl_selected_no  = ""
     else
        lcl_selected_yes = ""
        lcl_selected_no  = " selected=""selected"""
     end if

     response.write "            <tr>" & vbcrlf
     response.write "                <td>Display on Staff Directory: </td>" & vbcrlf
     response.write "                <td>" & vbcrlf
     response.write "                    <select name=""staff_dir_display"">" & vbcrlf
     response.write "                      <option value=""Y""" & lcl_selected_yes & ">Yes</option>" & vbcrlf
     response.write "                      <option value=""N""" & lcl_selected_no  & ">No</option>" & vbcrlf
     response.write "                    </select>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  else
     response.write "            <input type=""" & lcl_hidden & """ name=""staff_dir_display"" value=""N"" size=""1"" maxlength=""1"" />" & vbcrlf
  end if

 'Image Filename
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Image:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <input type=""input"" name=""imagefilename"" id=""imagefilename"" value=""" & lcl_imagefilename & """ size=""50"" maxlength=""500"" onchange=""clearMsg('findImageButton');"" />&nbsp;" & vbcrlf
  response.write "                    <input type=""button"" name=""findImageButton"" id=""findImageButton"" value=""Find Image"" class=""button"" onclick=""clearMsg('findImageButton');doPicker('UpdateUser.imagefilename');"" />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

	'Class Supervisor
	If lcl_orghasfeature_class_supervisors Then 
		If lcl_isclasssupervisor Then 
			lcl_selected = " checked=""checked"" "
		Else 
			lcl_selected = ""
		End If 

		response.write vbcrlf & "<tr>"
		response.write "<td>&nbsp;</td>"
		response.write "<td><input type=""checkbox"" name=""isclasssupervisor"" " & lcl_selected & "onClick=""ChangeSupervisor(" & lcl_userid & ");"" /> &nbsp; Class Supervisor</td>"
		response.write "</tr>"
	End If 

	'Rental Supervisor
	If lcl_orghasfeature_rental_supervisors Then 
		If lcl_isrentalsupervisor Then 
			lcl_selected = " checked=""checked"" "
		Else 
			lcl_selected = ""
		End If 

		response.write vbcrlf & "<tr>"
		response.write "<td>&nbsp;</td>"
		response.write "<td><input type=""checkbox"" name=""isrentalsupervisor"" " & lcl_selected & "onClick=""ChangeRentalSupervisor(" & lcl_userid & ");"" /> &nbsp; Rental Supervisor</td>"
		response.write "</tr>"
	End If 

	'Permit Inspector
	'If lcl_orghasfeature_permit_inspection_types Then 
	If lcl_orghasfeature_edit_permits Then 
		If lcl_ispermitinspector Then 
			lcl_selected = "checked=""checked"" "
		Else 
			lcl_selected = ""
		End If 

		response.write vbcrlf & "<tr>"
		response.write "<td>&nbsp;</td>"
		response.write "<td><input type=""checkbox"" name=""ispermitinspector"" " & lcl_selected & "onClick=""ChangePermitInspector(" & lcl_userid & ");"" /> &nbsp; Permit Inspector</td>"
		response.write "</tr>"
	End If 

	'Permit Reviewer
	'If lcl_orghasfeature_permit_review_types Then 
	If lcl_orghasfeature_edit_permits Then 
		If lcl_ispermitreviewer Then 
			lcl_selected = "checked=""checked"" "
		Else 
			lcl_selected = ""
		End If 

		response.write vbcrlf & "<tr>"
		response.write "<td>&nbsp;</td>"
		response.write "<td><input type=""checkbox"" name=""ispermitreviewer"" " & lcl_selected & "onClick=""ChangePermitReviewer(" & lcl_userid & ");"" /> &nbsp; Permit Reviewer</td>"
		response.write "</tr>"
	End If 

	response.write vbcrlf & "</table>"
	response.write "</td>"
	response.write "</tr>"

	response.write "  <tr valign=""top"">" & vbcrlf
	response.write "      <td>" & vbcrlf
	displayButtons
	response.write "      </td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
	response.write "</table>" & vbcrlf

	response.write "<input type=""hidden"" name=""username_o"" value=""" & lcl_username & """ />" & vbcrlf
	response.write "<input type=""hidden"" name=""email_original"" value=""" & lcl_email & """ />" & vbcrlf
	response.write "</form>" & vbcrlf

	response.write "      </td>" & vbcrlf
	response.write "      <td width=""200"">&nbsp;</td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
	response.write "</table>" & vbcrlf
	response.write "  </div>" & vbcrlf
	response.write "</div>" & vbcrlf
%>

<!--#Include file="../admin_footer.asp"-->  

<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf


'------------------------------------------------------------------------------
Sub ShowLocations( ByVal iLocationId )
	Dim sSql, oRs

	sSql = "SELECT locationid, name "
	sSql = sSql & " FROM egov_class_location "
	sSql = sSql & " WHERE orgid = " & Session("OrgID")
	sSql = sSql & " ORDER BY name"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write "<td><select name=""locationid"">"
	response.write vbcrlf & "<option value=""0"" "

	If clng(iLocationId) = clng(0) Then 
		response.write "selected=""selected"""
	End If 

	response.write ">Select a Location...</option>"

	Do While Not oRs.EOF
		If clng(iLocationId) = clng(oRs("locationid")) Then 
			lcl_selected_location = " selected=""selected"""
		Else 
			lcl_selected_location = ""
		End If 

		response.write vbcrlf & "<option value=""" & oRs("locationid") & """" & lcl_selected_location & ">" & oRs("name") & "</option>" 

		oRs.MoveNext
	Loop 

	response.write vbcrlf & "</select>"
	response.write "</td>"

	oRs.Close
	set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------
sub display_organizational_groups(p_parent_org_group_id, p_org_level, p_userid)

'Retrieve all of the sub organizational groups
 sSqlg = "SELECT org_group_id, org_name "
 sSqlg = sSqlg & " FROM egov_staff_directory_groups "
 sSqlg = sSqlg & " WHERE orgid=" & session("orgid")

 if p_parent_org_group_id = "" then
    sSqlg = sSqlg & " AND (parent_org_group_id IS NULL OR parent_org_group_id = 0) "
 else
    sSqlg = sSqlg & " AND parent_org_group_id = " & p_parent_org_group_id
 end if

 sSqlg = sSqlg & " AND active_flag = 'Y' "
 sSqlg = sSqlg & " ORDER BY UPPER(org_name) "

 set rsg = Server.CreateObject("ADODB.Recordset")
 rsg.Open sSqlg, Application("DSN"), 3, 1

 if not rsg.eof then
    while not rsg.eof
      'Determine how far to indent the org_name
       lcl_indent = (p_org_level * 5)

       lcl_indent_spaces = ""
       for x = 1 to lcl_indent
           lcl_indent_spaces = lcl_indent_spaces & "&nbsp;"
       next

      'Determine if the current record is selected
      'Retrieve all of the org_group_ids from the assignment table for the user
       sSqlu = "SELECT org_group_id "
       sSqlu = sSqlu & " FROM egov_staff_directory_usergroups "
       sSqlu = sSqlu & " WHERE userid = " & p_userid
       sSqlu = sSqlu & " AND org_group_id = " & rsg("org_group_id")

       set rsu = Server.CreateObject("ADODB.Recordset")
       rsu.Open sSqlu, Application("DSN"), 3, 1

       if not rsu.eof then
          lcl_selected = " selected"
       else
          lcl_selected = ""
       end if

       response.write "<option value=""" & rsg("org_group_id") & """" & lcl_selected & ">" & lcl_indent_spaces & rsg("org_name") & "</option>" & vbcrlf

      'Retrieve sub-organizational groups
       display_organizational_groups rsg("org_group_id"), p_org_level+1, p_userid

       rsg.movenext
    wend
 end if
end sub


'------------------------------------------------------------------------------
sub displayButtons()

	if lcl_userhaspermission_edit_users then
		response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='display_member.asp?sc_firstname=" & lcl_sc_firstname & "&sc_lastname=" & lcl_sc_lastname & "';"" />" & vbcrlf
	end if

	response.write "<input type=""button"" name=""updateButton"" id=""updateButton"" value=""Update"" class=""button"" onclick=""validate();"" />" & vbcrlf
	response.write "<input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" onclick=""confirmDelete();"" />" & vbcrlf

	if lcl_userhaspermission_user_permission then
		response.write "<input type=""button"" name=""userPermissionsButton"" id=""userPermissionsButton"" value=""User Permissions"" class=""button"" onclick=""location.href='../security/edit_user_security.asp?iuserid=" & clng(trim(request("userid"))) & "'"" />" & vbcrlf
	end if

end sub


'------------------------------------------------------------------------------
function displayGroups(iUserID)

  lcl_return = ""


  sSql = "SELECT g.GroupName, g.GroupID, ug.IsPrimaryGroup "
  sSql = sSql & " FROM Groups [g] LEFT OUTER JOIN UsersGroups [ug] ON ug.GroupID = g.GroupID "
  sSql = sSql & " AND ug.UserID = " & iUserID
  response.write sSql

  set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open sSql, Application("DSN"), 3, 1

  if not rs.eof then
     lcl_return = "<select name=""PrimaryGroup"">" & vbcrlf
     lcl_return = lcl_return & "  <option value=""0"">(None)</option>" & vbcrlf

     do while not rs.eof
        if rs("IsPrimaryGroup") then
           lcl_selected_primarygroup = " selected=""selected"""
        else
           lcl_selected_primarygroup = ""
        end if

        lcl_return = lcl_return & "  <option value=""" & rs("groupid") & """" & lcl_selected_primarygroup & ">" & rs("GroupName") & "</option>" & vbcrlf

        rs.MoveNext
     loop

     lcl_return = lcl_return & "</select>" & vbcrlf
  end if

  rs.close
  set rs = nothing

  displayGroups = lcl_return

end function


'------------------------------------------------------------------------------
Function setupScreenMsg( ByVal iSuccess )
	Dim lcl_return

	lcl_return = ""

	if iSuccess <> "" then
		iSuccess = UCASE(iSuccess)

		if iSuccess = "SU" then
			lcl_return = "Successfully Updated..."
		elseif iSuccess = "SA" then
			lcl_return = "Successfully Created..."
		elseif iSuccess = "SR" then
			lcl_return = "Successfully Reordered..."
		elseif iSuccess = "SD" then
			lcl_return = "Successfully Deleted..."
		end if
	end if

	setupScreenMsg = lcl_return

End Function



%>
