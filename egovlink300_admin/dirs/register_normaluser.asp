<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<% 
 sLevel = "../"  'Override of value from common.asp

 if not UserHasPermission( Session("UserId"), "add users" ) Then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Check for org features
 lcl_orghasfeature_admin_locations = orghasfeature("admin locations")

'Check for user permissions
 lcl_userhaspermission_edit_organizational_groups = userhaspermission(session("userid"),"edit_organizational_groups")
 lcl_userhaspermission_staff_directory            = userhaspermission(session("userid"),"staff_directory")

 lcl_hidden = "HIDDEN"   'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide

'Set the widths and height of the columns
 lcl_width_column1 = "143"
 lcl_width_column2 = "336"
 lcl_column_height = "23"

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
<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

  <title>E-Gov Administration Consule {Add User}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/removespaces.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
<!--

function CheckMemberRegister() {
  var lcl_focus       = "";
  var lcl_false_count = 0;

  //Image
		if (document.getElementById("imagefilename").value!="") {
      lcl_imagefilename = document.getElementById("imagefilename").value.toUpperCase();
      lcl_ext_start_pos = lcl_imagefilename.indexOf(".");
      lcl_ext           = lcl_imagefilename.substr(lcl_ext_start_pos+1,lcl_imagefilename.length);

      if(<%=lcl_imgTypes%>) {
         clearMsg("findImageButton");
     }else{
         inlineMsg(document.getElementById("findImageButton").id,'<strong>Invalid Value: </strong> The image file extension is not valid. Valid file extensions:<br /><strong><%=lcl_imgTypesDisplay%></strong>',10,'findImageButton');
         lcl_false_count = lcl_false_count + 1;
         lcl_focus       = "imagefilename";
     }
  }else{
     clearMsg("findImageButton");
  }

  //Last Name
		if (document.getElementById("lastname").value == "") {
      inlineMsg(document.getElementById("lastname").id,'<strong>Required Field Missing: </strong>Last Name',10,'lastname');
      lcl_false_count = lcl_false_count + 1;
      lcl_focus       = "lastname";
  }else{
      clearMsg("lastname");
		}

  //First Name
		if (document.getElementById("firstname").value == "") {
      inlineMsg(document.getElementById("firstname").id,'<strong>Required Field Missing: </strong>First Name',10,'firstname');
      lcl_false_count = lcl_false_count + 1;
      lcl_focus       = "firstname";
  }else{
      clearMsg("firstname");
		}

  //Password
		if (document.getElementById("password").value == "") {
      inlineMsg(document.getElementById("password").id,'<strong>Required Field Missing: </strong>Password',10,'password');
      lcl_false_count = lcl_false_count + 1;
      lcl_focus       = "password";
  }else{
      clearMsg("password");
		}

  //User Name
		if (document.getElementById("username").value == "") {
      inlineMsg(document.getElementById("username").id,'<strong>Required Field Missing: </strong>User Name',10,'username');
      lcl_false_count = lcl_false_count + 1;
      lcl_focus       = "username";
  }else{
      clearMsg("username");
		}

  if(lcl_false_count > 0) {
     document.getElementById(lcl_focus).focus();
     return false;
  }else{
     document.RegisterNormal.submit();
   		return true;
  }
}

function validateDate(str) {
  var re;
  re = /^(0[1-9]|1[012]|[1-9])\/(3[01]|0[1-9]|[1-9]|[12]\d)\/\d{2}$/;        
  if (re.test(str) == true) {
     	return true;
 	} else {
      alert('Invalid date');
      //       document.all.RegisterNormal.birthday.focus();
	   		return false;
  }
}

function validatePhoneNumber(str) {
  var re;
  re = /^\(?\d{3}\)?([-\/\.])?\d{7}(-\d{4})?$/;       
  if (re.test(str) == true) {
      //  alert('Valid date');
    		return true;
		} else {
      alert("<%=langInvalidPhone%>");
      //       document.all.RegisterNormal.birthday.focus();
   			return false;
  }
}

function validateUserName(str) {
  var re;
  re = /^[0-9a-zA-Z_-]{3,20}$/;
  if (re.test(str) == true) {
      //  alert('Valid date');
    		return true;
		} else {
      alert('<%=langUserNameRequired%>');
      //  document.all.RegisterNormal.username.focus();
     	return false;
  }
}

function doPicker(sFormField) {
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

function storeCaret (textEl) {
  if (textEl.createTextRange)
      textEl.caretPos = document.selection.createRange().duplicate();
}

function insertAtCaret (textEl, text) {
  if (textEl.createTextRange && textEl.caretPos) {
      var caretPos = textEl.caretPos;
      caretPos.text =
      caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
      text + ' ' : text;
  }
   else
      textEl.value  = text;
}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "";
}
//--> 
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
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
<table border="0" cellpadding="10" cellspacing="0" width="100%">
  <tr>
      <td><font size="+1"><strong><%=langRegisterUserTitle%></strong></font></td>
      <td>&nbsp;</td>
  </tr>
  <tr valign="top">
      <td>
          <form method="post" name="RegisterNormal" action="insert_normaluser.asp">
            <input type="hidden" name="orgid" value="<%=Session("OrgID")%>" />

          <table border="0" cellspacing="0" cellpadding="0" width="100%">
            <tr>
                <td align="left" style="font-size:10px;">
                    <% displayButtons "TOP" %>
                </td>
                <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
            </tr>
          </table>

          <table border="0" class="tablelist" cellpadding="3" cellspacing="0" id="newuserinput">
            <tr>
                <th width="<%=lcl_width_column1%>" align="left"><%=langProperty%></th>
                <th width="<%=lcl_width_column2%>" align="left"><%=langValue%></th>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langUserName%></td>
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                    <input type=text name="username" id="username" size="15" maxlength="15" />
                    <font color="#FF0000">*</font>
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langPassword%></td>
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                    <input type=text name="password" id="password" size="15" maxlength="15" />
                    <font color="#FF0000">*</font>
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langFirstname%></td>
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                    <input type=text name="firstname" id="firstname" size="25" maxlength="25" />
                    <font color="#FF0000">*</font>
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="19"><%=langMiddleInitial%></td>
                <td width="<%=lcl_width_column2%>" height="19">
                    <input type="text" name="middleinitial" size="1" maxlength="1" />
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langLastName%></td>
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                    <input type="text" name="lastname" id="lastname" size="25" maxlength="25" />
                    <font color="#FF0000">*</font>
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="19"><%=langNickname%></td>
                <td width="<%=lcl_width_column2%>" height="19">
                    <input type="text" name="nickname" size="25" maxlength="25" />
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langTypeEmail%></td>
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                    <input type="text" name="email" size="50" maxlength="50" />
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langTypeEmail%>&nbsp;(Alternate)</td>
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                    <input type="text" name="email2" size="50" maxlength="50" />
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langcompanyname%></td>
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                    <input type="text" name="companyname" size="50" maxlength="50" />
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langJobTitle%></td>
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                		  <input type="text" name="jobtitle" size="50" maxlength="50" />
                </td>
            </tr>
<%
 'Organizational Group (Staff Directory)
  if lcl_userhaspermission_edit_organizational_groups AND lcl_userhaspermission_staff_directory then
     response.write "            <tr valign=""top"">" & vbcrlf
     response.write "                <td>Organizational Group: </td>" & vbcrlf
     response.write "                <td>" & vbcrlf
     response.write "                    <select name=""department"" size=""10"" MULTIPLE>" & vbcrlf
     response.write "                      <option value=""""></option>" & vbcrlf

     display_organizational_groups "", 0

     response.write "                    </select>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  else
     response.write "            <input type=""" & lcl_hidden & """ name=""department"" value="""" size=""4"" maxlength=""10"" />" & vbcrlf
  end if
%>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langHomeAddress%></td>
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                		  <input type="text" name="homeaddress" size="50" maxlength="250" />
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langBusinessAddress%></td>
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                		  <input type="text" name="businessaddress" size="50" maxlength="250" />
                </td>
            </tr>
<%
 'Location
  if lcl_orghasfeature_admin_locations then
     response.write "            <tr>" & vbcrlf
     response.write "                <td width=""" & lcl_width_column1 & """ height=""" & lcl_column_height & """>Location</td>" & vbcrlf
     response.write "                <td width=""" & lcl_width_column2 & """ height=""" & lcl_column_height & """>" & vbcrlf

     ShowLocations

     response.write "             			</td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  else
     response.write "           	<input type=""hidden"" name=""locationid"" value=""0"" />" & vbcrlf
  end if
%>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langHomePhone%></td> 
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                    <input type="text" name="homenumber" size="20" maxlength="20" />
             		 </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langBusinessPhone%></td> 
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                    <input type="text" name="businessnumber" size="20" maxlength="20" />
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langMobileNumber%></td> 
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                    <input type="text" name="mobilenumber" size="20" maxlength="20" />
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langPagerNumber%></td>
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                    <input type="text" name="pagernumber" size="20" maxlength="20" />
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langFaxNumber%></td> 
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                    <input type="text" name="faxnumber" size="20" maxlength="20" />
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langBirthday%></td>
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                    <input type="text" name="birthday" size="10" maxlength="10" />&nbsp;<i><%=langDateFormat%></i>
                </td>
            </tr>
            <tr>
                <td width="<%=lcl_width_column1%>" height="<%=lcl_column_height%>"><%=langWebPage%></td>
                <td width="<%=lcl_width_column2%>" height="<%=lcl_column_height%>">
                    <input type="text" name="webpage" size="50" maxlength="250" />
                </td>
            </tr>
<%
 'Display on Staff Directory
  if lcl_userhaspermission_edit_organizational_groups AND lcl_userhaspermission_staff_directory then
     response.write "            <tr>" & vbcrlf
     response.write "                <td>Display on Staff Directory: </td>" & vbcrlf
     response.write "                <td>" & vbcrlf
     response.write "                    <select name=""staff_dir_display"">" & vbcrlf
     response.write "                      <option value=""Y"">Yes</option>" & vbcrlf
     response.write "                      <option value=""N"" selected>No</option>" & vbcrlf
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
  response.write "                    <input type=""button"" name=""findImageButton"" id=""findImageButton"" value=""Find Image"" class=""button"" onclick=""clearMsg('findImageButton');doPicker('RegisterNormal.imagefilename');"" />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

%>
          </table>

          <% displayButtons "BOTTOM" %>
          <p><font color="#ff0000">* </font><%=lanRegister_requred %></p>
          </form>
      </td>
  </tr>
</table>

  </div>
</div>

<!--#include file="../admin_footer.asp"-->  
<!--#include file="footer.asp"-->
<%
'------------------------------------------------------------------------------
Sub ListOrganization( )
 	Dim conn, rs, openstr

	 set conn = Server.CreateObject("ADODB.Connection")
 	conn.Open Application("DSN")

	 set rs = Server.CreateObject("ADODB.Recordset")
 	set rs.ActiveConnection = conn

	 rs.CursorLocation = 3 'adUseClient
	 rs.CursorType     = 3 'adOpenStatic
 	openstr="select OrgID,OrgName from Organizations"
	 'response.write openstr

	 rs.Open openstr,,, 2 'adCmdTable
	 rs.movefirst

	 Do While Not rs.EOF
		   response.write "<option value="""&rs("orgid")&""">"&rs("orgname")&"</option>"
   		rs.movenext
 	Loop

	rs.close
	set rs = Nothing
	conn.close
	set conn = Nothing
End Sub

'------------------------------------------------------------------------------
Sub ShowLocations( )
 	Dim sSql, oLocations

	 sSql = "SELECT locationid, name "
  sSql = sSql & " FROM egov_class_location "
  sSql = sSql & " WHERE orgid = " & Session("OrgID")
  sSql = sSql & " ORDER BY name"

 	Set oLocations = Server.CreateObject("ADODB.Recordset")
 	oLocations.Open sSQL, Application("DSN"), 1, 3

	 response.write vbcrlf & "<select name=""locationid"">"
	 response.write vbcrlf & "<option value=""0"" selected=""selected"">Select a Location...</option>" & vbcrlf

  while NOT oLocations.EOF
		   response.write "<option value=""" & oLocations("locationid") & """>" & oLocations("name") & "</option>" & vbcrlf
   		oLocations.MoveNext
 	wend

	 response.write "</select>"

	 oLocations.close
	 Set oLocations = Nothing
End Sub

'--------------------------------------------------------------------------
sub display_organizational_groups(p_parent_org_group_id, p_org_level)

'Retrieve all of the sub organizational groups
 sSQLg = "SELECT org_group_id, org_name "
 sSQLg = sSQLg & " FROM egov_staff_directory_groups "
 sSQLg = sSQLg & " WHERE orgid=" & session("orgid")

 if p_parent_org_group_id = "" then
    sSQLg = sSQLg & " AND (parent_org_group_id IS NULL OR parent_org_group_id = 0) "
 else
    sSQLg = sSQLg & " AND parent_org_group_id = " & p_parent_org_group_id
 end if

 sSQLg = sSQLg & " AND active_flag = 'Y' "
 sSQLg = sSQLg & " ORDER BY UPPER(org_name) "

 set rsg = Server.CreateObject("ADODB.Recordset")
 rsg.Open sSQLg, Application("DSN"), 3, 1

 if not rsg.eof then
    while not rsg.eof
      'Determine how far to indent the org_name
       lcl_indent = (p_org_level * 5)

       lcl_indent_spaces = ""
       for x = 1 to lcl_indent
           lcl_indent_spaces = lcl_indent_spaces & "&nbsp;"
       next

       response.write "<option value=""" & rsg("org_group_id") & """>" & lcl_indent_spaces & rsg("org_name") & "</option>" & vbcrlf

      'Retrieve sub-organizational groups
       display_organizational_groups rsg("org_group_id"), p_org_level+1

       rsg.movenext
    wend
 end if
end sub

'------------------------------------------------------------------------------
sub displayButtons(iTopBottom)

  if iTopButton = "BOTTOM" then
     lcl_divStyle = "margin-top"
  else
     lcl_divStyle = "margin-bottom"
  end if

  response.write "<div style=""" & lcl_divStyle & ":5px;"">" & vbcrlf
  'response.write "<img src=""../images/cancel.gif"" align=""absmiddle"" />&nbsp;" & vbcrlf
  'response.write "<a href=""javascript:document.all.RegisterNormal.reset();"">" & langCancel & "</a>" & vbcrlf
  'response.write "&nbsp;&nbsp;&nbsp;&nbsp;" & vbcrlf
  'response.write "<img src=""../images/go.gif"" align=""absmiddle"" />&nbsp;" & vbcrlf
  'response.write "<a href=""javascript:document.RegisterNormal.submit();"" onclick=""return CheckMemberRegister();"">" & langCreate & "</a>" & vbcrlf
  response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""document.all.RegisterNormal.reset();"" />" & vbcrlf
  response.write "<input type=""button"" name=""createButton"" id=""createButton"" value=""Create"" class=""button"" onclick=""return CheckMemberRegister();"" />" & vbcrlf
  response.write "</div>" & vbcrlf

end sub
%>