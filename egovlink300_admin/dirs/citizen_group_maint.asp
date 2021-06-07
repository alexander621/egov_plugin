<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="citizen_global_functions.asp" //-->
<!-- #include file="dir_constants.asp" //-->
<% 
 sLevel = "../" ' Override of value from common.asp

 if not userhaspermission(session("userid"), "groups") then
	   response.redirect sLevel & "permissiondenied.asp"
 end if

'Check for user permissions
 lcl_userhaspermission_groups        = userhaspermission(session("userid"), "groups")
 lcl_userhaspermission_edit_citizens = userhaspermission(session("userid"), "edit citizens")

 lcl_groupid     = 0
 lcl_screen_mode = "ADD"
 lcl_success     = request("success")
 lcl_onload      = "setMaxLength();"

 if request("groupid") <> "" then
    lcl_groupid     = request("groupid")
    lcl_screen_mode = "EDIT"
 end if

'Check for a screen message
 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if

'Set up the page variables
 lcl_groupname        = ""
 lcl_groupdescription = ""
 lcl_grouptype        = ""
 lcl_orgid            = session("orgid")

 if lcl_groupid > 0 then
  		sSQL = "SELECT groupname, "
    sSQL = sSQL & " groupdescription, "
    sSQL = sSQL & " grouptype, "
    sSQL = sSQL & " orgid "
    sSQL = sSQL & " FROM citizengroups "
    sSQL = sSQL & " WHERE groupid = " & clng(trim(lcl_groupid))

    set oGetGroupInfo = Server.CreateObject("ADODB.Recordset")
    oGetGroupInfo.Open sSQL, Application("DSN"), 3, 1

    if not oGetGroupInfo.eof then
       lcl_groupname        = oGetGroupInfo("groupname")
       lcl_groupdescription = oGetGroupInfo("groupdescription")
       lcl_grouptype        = oGetGroupInfo("grouptype")
       lcl_orgid            = oGetGroupInfo("orgid")
    end if

    oGetGroupInfo.close
    set oGetGroupInfo = nothing
 end if
%>
<html>
<head>
	 <title><%=langBSCommittees%></title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />

	 <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/textareamaxlength.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
<!--
function UpdateFamily( sUserId ) {
  location.href='../dirs/family_members.asp?userid=' + sUserId;
}

function CheckCommitteeField() {
  var lcl_false_count    = 0;

  if(document.getElementById("groupname").value=="") {
     inlineMsg(document.getElementById("groupname").id,'<strong>Required Field Missing: </strong> Group',10,'groupname');
     lcl_focus       = document.getElementById("groupname");
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("groupname");
  }

  if(lcl_false_count > 0) {
     lcl_focus.focus();
     return false;
  }else{
     document.getElementById("committee_maint").submit();
     return true;
  }
}

function doGroupsAccess() {
  x   = (screen.width-450)/2;
  y   = (screen.height-400)/2;
  win = window.open("ManageCommitteeAccess2.asp?groupid=<%=lcl_groupid%>", "disc_members", "width=450,height=350,status=0,menubar=0,scrollbars=1,toolbar=0,left="+x+",top="+y+",z-lock=yes");
  win.focus();
}

function openWin2(url, name) {
  popupWin = window.open(url, name,"resizable,width=380,height=300");
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

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">
<%
  response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" class=""menu"">" & vbcrlf
  response.write "    <tr>" & vbcrlf
  response.write "      <td background=""../images/back_main.jpg"">" & vbcrlf

  ShowHeader sLevel
%>
			<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" width=""100%"">" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td width=""60%"">" & vbcrlf
  response.write "          <font size=""+1""><strong>Citizen Groups: " & lcl_screen_mode & "</strong></font><br />" & vbcrlf
  response.write "          <input type=""button"" name=""backButton"" id=""backButton"" class=""button"" value=""Back to Citizen Groups"" onclick=""location.href='display_citizen_groups.asp'"" />" & vbcrlf
  response.write "	     </td>" & vbcrlf
  response.write "      <td align=""right"" nowrap=""nowrap"">" & vbcrlf
  response.write "          <span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf
                            displayButtons "TOP", lcl_screen_mode, RootPath, lcl_userhaspermission_groups, lcl_userhaspermission_edit_citizens
  response.write "          <table border=""0"" width=""100%"" class=""tablelist"" cellpadding=""5"" cellspacing=""0"">" & vbcrlf
  response.write "            <form name=""committee_maint"" id=""committee_maint"" method=""post"" action=""citizen_group_action.asp"">" & vbcrlf
  response.write "              <input type=""hidden"" name=""groupid"" id=""groupid"" value=""" & lcl_groupid & """ />" & vbcrlf
  response.write "              <input type=""hidden"" name=""grouptype"" id=""grouptype"" value=""" & lcl_grouptype & """ />" & vbcrlf
  response.write "              <input type=""hidden"" name=""orgid"" id=""orgid"" value=""" & lcl_orgid & """ />" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <th colspan=""2"" width=""100"" align=""left"">" & langUpdate & "&nbsp;" & langCommittee & "</th>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "          		<tr>" & vbcrlf
  response.write "          		    <td width=""10%"" valign=""top"">" & langGroup & ":</td>" & vbcrlf
  response.write "          		    <td width=""80%""><input type=""text"" name=""groupname"" id=""groupname"" value=""" & lcl_groupname & """ size=""50"" maxlength=""50"" onchange=""clearMsg('groupname');"" /></td>" & vbcrlf
  response.write "          		</tr>" & vbcrlf
  response.write "          		<tr>" & vbcrlf
  response.write "          		    <td width=""10%"" valign=""top"">" & langDescription & ":</td>" & vbcrlf
  response.write "          		    <td width=""80%""><textarea rows=""2"" cols=""50"" name=""groupdescription"" id=""groupdescription"" maxlength=""150"">" & lcl_groupdescription & "</textarea></td>" & vbcrlf
  response.write "          		</tr>" & vbcrlf
  response.write "            </form>" & vbcrlf
  response.write "          </table>" & vbcrlf
                            displayButtons "BOTTOM", lcl_screen_mode, RootPath, lcl_userhaspermission_groups, lcl_userhaspermission_edit_citizens
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#include file="../admin_footer.asp"-->  
<!--#include file='footer.asp'-->
<%
'------------------------------------------------------------------------------
sub displayButtons(iLocation, iScreenMode, iRootPath, iUserHasPermission_Groups, iUserHasPermission_EditCitizens)

  if iLocation = "BOTTOM" then
     lcl_style_padding = "padding-top"
  else
     lcl_style_padding = "padding-bottom"
  end if

  if iScreenMode = "ADD" then
     lcl_label_actionButton = "Add"
  else
     lcl_label_actionButton = "Save Changes"
  end if

  response.write "<div style=""" & lcl_style_padding & ":5px;"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & vbcrlf
  response.write "  <tr valign=""bottom"">" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <input type=""button"" name=""cancelButton"" id=""cancelButton"" class=""button"" value=""Cancel"" onclick=""history.back();"" />" & vbcrlf
  response.write "          <input type=""button"" name=""updateButton"" id=""updateButton"" class=""button"" value=""" & lcl_label_actionButton & """ onclick=""CheckCommitteeField();"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right"">" & vbcrlf
  response.write "          <strong>Registration Links</strong>" & vbcrlf

  if iUserHasPermission_Groups then
     response.write "  &nbsp;&nbsp;" & vbcrlf
     response.write "  <img src=""" & iRootPath & "images/newgroup.gif"" width=""16"" height=""16"" align=""absmiddle"" />&nbsp;" & vbcrlf
     response.write "  <a href=""" & iRootPath & "dirs/display_citizen_groups.asp"">All Citizen Groups</a>" & vbcrlf
  end if

  if iUserHasPermission_EditCitizens then
     response.write "  &nbsp;&nbsp;" & vbcrlf
     response.write "  <img src=""" & iRootPath & "images/newuser.gif"" width=""16"" height=""16"" align=""absmiddle"" />&nbsp;" & vbcrlf
     response.write "  <a href=""" & iRootPath & "dirs/display_citizen.asp"">All Citizens</a>" & vbcrlf
  end if

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf

end sub
%>
