<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="citizen_global_functions.asp" //-->
<!-- #include file="dir_constants.asp" //-->
<% 
 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"), "groups") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Check for org features
 lcl_orghasfeature_hasfamily = orghasfeature("hasfamily")

'Check for a screen message
 lcl_success = request("success")
 lcl_onload  = ""

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if
%>
<html>
<head>
	<title><%=langBSCommittees%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script src="../scripts/selectAll.js"></script>

	<script language="javascript">
	  <!--

function UpdateFamily( sUserId ) {
			location.href='../dirs/family_members.asp?userid=' + sUserId;
}

function confirmDelete() {
   if(confirm('Performing the action will remove all roles assigned to the directory.\nAre you sure you want to proceed?')) {
      document.getElementById("action").value = "DELETE";
      document.getElementById("DeleteCommittee").submit();
   }
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
  response.write "  <tr>" & vbcrlf
  response.write "      <td background=""../images/back_main.jpg"">" & vbcrlf

  ShowHeader sLevel
%>
		<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "<div id=""content"">" & vbcrlf
  response.write "	 <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" width=""100%"">" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td><font size=""+1""><strong>Citizen Groups</strong></font></td>" & vbcrlf
  response.write "      <td align=""right"">" & vbcrlf
  response.write "          <span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf
                            displayButtons
                            displayCitizenGroups session("orgid"), lcl_orghasfeature_hasfamily
                            displayButtons
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayCitizenGroups(iOrgID, iOrgHasFeature_HasFamily)

  lcl_bgcolor = "#ffffff"

  response.write "<div class=""shadow"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"" class=""tablelist"">" & vbrlf
  response.write "  <form name=""DeleteCommittee"" id=""DeleteCommittee"" method=""post"" action=""citizen_group_action.asp"">" & vbcrlf
  response.write "    <input type=""hidden"" name=""action"" id=""action"" value="""" />" & vbcrlf
  response.write "  <tr style=""height:25px;"" align=""left"">" & vbcrlf
  response.write "      <th><input type=""checkbox"" name=""chkSelectAll"" id=""chkSelectAll"" class=""listcheck"" onclick=""selectAll('DeleteCommittee', this.checked, 'delete')"" /></th>" & vbcrlf
 	response.write "      <th width=""1"">&nbsp;</th>" & vbcrlf
 	response.write "      <th colspan=""3"">"    & langDirectory   & "</th>" & vbcrlf
 	response.write "      <th>"                  & langDescription & "</th>" & vbcrlf
 	response.write "      <th align=""center"">" & langEntries     & "</th>" & vbcrlf
 	response.write "  </tr>" & vbcrlf

		sSQL = "SELECT groupid, orgid, groupname, groupdescription "
  sSQL = sSQL & " FROM citizengroups g "
  sSQL = sSQL & " WHERE g.orgid=" & iOrgID
  sSQL = sSQL & " ORDER BY groupname"

 	set oCitizenGroups = Server.CreateObject("ADODB.Recordset")
 	oCitizenGroups.Open sSQL, Application("DSN"), 3, 1

  if not oCitizenGroups.eof then
     'call statistics

  			do while not oCitizenGroups.eof
        lcl_bgcolor      = changeBGColor(lcl_bgcolor, "#eeeeee", "#ffffff")
        lcl_groupid      = oCitizenGroups("groupid")
        lcl_group_emails = getCommitteeEmails(iOrgID, lcl_groupid)
        lcl_totalUsers   = getTotalUsersInGroup(iOrgID, lcl_groupid, "", iOrgHasFeature_HasFamily)
        lcl_description  = ""

        if oCitizenGroups("groupdescription") <> "" then
           lcl_description = left(oCitizenGroups("groupdescription"),100)
        end if

     			response.write "  <tr bgcolor=" & lcl_bgcolor & ">" & vbcrlf
   			  response.write "      <td><input name=""delete"" id=""delete"" class=""listcheck"" type=""checkbox"" value=""" & lcl_groupid & """ /></td>" & vbcrlf
     			response.write "      <td><img src=""../images/newgroup.gif"" border=""0"" /></td>" & vbcrlf
     			response.write "      <td><a href='display_citizen.asp?groupid=" & lcl_groupid & "'>" & oCitizenGroups("groupname") & "</a></td>" & vbcrlf
     			response.write "      <td><a href=""citizen_group_maint.asp?groupid=" & lcl_groupid & """><img src=""../images/edit.gif"" align=""absmiddle"" border=""0"" title=""Click to Edit"" /></a></td>" & vbcrlf
     			response.write "      <td><a href=""" & lcl_group_emails & """><img src=""../images/newmail_small.gif"" border=""0"" title=""Click to Send Email"" /></a>" & vbcrlf
        response.write "      <td>" & lcl_description & "</td>" & vbcrlf
     			response.write "      <td align=""center"">" & lcl_totalUsers & "</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        oCitizenGroups.movenext
    	loop
  else
  			response.write "  <tr><td>No Citizen Groups Exist</td></tr>" & vbcrlf
  end if

  oCitizenGroups.close
  set oCitizenGroups = nothing

		response.write "  </form>" & vbcrlf
  response.write "</table>" & vbcrlf
		response.write "</div>" & vbcrlf

end sub 

'------------------------------------------------------------------------------
sub displayButtons()

  response.write "<div style=""padding-bottom:5px;"">" & vbcrlf
  response.write "  <input type=""button"" name=""deleteButton"" id=""deleteButton"" class=""button"" value=""Delete"" onclick=""confirmDelete();"" />" & vbcrlf
  response.write "  <input type=""button"" name=""newGroupButton"" id=""newGroupButton"" class=""button"" value=""New Group"" onclick=""location.href='citizen_group_maint.asp'"" />" & vbcrlf
	 response.write "</div>"

end sub
%>