<!-- #include file="../includes/common.asp" //-->
<!-- #include file="dir_constants.asp" //-->
<%
 lcl_groupid   = ""
 lcl_groupname = ""

 if request("groupid") <> "" then
    lcl_groupid = request("groupid")
 end if

 if lcl_groupid <> "" then
    sSQL = "SELECT groupname "
    sSQL = sSQL & " FROM citizengroups g "
    sSQL = sSQL & " WHERE g.groupid = " & lcl_groupid

   	set oGetCitizenGroupInfo = Server.CreateObject("ADODB.Recordset")
  	 oGetCitizenGroupInfo.Open sSQL, Application("DSN"), 3, 1

    if not oGetCitizenGroupInfo.eof then
       lcl_groupname = oGetCitizenGroupInfo("groupname")
    end if

    oGetCitizenGroupInfo.close
    set oGetCitizenGroupInfo = nothing
 end if
%>
<html>
<head>
  <title><%=langBSCommittees%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">

<script language="javascript">

function modifyList(iAction) {
  if(iAction == "REMOVE") {
     document.getElementById("c1").submit();
  } else {
     document.getElementById("r1").submit();
  }
}
</script>
</head>
<body onload="javasript:opener.location.reload(true);" bgcolor="#c9def0">
<%
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" width=""100%"" bgcolor=""#c9def0"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"" align=""center"" valign=""top"">" & vbcrlf
  response.write "          <font size=""10px"">Directory: <strong>" & lcl_groupname & "</strong></font><br />" & vbcrlf
  response.write "          <input type=""button"" name=""closeButton"" id=""closeButton"" class=""button"" value=""Close Window"" onclick=""self.close();"" />" & vbcrlf
  response.write "          <table cellpadding=""10"" cellspacing=""0"" width=""350"" border=""0"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      'CommitteeMemberList session("orgid"), lcl_groupid
                                      buildMemberList session("orgid"), lcl_groupid, "EXISTING"
  response.write "                </td>" & vbcrlf
  response.write "                <td align=""center"">" & vbcrlf
  response.write "                    <a href=""javascript:modifyList('REMOVE');""><img src=""../images/ieforward.gif"" align=""absmiddle"" border=""0"" /></a>" & vbcrlf
  response.write "                    <br /><br />" & vbcrlf
  response.write "                    <a href=""javascript:modifyList('ADD');""><img src=""../images/ieback.gif"" align=""absmiddle"" border=""0"" /></a>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      'TheRemainingMemberList session("orgid"), lcl_groupid
                                      buildMemberList session("orgid"), lcl_groupid, "AVAILABLE"
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub buildMemberList(iOrgID, iGroupID, iListType)

  if iListType = "EXISTING" then
     lcl_list_title        = "Existing Members"
     lcl_formname          = "c1"
     lcl_form_action       = "citizengroup_deletemember.asp?groupid=" & lcl_groupid
     lcl_dropdownlist_name = "committeelist"

     sSQL = "SELECT u.userid, userlname, userfname "
     sSQL = sSQL & " from egov_users u, vwCitizengroups ug "
     sSQL = sSQL & " WHERE u.userid = ug.citizenid "
     sSQL = sSQL & " AND u.orgid=" & iOrgID
     sSQL = sSQL & " AND ug.groupid = " & iGroupID
     sSQL = sSQL & " ORDER BY userlname "

  else  'AVAILABLE
     lcl_list_title        = "Available Members"
     lcl_formname          = "r1"
     lcl_form_action       = "citizengroup_addmember.asp?groupid=" & lcl_groupid
     lcl_dropdownlist_name = "OtherList"

     sSQL = "SELECT u.userid, userlname, userfname "
     sSQL = sSQL & " FROM egov_users u "
     sSQL = sSQL & " WHERE u.userid not in (select citizenid "
     sSQL = sSQL &                        " from vwCitizengroups ug "
     sSQL = sSQL &                        " where ug.groupid = " & iGroupID & ") "
     sSQL = sSQL & " AND u.orgid = " & iOrgID
     sSQL = sSQL & " AND u.userregistered > 0 "
     sSQL = sSQL & " ORDER BY userlname"

  end if

  response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""130"">" & vbcrlf
  response.write "  <form name=""" & lcl_formname & """ id=""" & lcl_formname & """ method=""post"" action=""" & lcl_form_action & """>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td height=""20""><strong>" & lcl_list_title & "</strong></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <p>" & vbcrlf
  response.write "          <select name=""" & lcl_dropdownlist_name & """ id=""" & lcl_dropdownlist_name & """ size=""15"" border=""0"" style=""width:140px"" multiple>" & vbcrlf

 	set oBuildMemberList = Server.CreateObject("ADODB.Recordset")
	 oBuildMemberList.Open sSQL, Application("DSN"), 3, 1

  if not oBuildMemberList.eof then
     do while not oBuildMemberList.eof
        lcl_membername = ""

        if trim(oBuildMemberList("userlname")) <> "" then
           if trim(oBuildMemberList("userfname")) <> "" then
              lcl_membername = trim(oBuildMemberList("userlname")) & ", " & trim(oBuildMemberList("userfname"))
           else
              lcl_membername = trim(oBuildMemberList("userlname"))
           end if
        else
           if trim(oBuildMemberList("userfname")) <> "" then
              lcl_membername = trim(oBuildMemberList("userlname"))
           end if
        end if

    	   if trim(lcl_membername) = "" then
           lcl_membername = "** " & oBuildMemberList("userid") & " **"
        end if

        response.write "            <option value=""" & oBuildMemberList("userid")& """>" & lcl_membername & "</option>" & vbcrlf

        oBuildMemberList.movenext
     loop
  end if

  oBuildMemberList.close
  set oBuildMemberList = nothing

 	response.write "            <option value="""">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>" & vbcrlf
  response.write "          </select>" & vbcrlf
  response.write "          </p>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  </form>"
  response.write "</table>" & vbcrlf
end sub
%>