<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="dir_constants.asp" //-->
<%
 'dim conn,rs1,rs2,rs,strSQL1,strSQL2,strSQL,i,thisname,committeeName,memberstr
 dim thisname, sCommitteeName

 sGroupID = ""

 if request.querystring("groupid") <> "" then
    if not containsApostrophe(request.querystring("groupid")) then
       sGroupID = request.querystring("groupid")
       sGroupID = trim(sGroupID)
       sGroupID = clng(sGroupID)
    end if
 end if

 thisname       = request.servervariables("script_name")
 sCommitteeName = getGroupName(sGroupID)

'	if not HasPermission("CanEditCommittee") and not HasPermission("CanEdit"&CommitteeName) then
'		response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleEditCommittee)
'	end if

'Build body ONLOAD
 lcl_onload = "opener.location.reload(true);"
%>
<html>
<head>
  <title>E-Gov Administration Console {Maintain Departments}</title>
  <link href="../global.css" rel="stylesheet" type="text/css" />

<style type="text/css">
body
{
   background-color: #c9def0;
}

#departmentListsTable
{
   margin: 10px auto;
}

#departmentHeader
{
   margin: 5px 0px 10px; 0px;
   text-align: center;
   font-size: 1.25em;
}

#buttonCloseDiv
{
   text-align: center;
}

#buttonClose
{
   cursor: pointer;
}

.departmentName
{
   font-weight: bold;
}

#buttonArrows
{
   padding: 10px;
}

#buttonArrows img
{
   cursor: pointer;
}

#committeelist,
#OtherList
{
   width: 100%;
}
</style>

  <script type="text/javascript" src="../scripts/jquery-1.9.1.min.js"></script>

<script type="text/javascript">
$(document).ready(function() {
   $('#buttonExistingListRemove').click(function() {
      $('#c1').submit();
   });

   $('#buttonExistingListAdd').click(function() {
      $('#r1').submit();
   });

   $('#buttonClose').click(function() {
      window.close();
   });
});
</script>

</head>
<body onload="<%=lcl_onload%>">
<%
  thisname = request.servervariables("script_name")

  'set conn = Server.CreateObject("ADODB.Connection")
  'conn.Open Application("DSN")

  'set rs1 = Server.CreateObject("ADODB.Recordset")
  'set rs1.ActiveConnection = conn

  'set rs2 = Server.CreateObject("ADODB.Recordset")
  'set rs2.ActiveConnection = conn

  'rs1.CursorLocation = 3 
  'rs1.CursorType     = 3 

  'rs2.CursorLocation = 3 
  'rs2.CursorType     = 3 

  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" width=""100%"" bgcolor=""#c9def0"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf
  response.write "          <div id=""departmentHeader"">" & vbcrlf
  response.write "            Department: <span class=""departmentName"">" & sCommitteeName & "</span>" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "          <table id=""departmentListsTable"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      CommitteeMemberList session("orgid"), _
                                                          sGroupID
  response.write "                </td>" & vbcrlf
  response.write "                <td id=""buttonArrows"">" & vbcrlf
  response.write "                    <img id=""buttonExistingListRemove"" src=""../images/ieforward.gif"" align=""absmiddle"" border=""0"" />" & vbcrlf
  response.write "                    <br /><br />" & vbcrlf
  response.write "                    <img id=""buttonExistingListAdd"" src=""../images/ieback.gif"" align=""absmiddle"" border=""0"" />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      remainingMembersList session("orgid"), _
                                                           sGroupID
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "          <div id=""buttonCloseDiv"">" & vbcrlf
  response.write "            <input type=""button"" name=""buttonClose"" id=""buttonClose"" value=""Close Window"" />" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
%>
<!--#include file='footer.asp'-->
<%
'------------------------------------------------------------------------------
function getGroupName(iGroupID)

  dim lcl_groupid, sSQL, oGroupName, lcl_return

  lcl_groupid = ""
  lcl_return  = ""

  if iGroupID <> "" then
     if not containsApostrophe(iGroupID) then
        lcl_groupid = iGroupID
        lcl_groupid = trim(lcl_groupid)
        lcl_groupid = clng(lcl_groupid)
     end if
  end if

  sSQL = "SELECT groupname "
  sSQL = sSQL & " FROM groups g "
  sSQL = sSQL & " WHERE g.groupid = " & lcl_groupid

  set oGroupName = Server.CreateObject("ADODB.Recordset")
  oGroupName.Open sSQL, Application("DSN"), 3, 1

  if not oGroupName.eof then
     lcl_return = oGroupName("groupname")
  end if

  oGroupName.close
  set oGroupName = nothing

  getGroupName = lcl_return

end function

'------------------------------------------------------------------------------
sub CommitteeMemberList(iOrgID, _
                        iGroupID)

  dim lcl_groupid, lcl_orgid, lcl_existing_members, lcl_member_name, sSQL, rs1, i

  lcl_groupid          = 0
  lcl_orgid            = 0
  lcl_existing_members = ""
  lcl_member_name      = ""

  if iOrgID <> "" then
     if not containsApostrophe(iOrgID) then
        lcl_orgid = trim(iOrgID)
        lcl_orgid = clng(lcl_orgid)
     end if
  end if

  if iGroupID <> "" then
     if not containsApostrophe(iGroupID) then
        lcl_groupid = trim(iGroupID)
        lcl_groupid = clng(lcl_groupID)
     end if

     sSQL = "SELECT u.userid, "
     sSQL = sSQL & " lastname, "
     sSQL = sSQL & " firstname "
     sSQL = sSQL & " FROM users u, "
     sSQL = sSQL &      " usersgroups ug "
     sSQL = sSQL & " WHERE u.orgid = " & lcl_orgid
     sSQL = sSQL & " AND  u.userid = ug.userid "
     sSQL = sSQL & " AND ug.groupid = " & lcl_groupid
     sSQL = sSQL & " AND u.isdeleted = 0 "
     sSQL = sSQL & " ORDER BY lastname "

     set rs1 = Server.CreateObject("ADODB.Recordset")
     rs1.Open sSQL, Application("DSN"), 3, 1
  end if

	 for i=0 to rs1.recordcount-1
    	lcl_member_name = trim(rs1("lastname")) & ", " & trim(rs1("firstname"))

   		if trim(lcl_member_name) = "" OR trim(lcl_member_name) = ", " then
        lcl_member_name = "** " & rs1("userid") & " **"
     end if

    	if not UserIsRootAdmin( rs1("userid") ) then
     	  lcl_existing_members = lcl_existing_members & "<option value=""" & rs1("userid") & """>" & lcl_member_name & "</option>" & vbcrlf
  			end if

	  		rs1.movenext
		next

  response.write "<div>" & vbcrlf
  response.write "  <form name=""c1"" id=""c1"" method=""post"" action=""Committee_deletemember.asp?groupid=" & lcl_groupid & """>" & vbclf
  response.write "    <strong>Existing Members</strong><br />" & vbcrlf
  response.write "    <select name=""committeelist"" id=""committeelist"" size=""15"" multiple=""multiple"">" & vbcrlf
  response.write        lcl_existing_members
  response.write "    </select>" & vbcrlf
  response.write "  </form>" & vbcrlf
  response.write "</div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub remainingMembersList(iOrgID, _
                         iGroupID)

  dim lcl_groupid, lcl_remaining_members, lcl_member_name, lcl_orgid, i
  dim sSQL, rs2

  lcl_groupid           = 0
  lcl_orgid             = 0
  lcl_remaining_members = ""
  lcl_member_name       = ""

  if iGroupID <> "" then
     if not containsApostrophe(iGroupID) then
        lcl_groupid = trim(iGroupID)
        lcl_groupid = clng(lcl_groupID)
     end if

     if iOrgID <> "" then
        if not containsApostrophe(iOrgID) then
           lcl_orgid = trim(iOrgID)
           lcl_orgid = clng(lcl_orgid)
        end if
     end if

     sSQL = "SELECT * "
     sSQL = sSQL & " FROM users u "
     sSQL = sSQL & " WHERE u.userid NOT IN (SELECT userid "
     sSQL = sSQL &                        " FROM usersgroups ug "
     sSQL = sSQL &                        " WHERE ug.groupid = " & lcl_groupid & ")"
     sSQL = sSQL & " AND u.orgid=" & lcl_orgid
     sSQL = sSQL & " AND u.isdeleted = 0 "
     sSQL = sSQL & " ORDER BY lastname "

     set rs2 = Server.CreateObject("ADODB.Recordset")
     rs2.Open sSQL, Application("DSN"), 3, 1
  end if

  for i=0 to rs2.recordcount-1
   		lcl_member_name = trim(rs2("lastname")) & ", " & trim(rs2("firstname"))

  			if trim(lcl_member_name) = "" then
        lcl_member_name = "** "&rs2("userid") &" **"
     end if

  			if not UserIsRootAdmin( rs2("userid") ) then
			    	lcl_remaining_members = lcl_remaining_members & "<option value=" & rs2("userid") & ">" & lcl_member_name & "</option>" & vbcrlf
   		end if

  			rs2.movenext
		next

  response.write "<div>" & vbcrlf
  response.write "  <form name=""r1"" id=""r1"" method=""post"" action=""Committee_AddMember.asp?groupid=" & lcl_groupid & """>" & vbclf
  response.write "    <strong>Available Members</strong><br />" & vbcrlf
  response.write "    <select name=""OtherList"" id=""OtherList"" size=""15"" multiple=""multiple"">" & vbcrlf
  response.write        lcl_remaining_members
  response.write "    </select>" & vbcrlf
  response.write "  </form>" & vbcrlf
  response.write "</div>" & vbcrlf

end sub
%>