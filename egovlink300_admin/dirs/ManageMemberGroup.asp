<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<html>
<head>
  <title><%=langBSCommittees%></title>
  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

  <script src="../scripts/tooltip_new.js"></script>
  
</head>

<!-- #include file="dir_constants.asp"-->

<body onload="javasript:opener.location.reload(true);" bgcolor="#c9def0">
<%
if trim(request.querystring("userid"))="" then
   response.write "<br />No User ID is entered, end program here"
   response.end
end if

thisname = request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 
rs.CursorType     = 3 

strSQL = "SELECT firstname, lastname, userid "
strSQL = strSQL & " FROM users u "
strSQL = strSQL & " WHERE u.userid = " & clng(request.querystring("userid"))

rs.Open strSQL

Title = "<strong>" & rs("firstname") & " " & rs("lastname") & "(" & rs("userid") & ")</strong>" & vbcrlf

rs.close
set rs = nothing
%>
<table border="0" cellspacing="0" cellpadding="5" align="center">
  <tr>
      <td align="center">
          <font size="+1">Manage User: <strong> <%=Title %> </strong></font>
      <td>
  </tr>
  <tr>
      <td>
<%
thisname = request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set rs1 = Server.CreateObject("ADODB.Recordset")
set rs1.ActiveConnection = conn
rs1.CursorLocation = 3 
rs1.CursorType     = 3 
set rs2 = Server.CreateObject("ADODB.Recordset")
set rs2.ActiveConnection = conn
rs2.CursorLocation = 3 
rs2.CursorType     = 3 

strSQL1 = "SELECT ugp.groupid, ugp.groupname "
strSQL1 = strSQL1 & " FROM UsersGroupsPlus ugp "
strSQL1 = strSQL1 & " WHERE ugp.userid = " & clng(request.querystring("userid"))

strSQL2 = "SELECT g.groupid,g.groupname "
strSQL2 = strSQL2 & " FROM Groups g "
strSQL2 = strSQL2 & " WHERE g.orgid = " & session("orgid")
strSQL2 = strSQL2 & " AND g.groupid NOT IN (select ugp.groupid "
strSQL2 = strSQL2 &                       " from users u "
strSQL2 = strSQL2 &                            " INNER JOIN UsersGroupsPlus ugp ON u.userid = ugp.userid "
strSQL2 = strSQL2 &                       " where g.orgid = " & session("orgid")
strSQL2 = strSQL2 &                       " and u.userid = " & clng(request.querystring("userid")) & ")"

rs1.Open strSQL1
rs2.Open strSQL2

response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""350"" class=""tablelist"">" & vbcrlf
response.write "  <tr>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          Already in Groups" & vbcrlf

call ExistingRoleList

response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          &nbsp;<a href=""javascript:document.c1.submit();""><img src=""../images/ieforward.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Remove Group');"" onmouseout=""tooltip.hide();""></a>" & vbcrlf
response.write "          <br />" & vbcrlf
response.write "          <a href=""javascript:document.r1.submit();""><img src=""../images/ieback.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Add Group');"" onmouseout=""tooltip.hide();""></a>" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          Available Groups" & vbcrlf

call TheRemainingMemberList

response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf
response.write "</table>" & vbcrlf

rs1.close
rs2.close
conn.close

set rs1  = nothing
set rs2  = nothing
set conn = nothing
%>
      </td>
  </tr>
</table>
<div align="center">
  <input type="button" name="closewindow" value="Close Window" onClick="parent.close();" />
</div>
</body>
</html>
<%
'------------------------------------------------------------------------------
sub ExistingRoleList
  response.write "<form name=""c1"" method=""POST"" action=""DeleteGroup.asp?userid=" & trim(request.querystring("userid")) & """>" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "<select size=""10"" width=""150"" style=""width:140px"" name=""ExistingList"" multiple>" & vbcrlf

  for i=0 to rs1.recordcount-1

   	response.write "  <option value=""" & rs1("GroupID") & """>" & rs1("GroupName") & "(" & rs1("GroupID") & ")" & "</option>" & vbcrlf

  		rs1.movenext
		next 

  response.write "</select>" & vbcrlf
  response.write "</p>" & vbcrlf
  response.write "</form>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub TheRemainingMemberList
  response.write "<form name=""r1"" method=""POST"" action=""AddGroup.asp?userid=" & trim(request.querystring("userid")) & """>" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "<select size=""10"" width=""150"" style=""width:140px"" name=""RemainingList"" multiple>" & vbcrlf

  for i=0 to rs2.recordcount-1

   	response.write "  <option value=""" & rs2("GroupID") & """>" & rs2("GroupName") & "(" & rs2("GroupID") & ")" & "</option>" & vbcrlf

				rs2.movenext
		next

  response.write "</select>" & vbcrlf
  response.write "</p>"	& vbcrlf
  response.write "</form>" & vbcrlf

end sub
%>
