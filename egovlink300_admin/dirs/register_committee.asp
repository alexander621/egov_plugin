<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<% 
 sLevel = "../" ' Override of value from common.asp

 if NOT UserHasPermission( Session("UserId"), "departments" ) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 if UCASE(request("sAction")) = "CREATE" then
    createRecord()
 end if

 lcl_groupname        = request("groupname")
 lcl_groupdescription = request("groupdescription")
%>
<html>
<head>
  <title><%=langBSCommittees%></title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

  <script src="../scripts/selectAll.js"></script>
  <script src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="JavaScript">
<!--
	function CheckCommitteeField()
	{
		if (document.RegisterCommittee.groupname.value == "")
		{
//			alert("Group Name is required");
			 inlineMsg(document.getElementById("groupname").id,'<strong>Required Field Missing: </strong>Group Name',10,'groupname');
			document.RegisterCommittee.groupname.focus();
			return false;				
		}					
		return true;
	}
//-->
</script>

</head>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="document.getElementById('groupname').focus()">

<!-- #include file="dir_constants.asp"-->

<div id="content">
  <div id="centercontent">

<table border="0" cellpadding="10" cellspacing="0" width="100%">
  <tr>
      <td>
          <font size="+1"><b>New Department</b></font><br />
        		<img src="../images/arrow_2back.gif" align="absmiddle" />&nbsp;<a href="display_committee.asp">Back To Department List</a>
      </td>
      <td width="40%" align="right">
      <%
        lcl_message = ""

        if request("success") = "EXISTS" then
           lcl_message = "<b style=""color:#FF0000"">*** Group Name already exists... ***</b>"
        else
           lcl_message = "&nbsp;"
        end if

        if lcl_message <> "" then
           response.write lcl_message
        end if
      %>
      </td>
  </tr>
  <tr>
      <td valign="top">
          <form method="post" name="RegisterCommittee" action="register_committee.asp" onsubmit="return CheckCommitteeField();">
            <input type="hidden" name="orgid" value="<%= Session("OrgID") %>">

          <% displayButtons %>

          <table border="0" width="478" height="100"  class="tablelist" cellpadding="5" cellspacing="0">
            <tr>
                <th width="91" align="left"><%=langProperty%></th>
                <th width="284" align="left"><%=langValue%></th>
            </tr>
            <tr>
                <td width="91" height="22"><strong><%=langGroup%> Name</strong></td>
                <td width="284" height="22">
                    <input type="text" name="groupname" id="groupname" size="42" maxlength="50" value="<%=lcl_groupname%>"> <font color="#FF0000">*</font>
                </td>
            </tr>
            <tr>
               	<td width="91" height="22"><strong><%=langDescription%></strong></td>
                <td width="284" height="22" valign="middle">
                    <textarea name="groupdescription" rows="6" cols="42"><%=lcl_groupdescription%></textarea>
                </td>
            </tr>
           	<tr>
              		<td width="91" height="22"><strong>Group Type</strong></td>
              		<td width="284" height="22" valign="middle">
                  		<select name="grouptype">
                   			<option value="2" selected="selected">Department</option>
                  		</select>
              		</td>
            </tr>
          </table>

          <% displayButtons %>

          <p><font color="#FF0000">* </font><%=lanRegister_requred%></p>
          </form>
     	</td>
      <td width="200">&nbsp;</td>
  </tr>
</table>

  </div>
</div>

<!--#include file="../admin_footer.asp"-->  

</body>
</html>
<%
'------------------------------------------------------------------------------
sub ListOrganization
	 sSQL = "SELECT OrgID, OrgName FROM Organizations "

 	set rs = Server.CreateObject("ADODB.Recordset")
	 rs.Open sSQL, Application("DSN") , 3, 1

 	response.write "  <option value=""0"">Choose One</option>" & vbcrlf

  if not rs.eof then
    	do while not rs.eof
	      	response.write "  <option value=""" & rs("orgid") & """>" & rs("orgname") & "</option>" & vbcrlf
      		rs.movenext
    	loop
  end if

	rs.close
	set rs = nothing

end sub 

'------------------------------------------------------------------------------
sub displayButtons()

response.write "<div style=""font-size:10px; padding-bottom:5px;"">" & vbcrlf
response.write "  <input type=""button"" value=""" & langCancel & """ onClick=""document.all.RegisterCommittee.reset();"" />" & vbcrlf
response.write "  <input type=""submit"" value=""" & langCreate & """ name=""sAction"" onClick=""return CheckCommitteeField();"" />" & vbcrlf

'response.write "  <img src=""../images/cancel.gif"" align=""absmiddle"" />&nbsp;<a href=""javascript:document.all.RegisterCommittee.reset();"">" & langCancel & "</a>" & vbcrlf
'response.write "  &nbsp;&nbsp;&nbsp;&nbsp;" & vbcrlf
'response.write "  <img src=""../images/go.gif"" align=""absmiddle"">&nbsp;<a href=""javascript:document.all.RegisterCommittee.submit();"" onclick=""return CheckCommitteeField();"">" & langCreate & "</a>" & vbcrlf
response.write "</div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub createRecord()

  sSQL = "SELECT groupid FROM groups "
  sSQL = sSQL & " WHERE rtrim(ltrim(groupname)) = '" & dbsafe(request("groupname")) & "' "
  sSQL = sSQL & " AND orgid = " & session("orgid")
  sSQL = sSQL & " AND isInactive <> 1 "

  set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open sSQL, Application("DSN") , 3, 1

  if not rs.eof then
     response.redirect "register_committee.asp?success=EXISTS&groupname=" & request("groupname") & "&groupdescription=" & request("groupdescription") & "&grouptype=" & request("grouptype")
  else
     sSQLi = "INSERT INTO groups (orgid, groupname, groupdescription, groupimage, grouptype) values ("
     sSQLi = sSQLi & session("orgid") & ", "
     sSQLi = sSQLi & "'" & dbsafe(request("groupname")) & "', "
     sSQLi = sSQLi & "'" & dbsafe(request("groupdescription")) & "', "
     sSQLi = sSQLi & "'', "
     sSQLi = sSQLi & "'" & dbsafe(request("grouptype")) & "' "
     sSQLi = sSQLi & ")"

     set rsi = Server.CreateObject("ADODB.Recordset")
     rsi.Open sSQLi, Application("DSN") , 3, 1

     set rsi = nothing

     response.redirect "display_committee.asp?success=SN"
  end if

  set rs = nothing

end sub
%>