<!--#include file='header.asp'-->
<% if not HasPermission("CanRegisterRole") then
response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleRegisterCommittee)
 end if %> 
  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_directory.jpg"></td>
      <td><font size="+1"><b><%= langRegisterRoleTitle%></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<A HREF="display_committee.asp"><%=langBackToCommittee%></A></td>
    </tr>
    <tr>
      <td valign="top">
        <!--#include file='quicklink.asp'-->      
      </td>
      <td  valign="top">

        <!--#include file='register_role.html'-->
	</td>

    </tr>
	
  </table>
           <!--#include file='footer.asp'-->