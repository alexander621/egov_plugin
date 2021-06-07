<!--#include file='header.asp'-->
<% if not HasPermission("CanRegisterContact") then
response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleRegisterContact)
 end if %>   
  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_directory.jpg"></td>
      <td><font size="+1"><b><%=langRegisterContactTitle %></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<A HREF="display_contact.asp"><%=langBackToContactDisplay%></A></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
        <!--#include file='quicklink.asp'-->      
      </td>
      <td  valign="top">

		   <!--#include file='register_contactuser.html'-->
	</td>
  <td width='200'>&nbsp;</td>
    </tr>	
  </table>
           <!--#include file='footer.asp'-->   
 
		   
