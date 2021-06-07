<!-- #include file="../includes/common.asp" //-->
<%
dim index, arrColors(2), hasShownLink
hasShownLink=false
arrColors(0)="#ffffff"
arrColors(1)="#eeeeee"
index=0
%>

<html>
<head>
  <title><%=langBSAdmin%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <SCRIPT LANGUAGE="JavaScript">
		<!--// This function will open page in new window
		function NewWindow(page) {
			OpenWin = window.open(page, "CtrlWindow", "width=420,height=120,status=no,toolbar=no,menubar=no,top=200,left=300,z-lock=no");
	   	}
		// End -->
  </SCRIPT>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabAdmin,1%>

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_admin.jpg"></td>
      <td><font size="+1"><b><%=langAdminLinks%></b></font><br><br></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
        <% Call DrawQuicklinks("", 1) %>
      </td>
      <td colspan="2" valign="top">
        <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
          <tr style="height:22px;">
            <th width="1">&nbsp;</th>
            <th align="left"><%=langAdminTask%>&nbsp;</th> 
            <th align="left" width="80%"><%=langDescription%>&nbsp;</th> 
          </tr>
          <%If HasPermission("CanEditAnnouncements") then%>
          <tr bgcolor="<%=arrColors(index)%>">
            <td><img src="../images/newannounce.gif"></td>
            <td nowrap><a href="../announcements"><%=langAdminEditAnnounce%></a></td>
            <td><%=langAdminEditAnnounceDesc%></td>
          </tr>
          <%
          hasShownLink=true
          index=1-index
          END IF
          %>
          <%If HasPermission("CanEditEvents") then%>
          <tr bgcolor="<%=arrColors(index)%>">
            <td><img src="../images/calendar.gif"></td>
            <td nowrap><a href="../events"><%=langAdminEditEvents%></a></td>
            <td><%=langAdminEditEventsDesc%></td>
          </tr>
          <%
          hasShownLink=true
          index=1-index
          END IF
          %>
          <%If HasPermission("CanEditFavorites") then%>
          <tr bgcolor="<%=arrColors(index)%>">
            <td><img src="../images/newfav.gif"></td>
            <td nowrap><a href="../favorites/default.asp?action=C"><%=langAdminEditFav%></a></td>
            <td><%=langAdminEditFavDesc%></td>
          </tr>
          <%
          hasShownLink=true
          index=1-index
          END IF
          %>
          <%If Session("UserID") > 0 then%>
          <tr bgcolor="<%=arrColors(index)%>">
            <td><img src="../images/newpersonalfav.gif"></td>
            <td nowrap><a href="../favorites/default.asp?action=U"><%=langAdminEditPersonalFav%></a></td>
            <td><%=langAdminEditPersonalFavDesc%></td>
          </tr>
          <%
          hasShownLink=true
          index=1-index
          END IF
          %>
          <%If HasPermission("CanEditDocuments") then%>
          <tr bgcolor="<%=arrColors(index)%>">
            <td><img src="../images/document_home.gif"></td>
            <td nowrap><a href="../docs"><%=langAdminEditDoc%></a></td>
            <td><%=langAdminEditDocDesc%></td>
          </tr>
          <%
          hasShownLink=true
          index=1-index
          END IF
          %>
          <%If HasPermission("CanRegisterCommittee") then%>
          <tr bgcolor="<%=arrColors(index)%>">
            <td><img src="../images/newgroup.gif"></td>
            <td nowrap><a href="../dirs/register_committee.asp"><%=langAdminRegDirectory%></a></td>
            <td><%=langAdminRegDirectoryDesc%></td>
          </tr>
          <%
          hasShownLink=true
          index=1-index
          END IF
          %>
          <%If HasPermission("CanRegisterUser") then%>
          <tr bgcolor="<%=arrColors(index)%>">
            <td><img src="../images/newuser.gif"></td>
            <td nowrap><a href="../dirs/register_normaluser.asp"><%=langAdminRegUser%></a></td>
            <td><%=langAdminRegUserDesc%></td>
          </tr>
          <%
          hasShownLink=true
          index=1-index
          END IF
          %>
          <%If Session("UserID") > 0 Then%>
          <tr bgcolor="<%=arrColors(index)%>">
            <td><img src="../images/newprofile.gif"></td>
            <td nowrap><a href="../dirs/update_user.asp?userid=<%= Session("UserID") %>"><%=langEditRegUser%></a></td>
            <td><%=langEditRegUserDesc%></td>
          </tr>
          <%
          hasShownLink=true
          index=1-index
          END IF
          %>
          <%If Session("UserID") > 0 Then%>
          <tr bgcolor="<%=arrColors(index)%>">
            <td><img src="../images/stocks.gif"></td>
            <td nowrap><a href="ChangePersonalSettings.asp"><%=langEditPersonalSettings%></a></td>
            <td><%=langEditPersonalSettingsDesc%></td>
          </tr>
          <%
          hasShownLink=true
          index=1-index
          END IF
          %>
          <%If HasPermission("CanRegisterContact") then%>
          <tr bgcolor="<%=arrColors(index)%>">
            <td><img src="../images/newcontact.gif"></td>
            <td nowrap><a href="../dirs/register_contactuser.asp"><%=langAdminRegContact%></a></td>
            <td><%=langAdminRegContactDesc%></td>
          </tr>
          <%
          hasShownLink=true
          index=1-index
          END IF
          %>

		  <%If HasPermission("CanEditRoles") then%>
          <tr bgcolor="<%=arrColors(index)%>">
            <td><img src="../images/newrole.gif"></td>
            <td nowrap><a href="../dirs/register_role.asp">New Role</a></td>
            <td>Create a new role</td>
          </tr>
          <%
          hasShownLink=true
          index=1-index
          END IF
          %>
          
          <%If HasPermission("CanEditCustomizations") then%>
          <tr bgcolor="<%=arrColors(index)%>">
            <td>
              <%If LCase(Application("Language"))= "english" Then%>
                <img src="../images/USA.gif">
              <%Else%>
                <img src="../images/SPAIN.gif">
              <% End If %>
            </td>
            <td nowrap><a href="../admin/ChangeLanguage.asp"><%=langAdminEditLang%></a></td>
            <td><%=langAdminCustomDescript%></td>
          </tr>
          <%
          hasShownLink=true
          index=1-index
          END IF
          %>
          
          <%If HasPermission("CanEditCustomizations") then%>
          <tr bgcolor="<%=arrColors(index)%>">
            <td>
			  <img src="../images/AboutMe.gif">
            </td>
            <td nowrap><a href="javascript:NewWindow('AboutBoardsite.asp');">About Me</a></td>
            <td>Information about this program.</td>
          </tr>
          <%
          hasShownLink=true
          index=1-index
          END IF
          %>


          <%If hasShownLink=false then%>
          <tr bgcolor="<%=arrColors(index)%>">
            <td>&nbsp;</td>
            <td nowrap colspan="3"><i><%=langAdminNoLinks%></i></td>
          </tr>
          <%
          END IF
          %>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>