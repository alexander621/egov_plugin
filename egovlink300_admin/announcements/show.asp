<!-- #include file="../includes/common.asp" //-->
<%
Dim oCmd, oRst, sSubject, sMessage, sEmail, dDate, intID, sLinks, bShown

Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
  .ActiveConnection = Application("DSN")
  .CommandText = "GetAnnouncement"
  .CommandType = adCmdStoredProc
  .Parameters.Append oCmd.CreateParameter("AnnounceID", adInteger, adParamInput, 4, Request.QueryString("id"))
End With
  
Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open oCmd
End With
Set oCmd = Nothing  
  
If not oRst.EOF then
  sSubject = oRst("Subject")
  sEmail = "<a href=""mailto:" & oRst("Email") & """>" & oRst("FullName") & "</a>"
  sMessage = AsciiToHtml(oRst("Message"))
  'sMessage = oRst("Message")
  dDate = oRst("ModifiedDate")
  intID=Request.QueryString("id")
End If
  
If oRst.State=1 then oRst.Close
Set oRst = Nothing
%>

<html>
<head>
  <title><%=langBSAnnouncements%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabHome,1%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_home.jpg"></td>
      <td><font size="+1"><b><%=langAnnouncement%>: <%= sSubject %></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="../"><%=langBackToStart%></a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top" nowrap>
       <!-- START: QUICK LINKS MODULE //-->
        
        <%
        sLinks = "<div style=""padding-bottom:8px;""><b>" & langAnnouncementLinks & "</b></div>"

        If HasPermission("CanEditAnnouncement") Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/newannounce.gif"" align=""absmiddle"">&nbsp;<a href=""newannouncement.asp"">" & langNewAnnouncement & "</a></div>"
          bShown = True
        End If
        
        If HasPermission("CanEditAnnouncement") Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/newannounce.gif"" align=""absmiddle"">&nbsp;<a href=""../announcements"">" & langEditAnnouncements & "</a></div>"
          bShown = True
        End If
        
        If bShown Then
          Response.Write sLinks & "<br>"
        End If
        %>

        <% Call DrawQuicklinks("", 1) %>
        <!-- END: QUICK LINKS MODULE //-->
      </td>

      <td colspan="2" valign="top">
        <form name="DelAnnouncement" method=post action="deleteannouncements.asp" method="post">
          <input type="hidden" name="del_<%= Request.QueryString("id") %>">
          <% If HasPermission("CanEditAnnouncements") Then %>
            <div style="font-size:10px; padding-bottom:5px;">
              <img src="../images/edit.gif" align="absmiddle">&nbsp;<a href="updateannouncement.asp?<%= Request.QueryString() %>"><%=langEdit%></a>
              &nbsp;&nbsp;&nbsp;&nbsp;
              <img src="../images/small_delete.gif" align="absmiddle">&nbsp;
              <a href="javascript:document.all.DelAnnouncement.submit();"><%=langDelete%></a>
            </div>
          <%End If%>
          <table border="0" cellpadding="5" cellspacing="0" width="95%" class="tablelist">
            <tr>
                <th align="left" ><%= sSubject %></th>
                <th align=right><%= dDate%></th>
            </tr>
            <tr class="subtablelist">
              <th align=left colspan="2">by <%= sEmail %></th>
            </tr>
            <tr>
              <td colspan=2 style="padding:10px;"><%= sMessage %></td>
            </tr>          
          </table>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>