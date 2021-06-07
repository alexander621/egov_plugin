<!-- #include file="../includes/common.asp" //-->

<%
Dim oCmd, sLinks, bShown

If Not (HasPermission("CanEditAnnouncements")) Then Response.Redirect "../default.asp"

If Request.Form("_task") = "newannounce" Then

  sTmp = Request.Form("Message")

  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "NewAnnouncement"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
    .Parameters.Append .CreateParameter("CreatorID", adInteger, adParamInput, 4, Session("UserID"))
    .Parameters.Append .CreateParameter("Subject", adVarChar, adParamInput, 50, Request.Form("Subject"))
    .Parameters.Append .CreateParameter("Message", adLongVarChar, adParamInput, Len(sTmp), sTmp)
    .Execute
  End With
  Set oCmd = Nothing

  Response.Redirect "../announcements"

End If
%>

<html>
<head>
  <title><%=langBSAnnouncements%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  
  <script language="Javascript">
  function verifySubmit()
  {
    if(document.all.NewAnnouncement.Subject.value == "" || document.all.NewAnnouncement.Message.value == "")
    {
      var msg="";
      if(document.all.NewAnnouncement.Subject.value == "")
        msg+="Subject cannot be blank.\n"
      if(document.all.NewAnnouncement.Message.value == "")
        msg+="Message cannot be blank.\n"
      alert(msg);
    }
    else
    {
      document.all.NewAnnouncement.submit();
    }
  }
  </script>
  
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabHome,1%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_home.jpg"></td>
      <td><font size="+1"><b><%=langAnnouncements%>: <%=langNew%></b></font><br><br></a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top" nowrap>

         <!-- START: QUICK LINKS MODULE //-->
        
        <%
        sLinks = "<div style=""padding-bottom:8px;""><b>" & langAnnouncementLinks & "</b></div>"

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
        <!-- START: NEW ANNOUNCEMENT -->
      <td colspan="2" valign="top">
        <form name="NewAnnouncement" method=post action="newannouncement.asp" method="post">
          <input type="hidden" name="_task" value="newannounce">

          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();">Cancel</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:verifySubmit();"><%=langCreate%></a></div>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
            <tr>
              <th align="left" colspan="2"><%=langNewAnnouncement%></th>
            </tr>
            <tr>
              <td width="1" valign="top"><%=langSubject%>:</td>
              <td><input type="text" name="Subject" style="width:400px;" maxlength="50"></td>
            </tr>
            <tr>
              <td width="1" valign="top"><%=langMessage%>:&nbsp;</td>
              <td><textarea name="Message" rows="15" style="width:400px;"></textarea></td>
            </tr>
          </table>
          <div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();">Cancel</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:verifySubmit();"><%=langCreate%></a></div>
        </form>
      </td>
        <!-- END: NEW ANNOUNCEMENT -->
    </tr>
  </table>
</body>
</html>