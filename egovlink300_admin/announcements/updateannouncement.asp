<!-- #include file="../includes/common.asp" //-->

<%
Dim oCmd, oRst, sSubject, intID, sMessage, sLinks, bShown

If Not HasPermission("CanEditAnnouncements") Then Response.Redirect "../default.asp"

If Request.Form("_task") <> "" Then

  sTmp = Request.Form("Message")

  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "UpdateAnnouncement"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("AnnounceID", adInteger, adParamInput, 4, Request.Form("id"))
    .Parameters.Append .CreateParameter("UserID", adInteger, adParamInput, 4, Session("UserID"))
    .Parameters.Append .CreateParameter("Subject", adVarChar, adParamInput, 50, Request.Form("Subject"))
    .Parameters.Append .CreateParameter("Message", adLongVarChar, adParamInput, Len(sTmp), sTmp)
    .Execute
  End With
  Set oCmd = Nothing

  Response.Redirect "../announcements"

Else

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
    intID = clng(oRst("AnnouncementID"))
    sSubject = oRst("Subject")
    sMessage = oRst("Message")
  End If
  
  If oRst.State=1 then oRst.Close
  Set oRst = Nothing  

End If
%>

<html>
<head>
  <title><%=langBSAnnouncements%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  
  <script language="Javascript">
  function verifySubmit()
  {
    if(document.all.UpdateAnnouncement.Subject.value == "" || document.all.UpdateAnnouncement.Message.value == "")
    {
      var msg="";
      if(document.all.UpdateAnnouncement.Subject.value == "")
        msg+="Subject cannot be blank.\n"
      if(document.all.UpdateAnnouncement.Message.value == "")
        msg+="Message cannot be blank.\n"
      alert(msg);
    }
    else
    {
      document.all.UpdateAnnouncement.submit();
    }
  }
  </script>
  
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%
  DrawTabs tabHome,1
  %>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_home.jpg"></td>
      <td><font size="+1"><b><%=langAnnouncements%>: <%=langUpdate%></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langBackToAnnouncementList%></a></td>
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
        <!-- START: UPDATE ANNOUNCEMENT -->
      <td colspan="2" valign="top">
        <form name="UpdateAnnouncement" method=post action="updateannouncement.asp" method="post">
          <input type="hidden" name="_task" value="update">
          <input type="hidden" name="id" value=<%=intID%>>

          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();">Cancel</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:verifySubmit();">Update</a></div>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
            <tr>
              <th align="left" colspan="2"><%=langUpdateAnnouncement%></th>
            </tr>
            <tr>
              <td width="1" valign="top"><%=langSubject%>:</td>
              <td><input type="text" name="Subject" style="width:400px;" maxlength="50" value="<%=sSubject%>"></td>
            </tr>
            <tr>
              <td width="1" valign="top"><%=langMessage%>:&nbsp;</td>
              <td><textarea name="Message" rows="15" style="width:400px;"><%=sMessage%></textarea></td>
            </tr>
          </table>
          <div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();">Cancel</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:verifySubmit();">Update</a></div>
        </form>
      </td>
        <!-- END: UPDATE ANNOUNCEMENT -->
    </tr>
  </table>
</body>
</html>
