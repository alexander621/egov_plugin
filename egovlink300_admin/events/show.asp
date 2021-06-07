<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
Dim oCmd, oRst, sSubject, sMessage, sDate, iDuration, sDuration, sLinks, bShown

Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "GetEvent"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("EventID", adInteger, adParamInput, 4, Request.QueryString("id"))
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
  sDate = Split(MyFormatDateTime(oRst("EventDate"), ";" ), ";")
  sDate(1)=sDate(1) & " " & oRst("TZAbbreviation")

  iDuration = CLng(oRst("EventDuration"))
  If iDuration >= 10080 Then
    sDuration = sDuration & iDuration \ 10080 & " weeks, "
    iDuration = iDuration Mod 10080
  End If
  If iDuration >= 1440 Then
    sDuration = sDuration & iDuration \ 1440 & " days, "
    iDuration = iDuration Mod 1440
  End If
  If iDuration >= 60 Then
    sDuration = sDuration & iDuration \ 60 & " hours and "
    iDuration = iDuration Mod 60
  End If
  If iDuration > 0 Then
    sDuration = sDuration & iDuration & " minutes"
  End If

  If Right(sDuration, 2) = ", " Then
    sDuration = Left(sDuration, Len(sDuration)-2)
  ElseIf Right(sDuration, 4) = "and " Then
    sDuration = Left(sDuration, Len(sDuration)-4)
  End If

  sMessage = oRst("Message")
End If

If oRst.State=1 then oRst.Close
Set oRst = Nothing
%>

<html>
<head>
  <title><%=langBSEvents%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabCalendar,1%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_home.jpg"></td>
      <td><font size="+1"><b><%=langEvent%>: <%= sSubject %></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="../"><%=langBackToStart%></a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top" nowrap>
      <!-- START: QUICK LINKS MODULE //-->

        <%
        sLinks = "<div style=""padding-bottom:8px;""><b>" & langEventLinks & "</b></div>"

        If HasPermission("CanEditEvent") Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/calendar.gif"" align=""absmiddle"">&nbsp;<a href=""newevent.asp"">" & langNewEvent & "</a></div>"
          bShown = True
        End If

        If HasPermission("CanEditEvent") Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/calendar.gif"" align=""absmiddle"">&nbsp;<a href=""../events"">" & langEditEvents & "</a></div>"
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
        <form name="DelEvent" method=post action="deleteevents.asp" method="post">
          <input type="hidden" name="del_<%= Request.QueryString("id") %>">
          <% If HasPermission("CanEditEvents") Then %>
            <div style="font-size:10px; padding-bottom:5px;">
              <img src="../images/edit.gif" align="absmiddle">&nbsp;<a href="updateevent.asp?<%= Request.QueryString() %>"><%=langEdit%></a>
              &nbsp;&nbsp;&nbsp;&nbsp;
              <img src="../images/small_delete.gif" align="absmiddle">&nbsp;
              <a href="javascript:document.all.DelEvent.submit();"><%=langDelete%></a>
            </div>
          <%End If%>
          <table border="0" cellpadding="5" cellspacing="0" width="100%" class="tableadmin">
            <tr>
              <th colspan="2" align="left">Event Information</th>
            </tr>
            <tr>
              <td style="color:#336699;"><%=langEvent%>:&nbsp;&nbsp;&nbsp;</td>
              <td width="100%"><%= sSubject %></td>
            </tr>
            <tr bgcolor="#eeeeee">
              <td style="color:#336699;">Date:</td>
              <td><%= sDate(0) %></td>
            </tr>
            <tr>
              <td style="color:#336699;">Time:</td>
              <td><%= sDate(1) %></td>
            </tr>
            <% If Len(sDuration) > 0 Then %>
              <tr bgcolor="#eeeeee">
                <td style="color:#336699;"><%=langDuration%>:&nbsp;&nbsp;</td>
                <td><%= sDuration %></td>
              </tr>
            <% End If %>
            <tr>
              <td colspan="2" style="border-top:1px solid #336699; padding:10px;"><%= AsciiToHtml(sMessage) %></td>
            </tr>
          </table>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>