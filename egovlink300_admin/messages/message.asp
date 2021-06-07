<!-- #include file="../includes/common.asp" //-->
<%
'if not logged in dont allow access, no matter what
If Session("UserID") = 0 Then Response.Redirect RootPath

Dim oCmd, oRst, sFrom, sTo, sCc, sDate, sSubject, sMessage, sError, iNextMessage, iPrevMessage
Dim iCreator, sBackPage, iBackID

iBackID=clng(Request.QueryString("backid"))

Select Case iBackID
  Case MAILBOX_IN
    iCreator=0
    sBackPage="default.asp"
  Case MAILBOX_DRAFT
    iCreator=0
    sBackPage="drafts.asp"
  Case MAILBOX_SENT
    iCreator=1
    sBackPage="sentmail.asp"
  Case ELSE 'treat like MAILBOX_IN
    iCreator=0
    sBackPage="default.asp"
End Select

Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
  .ActiveConnection = Application("DSN")
  .CommandText = "GetMailMessage"
  .CommandType = adCmdStoredProc
  .Parameters.Append oCmd.CreateParameter("UserID", adInteger, adParamInput, 4, Session("UserID"))
  .Parameters.Append oCmd.CreateParameter("PmailID", adInteger, adParamInput, 4, Request.QueryString("pid"))
  .Parameters.Append oCmd.CreateParameter("BoxType", adInteger, adParamInput, 4, iBackID)
End With

Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open oCmd
End With
Set oCmd = Nothing

If Not oRst.EOF Then
  If Not oRst.EOF Then
    sFrom     = oRst("SentFrom")
    sTo       = oRst("SentTo")
    sCc       = oRst("SentCc")
    sDate     = oRst("Date")
    sSubject  = oRst("Subject")
    sMessage  = oRst("Message")
    iNextMessage= oRst("NextMessage")
    iPrevMessage= oRst("PrevMessage")
  End If
  oRst.Close
Else
  sError = "<font color=red><b>Error retrieving message.</b></font><br>Please contact your system administrator.<br><br><br>"
End If

Set oRst = Nothing
%>

<html>
<head>
  <title><%= langBSMessages %></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <style type="text/css">
  <!--
    .nomargin {margin:-4px;}
  //-->
  </style>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabMessages, 1%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_messages.jpg"></td>
      <td><font size="+1"><b><%=langMessage%>: <%= sSubject %></b></font><br>
      <%
      Select Case Right(Request.ServerVariables("HTTP_REFERER"),12)
        Case "sentmail.asp"
          Response.Write "<img src=""../images/arrow_2back.gif"" align=""absmiddle"">&nbsp;<a href=""sentmail.asp"">" & langBackTo & " " & langSentMail & "</a></td>"
        Case "s/drafts.asp"
          Response.Write "<img src=""../images/arrow_2back.gif"" align=""absmiddle"">&nbsp;<a href=""drafts.asp"">" & langBackTo & " " & langDrafts & "</a></td>"
        Case Else
          Response.Write "<img src=""../images/arrow_2back.gif"" align=""absmiddle"">&nbsp;<a href=""../messages"">" & langBackTo & " " & langInbox & "</a></td>"
      End Select
      %>
    </tr>
    <tr>
      <td valign="top">

        <!-- START: QUICK LINKS MODULE //-->
        <div style="padding-bottom:8px;"><b><%=langMessageLinks%></b></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="../images/newmail_small.jpg" width="16" height="16" align="absmiddle">&nbsp;<a href="compose.asp"><%=langNewMessage%></a></div>
        <br>
        <div style="padding-bottom:8px;"><b><%=langMessageBoxes%></b></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="../images/folder_closed.gif" width="16" height="16" align="absmiddle">&nbsp;<a href="../messages"><%=langInbox%></a></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="../images/folder_closed.gif" width="16" height="16" align="absmiddle">&nbsp;<a href="drafts.asp"><%=langDrafts%></a></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="../images/folder_closed.gif" width="16" height="16" align="absmiddle">&nbsp;<a href="sentmail.asp"><%=langSentMail%></a></div>
        <br>
        <div style="padding-bottom:3px;"><%=langSearchMessages%>:</div>
        <input type="text" style="background-color:#eeeeee; border:1px solid #000000; width:144px;"><br>
        <div class="quicklink" align="right"><a href="#"><img src="../images/shortcut.jpg" border="0"><%=langGo%></a>&nbsp;&nbsp;</div>
        <!-- END: QUICK LINKS MODULE //-->

      </td>
      <td valign="top">
        <%
        If sError & "" <> "" Then
          Response.Write sError
          Response.End
        Else
        %>
        <form name="frmDelete" action="delete.asp" method=post>
          <input type=hidden name="BackPage" value="<%=sBackPage%>">
          <input type=hidden name="del_<%=Request.QueryString("pid")%>" value="checked">
          <input type=hidden name="IsCreator" value="<%=iCreator%>">
        </form>
        <div style="font-size:10px; padding-bottom:5px;"><img src="../images/arrow_back.gif" align="absmiddle">
        <%
          If iPrevMessage > 0 Then
            Response.Write "<a href=""message.asp?pid=" & iPrevMessage & "&backid=" & iBackID & """>" & langPrevMessage & "</a>&nbsp;&nbsp;"
          Else
            Response.Write "<font color=""#999999"">" & langPrevMessage & "</font>&nbsp;&nbsp;"
          End If

          If iNextMessage > 0 Then
            Response.Write "<a href=""message.asp?pid=" & iNextMessage & "&backid=" & iBackID & """>" & langNextMessage & "</a>"
          Else
            Response.Write "<font color=""#999999"">" & langNextMessage & "</font>"
          End If
        %>
        
        <img src="../images/arrow_forward.gif" align="absmiddle">&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/reply.gif" align="absmiddle">&nbsp;<a href="compose.asp?pid=<%=Request.QueryString("pid")%>&type=<%=COMPOSE_TYPE_REPLY%>"><%=langReply%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/replyall.gif" align="absmiddle">&nbsp;<a href="compose.asp?pid=<%=Request.QueryString("pid")%>&type=<%=COMPOSE_TYPE_REPLYALL%>"><%=langReplyAll%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/forward.gif" align="absmiddle">&nbsp;<a href="compose.asp?pid=<%=Request.QueryString("pid")%>&type=<%=COMPOSE_TYPE_FOWARD%>"><%=langForward%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/small_delete.gif" align="absmiddle">&nbsp;<a href="javascript:document.frmDelete.submit();"><%=langDelete%></a></div>
        <table width="100%" border="0" cellpadding="5" cellspacing="0" class="messagehead">
          <tr>
            <th align="left">
              <table border="0" cellpadding="3" cellspacing="0">
                <tr>
                  <td style="font-weight:bold; color:#003366;"><%=langFrom%>:</td>
                  <td><%= sFrom %></td>
                </tr>
                <tr>
                  <td style="font-weight:bold; color:#003366;"><%=langTo%>:</td>
                  <td><%= sTo %></td>
                </tr>
                <% If sCc <> "" Then %>
                <tr>
                  <td style="font-weight:bold; color:#003366;"><%=langCc%>:</td>
                  <td><%= sCc %></td>
                </tr>
                <% End If %>
                <tr>
                  <td style="font-weight:bold; color:#003366;"><%=langDate%>:</td>
                  <td><%= sDate %></td>
                </tr>
                <tr>
                  <td style="font-weight:bold; color:#003366;"><%=langSubject%>:&nbsp;&nbsp;</td>
                  <td><%= sSubject %></td>
                </tr>
              </table>
            </th>
          </tr>
          <tr>
            <td style="font-family:Arial; font-size:13px;" class="rtf"><%= sMessage %></td>
          </tr>
        </table>
        <%
        End If
        %>
      </td>
    </tr>
  </table>
</body>
</html>
