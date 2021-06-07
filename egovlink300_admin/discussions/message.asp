<!-- #include file="../includes/common.asp" //-->
<%
Response.Buffer = True
Dim oCmd, oRst, sMsgs, iCount, iTotal, sTopic
Dim bShown, bCanEdit, sLinks, bLoggedIn

If Session("UserID") > 0 Then
  bLoggedIn = True
End If

bCanEdit = HasPermission("CanEditDiscussionMessages")

'-------------------------------------------------------------show message
Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
  .ActiveConnection = Application("DSN")
  .CommandText = "ListDiscussionMessage"
  .CommandType = adCmdStoredProc
  .Parameters.Append oCmd.CreateParameter("DiscussionID", adInteger, adParamInput, 4, Request.QueryString("msgid"))
End With
            
Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open oCmd
End With
Set oCmd = Nothing

sMsgs = ""
iCount = 1
iTotal = oRst.RecordCount

If Not oRst.EOF Then
  sTopic = oRst("Subject")

  Do While Not oRst.EOF
    sMsgs = sMsgs & "<table width=""100%"" border=""0"" cellpadding=""5"" cellspacing=""0"" class=""messagehead""><tr>"
    sMsgs = sMsgs & "<th style=""background-color:#ccddff; border-right:1px solid #336699;"">"
    
    If bLoggedIn Then
      sMsgs = sMsgs & "<a href=""newreply.asp?" & Request.QueryString() & "&topic=" & oRst("Subject") & """>" & langReply & "</a></th>"
    Else
      sMsgs = sMsgs & "&nbsp;</th>"
    End If
    
    If (bCanEdit Or oRst("UserID") = Session("UserID")) Then
      sMsgs = sMsgs & "<th width=""88%"" align=""right""><img src=""../images/small_delete.gif"" align=""absmiddle"">&nbsp;<a href=""deletemessage.asp?tid=" & Request.QueryString("tid") & "&tp=" & Request.QueryString("tp") & "&tn=" & Request.QueryString("tn") & "&delid=" & oRst("DiscussionID") & """>" & langDelete & "</a>&nbsp;&nbsp;</th></tr>"
    Else
      sMsgs = sMsgs & "<th width=""88%"" align=""right"">&nbsp;</th></tr>"
    End If
      
    sMsgs = sMsgs & "<tr><td colspan=""2""><table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%""><tr>"
    sMsgs = sMsgs & "<td style=""font-size:10px;"">From: <a href=""../dirs/display_individual.asp?userid=" & oRst("UserID") & """ target=""_userinfo"">" & oRst("Creator") & "</a></td>"
    sMsgs = sMsgs & "<td style=""font-size:10px;"" align=""right"">" & langSent & ": " & oRst("DateOfPost") & "</td></tr>"
    If oRst("HasBeenDeleted") = 0 Then
      sMsgs = sMsgs & "<tr><td colspan=""2"" style=""padding:8px;"" class=""rtf"">" & oRst("Message") & "</td></tr></table></td></tr></table><br>"
    Else
      sMsgs = sMsgs & "<tr><td colspan=""2"" style=""padding:8px;""><b><i>" & langMessageDeleted & "</i></b></td></tr></table></td></tr></table><br>"
    End If

    iCount = iCount + 1
    oRst.MoveNext
  Loop
  oRst.Close
Else
  sMsgs = "<font color=red><b>" & langErrorRetrievingTopic & "</b></font><br>"
End If

'-------------------------------------------------------------list replies
Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
  .ActiveConnection = Application("DSN")
  .CommandText = "ListDiscussionMessageReplies"
  .CommandType = adCmdStoredProc
  .Parameters.Append oCmd.CreateParameter("DiscussionID", adInteger, adParamInput, 4, Request.QueryString("msgid"))
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
  sMsgs = sMsgs & "<br><b>" & langReplies & ":</b><br><br><div style=""padding-left:25px;"">"
  Do While Not oRst.EOF
    sMsgs = sMsgs & "<a href=""message.asp?msgid=" & oRst("DiscussionID") & """>" & oRst("Subject") & "</a><br>"
    oRst.MoveNext
  Loop
  sMsgs = sMsgs & "</div>"
  oRst.Close
End If

Set oRst = Nothing
%>

<html>
<head>
  <title><%=langBSDiscussions%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabDiscussions,1%>

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_discussions.jpg"></td>
      <td><font size="+1"><b><%=langMessage%>:  <%= sTopic %></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">
      <a href="javascript:history.back();"><%
      If Request.QueryString("tid") > 0 Then
        Response.Write langBackToTopic
      Else
        Response.Write langBackToSearch
      End If
      %></a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
 
        <!-- START: QUICK LINKS MODULE //-->
        <%
        sLinks = "<div style=""padding-bottom:8px;""><b>" & langDiscussionLinks & "</b></div>"

        If HasPermission("CanCreateDiscussionGroup") Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/newdiscgroup.gif"" align=""absmiddle"">&nbsp;<a href=""newdiscgroup.asp"">" & langNewDiscussionGroup & "</a></div>"
          bShown = True
        End If
        
        If bShown Then
          Response.Write sLinks & "<br>"
        End If
        %>

        <% Call DrawQuicklinks(langSearchDiscussions, 1) %>
        <!-- END: QUICK LINKS MODULE //-->

      </td>
      <td colspan="2" valign="top">
        <div style="font-size:10px; padding-bottom:5px;"><img src="../images/arrow_back.gif" align="absmiddle"><font color="#999999"><%=langPrevMessage%></font>&nbsp;&nbsp;<font color="#999999"><%=langNextMessage%></font><img src="../images/arrow_forward.gif" align="absmiddle"></div>
        <%= sMsgs %>
      </td>
    </tr>
  </table>
</body>
</html>