<!-- #include file="../includes/common.asp" //-->
<%
Response.Buffer = True
Dim sSql, oRst, sMsgs, iCount, iTotal, sTopic, sPage, iPage, iNumMoreRecords, iTotalPages
Dim bShown, bCanEdit, sLinks, bLoggedIn

If Session("UserID") > 0 Then
  bLoggedIn = True
End If

sPage = Request.QueryString("mp")
If sPage & "" <> "" Then
  iPage = clng(sPage)
Else
  iPage = 1
End If

bCanEdit = HasPermission("CanEditDiscussionMessages")

sSql = "EXEC ListDiscussionMessages " & Request.QueryString("mid") & "," & Session("PageSize") & "," & iPage

Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .ActiveConnection = Application("DSN")
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open sSql
  .ActiveConnection = Nothing
End With

sMsgs = ""
iCount = ((iPage-1) * Session("PageSize")) + 1

If Not oRst.EOF Then
  sTopic = oRst("Subject")
  iNumMoreRecords = oRst("NumMoreRecords")
  iTotal = (iCount - 1) + oRst.RecordCount + iNumMoreRecords

  Do While Not oRst.EOF
    If oRst("HasBeenDeleted") = 1 Then
      sMsgs = sMsgs & "<table width=""100%"" border=""0"" cellpadding=""5"" cellspacing=""0"" class=""messagehead""><tr>"
      sMsgs = sMsgs & "<th align=""left"" style=""background-color:#ccddff; border-right:1px solid #336699;"">&nbsp;</th><th width=""88%"" align=""right"">" & langMessage & " " & iCount & " " & langOf & " " & iTotal & " " & langIn & " " & langDiscussion & "</th></tr>"
      sMsgs = sMsgs & "<tr><td colspan=2><i>(" & langMessageDeleted & ")</i></td></tr></table><br>"
    Else
      sMsgs = sMsgs & "<table width=""100%"" border=""0"" cellpadding=""5"" cellspacing=""0"" class=""messagehead""><tr>"
      sMsgs = sMsgs & "<th style=""background-color:#ccddff; border-right:1px solid #336699;"">"
      
      If bLoggedIn Then
        sMsgs = sMsgs & "<a href=""newreply.asp?" & Request.QueryString() & "&msgid=" & oRst("DiscussionID") & "&topic=" & oRst("Subject") & """>" &langReply & "</a></th>"
      Else
        sMsgs = sMsgs & "&nbsp;</th>"
      End If
      
      If (bCanEdit Or oRst("UserID") = Session("UserID")) Then
        sMsgs = sMsgs & "<th width=""88%"" align=""right""><img src=""../images/small_delete.gif"" align=""absmiddle"">&nbsp;<a href=""deletemessage.asp?tid=" & Request.QueryString("tid") & "&tp=" & Request.QueryString("tp") & "&tn=" & Request.QueryString("tn") & "&delid=" & oRst("DiscussionID") & """>Delete</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""message.asp?" & Request.QueryString() & "&msgid=" & oRst("DiscussionID") & """>" & langMessage & " " & iCount & "</a> " & langOf & " " & iTotal & " " & langIn & " " & langDiscussion & "</th></tr>"
      Else
        sMsgs = sMsgs & "<th width=""88%"" align=""right""><a href=""message.asp?" & Request.QueryString() & "&msgid=" & oRst("DiscussionID") & """>" & langMessage & " " & iCount & "</a> " & langOf & " " & iTotal & " " & langIn & " " & langDiscussion & "</th></tr>"
      End If
      
      sMsgs = sMsgs & "<tr><td colspan=""2""><table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%""><tr>"
      sMsgs = sMsgs & "<td style=""font-size:10px;"">From: <a href=""../dirs/display_individual.asp?userid=" & oRst("UserID") & """ target=""_userinfo"">" & oRst("Creator") & "</a></td>"
      sMsgs = sMsgs & "<td style=""font-size:10px;"" align=""right"">" & langSent & ": " & oRst("DateOfPost") & "</td></tr>"
      sMsgs = sMsgs & "<tr><td colspan=""2"" style=""padding:8px;"" class=""rtf"">" & oRst("Message") & "</td></tr></table></td></tr></table><br>"
    End If

    iCount = iCount + 1
    oRst.MoveNext
  Loop
  oRst.Close

  'find total number of pages for display reasons only
  iTotalPages = iNumMoreRecords \ Session("PageSize")
  If iNumMoreRecords Mod Session("PageSize") > 0 Then
    iTotalPages = iTotalPages + 1
  End If
  iTotalPages = iTotalPages + iPage

Else
  sMsgs = "<font color=red><b>" & langErrorRetrievingTopic & "</b></font><br>"
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
      <td><font size="+1"><b><%=langTopic%>: <%= sTopic %></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="topics.asp?gn=<%= Request.QueryString("gn") %>&tid=<%= Request.QueryString("tid") %>&tp=<%= Request.QueryString("tp") %>"><%=langBackToTopicsList%></a></td>
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
        <div style="font-size:10px; padding-bottom:5px;"><img src="../images/arrow_back.gif" align="absmiddle">
        
        <%
        If iPage > 1 Then
          Response.Write "<a href=""messages.asp?tid=" & Request.QueryString("tid") & "&tp=" & Request.QueryString("tp") & "&mid=" & Request.QueryString("mid") & "&mp=" & iPage-1 & """>" & langPrev & " " & Session("PageSize") & "</a>&nbsp;&nbsp;"
        Else
          Response.Write "<font color=""#999999"">" & langPrev & " " & Session("PageSize") & "</font>&nbsp;&nbsp;"
        End If

        If iNumMoreRecords > 0 Then
          Response.Write "<a href=""messages.asp?tid=" & Request.QueryString("tid") & "&tp=" & Request.QueryString("tp") & "&mid=" & Request.QueryString("mid") & "&mp=" & iPage+1 & """>" & langNext & " " & Session("PageSize") & "</a>"
        Else
          Response.Write "<font color=""#999999"">" & langNext & " " & Session("PageSize") & "</font>"
        End If
        %>

        <img src="../images/arrow_forward.gif" align="absmiddle">
        &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/view.gif" align="absmiddle">&nbsp;<a href="thread.asp?<%= Request.QueryString() %>"><%=langThreadedView%></a></div>

        <%= sMsgs %>
        <div id="pageinfo" style="margin-top:-7px; font-size:10px; color:#000000;">Page <%=sPage%> of <%=iTotalPages%></div>
      </td>
    </tr>
  </table>
</body>
</html>
