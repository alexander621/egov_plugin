<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
Dim sSql, oRst, sTopics, iTopicID, iTopicPage, iCount, iTotal, sBgcolor, iNumMessages, iPage, sPage, sTopic, iNumMoreRecords, iSetPage
Dim bShown, bCanCreate, bCanEdit, sLinks

If Request.QueryString("tid") & "" = "" Then
  Response.Redirect "default.asp"
End If

sTopic = Request.QueryString("gn")
sPage = Request.QueryString("tp")
If sPage & "" <> "" Then
  iPage = clng(sPage)
End If

bCanCreate = HasPermission("CanCreateDiscussionTopics")
bCanEdit = HasPermission("CanEditDiscussionTopics")

sSql = "EXEC ListDiscussionTopics " & Session("OrgID") & "," & Request.QueryString("tid") & "," & Session("UserID") & "," & Session("PageSize") & "," & iPage
            
Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .ActiveConnection = Application("DSN")
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open sSql
  .ActiveConnection = Nothing
End With

iTopicID = Request.QueryString("tid")
iTopicPage = Request.QueryString("tp")

sTopics = ""
iCount = 1
iTotal = oRst.RecordCount

If Not oRst.EOF Then
  iNumMoreRecords = oRst("NumMoreRecords")
  sBgcolor = "#ffffff"

  Do While Not oRst.EOF
    If bCanEdit Then
      sTopics = sTopics & "<tr bgcolor=""" & sBgcolor & """><td><input type=""checkbox"" class=""listcheck"" name=""del_" & oRst("DiscussionID") & """></td>"
    Else
      sTopics = sTopics & "<tr bgcolor=""" & sBgcolor & """><td>&nbsp;</td>"
    End If

    sTopics = sTopics & "<td style=""padding:0px;""><img src=""../images/newdisc.gif"" border=""0"">&nbsp;</td>"

    iNumMessages = oRst("NumMessages")
    If iNumMessages > 1 Then
      iSetPage = oRst("NumMessages") \ Session("PageSize")
      If oRst("NumMessages") Mod Session("PageSize") > 0 Then
        iSetPage = iSetPage + 1
      End If
      sTopics = sTopics & "<td nowrap><a href=""thread.asp?gn=" & Request.QueryString("gn") & "&tid=" & iTopicID & "&tp=" & iTopicPage & "&mid=" & oRst("DiscussionID") & "&mp=" & iSetPage & """>" & oRst("Subject") & "</td>"
    Else
      sTopics = sTopics & "<td nowrap><a href=""messages.asp?gn=" & Request.QueryString("gn") & "&tid=" & iTopicID & "&tp=" & iTopicPage & "&mid=" & oRst("DiscussionID") & "&mp=1"">" & oRst("Subject") & "</td>"
    End If

    sTopics = sTopics & "<td align=""center"">" & iNumMessages & "</td>"
    sTopics = sTopics & "<td nowrap>" & oRst("Creator") & "&nbsp;&nbsp;</td>"
    sTopics = sTopics & "<td nowrap>" & MyFormatDateTime(oRst("LastReply"), " ") & "</td></tr>"

    If sBgcolor = "#ffffff" Then sBgcolor = "#eeeeee" Else sBgcolor = "#ffffff"
    iCount = iCount + 1
    oRst.MoveNext
  Loop
  oRst.Close
Else
  sTopics = "<tr><td colspan=6 style=""border:0px;"">" & langNoTopics & "</td></tr>"
End If

Set oRst = Nothing
%>

<html>
<head>
  <title><%=langBSDiscussions%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabDiscussions,1%>

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_discussions.jpg"></td>
      <td colspan="2"><font size="+1"><b><%=langDiscussionGroup%>: <%= sTopic %></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="../discussions"><%=langBackToGroupsList%></a></td>
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

        If bCanCreate Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/newdisc.gif"" align=""absmiddle"">&nbsp;<a href=""newtopic.asp?" & Request.QueryString() & """>" & langNewTopic & "</a></div>"
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
        <form name="DelTopics" action="deletetopics.asp?<%= Request.QueryString() %>" method="post">
          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/arrow_back.gif" align="absmiddle">
          
          <%
          If IPage > 1 Then
            Response.Write "<a href=""topics.asp?tid=" & Request.QueryString("tid") & "&tp=" & iPage-1 & "&tn=" & sTopic & """>" & langPrev & " " & Session("PageSize") & "</a>&nbsp;&nbsp;"
          Else
            Response.Write "<font color=""#999999"">" & langPrev & " " & Session("PageSize") & "</font>&nbsp;&nbsp;"
          End If

          If iNumMoreRecords > 0 Then
            Response.Write "<a href=""topics.asp?tid=" & Request.QueryString("tid") & "&tp=" & iPage+1 & "&tn=" & sTopic & """>" & langNext & " " & Session("PageSize") & "</a>"
          Else
            Response.Write "<font color=""#999999"">" & langNext & " " & Session("PageSize") & "</font>"
          End If
          %>

          <img src="../images/arrow_forward.gif" align="absmiddle">
          <%
          If bCanCreate Then
          %>
            &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/newdisc.gif" align="absmiddle">&nbsp;<a href="newtopic.asp?<%= Request.QueryString() %>"><%=langNewTopic%></a>
          <%
          End If
          If bCanEdit Then
          %>
            &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/small_delete.gif" align="absmiddle">&nbsp;<a href="javascript:if (confirm('<%=langConfirmDeleteTopic%>')){document.all.DelTopics.submit();}"><%=langDelete%></a>
          <%
          End If
          %>
          </div>

          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
            <tr>
              <th align=left>
              <%If bCanEdit then%>
              <input class="listCheck" type=checkbox name="chkSelectAll" onClick="selectAll('DelTopics', this.checked)">
              <%Else%>
              &nbsp;
              <%End If%>
              </th>
              <th width="1">&nbsp;</th>
              <th align="left" width="70%"><%=langTopic%></th>
              <th align="center" nowrap><%=langMessages%></th>
              <th align="left" nowrap><%=langStartedBy%></th>
              <th align="left" nowrap><%=langLastReply%></th>
            </tr>
            <%= sTopics %>
          </table>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>
