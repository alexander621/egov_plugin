<!-- #include file="../includes/common.asp" //-->
<%
Dim sSql, oRst, sMsgs, iTopicID, iTopicPage, iCount, iTotal, sBgcolor, sTopic, sPage, iPage, iNumMoreRecords, bDeleteShown, iTotalPages
Dim bShown, bCanCreate, bCanEdit, sLinks

sPage = Request.QueryString("mp")
If sPage & "" <> "" Then
  iPage = clng(sPage)
Else
  iPage = 1
End If

bCanEdit = HasPermission("CanEditDiscussionMessages")

sSql = "EXEC ListDiscussionThread " & Request.QueryString("mid") & "," & Session("PageSize") & "," & iPage
 
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

sMsgs = ""
iCount = 1
iTotal = oRst.RecordCount

If Not oRst.EOF Then
  bDeleteShown = False
  sTopic = oRst("Subject")
  sBgcolor = "#ffffff"
  iNumMoreRecords = oRst("NumMoreRecords")

  Do While Not oRst.EOF
    If oRst("HasBeenDeleted") = 1 Then
      sMsgs = sMsgs & "<tr bgcolor=""" & sBgcolor & """><td width=1>&nbsp;</td><td colspan=""3"" " & sStyle & "><img src=""../images/spacer.gif"" width=""" & oRst("PostLevel") * 16 & """ height=""1"" align=""absmiddle""><i>(" & langMessageDeleted & ")</i></td>"
    Else
      If (bCanEdit Or oRst("CreatorID") = Session("UserID")) Then
        sMsgs = sMsgs & "<tr bgcolor=""" & sBgcolor & """><td width=1><input type=""checkbox"" class=""listcheck"" name=""del_" & oRst("DiscussionID") & """></td>"
        bDeleteShown = True
      Else
        sMsgs = sMsgs & "<tr bgcolor=""" & sBgcolor & """><td width=1>&nbsp;</td>"
      End If
      sMsgs = sMsgs & "<td><img src=""../images/spacer.gif"" width=""" & oRst("PostLevel") * 16 & """ height=""1"" align=""absmiddle"">"
      sMsgs = sMsgs & "<a href=""message.asp?" & Request.QueryString() & "&msgid=" & oRst("DiscussionID") & """>" & oRst("Subject") & "</td>"
      sMsgs = sMsgs & "<td>" & oRst("Creator") & "</td>"
      sMsgs = sMsgs & "<td nowrap>" & oRst("DateOfPost") & "</td>"
    End If

    If sBgcolor = "#ffffff" Then sBgcolor = "#eeeeee" Else sBgcolor = "#ffffff"
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
  sMsgs = "<tr><td colspan=6 style=""border:0px;"">" & langNoTopics & "</td></tr>"
End If

Set oRst = Nothing
%>

<html>
<head>
  <title><%=langBSDiscussions%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
  <script language="Javascript">
  <!--
    function doReply() {
      x = (screen.width - 475)/2;
      y = (screen.height - 515)/2;
      window.open("htmleditor/newdiscmsg.html", "newreply", "width=475,height=515,scrollbars=no,status=no,toolbar=no,menubar=no,left="+x+",top="+y);
    }
  //-->
  </script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabDiscussions,1%>

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_discussions.jpg"></td>
      <td colspan="2"><font size="+1"><b>Topic: <%= sTopic %></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="topics.asp?gn=<%= Request.QueryString("gn") %>&tid=<%= Request.QueryString("tid") %>&tp=<%= Request.QueryString("tp") %>"><%=langBackToTopicsList%></a></td>
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
        <form name="DelMessages" action="deletemessages.asp?<%= Request.QueryString() %>" method="post">
          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/arrow_back.gif" align="absmiddle">
          
          <%
          If iPage > 1 Then
            Response.Write "<a href=""thread.asp?tid=" & Request.QueryString("tid") & "&tp=" & Request.QueryString("tp") & "&mid=" & Request.QueryString("mid") & "&mp=" & iPage-1 & """>" & langPrev & " " & Session("PageSize") & "</a>&nbsp;&nbsp;"
          Else
            Response.Write "<font color=""#999999"">" & langPrev & " " & Session("PageSize") & "</font>&nbsp;&nbsp;"
          End If

          If iNumMoreRecords > 0 Then
            Response.Write "<a href=""thread.asp?tid=" & Request.QueryString("tid") & "&tp=" & Request.QueryString("tp") & "&mid=" & Request.QueryString("mid") & "&mp=" & iPage+1 & """>" & langNext & " " & Session("PageSize") & "</a>"
          Else
            Response.Write "<font color=""#999999"">" & langNext & " " & Session("PageSize") & "</font>"
          End If
          %>
          
          <img src="../images/arrow_forward.gif" align="absmiddle">
          &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/view.gif" align="absmiddle">&nbsp;<a href="messages.asp?<%= Request.QueryString() %>"><%=langMessageView%></a>
          <%
          If bDeleteShown Then
          %>
            &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/small_delete.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.DelMessages.submit();"><%=langDelete%></a>
          <%
          End If
          %>
          </div>

          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
            <tr>
              <th align=left>
              <%If bCanEdit then%>
              <input class="listCheck" type=checkbox name="chkSelectAll" onClick="selectAll('DelMessages', this.checked)">
              <%Else%>
              &nbsp;
              <%End If%>
              </th>
              <th align="left" width="50%"><%=langTopic%></th>
              <th align="left"><%=langCreator%></th>
              <th align="left"><%=langDate%></th>
            </tr>
            <%= sMsgs %>
          </table>
          <div id="pageinfo" style="padding-top:5px; font-size:10px; color:#000000;">Page <%=sPage%> of <%=iTotalPages%></div>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>
