<!-- #include file="../includes/common.asp" //-->
<%
Dim sSql, oRst, sTopics, iTopicID, iTopicPage, iCount, iTotal, sBgcolor, iNumMessages, iPage, sPage, sTopic, iNumMoreRecords
Dim bShown, bCanCreate, bCanEdit, sLinks, sSearch, sEE

sPage = Request.QueryString("p")
If sPage & "" <> "" Then
  iPage = clng(sPage)
End If

bCanEdit = HasPermission("CanEditDiscussionTopics")
sSearch  = Request.QueryString("s")

sSearch  = Replace(sSearch, "'", "''")
sSearch  = Replace(sSearch, "%", "[%]")

sSql = "EXEC SearchDiscussions '" & sSearch & "'," & Session("PageSize") & "," & iPage
            
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

    sTopics = sTopics & "<td width=1 style=""padding:0px;""><img src=""../images/newdisc.gif"" border=""0""></td>"
    sTopics = sTopics & "<td nowrap><a href=""message.asp?msgid=" & oRst("DiscussionID") & """>" & oRst("Subject") & "</a>&nbsp;&nbsp;</td>"
    sTopics = sTopics & "<td nowrap>" & oRst("Creator") & "&nbsp;&nbsp;</td>"
    sTopics = sTopics & "<td nowrap>" & oRst("DateOfPost") & "</td></tr>"

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
      <td><font size="+1"><b><%=langDiscussions%>: <%= langSearch %> for <%=Request.QueryString("s")%></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="../discussions"><%=langBackToGroupsList%></a></td>
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
          If bCanEdit Then
          %>
            &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/small_delete.gif" align="absmiddle">&nbsp;<a href="javascript:if (confirm('<%=langConfirmDeleteTopic%>')){document.all.DelTopics.submit();}"><%=langDelete%></a>
          <%
          End If
          %>
          </div>

          <% If sEE = "" Then %>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
            <tr>
              <th width="1%">
                <%If bCanEdit then%>
                  <input class="listCheck" type=checkbox name="chkSelectAll" onClick="selectAll('DelTopics', this.checked)">
                <%Else%>
                  &nbsp;
                <%End If%>
              </th>
              <th width="1">&nbsp;</th>
              <th align="left" width="70%"><%=langTopic%></th>
              <th align="left"><%=langCreator%></th>
              <th align="left"><%=langDate%></th>
            </tr>
            <%= sTopics %>
          </table>
          <% Else %>
            <h2>Your words ring loud and true my brother.</h2>
          <% End If %>
        </form>
      </td>
    </tr>
  </table>
  <%=sEE%>
</body>
</html>
