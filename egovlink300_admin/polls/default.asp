<!-- #include file="../includes/common.asp" //-->
<%
Response.Buffer = True

Dim sSql, oRst, sVotes, iCount, iTotal, sBgcolor, iPage, sPage, sTopic, iNumMoreRecords
Dim bCanCreate, bCanEdit

sPage = Request.QueryString("gp")
If sPage & "" <> "" Then
  iPage = clng(sPage)
Else
  iPage = 1
End If

bCanCreate = HasPermission("CanCreatePolls")
bCanEdit = HasPermission("CanEditPolls")

sSql = "EXEC ListVotes " & Session("OrgID") & "," & Session("UserID") & "," & Session("PageSize") & "," & iPage

Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .ActiveConnection = Application("DSN")
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open sSql
  .ActiveConnection = Nothing
End With

sVotes = ""
iCount = 1
iTotal = oRst.RecordCount

If Not oRst.EOF Then
  sBgcolor = "#ffffff"
  iNumMoreRecords = oRst("NumMoreRecords")

  Do While Not oRst.EOF
    If bCanEdit Then
      sVotes = sVotes & "<tr bgcolor=""" & sBgcolor & """><td><input type=""checkbox"" class=""listcheck"" name=""del_" & oRst("VoteID") & """></td>"
    Else
      sVotes = sVotes & "<tr bgcolor=""" & sBgcolor & """><td>&nbsp;</td>"
    End If
    
    sVotes = sVotes & "<td style=""padding:0px;""><img src=""../images/newpoll.gif"" border=""0"">&nbsp;</td>"
    
    If oRst("HasVoted") > 0 Then
      sVotes = sVotes & "<td nowrap><a href=""viewpoll.asp?id=" & oRst("VoteID") & """>" & oRst("Subject") & "</a></td>"
    ElseIf Session("UserID") > 0 Then
      sVotes = sVotes & "<td nowrap><a href=""takepoll.asp?id=" & oRst("VoteID") & """>" & oRst("Subject") & "</a></td>"
    Else
      sVotes = sVotes & "<td nowrap>" & oRst("Subject") & "</td>"
    End If
    
    'If bCanEdit Then
    '  sVotes = sVotes & "&nbsp;<a href=""updatepoll.asp?id=" & oRst("VoteID") & """ style=""font-family:Arial,Tahoma; font-size:10px;""><img src=""../images/edit.gif"" align=""absmiddle"" border=0 alt=""Edit Vote""></a>"
    'End If
    
    sVotes = sVotes & "<td nowrap>" & oRst("VoteType") & "</td>"
    sVotes = sVotes & "</td><td align=""center"">" & oRst("NumResponses") & "</td>"

    If oRst("Status") = 1 Then
      sVotes = sVotes & "<td nowrap>Open</td>"
    Else
      sVotes = sVotes & "<td nowrap>Closed</td>"
    End If

    If oRst("AccessID") > 0 And bCanEdit Then
      sVotes = sVotes & "<td nowrap><img src=""../images/locked.gif"" border=""0"" alt=""This voting poll is both hidden & locked to unauthorized users.""></td></tr>"
    Else
      sVotes = sVotes & "<td nowrap>&nbsp;</td></tr>"
    End If

    If sBgcolor = "#ffffff" Then sBgcolor = "#eeeeee" Else sBgcolor = "#ffffff"
    iCount = iCount + 1
    oRst.MoveNext
  Loop
  oRst.Close
Else
  sVotes = "<tr><td colspan=""6"">No new voting polls.</td></tr>"
End If

Set oRst = Nothing
%>
<html>
  <head>
    <title>
      <%=langBSVoting%>
    </title>
    <link href="../global.css" rel="stylesheet" type="text/css">
      <script src="../scripts/selectAll.js"></script>
  </head>
  <body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
    <% Call DrawTabs(tabVoting,1) %>
    <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
      <tr>
        <td width="151" align="center"><img src="../images/icon_voting.jpg"></td>
        <td colspan="2"><font size="+1"><b><%=langVotingPolls%></b></font><br>
          <br>
        </td>
      </tr>
      <tr>
        <td valign="top">
          <!-- #include file="quicklinks.asp" //-->
          <% Call DrawQuicklinks("",1) %>
        </td>
        <td colspan="2" valign="top">
          <form name="DelPolls" method="post" action="deletepolls.asp" method="post">
            <div style="font-size:10px; padding-bottom:5px;"><img src="../images/arrow_back.gif" align="absmiddle">
              <%
            If IPage > 1 Then
              Response.Write "<a href=""default.asp?gp=" & iPage-1 & """>" & langPrev & "  " & Session("PageSize") & "</a>&nbsp;&nbsp;"
            Else
              Response.Write "<font color=""#999999"">Prev " & Session("PageSize") & "</font>&nbsp;&nbsp;"
            End If

            If iNumMoreRecords > 0 Then
              Response.Write "<a href=""default.asp?gp=" & iPage+1 & """>" & langNext & " " & Session("PageSize") & "</a>"
            Else
              Response.Write "<font color=""#999999"">Next " & Session("PageSize") & "</font>"
            End If
          %>
              <img src="../images/arrow_forward.gif" align="absmiddle">
              <%
          If bCanCreate Then
          %>
              &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/newpoll.gif" align="absmiddle">&nbsp;<a href="newpoll.asp"><%=langNewPoll%></a>
              <%
          End If
          If bCanEdit Then
          %>
              &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/small_delete.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.DelPolls.submit();"><%=langDelete%></a>
              <%
          End If
          %>
            </div>
            <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
              <tr>
                <th align="left">
                  <%If bCanEdit then%>
                  <input class="listCheck" type="checkbox" name="chkSelectAll" onClick="selectAll('DelPolls', this.checked)">
                  <%Else%>
                  &nbsp;
                  <%End If%>
                </th>
                <th>
                  &nbsp;</th>
                <th align="left" width="70%">
                  <%=langTopic%>
                </th>
                <th align="left" width="100" nowrap>
                  <%=langType%>
                </th>
                <th align="center" width="100" nowrap>
                  <%=langResponses%>
                </th>
                <th align="left" width="100" nowrap>
                  <%=langStatus%>
                </th>
                <th width="1">
                  &nbsp;</th>
              </tr>
              <%= sVotes %>
            </table>
          </form>
        </td>
      </tr>
    </table>
  </body>
</html>
