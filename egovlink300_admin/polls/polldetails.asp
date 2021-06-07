<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
Response.Buffer = True
Dim sSql, oRst, sResponses, iCount, iTotal, sBgcolor, iPage, sPage, sTopic, iNumMoreRecords, bCanEdit, sSubject, iVoteID

bCanEdit = HasPermission("CanEditPolls")

sPage = Request.QueryString("p")
If sPage & "" <> "" Then
  iPage = clng(sPage)
Else
  iPage = 1
End If

iVoteID = Request("id")

sSql = "EXEC ListVoteDetails " & iVoteID & ", " & Session("PageSize") & ", " & iPage

Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .ActiveConnection = Application("DSN")
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open sSql
  .ActiveConnection = Nothing
End With

sResponses = ""
iCount = 1
iTotal = oRst.RecordCount

If Not oRst.EOF Then
  iNumMoreRecords = oRst("NumMoreRecords")
  sBgcolor = "#ffffff"
  sSubject = oRst("Subject")

  Do While Not oRst.EOF
    If bCanEdit Then
      sResponses = sResponses & "<tr bgcolor=""" & sBgcolor & """><td><input type=""checkbox"" class=""listcheck"" name=""del_" & oRst("VoteResponseID") & """></td>"
    Else
      sResponses = sResponses & "<tr bgcolor=""" & sBgcolor & """><td>&nbsp;</td>"
    End If
    
    sResponses = sResponses & "<td style=""padding:0px;""><img src=""../images/newuser.gif"" border=""0"">&nbsp;</td>"
    sResponses = sResponses & "<td nowrap>" & oRst("Fullname") & "</td>"
    sResponses = sResponses & "<td>" & oRst("AnswerDescription") & "</td>"
    sResponses = sResponses & "<td>" & MyFormatDateTime(oRst("ResponseDateTime"), " ") & "</td></tr>"

    If sBgcolor = "#ffffff" Then sBgcolor = "#eeeeee" Else sBgcolor = "#ffffff"
    iCount = iCount + 1
    oRst.MoveNext
  Loop
  oRst.Close
Else
  sResponses = "<tr><td colspan=""6"">No new voting polls.</td></tr>"
End If

Set oRst = Nothing
%>

<html>
<head>
  <title><%=langBSVoting%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  
  <% Call DrawTabs(tabVoting,1) %>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_voting.jpg"></td>
      <td colspan="2"><font size="+1"><b><%=langVotingPollDetails%>: <%=sSubject%></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="viewpoll.asp?<%=Request.QueryString()%>">Back To <%=sSubject%></a></td>
    </tr>
    <tr>
      <td valign="top">
        <!-- #include file="quicklinks.asp" //-->
        <% Call DrawQuicklinks("",1) %>
      </td>
      <td colspan="2" valign="top">
        <form name="DelResponses" method=post action="deleteresponses.asp?<%=Request.QueryString()%>" method="post">
          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/arrow_back.gif" align="absmiddle">
          <%
            If IPage > 1 Then
              Response.Write "<a href=""polldetails.asp?id=" & iVoteID & "&p=" & iPage-1 & """>" & langPrev & "  " & Session("PageSize") & "</a>&nbsp;&nbsp;"
            Else
              Response.Write "<font color=""#999999"">Prev " & Session("PageSize") & "</font>&nbsp;&nbsp;"
            End If

            If iNumMoreRecords > 0 Then
              Response.Write "<a href=""polldetails.asp?id=" & iVoteID & "&p=" & iPage+1 & """>" & langNext & " " & Session("PageSize") & "</a>"
            Else
              Response.Write "<font color=""#999999"">Next " & Session("PageSize") & "</font>"
            End If
          %>
          <img src="../images/arrow_forward.gif" align="absmiddle">
          <%
          If bCanEdit Then
          %>
            &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/small_delete.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.DelResponses.submit();"><%=langDelete%>&nbsp;<%=langVote%></a>
          <%
          End If
          %>
          </div>

          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
            <tr>
              <th>
                <%If bCanEdit then%>
                  <input class="listCheck" type=checkbox name="chkSelectAll" onClick="selectAll('DelResponses', this.checked)">
                <%Else%>
                  &nbsp;
                <%End If%>
              </th>
              <th>&nbsp;</th>
              <th align="left" nowrap width="150"><%=langUser%></th>
              <th align="left" nowrap width="50%"><%=langResponse%></th>
              <th align="left" nowrap width="200"><%=langDateTime%></th>
            </tr>
            <%= sResponses %>
          </table>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>
