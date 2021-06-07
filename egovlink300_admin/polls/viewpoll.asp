<!-- #include file="../includes/common.asp" //-->
<%
Response.Buffer = True
Dim sSql, oRst, iVoteCount, iCount, iTotal, iUserID, sFullname, sEmail, sSubject, sDescription, sQuestion
Dim iStatus, sCreateDateTime, sVoteType, sVotes, sComments, bCanEdit, iPage, iNumMoreRecords

bCanEdit = HasPermission("CanEditPolls")

sSql = "EXEC GetVote " & Request("id") & "," & Session("OrgID") & "," & Session("UserID")

Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .ActiveConnection = Application("DSN")
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open sSql
  .ActiveConnection = Nothing
End With

If Not oRst.EOF Then
  If oRst("HasVoted") = 0 And oRst("Status") = 1 Then 'if user has NOT voted and poll is still force them to vote
    Response.Redirect "takepoll.asp?" & Request.QueryString()
  End If

  iUserID = oRst("UserID")
  iStatus = oRst("Status")
  sFullname = oRst("Fullname")
  sEmail = oRst("Email")
  sSubject = oRst("Subject")
  sDescription = oRst("Description")
  sQuestion = oRst("Question")
  sCreateDateTime = oRst("CreateDateTime")
  sVoteType = LCase(oRst("VoteTypeDescription"))
  oRst.Close

  If sVoteType <> "private" Or iUserID = Session("UserID") Then
    sSql = "EXEC GetVoteResults " & Request("id")
    With oRst
      .ActiveConnection = Application("DSN")
      .Open sSql
      .ActiveConnection = Nothing
    End With

    If Not oRst.EOF Then
      iVoteCount = 0
      iCount = 0
      iTotal = oRst.RecordCount

      ReDim sVotes(iTotal,2)

      Do While Not oRst.EOF
        iVoteCount = iVoteCount + oRst("NumVotes")
        sVotes(iCount,0) = oRst("AnswerDescription")
        sVotes(iCount,1) = oRst("NumVotes")
        iCount = iCount + 1
        oRst.MoveNext
      Loop
      oRst.Close
    End If

    sSql = "EXEC GetVoteComments " & Request("id")
    With oRst
      .ActiveConnection = Application("DSN")
      .Open sSql
      .ActiveConnection = Nothing
    End With

    If Not oRst.EOF Then
      Do While Not oRst.EOF
        sComments = sComments & "<br><table width=""100%"" border=0 cellpadding=5 cellspacing=0 class=""messagehead"">"
        sComments = sComments & "<tr><th align=""left"" style=""font-weight:normal;"">Voted: <b>" & oRst("AnswerDescription") & "</b></th></tr>"
        sComments = sComments & "<tr><tr><td><table border=0 cellpadding=0 cellspacing=0 width=""100%"">"

        If sVoteType <> "anonymous" Then
          sComments = sComments & "<tr><td style=""font-size:10px;"">From: <a href=""mailto:" & oRst("Email") & """>" & oRst("FullName") & "</a></td></tr>"
        End If

        sComments = sComments & "<tr><td style=""padding:8px;"">" & oRst("Comments") & "</td></tr></table></td></tr></table>"
        oRst.MoveNext
      Loop
      oRst.Close
    End If
  End If
End If

Set oRst = Nothing
%>

<html>
<head>
  <title><%=langBSVoting%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  
  <% Call DrawTabs(tabVoting,1) %>

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_voting.jpg"></td>
      <td><font size="+1"><b><%=langVotingPoll%>: <%=sSubject%></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="../polls/">Back To Voting Poll List</a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
        <!-- #include file="quicklinks.asp" //-->
        <% Call DrawQuicklinks("",1) %>
      </td>
      <td colspan="2" valign="top">        
        <div style="font-size:10px; padding-bottom:5px;"><img src="../images/arrow_back.gif" align="absmiddle">
          <%
            If iPage > 1 Then
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
          If (sVoteType = "private" And iUserID = Session("UserID")) Or (sVoteType <> "private" And bCanEdit) Then
            If iStatus = 1 Then
            %>
              &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/close.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmTogglePoll.submit();"><%=langClose%>&nbsp;<%=langVotingPoll%></a>
            <%
            Else
            %>
              &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/close.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmTogglePoll.submit();"><%=langOpen%>&nbsp;<%=langVotingPoll%></a>
            <%
            End If
            %>
          <%
          End If

          If sVoteType = "public" Or iUserID = Session("UserID") Then
          %>
            &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/view.gif" align="absmiddle">&nbsp;<a href="polldetails.asp?id=<%=Request("id")%>&p=1"><%=langViewDetails%></a>
          <%
          End If
          %>
        </div>
        <table width="100%" border="0" cellpadding="5" cellspacing="0" class="messagehead">
          <tr>
            <th align="left" colspan="2"><%=langQuestion%></th>
          </tr>
          <tr>
            <td colspan="2">
              <table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr>
                  <td style="font-size:10px;"><%=langFrom%>: <a href="mailto:<%=sEmail%>"><%=sFullname%></a></td>
                  <td style="font-size:10px;" align="right"><%=langCreate%>: <%=sCreateDateTime%></td>
                </tr>
                <tr>
                  <td colspan="2" style="padding:8px;">
                    <%=sDescription%>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" style="padding:8px;">
                    <b><%=sQuestion%></b><br>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
            <td width="100%">
              <table border="0" cellpadding="6" cellspacing="0">
                <%
                Dim dPercent
                dPercent = 0
                If IsArray(sVotes) Then
                  For iCount = 0 To UBound(sVotes,1)-1

                    If iVoteCount <> 0 Then
                      dPercent = Round((sVotes(iCount,1) / iVoteCount * 100))
                    End If
                  %>
                    <tr>
                      <td align="right" nowrap><%=sVotes(iCount,0)%></td>
                      <% If dPercent <> 100 And dPercent <> 0 Then %>
                        <td nowrap><span style="font-size:8px; width:200; border:1px solid #006600;"><span style="border-right:1px solid #006600; background-color:#66cc99; width:<%=dPercent%>%"></span></span><font style="font-family:Tahoma,Arial;font-size:9px;">&nbsp;<%=dPercent%>%</font></td><td nowrap><font style="font-family:Tahoma,Arial;font-size:9px;">( <%=sVotes(iCount,1)%> / <%=iVoteCount%> )</font></td>
                      <% Else %>
                        <td nowrap><span style="font-size:8px; width:200; border:1px solid #006600;"><span style="background-color:#66cc99; width:<%=dPercent%>%"></span></span><font style="font-family:Tahoma,Arial;font-size:9px;">&nbsp;<%=dPercent%>%</font></td><td nowrap><font style="font-family:Tahoma,Arial;font-size:9px;">( <%=sVotes(iCount,1)%> / <%=iVoteCount%> )</font></td>
                      <% End If %>
                    </tr>
                  <%
                  Next
                Else
                  Response.Write "<tr><td><font color=""#ff0000"">This vote is private and only the creator can view the results.</font></td></tr>"
                End If
                %>
              </table>
              <br>
            </td>
          </tr>
        </table>
        <%=sComments%>
        <form name="frmTogglePoll" action="togglepollstatus.asp?id=<%=Request("id")%>" method="post">
        </form>
      </td>
    </tr>
  </table>
</body>
</html>