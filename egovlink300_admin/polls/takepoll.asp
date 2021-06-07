<!-- #include file="../includes/common.asp" //-->
<%
Response.Buffer = True
Dim sSql, oCmd, oRst, sFullname, sEmail, sSubject, sDescription, sCreateDateTime, sQuestion

If Request.Form("_task") = "castvote" Then
  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "NewVoteResponse"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("UserID", adInteger, adParamInput, 4, Session("UserID"))
    .Parameters.Append oCmd.CreateParameter("VoteID", adInteger, adParamInput, 4, Request.Form("VoteID"))
    .Parameters.Append oCmd.CreateParameter("AnswerID", adInteger, adParamInput, 4, Request.Form("AnswerID"))
    .Parameters.Append oCmd.CreateParameter("Comments", adVarChar, adParamInput, 1000, SQLText(Request.Form("Comments")))
    .Execute
  End With
  Response.Redirect "../polls/viewpoll.asp?id=" & Request.Form("VoteID")
End If

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
  If oRst("HasVoted") > 0 Or oRst("Status") = 0 Then   'if the user has voted, or the vote is closed then redirect to view results
    Response.Redirect "viewpoll.asp?" & Request.QueryString()
  End If
  
  sFullname = oRst("Fullname")
  sEmail = oRst("Email")
  sSubject = oRst("Subject")
  sDescription = oRst("Description")
  sQuestion = oRst("Question")
  sCreateDateTime = oRst("CreateDateTime")

  oRst.Close
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
      <td colspan="2"><font size="+1"><b><%=langVotingPoll%>: <%=sSubject%></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="../polls/">Back To Voting Poll List</a></td>
    </tr>
    <tr>
      <td valign="top">
        <!-- #include file="quicklinks.asp" //-->
        <% Call DrawQuicklinks("",1) %>
      </td>
      <td colspan="2" valign="top">
        <form name="frmTakePoll" method=post action="takepoll.asp" method="post">
          <input type="hidden" name="_task" value="castvote">
          <input type="hidden" name="VoteID" value="<%=Request("id")%>"> 

          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmTakePoll.submit();"><%=langSubmitVote%></a></div>
          <table width="100%" border="0" cellpadding="5" cellspacing="0" class="messagehead">
            <tr>
              <th align="left"><%=langQuestion%></th>
            </tr>
            <tr>
              <td colspan="2">
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                  <tr>
                    <td style="font-size:10px;"><%=langFrom%>: <a href="mailto:<%=sEmail%>"><%=sFullname%></a></td>
                    <td style="font-size:10px;" align="right">Created: <%=sCreateDateTime%></td>
                  </tr>
                  <tr>
                    <td colspan="2" style="padding:8px;">
                      <%=sDescription%>
                    </td>
                  </tr>
                  <tr>
                    <td style="padding:8px;" width=75%>
                      <b><%=sQuestion%></b><br>
                      <br>
                      <input type="radio" name="AnswerID" value="1">Yes<br>
                      <input type="radio" name="AnswerID" value="2">No<br>
                      <input type="radio" name="AnswerID" value="3">Needs Discussion<br>
                      <input type="radio" name="AnswerID" value="4">Abstain<br>
                    </td>
                    <td style="padding:8px;" width="100%">
                      <br><br>Comments:<br><textarea style="height:70px; width:100%; font-family:Arial; font-size:11px;" name="Comments"></textarea>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
          <div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmTakePoll.submit();"><%=langSubmitVote%></a></div>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>
