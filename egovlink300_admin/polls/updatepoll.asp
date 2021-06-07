<!-- #include file="../includes/common.asp" //-->
<%
Dim sSql, oCmd, oRst, sSubject, sDescription, sVoteType

If Request.Form("_task") = "updatepoll" Then
  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "UpdatePoll"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("VoteID", adInteger, adParamInput, 4, Request.Form("VoteID"))
    .Parameters.Append oCmd.CreateParameter("VoteTypeID", adInteger, adParamInput, 4, Request.Form("VoteTypeID"))
    .Parameters.Append oCmd.CreateParameter("Description", adVarChar, adParamInput, 1000, Request.Form("Description"))
    .Execute
  End With
  Response.Redirect "../polls/"
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
  sSubject = oRst("Subject")
  sDescription = oRst("Description")
  sQuestion = oRst("Question")
  sCreateDateTime = oRst("CreateDateTime")
  sVoteType = LCase(oRst("VoteTypeDescription"))
  oRst.Close
End If
%>

<html>
<head>
  <title><%=langBSVoting%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="document.all.Topic.focus();">
  
  <% Call DrawTabs(tabVoting,1) %>

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_discussions.jpg"></td>
      <td><font size="+1"><b><%=langVotingPolls%>: <%=langUpdate%></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langBackToVotingPoll%></a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
        <!-- #include file="quicklinks.asp" //-->
        <% Call DrawQuicklinks("",1) %>
      </td>
      <td colspan="2" valign="top">
        <form name="frmUpdatePoll" method=post action="updatepoll.asp" method="post">
          <input type="hidden" name="_task" value="updatepoll">
          <input type="hidden" name="VoteID" value="<%=Request("id")%>">

          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmUpdatePoll.submit();"><%=langUpdate%></a></div>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
            <tr>
              <th align="left" colspan="2"><%=langUpdate%>&nbsp;<%=langVotingPoll%></th>
            </tr>
            <tr>
              <td width="1" valign="top" nowrap>Vote&nbsp;Type:</td>
              <td>
                <select name="VoteTypeID">
                  <%
                  If sVoteType = "anonymous" Then
                    Response.Write "<option value=""1"" selected>Anonymous</option>"
                  Else
                    Response.Write "<option value=""1"">Anonymous</option>"
                  End If

                  If sVoteType = "private" Then
                    Response.Write "<option value=""2"" selected>Private</option>"
                  Else
                    Response.Write "<option value=""2"">Private</option>"
                  End If

                  If sVoteType = "public" Then
                    Response.Write "<option value=""3"" selected>Public</option>"
                  Else
                    Response.Write "<option value=""3"">Public</option>"
                  End If
                  %>
                </select>
              </td>
            </tr>
            <tr>
              <td width="1" valign="top" height="30"><%=langTopic%>:</td>
              <td><%=sSubject%></td>
            </tr>
            <tr>
              <td width="1" valign="top" height="30"><%=langQuestion%>:</td>
              <td><%=sQuestion%></td>
            </tr>
            <tr>
              <td width="1" valign="top"><%=langDescription%>:&nbsp;</td>
              <td><textarea name="Description" rows="3" style="width:430px;"><%=sDescription%></textarea></td>
            </tr>
          </table>
          <div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmUpdatePoll.submit();"><%=langUpdate%></a></div>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>