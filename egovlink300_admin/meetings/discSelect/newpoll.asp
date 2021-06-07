<!-- #include file="../includes/common.asp" //-->
<%
Dim oCmd, sSql, oRst
Dim sMembers, sMemberNames, iVoteType, sTopic, sDesc, sQuestion

sMembers = Request.Form("Members") & ""
If sMembers = "" Then
  sMembers = NULL
Else
  sMembers = Replace(sMembers, ",0", "")
  sMembers = Replace(sMembers, "0,", "")
  sMembers = Replace(sMembers, "0", "")
End If

iVoteType = Request.Form("VoteType")
sTopic = Request.Form("Topic")
sDesc = Request.Form("Description")
sQuestion = Request.Form("Question")

If Request.Form("_task") = "newpoll" Then
  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "NewVote"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
    .Parameters.Append oCmd.CreateParameter("UserID", adInteger, adParamInput, 4, Session("UserID"))
    .Parameters.Append oCmd.CreateParameter("VoteType", adInteger, adParamInput, 4, iVoteType)
    .Parameters.Append oCmd.CreateParameter("Subject", adVarChar, adParamInput, 30, sTopic)
    .Parameters.Append oCmd.CreateParameter("Question", adVarChar, adParamInput, 250, sQuestion)
    .Parameters.Append oCmd.CreateParameter("Description", adVarChar, adParamInput, 1000, sDesc)
    .Parameters.Append oCmd.CreateParameter("GroupIDs", adVarChar, adParamInput, 1000, sMembers)
    .Execute
  End With
  Response.Redirect "../polls/"
End If

sSql = "EXEC GetGroupNames '" & sMembers & "'"
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
  sMemberNames = oRst("GroupNames") & ""
  oRst.Close
End If
Set oRst = Nothing

If sMemberNames = "" Then
  sMemberNames = "Everyone"
End If
%>

<html>
<head>
  <title><%=langBSVoting%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script language="Javascript">
  <!--
    function doMembers() {
      x = (screen.width-450)/2;
      y = (screen.height-400)/2;
      mem = document.all.Members.value;
      win = window.open("members_buffered.asp?mem="+mem, "members", "width=450,height=350,status=0,menubar=0,scrollbars=0,toolbar=0,left="+x+",top="+y+",z-lock=yes");
      win.focus();
    }
  //-->
  </script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  
  <% Call DrawTabs(tabVoting,1) %>

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_discussions.jpg"></td>
      <td><font size="+1"><b><%=langVotingPolls%>: <%=langNew%></b></font><br><br></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
        <!-- #include file="quicklinks.asp" //-->
        <% Call DrawQuicklinks("",1) %>
      </td>
      <td colspan="2" valign="top">
        <form name="frmNewPoll" method=post action="newpoll.asp" method="post">
          <input type="hidden" name="_task" value="newpoll">
          <input type="hidden" name="Members" value="<%=sMembers%>">

          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="../polls"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmNewPoll.submit();"><%=langCreate%></a></div>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
            <tr>
              <th align="left" colspan="2"><%=langNew%>&nbsp;<%=langVotingPoll%></th>
            </tr>
            <tr>
              <td width="1" valign="top" nowrap><%=langType%>:</td>
              <td>
                <select name="VoteType">
                  <option value="1"<%If iVoteType=1 Then Response.Write " selected"%>>Anonymous</option>
                  <option value="2"<%If iVoteType=2 Then Response.Write " selected"%>>Private</option>
                  <option value="3"<%If iVoteType=3 Then Response.Write " selected"%>>Public</option>
                </select>
              </td>
            </tr>
            <tr>
              <td width="1" valign="top"><%=langTopic%>:</td>
              <td><input type="text" name="Topic" style="width:320px;" maxlength="30" value="<%=sTopic%>"> (30 character max)</td>
            </tr>
            <tr>
              <td width="1" valign="top"><%=langDescription%>:&nbsp;</td>
              <td><textarea name="Description" rows="3" style="width:430px;"><%=sDesc%></textarea></td>
            </tr>
            <tr>
              <td width="1" valign="top"><%=langQuestion%>:</td>
              <td><input type="text" name="Question" style="width:320px;" maxlength="250" value="<%=sQuestion%>"> (250 character max)</td>
            </tr>
            <tr>
              <td width="1"><%=langMembers%>:&nbsp;</td>
              <td>
                <%=sMemberNames%>&nbsp;&nbsp;&nbsp;&nbsp;
                <img src="../images/newpermission.gif" border="0" align="absmiddle">&nbsp;<a href="javascript:doMembers()"><%=langEditSecurity%>...</a>
              </td>
            </tr>
          </table>
          <div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="../polls"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmNewPoll.submit();"><%=langCreate%></a></div>
          <br>
          <br>
          <br>
          <font color="#ff0000"><b>NOTE</b>: Once a voting poll has been created, it can not be updated.  This is by design and ensures a fair and unchanging voting process.</font>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>