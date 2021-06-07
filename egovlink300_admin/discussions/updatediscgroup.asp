<!-- #include file="../includes/common.asp" //-->
<%
Dim bShown, sLinks

Dim oCmd, oRst, sSubject, intID, sMessage, sMembers, sName, sDesc
If Not HasPermission("CanEditDiscussionGroups") Then Response.Redirect "../default.asp"

If Request.Form("_task") <> "" Then
  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "UpdateDiscGroup"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("DiscGroupID", adInteger, adParamInput, 4, Request.Form("id"))
    .Parameters.Append oCmd.CreateParameter("UserID", adInteger, adParamInput, 4, Session("UserID"))
    .Parameters.Append oCmd.CreateParameter("Name", adVarChar, adParamInput, 50, Request.Form("Name"))
    .Parameters.Append oCmd.CreateParameter("Desc", adVarChar, adParamInput, 5000, Request.Form("Description"))
    .Execute
  End With
  Set oCmd = Nothing
  Response.Redirect "../discussions/"
Else

  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "GetDiscGroup"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("DiscGroupID", adInteger, adParamInput, 4, Request.QueryString("id"))
  End With
  
  Set oRst = Server.CreateObject("ADODB.Recordset")
  With oRst
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open oCmd
  End With
  Set oCmd = Nothing  
  
  If not oRst.EOF then
    intID = clng(oRst("DiscussionGroupID"))
    sName = oRst("Name")
    sDesc = oRst("Description")
    sMembers = oRst("Members")
  End If
  
  If oRst.State=1 then oRst.Close
  Set oRst = Nothing  

End If
%>

<html>
<head>
  <title><%=langBSDiscussions%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script language="Javascript">
  <!--
    function doMembers() {
      x = (screen.width-450)/2;
      y = (screen.height-400)/2;
      win = window.open("members.asp?id=<%=Request("id")%>", "disc_members", "width=450,height=350,status=0,menubar=0,scrollbars=0,toolbar=0,left="+x+",top="+y+",z-lock=yes");
      win.focus();
    }
  //-->
  </script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabDiscussions,1%>

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_discussions.jpg"></td>
      <td><font size="+1"><b><%=langDiscussionGroups%>: <%=langUpdate%></b></font><br><br></td>
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
        <form name="frmUpdateDiscGroup" method=post action="updatediscgroup.asp" method="post">
          <input type="hidden" name="_task" value="update">
          <input type="hidden" name="id" value=<%=intID%>>

          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmUpdateDiscGroup.submit();"><%=langUpdate%></a></div>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
            <tr>
              <th align="left" colspan="2"><%=langNewDiscussionGroup%></th>
            </tr>
            <tr>
              <td width="1" valign="top"><%=langName%>:</td>
              <td><input type="text" name="Name" style="width:320px;" maxlength="30" value="<%=sName%>"> (30 character max)</td>
            </tr>
            <tr>
              <td width="1" valign="top"><%=langDescription%>:&nbsp;</td>
              <td><textarea name="Description" rows="3" style="width:430px;"><%=sDesc%></textarea></td>
            </tr>
            <tr>
              <td width="1"><%=langMembers%>:&nbsp;</td>
              <td>
                <%=sMembers%>&nbsp;&nbsp;&nbsp;&nbsp;
                <img src="../images/newpermission.gif" border="0" align="absmiddle">&nbsp;<a href="javascript:doMembers()"><%=langEditSecurity%>...</a>
              </td>
            </tr>
          </table>
          <div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmUpdateDiscGroup.submit();"><%=langUpdate%></a></div>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>
