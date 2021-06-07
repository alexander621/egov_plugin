<!-- #include file="../includes/common.asp" //-->
<%
Dim oCmd, sSql, oRst
Dim bShown, sLinks, sMembers, sMemberNames, sName, sDesc

sMembers = Request.Form("Members") & ""
If sMembers = "" Then
  sMembers = NULL
Else
  sMembers = Replace(sMembers, ",0", "")
  sMembers = Replace(sMembers, "0,", "")
  sMembers = Replace(sMembers, "0", "")
End If

sName = Request.Form("Name")
sDesc = Request.Form("Description")

If Request.Form("_task") = "newdg" Then
  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "NewDiscussionGroup"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
    .Parameters.Append oCmd.CreateParameter("UserID", adInteger, adParamInput, 4, Session("UserID"))
    .Parameters.Append oCmd.CreateParameter("Name", adVarChar, adParamInput, 50, sName)
    .Parameters.Append oCmd.CreateParameter("Description", adVarChar, adParamInput, 200, sDesc)
    .Parameters.Append oCmd.CreateParameter("GroupIDs", adVarChar, adParamInput, 1000, sMembers)
    .Execute
  End With
  Response.Redirect "../discussions/"
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
  <title><%=langBSDiscussions%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script language="Javascript">
  <!--
    function doMembers() {
      x = (screen.width-450)/2;
      y = (screen.height-400)/2;
      mem = document.all.Members.value;
      win = window.open("members_buffered.asp?mem="+mem, "disc_members", "width=450,height=350,status=0,menubar=0,scrollbars=0,toolbar=0,left="+x+",top="+y+",z-lock=yes");
      win.focus();
    }
	function textCounter(field, countfield, maxlimit) {
	if (field.value.length > maxlimit) // if too long...trim it!
	field.value = field.value.substring(0, maxlimit);
	// otherwise, update 'characters left' counter
	else 
	countfield.value = maxlimit - field.value.length;
	}

  //-->
  </script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabDiscussions,1%>

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_discussions.jpg"></td>
      <td><font size="+1"><b><%=langDiscussionGroups%>: <%=langNew%></b></font><br><br></td>
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
        <form name="frmNewDiscGroup" method=post action="newdiscgroup.asp" method="post">
          <input type="hidden" name="_task" value="newdg">
          <input type="hidden" name="Members" value="<%=sMembers%>">

          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="../discussions"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmNewDiscGroup.submit();"><%=langCreate%></a></div>
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
              <td><textarea name="Description" rows="3" style="width:430px;"
		onKeyDown="textCounter(this.form.Description,this.form.remLen,200);" onKeyUp="textCounter(this.form.Description,this.form.remLen,200);"
	      ><%=sDesc%></textarea>
	      <input type=hidden name=remLen value=200>
	      </td>
            </tr>
            <tr>
              <td width="1"><%=langMembers%>:&nbsp;</td>
              <td>
                <%=sMemberNames%>&nbsp;&nbsp;&nbsp;&nbsp;
                <img src="../images/newpermission.gif" border="0" align="absmiddle">&nbsp;<a href="javascript:doMembers()"><%=langEditSecurity%>...</a>
              </td>
            </tr>
          </table>
          <div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="../discussions"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmNewDiscGroup.submit();"><%=langCreate%></a></div>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>
