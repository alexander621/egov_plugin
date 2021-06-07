<!-- #include file="../includes/common.asp" //-->
<%
Response.Buffer = True
Dim sSql, oRst, sGroups, iCount, iTotal, sBgcolor, iPage, sPage, sTopic, iNumMoreRecords
Dim bShown, bCanCreate, bCanEdit, sLink

sPage = Request.QueryString("gp")
If sPage & "" <> "" Then
  iPage = clng(sPage)
Else
  iPage = 1
End If

bCanCreate = HasPermission("CanCreateDiscussionGroups")
bCanEdit = HasPermission("CanEditDiscussionGroups")

sSql = "EXEC ListDiscussionGroups " & Session("OrgID") & "," & Session("UserID") & "," & Session("PageSize") & "," & iPage

Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .ActiveConnection = Application("DSN")
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open sSql
  .ActiveConnection = Nothing
End With

sGroups = ""
iCount = 1
iTotal = oRst.RecordCount

If Not oRst.EOF Then
  sBgcolor = "#ffffff"
  iNumMoreRecords = oRst("NumMoreRecords")

  Do While Not oRst.EOF
    If bCanEdit Then
      sGroups = sGroups & "<tr bgcolor=""" & sBgcolor & """><td><input type=""checkbox"" class=""listcheck"" name=""del_" & oRst("DiscussionGroupID") & """></td>"
    Else
      sGroups = sGroups & "<tr bgcolor=""" & sBgcolor & """><td>&nbsp;</td>"
    End If
    
    sGroups = sGroups & "<td style=""padding:0px;""><img src=""../images/newdiscgroup.gif"" border=""0"">&nbsp;</td>"
    sGroups = sGroups & "<td nowrap><a href=""topics.asp?tid=" & oRst("DiscussionGroupID") & "&tp=1&gn=" & Server.URLEncode(oRst("Name")) & """>" & oRst("Name") & "</a>"
    
    If bCanEdit Then
      sGroups = sGroups & "&nbsp;<a href=""updatediscgroup.asp?id=" & oRst("DiscussionGroupID") & """ style=""font-family:Arial,Tahoma; font-size:10px;""><img src=""../images/edit.gif"" align=""absmiddle"" border=0 alt=""Edit Discussion Board""></a>"
    End If
    
    sGroups = sGroups & "</td><td align=""center"">" & oRst("NumTopics") & "</td>"
    sGroups = sGroups & "<td>" & oRst("Description") & "&nbsp;</td>"

    If oRst("AccessID") > 0 And bCanEdit Then
      sGroups = sGroups & "<td nowrap><img src=""../images/locked.gif"" border=""0"" alt=""This discussion board is both hidden & locked to unauthorized users.""></td></tr>"
    Else
      sGroups = sGroups & "<td nowrap>&nbsp;</td></tr>"
    End If

    If sBgcolor = "#ffffff" Then sBgcolor = "#eeeeee" Else sBgcolor = "#ffffff"
    iCount = iCount + 1
    oRst.MoveNext
  Loop
  oRst.Close
Else
  sGroups = "<tr><td colspan=""6"">No discussion boards available.</td></tr>"
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
      <td colspan="2"><font size="+1"><b><%=langDiscussionGroups%></b></font><br><br></td>
    </tr>
    <tr>
      <td valign="top">
        
        <!-- START: QUICK LINKS MODULE //-->
        <%
        Dim sLinks

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
        <form name="DelGroups" action="deletegroups.asp?<%= Request.QueryString() %>" method="post">
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
            &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/newdiscgroup.gif" align="absmiddle">&nbsp;<a href="newdiscgroup.asp"><%=langNewDiscussionGroup%></a>
          <%
          End If
          If bCanEdit Then
          %>
            &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/small_delete.gif" align="absmiddle">&nbsp;<a href="javascript:if (confirm('<%=langConfirmDeleteDiscussion%>')){document.all.DelGroups.submit();}">Delete</a>
          <%
          End If
          %>
          </div>
          
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
            <tr>
              <th width="1%">
                <%If bCanEdit then%>
                  <input class="listCheck" type=checkbox name="chkSelectAll" onClick="selectAll('DelGroups', this.checked)">
                <%Else%>
                  &nbsp;
                <%End If%>
              </th>
              <th width="1%">&nbsp;</th>
              <th align="left" width="1%"><%=langDiscussionBoard%></th>
              <th align="center"><%=langTopics%></th>
              <th align="left" width="80%"><%=langDescription%></th>
              <th align="left">&nbsp;</th>
            </tr>	
            <%= sGroups %>
          </table>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>
