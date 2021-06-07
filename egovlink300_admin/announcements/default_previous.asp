<!-- #include file="../includes/common.asp" //-->

<%
Dim oCmd, oRst, sAnnouncements, index, arrColors(2), truncMessage, iTotal, iCount, sDesc, sLinks, bShown
Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
  .ActiveConnection = Application("DSN")
  .CommandText = "ListAnnouncements"
  .CommandType = adCmdStoredProc
  .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
  .Execute
End With

Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open oCmd
End With
Set oCmd = Nothing

arrColors(0)="ffffff"
arrColors(1)="eeeeee"
index=0

iCount = 1
iTotal = oRst.RecordCount

If iTotal = 0 then
   sAnnouncements = sAnnouncements & "<tr><td width='1%'>&nbsp</td><td width='1%'>&nbsp</td><td colspan=2>No new announcements.</td></tr>"
End if


Do while not oRst.EOF
  truncMessage = oRst("Message")
  If Len(truncMessage) > 250 Then
    truncMessage=Left(truncMessage,248) & "..."
  End If

  sAnnouncements = sAnnouncements & "<tr bgcolor='" & arrColors(index) & "'>"
  If HasPermission("CanEditAnnouncements") Then
	sAnnouncements = sAnnouncements & "<td valign='top'><input type ='checkbox' class='listcheck' name='del_" & oRst("AnnouncementID") & "'></td>"
	sAnnouncements = sAnnouncements & "<td valign='top'><img src=""../images/newannounce.gif"" align=""absmiddle""></td>"
	sAnnouncements = sAnnouncements & "<td valign='top'><a href='updateannouncement.asp?id=" & oRst("AnnouncementID") & "'>" & oRst("Subject") & "</a></td>"
  Else
	sAnnouncements = sAnnouncements & "<td valign='top'>&nbsp;</td>"
	sAnnouncements = sAnnouncements & "<td valign='top'><img src=""../images/newannounce.gif"" align=""absmiddle""></td>"
	sAnnouncements = sAnnouncements & "<td valign='top'>" & oRst("Subject") & "</td>"
  End If
 
  sAnnouncements = sAnnouncements & "<td valign='top'>" & truncMessage & "</td>"
  sAnnouncements = sAnnouncements & "<td valign='top'><a href='mailto:" & oRst("Email") & "'>" & oRst("FirstName") & " " & oRst("LastName") & "</a></td>"
  sAnnouncements = sAnnouncements & "</tr>"
  
  index = 1 - index 'flip the index
  iCount = iCount + 1

  oRst.MoveNext
Loop

If oRst.State=1 then oRst.Close
Set oRst = Nothing
%>

<html>
<head>
  <title><%=langBSAnnouncements%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%
  DrawTabs tabHome,1
  %>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_home.jpg"></td>
      <td><font size="+1"><b><%=langAnnouncements%></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="../"><%=langBackToStart%></a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top" nowrap>

         <!-- START: QUICK LINKS MODULE //-->
        
        <%
        sLinks = "<div style=""padding-bottom:8px;""><b>" & langAnnouncementLinks & "</b></div>"

        If HasPermission("CanEditAnnouncement") Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/newannounce.gif"" align=""absmiddle"">&nbsp;<a href=""newannouncement.asp"">" & langNewAnnouncement & "</a></div>"
          bShown = True
        End If

        If bShown Then
          Response.Write sLinks & "<br>"
        End If
        %>

        <% Call DrawQuicklinks("", 1) %>
        <!-- END: QUICK LINKS MODULE //-->

      </td>
        <!-- START: NEW ANNOUNCEMENT -->
      <td colspan="2" valign="top">
        <form name="DelAnnouncement" method=post action="deleteannouncements.asp" method="post">
		<% If HasPermission("CanEditAnnouncements") Then %>
          <div style="font-size:10px; padding-bottom:5px;">
			  <img src="../images/newannounce.gif" align="absmiddle">&nbsp;
			  <a href="newannouncement.asp"><%=langNewAnnouncement%></a>
			  &nbsp;&nbsp;&nbsp;&nbsp;
			  <img src="../images/small_delete.gif" align="absmiddle">&nbsp;
			  <a href="javascript:document.all.DelAnnouncement.submit();" ><%=langDelete%></a>
		  </div>
		<% End If %>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
            <tr>
              <th>
              <% If HasPermission("CanEditAnnouncements") Then %>
              <input class="listCheck" type=checkbox name="chkSelectAll" onClick="selectAll('DelAnnouncement', this.checked)"></th>
              <%Else%>
              &nbsp;
              <%End If%>
              </th>
              <th>&nbsp;</th>
              <th align="left" width="20%"><%=langSubject%></th>
              <th align="left" width="70%"><%=langMessage%></th>
              <th align="left" width="10%"><%=langCreator%></th>
            </tr>
            <tr><%=sAnnouncements%></tr>
          </table>
        </form>
      </td>
        <!-- END: NEW ANNOUNCEMENT -->
    </tr>
  </table>
</body>
</html>