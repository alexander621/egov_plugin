<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'if not logged in dont allow access, no matter what
If Session("UserID") = 0 Then Response.Redirect RootPath

Dim sSql, oRst, sMsgs, iCount, iTotal, iTotalNew, sBgcolor, iPage, iNumMoreRecords

iPage=1
sPage=Request.QueryString("page")
If sPage & "" <> "" Then
  iPage = clng(sPage)
End If

Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
  .ActiveConnection = Application("DSN")
  .CommandText = "ListInboxMail"
  .CommandType = adCmdStoredProc
  .Parameters.Append oCmd.CreateParameter("UserID", adInteger, adParamInput, 4, Session("UserID"))
  .Parameters.Append oCmd.CreateParameter("PageSize", adInteger, adParamInput, 4, Session("PageSize"))
  .Parameters.Append oCmd.CreateParameter("Page", adInteger, adParamInput, 4, iPage)
  .Parameters.Append oCmd.CreateParameter("NumTotal", adInteger, adParamOutput, 4, iTotal)
  .Parameters.Append oCmd.CreateParameter("NumNew", adInteger, adParamOutput, 4, iTotalNew)
End With
            
Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open oCmd
End With
iTotal=oCmd.Parameters("NumTotal").Value
iTotalNew = oCmd.Parameters("NumNew").Value
Set oCmd = Nothing
sMsgs = ""

If Not oRst.EOF Then
  iCount = 1
  sBgcolor = "#ffffff"
  iNumMoreRecords = oRst("NumMoreRecords")

  Do While Not oRst.EOF
    sMsgs = sMsgs & "<tr bgcolor=""" & sBgcolor & """><td width=""1%"">"
    sMsgs = sMsgs & "<input type=""checkbox"" class=""nomargin"" name=""del_" & oRst("PmailID") & """>"
    
    If oRst("HasBeenRead") = 0 Then
      sMsgs = sMsgs & "&nbsp;&nbsp;<img src=""../images/newmail.gif"" border=""0"" align=""absmiddle"">"
    End If
    
    sMsgs = sMsgs & "</td><td nowrap><a href=""message.asp?pid=" & oRst("PmailID") & "&backid=1"">" & oRst("From") & "</td>"
    sMsgs = sMsgs & "<td>" & oRst("Subject") & "</td>"
    sMsgs = sMsgs & "<td nowrap>" & MyFormatDateTime(oRst("SentDateTime"), " ") & "</td></tr>"

    If sBgcolor = "#ffffff" Then sBgcolor = "#eeeeee" Else sBgcolor = "#ffffff"
    iCount = iCount + 1
    oRst.MoveNext
  Loop
  oRst.Close
Else
  sMsgs = "<tr><td colspan=6 style=""border:0px;"">" & langNoNewMessages & "</td></tr>"
End If

Set oRst = Nothing
%>

<html>
<head>
  <title><%=langBSMessages%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
  <style type="text/css">
  <!--
    .nomargin {margin:-4px;}
  //-->
  </style>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabMessages,1%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_messages.jpg"></td>
      <td><font size="+1"><b><%=langMessages%>: <%=langInbox%></b></font><br><%=langYouHave%>&nbsp;<%= iTotalNew %>&nbsp;<%=langNewMessages%>&nbsp;&nbsp;(<%= iTotal %>&nbsp;<%=langTotal%>)</td>
    </tr>
    <tr>
      <td valign="top">

        <!-- START: QUICK LINKS MODULE //-->
        <div style="padding-bottom:8px;"><b><%=langMessageLinks%></b></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="../images/newmail_small.jpg" width="16" height="16" align="absmiddle">&nbsp;<a href="compose.asp"><%=langNewMessage%></a></div>
        <br>
        <div style="padding-bottom:8px;"><b><%=langMessageBoxes%></b></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="../images/folder_opened.gif" width="16" height="16" align="absmiddle">&nbsp;<a href="../messages"><i><%=langInbox%></i></a></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="../images/folder_closed.gif" width="16" height="16" align="absmiddle">&nbsp;<a href="drafts.asp"><%=langDrafts%></a></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="../images/folder_closed.gif" width="16" height="16" align="absmiddle">&nbsp;<a href="sentmail.asp"><%=langSentMail%></a></div>
        <br>        
        <div style="padding-bottom:3px;"><%=langSearchMessages%>:</div>
        <input type="text" style="background-color:#eeeeee; border:1px solid #000000; width:144px;"><br>
        <div class="quicklink" align="right"><a href="#"><img src="../images/go.gif" border="0"><%=langGo%></a>&nbsp;&nbsp;</div>
        <!-- END: QUICK LINKS MODULE //-->
      </td>
      <td valign="top">
        <form name="frmMsgList" action="delete.asp" method="post">
          <input type="hidden" name="IsCreator" value="0">
          <input type="hidden" name="BackPage" value="../messages">
          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/arrow_back.gif" align="absmiddle">
          
          <%
          If iPage > 1 Then
            Response.Write "<a href=""default.asp?page=" & iPage-1 & """>" & langPrev & "  " & Session("PageSize") & "</a>&nbsp;&nbsp;"
          Else
            Response.Write "<font color=""#999999"">Prev " & Session("PageSize") & "</font>&nbsp;&nbsp;"
          End If

          If iNumMoreRecords > 0 Then
            Response.Write "<a href=""default.asp?page=" & iPage+1 & """>" & langNext & " " & Session("PageSize") & "</a>"
          Else
            Response.Write "<font color=""#999999"">Next " & Session("PageSize") & "</font>"
          End If
          %>
          
          <img src="../images/arrow_forward.gif" align="absmiddle">
          &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/small_delete.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmMsgList.submit();">Delete</a>
          </div>

          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
            <tr>
              <th width="1%">&nbsp;</th>
              <th align="left"><%=langFrom%></th>
              <th align="left" width="80%"><%=langSubject%></th>
              <th align="left"><%=langDate%></th>
            </tr>
            <%= sMsgs %>
          </table>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>
