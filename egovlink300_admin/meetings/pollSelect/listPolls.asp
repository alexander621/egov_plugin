<!-- #include file="../../includes/common.asp" //-->
<%
Response.Buffer = True

Dim sSql, oRst, sVotes, iCount, iTotal, sBgcolor, iPage, sPage, sTopic, iNumMoreRecords
Dim bCanCreate, bCanEdit

sPage = Request.QueryString("gp")
If sPage & "" <> "" Then
  iPage = clng(sPage)
Else
  iPage = 1
End If

bCanCreate = HasPermission("CanCreatePolls")
bCanEdit = HasPermission("CanEditPolls")

sSql = "EXEC ListVotes " & Session("OrgID") & "," & Session("UserID") & "," & Session("PageSize") & "," & iPage

Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .ActiveConnection = Application("DSN")
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open sSql
  .ActiveConnection = Nothing
End With

sVotes = ""
iCount = 1
iTotal = oRst.RecordCount

If Not oRst.EOF Then
  sBgcolor = "#ffffff"
  iNumMoreRecords = oRst("NumMoreRecords")

  Do While Not oRst.EOF
    
    sVotes = sVotes & "<td style=""padding:0px;""><img src=""../../images/newpoll.gif"" border=""0"">&nbsp;</td>"
    
    sVotes = sVotes & "<td nowrap><a href=""#"" onClick=""saveSelection('" & oRst("VoteID") & "', '"& oRst("Subject") &"')"">" & oRst("Subject") & "</a></td>"
    
    sVotes = sVotes & "<td nowrap>" & oRst("VoteType") & "</td>"
    sVotes = sVotes & "</td><td align=""center"">" & oRst("NumResponses") & "</td>"

    If oRst("Status") = 1 Then
      sVotes = sVotes & "<td nowrap>Open</td>"
    Else
      sVotes = sVotes & "<td nowrap>Closed</td>"
    End If

    If oRst("AccessID") > 0 Then
      sVotes = sVotes & "<td nowrap><img src=""../../images/locked.gif"" border=""0"" alt=""This voting poll is both hidden & locked to unauthorized users.""></td></tr>"
    Else
      sVotes = sVotes & "<td nowrap>&nbsp;</td></tr>"
    End If

    If sBgcolor = "#ffffff" Then sBgcolor = "#eeeeee" Else sBgcolor = "#ffffff"
    iCount = iCount + 1
    oRst.MoveNext
  Loop
  oRst.Close
Else
  sVotes = "<tr><td colspan=""6"">No new voting polls.</td></tr>"
End If

Set oRst = Nothing
%>

<html>
<head>
  <title><%=langBSVoting%></title>
  <link href="../../global.css" rel="stylesheet" type="text/css">
  <script src="../../scripts/selectAll.js"></script>
  <script>
  function saveSelection(id, name)
  {
    var objParent=window.opener;
	objParent.addItem.link.value=name;
	objParent.addItem.itemID.value=id;
	if(objParent.addItem.title.value=="")objParent.addItem.title.value=name;
    window.close();
  }
  </script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td colspan="2" valign="top">
        <form name="DelPolls" method=post action="deletepolls.asp" method="post">
          <div style="font-size:10px; padding-bottom:5px;"><img src="../../images/arrow_back.gif" align="absmiddle">
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
          <img src="../../images/arrow_forward.gif" align="absmiddle">
          <%
          If bCanCreate Then
          %>
            &nbsp;&nbsp;&nbsp;&nbsp;<img src="../../images/newpoll.gif" align="absmiddle">&nbsp;<a href="newpoll.asp"><%=langNewPoll%></a>
          <%
          End If
          %>
          </div>

          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
            <tr>
              
              <th>&nbsp;</th>
              <th align="left" width="70%"><%=langTopic%></th>
              <th align="left" width="100" nowrap><%=langType%></th>
              <th align="center" width="100" nowrap><%=langResponses%></th>
              <th align="left" width="100" nowrap><%=langStatus%></th>
              <th width="1">&nbsp;</th>
            </tr>
            <%= sVotes %>
          </table>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>
