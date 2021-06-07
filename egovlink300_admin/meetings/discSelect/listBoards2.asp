<!-- #include file="../../includes/common.asp" //-->
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
sStyle = ""
iCount = 1
iTotal = oRst.RecordCount

If Not oRst.EOF Then
  sBgcolor = "#ffffff"
  iNumMoreRecords = oRst("NumMoreRecords")

  Do While Not oRst.EOF
    sGroups = sGroups & "<tr bgcolor=""" & sBgcolor & """><td>&nbsp;</td>"
    
    sGroups = sGroups & "<td style=""padding:0px;""><img src=""../../images/newdiscgroup.gif"" border=""0"">&nbsp;</td>"
    
    sGroups = sGroups & "<td nowrap><a href=""topics.asp?tid=" & oRst("DiscussionGroupID") & "&tp=1&gn=" & Server.URLEncode(oRst("Name")) & """>" & oRst("Name")
    
    sGroups = sGroups & "</td><td align=""center"">" & oRst("NumTopics") & "</td>"
    sGroups = sGroups & "<td nowrap>" & oRst("Description") & "&nbsp;</td>"

    If oRst("AccessID") > 0 Then
      sGroups = sGroups & "<td nowrap><img src=""../../images/locked.gif"" border=""0""></td></tr>"
    Else
      sGroups = sGroups & "<td nowrap>&nbsp;</td></tr>"
    End If

    If sBgcolor = "#ffffff" Then sBgcolor = "#eeeeee" Else sBgcolor = "#ffffff"
    iCount = iCount + 1
    oRst.MoveNext
  Loop
  oRst.Close
Else
  sGroups = "<tr><td colspan=""6"">No new messages.</td></tr>"
End If

Set oRst = Nothing
%>

<html>
<head>
  <title><%=langBSDiscussions%></title>
  <link href="<%=rootpath%>global.css" rel="stylesheet" type="text/css">
  <script src="<%=rootpath%>scripts/selectAll.js"></script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">


  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    
    <tr>
       <td colspan="2" valign="top">
        <form name="DelGroups" action="<%=rootpath%>deletegroups.asp?<%= Request.QueryString() %>" method="post">
          <div style="font-size:10px; padding-bottom:5px;"><img src="<%=rootpath%>images/arrow_back.gif" align="absmiddle">
  
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

          <img src="<%=rootpath%>images/arrow_forward.gif" align="absmiddle">
          
          </div>
          
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
            <tr>
              <th>&nbsp;</th>
              <th align="left"><%=langDiscussionBoard%></th>
              <th align="center"><%=langTopics%></th>
              <th align="center">&nbsp</th>
              <th align="center">&nbsp</th>
              <th align="center">&nbsp</th>
              </tr>	
            <%= sGroups %>
          </table>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>
