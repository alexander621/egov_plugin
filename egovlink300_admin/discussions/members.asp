<!-- #include file="../includes/common.asp" //-->
<%
Dim oCnn, oRst, i, sSql, sExisting, sAvailable, iID, sTask, sReload, bAdd, bDelete, bBack, bForward, SelectID

iID = Request.QueryString("id")
sTask = Request.QueryString("t")
sReload = "onload=""window.opener.location.reload();"""
bBack = True
bForward = True

If sTask = "add" Then
  Set oCnn = Server.CreateObject("ADODB.Connection")
  oCnn.Open Application("DSN")
  For Each SelectID in Request.Form("RemainingList")
    If CLng(SelectID) <> -1 Then
      sSql = "EXEC NewDiscussionAccessGroup " & iID & ", " & SelectID
      oCnn.Execute sSql
    End If
  Next
  oCnn.Close
  Set oCnn = Nothing
ElseIf sTask = "del" Then
  Set oCnn = Server.CreateObject("ADODB.Connection")
  oCnn.Open Application("DSN")
  For Each SelectID in Request.Form("ExistingList")
    If CLng(SelectID) <> -1 Then
      sSql = "EXEC DelDiscussionAccessGroup " & iID & ", " & SelectID
      oCnn.Execute sSql
    End If
  Next
  oCnn.Close
  Set oCnn = Nothing
Else
  sReload = ""
End If

Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .ActiveConnection = Application("DSN")
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
End With

'----------------------------Get existing member groups
sSql = "SELECT g.GroupID, g.GroupName FROM DiscussionGroups [dg] INNER JOIN FeatureAccess [fa] ON fa.AccessID = dg.AccessID INNER JOIN Groups [g] ON g.GroupID = fa.GroupID AND g.OrgID = " & Session("OrgID") & " WHERE dg.DiscussionGroupID = " & iID & " ORDER BY GroupName"
oRst.Open sSql

sExisting = ""
sExisting = sExisting & "<select size=""20"" style=""width:200px"" name='ExistingList' multiple>"
For i = 0 To oRst.RecordCount - 1
  sExisting = sExisting & "<option value=" & oRst("GroupID") & ">" & oRst("GroupName") & "</option>"
	oRst.MoveNext
Next
If i = 0 Then
  sExisting = sExisting & "<option value=""-1"">Everyone</option>"
  bForward = False
End If
sExisting = sExisting & "</select>"  
oRst.Close

'----------------------------Get available member groups
sSql = "SELECT g.GroupID, g.GroupName FROM Groups [g] WHERE OrgID=" & Session("OrgID") & " AND GroupID NOT IN (SELECT g2.GroupID FROM DiscussionGroups [dg] INNER JOIN FeatureAccess [fa] ON fa.AccessID = dg.AccessID INNER JOIN Groups [g2] ON g2.GroupID = fa.GroupID WHERE dg.DiscussionGroupID = " & iID & ") ORDER BY GroupName"
oRst.Open sSql

sAvailable = ""
sAvailable = sAvailable & "<select size=""20"" style=""width:200px;"" name='RemainingList' multiple>"
For i = 0 To oRst.RecordCount - 1
  sAvailable = sAvailable & "<option value=" & oRst("GroupID") & ">" & oRst("GroupName") & "</option>"
  oRst.MoveNext
Next
If i = 0 Then
  bBack = False
End If
sAvailable = sAvailable & "</select>"  
 
oRst.Close
Set oRst = Nothing
%>

<html>
<head>
  <title>Select Groups</title>
  <link href="../global.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#c9def0" <%=sReload%>>
  <table border="0" cellpadding="5" cellspacing="0">
    <tr>
      <td>
        <form name="c1" method="post" action="members.asp?t=del&id=<%=iID%>">
          <table border="0" cellpadding="0" cellspacing="0" width="130">
            <tr>
              <td><b><%= langMember & "&nbsp;" & langCommittees %></b></td>
            </tr>
            <tr><td><br></td></tr>
            <tr>
              <td><%= sExisting %></td>
            </tr>
          </table>
        </form>
      </td>
      <td>
        <% If bForward Then %>
          <a href="javascript:document.c1.submit();"><img src="../images/ieforward.gif" border="0"></a><br>
        <% Else %>
          <img src="../images/ieforward_disabled.gif" border="0"></a><br>
        <% End If %>
        <br>
        <% If bBack Then %>
          <a href="javascript:document.r1.submit();"><img src="../images/ieback.gif" border="0"></a><br>
        <% Else %>
          <img src="../images/ieback_disabled.gif" border="0"></a><br>
        <% End If %>
      </td>
      <td>
        <form name="r1" method="post" action="members.asp?t=add&id=<%=iID%>">
          <table border="0" cellpadding="0" cellspacing="0" width="130">
            <tr>
              <td><b><%=langCommittee%> List</b></td>
            </tr>
            <tr><td><br></td></tr>
            <tr>
              <td><%= sAvailable %></td>
            </tr>
          </table>
        </form>
      </td>
    </tr>
    <tr>
      <td colspan="3" align="center"><a href="javascript:window.close();">Close Window</a></td>
    </tr>
  </table>
</body>
</html>