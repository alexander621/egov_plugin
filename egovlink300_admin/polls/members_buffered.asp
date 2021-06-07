<!-- #include file="../includes/common.asp" //-->
<%
Dim oRst, i, sSql, sExisting, sAvailable, iID, sTask, sReload, bAdd, bDelete, sMembers, bBack, bForward, iSelectID

sTask = Request.QueryString("t")
bBack = True
bForward = True

sMembers = Request.QueryString("mem") & ""
If sMembers = "" Then
  sMembers = NULL
End If

If sTask = "add" Then
  For Each iSelectID in Request.Form("RemainingList")
    If Len(sMembers) > 0 Then
      sMembers = sMembers & "," & iSelectID
    Else
      sMembers = iSelectID
    End If
  Next
  sReload = "onload=""window.opener.document.all.Members.value='" & sMembers & "';window.opener.document.all._task.value='reload';window.opener.document.frmNewPoll.submit();"""
ElseIf sTask = "del" Then
  For Each iSelectID in Request.Form("ExistingList")
    sMembers = Replace(sMembers, "," & iSelectID, "")
    sMembers = Replace(sMembers, iSelectID & ",", "")
    sMembers = Replace(sMembers, iSelectID, "")
  Next
  sReload = "onload=""window.opener.document.all.Members.value='" & sMembers & "';window.opener.document.all._task.value='reload';window.opener.document.frmNewPoll.submit();"""
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

If sMembers & "" = "" Then
  sMembers = 0
End If

'----------------------------Get existing member groups
sSql = "SELECT GroupID, GroupName FROM Groups WHERE GroupID IN (" & sMembers & ") AND OrgID=" & Session("OrgID") & " ORDER BY GroupName"
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
sSql = "SELECT GroupID, GroupName FROM Groups WHERE GroupID NOT IN (" & sMembers & ") AND OrgID=" & Session("OrgID") & " ORDER BY GroupName"
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
        <form name="c1" method="post" action="members_buffered.asp?t=del&mem=<%=sMembers%>">
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
        <form name="r1" method="post" action="members_buffered.asp?t=add&mem=<%=sMembers%>">
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