<!-- #include file="../includes/common.asp" //-->
<%
Dim oCnn, oRst, i, sSql, sExisting, sAvailable, iID, sTask, sReload, bAdd, bDelete, sPath, bBack, bForward, iSelectID

sPath = Trim(Request.QueryString("path"))
sTask = Request.QueryString("t")
bBack = True
bForward = True

'Response.Write sPath
'Response.End

If sTask = "add" Then
  Set oCnn = Server.CreateObject("ADODB.Connection")
  oCnn.Open Application("DSN")
  For Each iSelectID in Request.Form("RemainingList")
    If CLng(iSelectID) <> -1 Then
      sSql = "EXEC NewFolderAccessGroup '" & sPath & "', " & iSelectID
      oCnn.Execute sSql
    End If
  Next
  oCnn.Close
  Set oCnn = Nothing
ElseIf sTask = "del" Then
  Set oCnn = Server.CreateObject("ADODB.Connection")
  oCnn.Open Application("DSN")
  For Each iSelectID in Request.Form("ExistingList")
    If CLng(iSelectID) <> -1 Then
      sSql = "EXEC DelFolderAccessGroup '" & sPath & "', " & iSelectID
      oCnn.Execute sSql
    End If
  Next
  oCnn.Close
  Set oCnn = Nothing
End If

Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .ActiveConnection = Application("DSN")
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
End With

'----------------------------Get existing member groups
sSql = "SELECT g.GroupID, g.GroupName FROM DocumentFolders [df] INNER JOIN FeatureAccess [fa] ON fa.AccessID = df.AccessID INNER JOIN Groups [g] ON g.GroupID = fa.GroupID WHERE df.FolderPath = '" & sPath & "' ORDER BY GroupName"
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
sSql = "SELECT g.GroupID, g.GroupName FROM Groups [g] WHERE OrgID=" & Session("OrgID") & " AND GroupID NOT IN (SELECT g2.GroupID FROM DocumentFolders [df] INNER JOIN FeatureAccess [fa] ON fa.AccessID = df.AccessID INNER JOIN Groups [g2] ON g2.GroupID = fa.GroupID WHERE df.FolderPath = '" & sPath & "') ORDER BY GroupName"
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
  <link href="global.css" rel="stylesheet">
</head>

<body bgcolor="#ffffff">
  
  <div style="padding-top:9px;"><font style="font-family:Verdana,Arial; font-size:18px;">
  <img src="../images/newpermission.gif">
  <font size="3"><b>Edit Security:</b> <%=Right(sPath,Len(sPath)-InStrRev(sPath,"/"))%></font>
  </font></div><img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='main.asp'>Back To Documents Main</a>
 
  
  <hr style="height:1px;"><br>
  <table border="0" cellpadding="5" cellspacing="0">
    <tr>
      <td>
        <form name="c1" method="post" action="members.asp?t=del&path=<%=sPath%>">
          <table border="0" cellpadding="0" cellspacing="0" width="130">
            <tr>
              <td><b><%= langMember & "&nbsp;" & langCommittees %></b></td>
            </tr>
            <tr><td><br></td></tr>
            <tr>
              <td><%= sExisting %></td>
            </tr>
          </table>
        <input type="hidden" name="sMsg" value="Permissions successfully changed.">
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
        <form name="r1" method="post" action="members.asp?t=add&path=<%=sPath%>">
          <table border="0" cellpadding="0" cellspacing="0" width="130">
            <tr>
              <td><b><%=langCommittee%> List</b></td>
            </tr>
            <tr><td><br></td></tr>
            <tr>
              <td><%= sAvailable %></td>
            </tr>
          </table>
        <input type="hidden" name="sMsg" value="Permissions successfully changed." ID="Hidden1">
        </form>
      </td>
    </tr>
    <tr><td class=success><%=request.form("sMSG")%></td></tr>
  </table>
</body>
</html>