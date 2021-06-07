<%
Response.Buffer = True
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
strPath = Request.Form("txtDir") & ""
%>
<html>
<head>
  <link href="global.css" rel="stylesheet">
  <%
  Sub DrawListBoxContents(dirPath, padding)
    Set objDir = objFSO.GetFolder(Server.MapPath(dirPath))
    For Each objFound in objDir.SubFolders
      If Left(objFound.Name,1) <> "_" Then
        If dirPath & "/" & objFound.Name & "/" = strPath Then
          Response.Write "<option value=""" & dirPath & "/" & objFound.Name & "/"" SELECTED>" & padding & objFound.Name
          DrawListBoxContents dirPath & "/" & objFound.Name, padding & "&nbsp;&nbsp;"
        Else
          Response.Write "<option value=""" & dirPath & "/" & objFound.Name & "/"">" & padding & objFound.Name
          DrawListBoxContents dirPath & "/" & objFound.Name, padding & "&nbsp;&nbsp;"
        End If
      End If
    Next
    Set objDir = Nothing
  End Sub
  
  Function GetFileList(dirPath)
  Dim strFileList
    
    strFileList = ""

    Set objDir = objFSO.GetFolder(Server.MapPath(dirPath))
    For Each objFile in objDir.Files
      strFileList = strFileList & objFile.Name & "/"
    Next

    GetFileList = strFileList
  End Function
  
  Sub DeleteFiles(dirPath, fileList)
    Set cnnECDB = Server.CreateObject("ADODB.Connection")
    cnnECDB.Open Application("ECapture_ConnectionString")

    loc = InStr(fileList,"/")
    Do While (loc > 0)
      file = Left(fileList,loc-1)
      fileList = Mid(fileList,loc+1)
      filePath = dirPath & file
      
      objFSO.DeleteFile(Server.MapPath(filePath))

      '---BEGIN: Update DB fields--------------------------------
      'strPath = Replace(dirPath, "/eclipse/tools/ecapture/Categories/", "")

      'sSQL = "DELETE FROM ActionLog WHERE Article='" & file & "' AND Category='" & strPath & "'"
      'cnnECDB.Execute sSQL
      
      'sSQL = "INSERT INTO ActionLog (ActionType,Article,Category,ActionBy,ActionByEmail,IPAddress) VALUES ('ARTICLE - DELETE'" _
      '                                    & ",'" & file & "','" & strPath & "','" & Session("UserName") &  "','" & Session("EmailName") _
      '                                    & "','" & Request.ServerVariables("REMOTE_ADDR") & "');"
      'cnnECDB.Execute sSQL
      '---END: Update DB fields----------------------------------
      
      loc = InStr(fileList, "/")
    Loop

    cnnECDB.Close
    Set cnnECDB = Nothing
  End Sub
  %>
  <script language="Javascript1.2">
  <!--
    function DrawTree(dirPath, fileList) {
      var padding = "";
    
      document.write("<form action='delarticle.asp' method=post>");
      path = dirPath;
      while ((loc = path.search("/")) != -1) {
        node = path.slice(0,loc);
        path = path.replace(node + "/", "");
        document.write("<table border=0 cellpadding=0 cellspacing=0><tr><td>" + padding);
        document.write("<img src='images/ftv2folderopen.gif' border=0></td><td valign=middle nowrap>" + node + "</td></tr></table>");
        padding = padding + "&nbsp;&nbsp;&nbsp;";
      }
      i=0;
      document.write("<table border=0 cellpadding=0 cellspacing=0>");
      while ((loc = fileList.search("/")) != -1) {
        file = fileList.slice(0,loc);
        fileList = fileList.replace(file + "/", "");
        fileName = file.slice(0, file.length-4);
        document.write("<tr><td>" + padding + "</td><td><input type=checkbox name='chk" + i + "' value='" + file + "'></td><td valign=middle nowrap>" + fileName + "</td></tr>");
        i++;
      }
      document.write("</table><input type=hidden name='hdnFileCount' value='" + i + "'></td></tr>");
    }
  //-->
  </script>
</head>

<body bgcolor="#ffffff">
<font size="3"><b>Delete an Article</b></font>
<hr style="height:1px;"><br>

<%
Select Case Request("task")

  Case "DELETE"
    strFiles = ""
    For i = 0 to Request.Form("hdnFileCount")
      If Request.Form("chk" & i) <> "" Then strFiles = strFiles & Request.Form("chk" & i) & "/"      
    Next
    DeleteFiles Request.Form("txtPath"), strFiles

    Response.Write "<b>Article(s) were deleted successfully.</b>"
    Response.Write "<script language=""Javascript"">parent.fraToc.RefreshPath();</script>"
    
  Case Else
    If Session("ECapture_Error") <> "" Then
      Response.Write "<h3>" & Session("ECapture_Error") & "</h3>"
      Session("ECapture_Error") = ""
    End If %>
    <form name="frmDraw" action="delarticle.asp" method="post">
      <input type="hidden" name="task" value="DRAW">
      <table border="0" cellpadding="2" cellspacing="0" width="90%">
      <% If Request.Form("task") = "DRAW" Then %>
          <tr>
            <td colspan="3"><b>Step 1: </b>Choose the category of the article(s)<br><br></td>
          </tr>
          <tr>
            <td width="10"><img src="image/spacer.gif" width="10" height="1" border="0"></td>
            <td colspan="2" width="100%">
              <font color="#999999">Category:</font><br>
              <%= Replace(strPath, Application("ECapture_ArticlesPath") & "/", "") %>
            </td>
          </tr>
          <tr>
            <td colspan="3"><br><br><b>Step 2: </b>Choose articles to delete<br><br></td>
          </tr>
          <tr>
            <td></td>
            <td colspan="2"> 
            <% 
              fileList = GetFileList(strPath)

              If Len(fileList) = 0 Then
                Response.Write "<font color=""#ff0000"">This category contains no articles.</font><br>" & vbCrLf
              %>
                <br><br>
                <input type="button" value="Try Again" onclick="history.back();">
                </td>
              </tr>
              <%
              Else
                Response.Write "<script>DrawTree(""Categories" & Replace(strPath, Application("ECapture_ArticlesPath"), "") & """, """ & fileList & """);</script>"
              %>
                <input type=hidden name="txtPath" value="<%= strPath %>">
                <script>document.frmDraw.task.value = "DELETE";</script>
                </td>
              </tr>
              <tr>
                <td colspan="3" align="left"><br><br><b>Step 3: </b>Delete the article(s)<br><br></td>
              </tr>   
              <tr>
                <td></td>
                <td colspan="2"><input type="submit" name="btnSubmit" value="Delete Article(s)"></td>
              </tr>
           <% End If %>
      <% Else %>
          <tr>
            <td colspan="3"><b>Step 1: </b>Choose the category of the article(s)<br><br></td>
          </tr>
          <tr>
            <td width="10"><img src="image/spacer.gif" width="10" height="1" border="0"></td>
            <td colspan="2" width="100%">
              <font color="#666666">Choose category:</font><br>
              <select name="txtDir" style="width:100%" onchange="document.frmDraw.submit();">
                <option>Choose one...</option>
                <% DrawListBoxContents Application("ECapture_ArticlesPath"), "" %>
              </select>
            </td>
          </tr>
       <% End If %>

      </table>
    </form>
<% End Select %>
</body>
</html>