<!-- #include file="../includes/common.asp" -->
<%
Response.Buffer = True

Dim objFSO, objDir, objFound, strPath, strDir, oCmd

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
strPath = Request.Form("txtDir")
%>

<html>
<head>
  <link href="global.css" rel="stylesheet">
  <%
  Sub DrawListBoxContents(dirPath, padding)
    Set objDir = objFSO.GetFolder(Server.MapPath(dirPath))
	
    
    On Error Resume Next


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
  %>
  <script language="Javascript" src="scripts/vbstring.js"></script>
  <script language="Javascript1.2">
  <!--
    function DrawTree(dirPath) {
      var output = "";
      var padding = "";

      path = "Categories" + replace(dirPath, "<%= Application("ECapture_ArticlesPath") %>", "");
      while ((loc = path.search("/")) != -1) {
        node = path.slice(0,loc);
        path = path.replace(node + "/", "");
        output = output + "<table border=0 cellpadding=0 cellspacing=0><tr><td>" + padding;
        output = output + "<img src='../images/ftv2folderopen.gif' border=0></td><td valign=middle nowrap style='color:#666666;'>" + node + "</td></tr></table>";
        padding = padding + "&nbsp;&nbsp;&nbsp;";
      }
      output = output + "<br><input type=checkbox name='txtDelete' onclick='if (this.checked) {document.frmDraw.btnSubmit.disabled=false;} else {document.frmDraw.btnSubmit.disabled=true;}'><font color='#ff0000'> Delete the &quot;"+ node +"&quot; directory and all directories below it?<font>";

      document.all.step2.innerHTML = output;
    }
  //-->
  </script>
</head>

<body>
<font size="3"><b><%=langDeleteAFolder%></b></font>
<hr style="height:1px;"><br>

<%
Select Case Request("task")

  Case "DEL"
    strDir = Request.Form("txtDir") & Request.Form("txtDirName")

	If objFSO.FolderExists(Server.MapPath(strDir)) Then
      objFSO.DeleteFolder(Server.MapPath(strDir))

      '---BEGIN: Update DB fields--------------------------------
      Set oCmd = Server.CreateObject("ADODB.Command")
      With oCmd
      .ActiveConnection = Application("DSN")
      .CommandText = "DelFolder"
      .CommandType = adCmdStoredProc
      .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
      .Parameters.Append oCmd.CreateParameter("FolderPath", adVarChar, adParamInput, 300, strDir)
      .Execute
      End With
      Set oCmd = Nothing
      '---END: Update DB fields----------------------------------

      Response.Write "<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='main.asp' style='font-family: Verdana, Arial, Helvetica; font-weight: bold; font-size: 8pt;' >Back To Documents Main</a><br><br>"
      Response.Write "<b>Folder was deleted successfully.</b>"
      Response.Write "<script language=""Javascript"">parent.fraToc.RefreshPath();</script>"
    Else
	  Response.Write "<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='main.asp' style='font-family: Verdana, Arial, Helvetica; font-weight: bold; font-size: 8pt;' >Back To Documents Main</a><br><br>"
      Session("ECapture_Error") = "The folder you attempted to delete does not exist.<br>Please try again."
      Response.Redirect "delcategory.asp?task=browse"
    End If

  Case Else
    If Session("eCapture_Error") <> "" Then
      Response.Write "<h3>" & Session("eCapture_Error") & "</h3>"
      Session("eCapture_Error") = ""
    End If
    %>
    <form name="frmDraw" action="delcategory.asp" method="post">
      <input type="hidden" name="task" value="DEL">
      <input type="hidden" name="txtDir" value="">
      <table border="0" cellpadding="2" cellspacing="0" width="90%">
        <tr>
          <td colspan="3">
            <table border=0 cellpadding=0 cellspacing=0>
              <tr>
                <td><b>Step 1:&nbsp;</b></td>
                <td><%=langChooseFolderDelete%></td>
              </tr>
            </table>
            <br>
        </tr>
        <tr>
          <td width="10"><img src="image/spacer.gif" width="10" height="1" border="0"></td>
          <td colspan="2" width="100%">
            <div id="step1">
              <font color="#666666">Choose folder:</font><br>
              <select name="selDir" style="width:100%" onchange="document.all.txtDir.value=this.value; DrawTree(this[selectedIndex].value);">
                <option>Choose one...</option>
                <% DrawListBoxContents Application("ECapture_ArticlesPath"), "" %>
              </select>
            </div>
          </td>
        </tr>

          <tr>
            <td colspan="3"><br><br><b>Step 2: </b><%=langDeleteVerify%><br><br></td>
          </tr>
          <tr>
            <td></td>
            <td colspan="2"><div id="step2"></div></td>
          </tr>
          <tr>
            <td colspan="3" align="left"><br><br><b>Step 3: </b><%=langDeleteTheFolder%><br><br></td>
          </tr>   
          <tr>
            <td></td>
            <td colspan="2"><input type="submit" name="btnSubmit" value="<%=langDeleteButton%>" disabled></td>
          </tr>

      </table>
    </form>
<% End Select %>
</body>
</html>

<%
strPath = Request("path") & ""
If strPath <> "" Then
  Response.Write "<script lanuage=""Javascript"">document.all.txtDir.value='"& EscapeSingleQuote(strPath) &"';document.all.step1.innerHTML='"& EscapeSingleQuote(Replace(strPath, Application("ECapture_ArticlesPath") & "/", "")) &"'; DrawTree('"& EscapeSingleQuote(strPath) & "/" &"');</script>"
End If

Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

Function EscapeSingleQuote(sValue)
	EscapeSingleQuote = replace(sValue,"'","\'")
End Function

%>