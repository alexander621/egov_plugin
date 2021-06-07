<!-- #include file="../includes/common.asp" -->
<%
Response.Buffer = True

Dim objFSO, strDir, oCmd, i

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
%>

<html>
<head>
  <link href="global.css" rel="stylesheet" type="text/css">
  <%
  Sub DeleteFiles(dirPath, fileList)
    
    ' DELETE FROM THE FILE SYSTEM
    objFSO.DeleteFile(Server.MapPath(Request.Form("txtArticlePath")))
  
    ' DELETE DOCUMENT FROM THE DATABASE
	strDir = Request.Form("txtArticlePath")
	If strDir <> "" Then 
	'---BEGIN: Update DB fields--------------------------------
      Set oCmd = Server.CreateObject("ADODB.Command")
      With oCmd
      .ActiveConnection = Application("DSN")
      .CommandText = "DelDocument"
      .CommandType = adCmdStoredProc
      .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
      .Parameters.Append oCmd.CreateParameter("DocumentURL", adVarChar, adParamInput, 255, strDir)
      .Execute
      End With
      Set oCmd = Nothing
      '---END: Update DB fields----------------------------------
     
	 End If
    
  End Sub
  %>
</head>

<body bgcolor="#ffffff" style="margin-left: 15pt; margin-top: 10pt">
<font size="3"><b>Delete an Document</b></font>
<hr style="height:1px;"><br>

<%
Select Case Request("task")

  Case "DELETE"
    Dim strFiles
    strFiles = ""
    For i = 0 to Request.Form("hdnFileCount")
      If Request.Form("chk" & i) <> "" Then strFiles = strFiles & Request.Form("chk" & i) & "/"      
    Next
    DeleteFiles Request.Form("txtPath"), strFiles
    Response.Write "<p><img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='main.asp' style='font-family: Verdana, Arial, Helvetica; font-weight: bold; font-size: 8pt;'>Back To Documents Main</a></P>"
    Response.Write "<b>Document(s) were deleted successfully.</b>"
    Response.Write "<script language=""Javascript"">parent.fraToc.document.location.reload();</script>"
    
  Case Else
    If Session("ECapture_Error") <> "" Then
      'Response.Write "<h3>" & Session("ECapture_Error") & "</h3>"
      Session("ECapture_Error") = ""
    End If
    
    Dim strDirPath, pos, strArticlePath
    strArticlePath = Request("path") & ""
    pos = InStrRev(strArticlePath, "/")
    If pos > 0 Then
      strDirPath = Left(strArticlePath, pos-1)
    End If
    %>
    <form name="frmDraw" action="delarticle.asp" method="post">
      <input type="hidden" name="task" value="DRAW">
      <table border="0" cellpadding="2" cellspacing="0" width="90%">

          <tr>
            <td colspan="3"><b>Step 1: </b>Confirm document deletion<br><br></td>
          </tr>
          <tr>
            <td></td>
            <td colspan="2"> 
              <input type="checkbox" onclick="document.all.btnSubmit.disabled = !this.checked;">
              <%= "<b>" & Mid(strArticlePath,pos+1) & "</b>&nbsp;&nbsp;(" & Replace(strArticlePath, Application("ECapture_ArticlesPath") & "/", "") & ")" %>
                <input type=hidden name="txtArticlePath" value="<%= strArticlePath %>">
                <script>document.frmDraw.task.value = "DELETE";</script>
                </td>
              </tr>
              <tr>
                <td colspan="3" align="left"><br><br><b>Step 2: </b>Delete the document(s)<br><br></td>
              </tr>   
              <tr>
                <td></td>
                <td colspan="2"><input type="submit" name="btnSubmit" value="Delete Document(s)" disabled></td>
              </tr>

      </table>
    </form>
<% End Select %>
</body>
</html>