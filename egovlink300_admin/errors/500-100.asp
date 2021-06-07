<!-- #include file="../includes/common.asp" //-->
<html>
<head>
  <title><%=langError%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <style type="text/css">
  <!--
    pre {font-size:12px;}
  //-->
  </style>
  <script language="Javascript">
  <!--
    function toggleError() {
      if (document.all.err_details.style.display == "none") {
        document.all.err_details.style.display = "";
        document.all.err_link.style.display = "none";
      }
    }
        
    //breakout of frames
    if (parent.frames.length)
      top.location.href = document.location;
  //-->
  </script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs 0,1%>

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_error.jpg"></td>
      <td><font size="+1"><b><%=langError%></b></font><br><br></td>
    </tr>
    <tr>
      <td valign="top" width="151"><img src="../images/spacer.gif" width="151" height="1" border="0"><br></td>
      <td>
        <%
        Dim oError, sDetail, sErrorNum, sErrorDesc, sAspCode, sAspDesc, sCategory, sFilename, sLineNum, sColNum, lColNum, sSourceCode
        Dim oCmd, sSql, sReferer

        'get error info
        Set oError = Server.GetLastError()

        sErrorNum   = CStr(oError.Number)
        sAspCode    = "'" & oError.ASPCode & "'"
        sErrorDesc  = oError.Description
        sAspDesc    = oError.ASPDescription
        sCategory   = oError.Category
        sFilename   = oError.File
        sLineNum    = oError.Line
        sColNum     = oError.Column
        sSourceCode = oError.Source
        sReferer    = Request.ServerVariables("HTTP_REFERER")
        Set oError = Nothing

        If sAspCode = "''" Then
          sAspCode = ""
        End If

        If IsNumeric(sColNum) Then
          lColNum = CLng(sColNum)
        Else
          lColNum = 0
        End If

        ' create the error string
        sDetail = "ASP Error " & sAspCode & " occurred " & Now()
        If Len(sCategory) Then
          sDetail = sDetail & " in " & sCategory
        End If
        sDetail = sDetail & vbCrLf & "Server: " & Request.ServerVariables("SERVER_NAME") & vbCrLf
        sDetail = sDetail & "Error Number: " & sErrorNum & " (0x" & Hex(sErrorNum) & ")" & vbCrLf

        If Len(sFilename) Then
          sDetail = sDetail & "File: " & sFileName
          If sLineNum > "0" Then
            sDetail = sDetail & ", line " & sLineNum
            If lColNum > 0 Then
              sDetail = sDetail & ", column " & lColNum
              If Len(sSourceCode) Then
                sDetail = sDetail & vbCrLf & sSourceCode & vbCrLf & String(lColNum-1, "-") & "^"
              End If
            End If
          End If
          sDetail = sDetail & vbCrLf
        End If
        sDetail = sDetail & sErrorDesc & vbCrLf
        If Len(sAspDesc) Then
          sDtail = sDetail & "ASP reports: " & sAspDesc & vbCrLf
        End If

        sSql = "EXEC NewAuditEvent " & Session("UserID") & ", 'Error','" & sFileName & "','" & Replace(sDetail,"'","''") & "'"

        Set oCnn = Server.CreateObject("ADODB.Connection")
        oCnn.Open Application("DSN")
        oCnn.Execute sSql
        oCnn.Close
        Set oCnn = Nothing

        Response.Write langErrorSorry & "<br><br><a id=""err_link"" href=""javascript:toggleError();"">View Error Details</a>"
        Response.Write "<div id=""err_details"" style=""display:none; padding:10px;""><pre>" & Server.HTMLEncode(sDetail) & "</pre></div>"
        %>

        <form action="<%=sReferer%>">
          <input type="submit" value="Return to Previous Page">
        </form>
      </td>
    </tr>
  </table>
</body>
</html>