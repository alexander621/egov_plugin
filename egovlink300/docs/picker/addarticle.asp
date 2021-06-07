<!-- #include file="functions.inc" //-->
<!-- #include file="../includes/common.asp" //-->
<%
Server.ScriptTimeout = 300    ' 5 minute timeout
Response.Buffer = True
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
%>

<html>
<head>
 
  <%
  Sub DrawListBoxContents(dirPath, padding)
    Set objDir = objFSO.GetFolder(Server.MapPath(dirPath))
    For Each objFound in objDir.SubFolders
      If Left(objFound.Name,1) <> "_" And (objFound.Name <> "z.Help") Then
        If padding = "" Then
          Response.Write "<option style=""background-color:#cccccc;"" value=""" & dirPath & "/" & objFound.Name & "/"">" & padding & objFound.Name
        Else
          Response.Write "<option value=""" & dirPath & "/" & objFound.Name & "/"">" & padding & objFound.Name
        End If
        DrawListBoxContents dirPath & "/" & objFound.Name, padding & "&nbsp;&nbsp;"
      End If
    Next
    Set objDir = Nothing
  End Sub
  %>

  <script language="Javascript">
  <!--
    function showMethod( layerName ) {
      if (layerName == "direct") {
        link.style.display = "none";
        upload.style.display = "none";
        direct.style.display = "";  
        frmAddArticle.encoding = "application/x-www-form-urlencoded";
        frmAddArticle.action = "addarticle.asp?task=ADD&method=direct";
      }
      else if (layerName == "upload") {
        link.style.display = "none";
        direct.style.display = "none";
        upload.style.display = "";
        frmAddArticle.encoding = "multipart/form-data";
        frmAddArticle.action = "upload.asp?task=UPLOAD";
      }
      else {
        direct.style.display = "none";
        upload.style.display = "none";
        link.style.display = "";
        frmAddArticle.encoding = "application/x-www-form-urlencoded";
        frmAddArticle.action = "addarticle.asp?task=ADD&method=link";
      } 
    }
  //-->
  </script>
  <link href="global.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#ffffff" style="margin-left: 15pt; margin-top: 10pt">
<font size="3"><b><%=langAddDocument1%></b></font>
<hr style="height:1px;"><br>

<%
Select Case Request("task")

  Case "ADD"
    If Request("method") = "upload" Then
      strTitle = Request("strTitle")

    Else
      If Request("txtMethod") = "link" Then
        strTitle = Request.Form("txtURLTitle") & ".htm"
        If Request("chkOverwrite") = "on" Then  blnOverwrite = True      Else blnOverwrite = False
        
        If objFSO.FileExists(Server.MapPath(Request.Form("txtTopic") & strTitle)) And Not blnOverwrite Then
          strTitle = ""
        Else
          Set objNewFile = objFSO.CreateTextFile(Server.MapPath(Request.Form("txtTopic") & strTitle), True)
          If Request("openNew") = "on" Then
            objNewFile.Write( "<html><body><script language=""Javascript"">window.open('" & Request.Form("txtURL") & "');</script></body></html>" )
          Else
            objNewFile.Write( "<META HTTP-EQUIV=refresh CONTENT=""0; URL=" & Request.Form("txtURL") & """>" )
          End If
          objNewFile.Close

          '---BEGIN: Update DB fields--------------------------------
          'Set cnnECDB = Server.CreateObject("ADODB.Connection")
          'cnnECDB.Open Application("ECapture_ConnectionString")

          'strPath = Replace(Request.Form("txtTopic"), "/eclipse/tools/ecapture/Categories/", "")
          'strTitle = Replace(strTitle, "'", "''")

          'sSQL = "INSERT INTO ActionLog (ActionType,Article,Category,ActionBy,ActionByEmail,IPAddress) VALUES ('ARTICLE - ADD'" _
          '                                & ",'" & strTitle & "','" & strPath & "','" & Session("UserName") &  "','" & Session("EmailName") _
          '                                & "','" & Request.ServerVariables("REMOTE_ADDR") & "');"
          'cnnECDB.Execute sSQL
          'cnnECDB.Close
          'Set cnnECDB = Nothing
          '---END: Update DB fields----------------------------------
        End If
      Else
        If Request.Form("txtTitle") <> "" And Request.Form("txtContent") <> "" Then
          If Request.Form("blnIsHTML") = "on" Then
            strTitle = Request.Form("txtTitle") & ".htm"
            strContent = Request.Form("txtContent")
          Else
            strTitle = Request.Form("txtTitle") & ".txt"
            strContent = CRsafe(Request.Form("txtContent"))
          End If

          If Request("chkOverwrite") = "on" Then    blnOverwrite = True      Else blnOverwrite = False

          If objFSO.FileExists(Server.MapPath(Request.Form("txtTopic") & strTitle)) And Not blnOverwrite Then
            strTitle = ""
          Else
            Set objNewFile = objFSO.CreateTextFile(Server.MapPath(Request.Form("txtTopic") & strTitle), True)
            objNewFile.Write( Request.Form("txtContent") )
            objNewFile.Close

            '---BEGIN: Update DB fields--------------------------------
            'Set cnnECDB = Server.CreateObject("ADODB.Connection")
            'cnnECDB.Open Application("ECapture_ConnectionString")

            'strPath = Replace(Request.Form("txtTopic"), "/eclipse/tools/ecapture/Categories/", "")

            'sSQL = "INSERT INTO ActionLog (ActionType,Article,Category,ActionBy,ActionByEmail,IPAddress) VALUES ('ARTICLE - ADD'" _
            '                              & ",'" & strTitle & "','" & strPath & "','" & Session("UserName") &  "','" & Session("EmailName") _
            '                              & "','" & Request.ServerVariables("REMOTE_ADDR") & "');"
            'cnnECDB.Execute sSQL
            'cnnECDB.Close
            'Set cnnECDB = Nothing
            '---END: Update DB fields----------------------------------
          End If
        End If
      End If
    End If

    If strTitle <> "" Then
      Response.Write "<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='main.asp' style='font-family: Verdana, Arial, Helvetica; font-weight: bold; font-size: 8pt;' >Back To Documents Main</a><br><br>"
	  Response.Write "<b>Document &quot;" & strTitle & "&quot; was added successfully.</b>"
      Response.Write "<script language=""Javascript"">parent.fraToc.RefreshPath();</script>"
    Else
	  Response.Write "<p><img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='main.asp' style='font-family: Verdana, Arial, Helvetica; font-weight: bold; font-size: 8pt;'>Back To Documents Main</a></p>"
      Response.Write "<b>Document was not added.  Another file by this name already exists.</b>"
    End If 

  Case Else %>
    <form name="frmAddArticle" action="upload.asp?task=UPLOAD" method="POST" enctype="multipart/form-data">
      <table border=0 cellpadding=2 cellspacing=0 width="90%">
        <tr>
          <td colspan="3"><b>Step 1: </b><%=langDocDestiny%><br><br></td>
        </tr>
        <tr>
          <td width="10"><img src="image/spacer.gif" width="10" height="1" border="0"></td>
          <td colspan="2">
            <font color="#666666"><%=langCreateNewDocIn%></font><br>
            <input type="hidden" name="txtTopic">
            <%
            Dim strPath
            strPath = Request("path") & "/"
            If strPath = "/" Then %>
              <select name="selTopic" style="width:100%" onchange="document.all.txtTopic.value=this.value;">
                <option selected><%=langChooseDoc%></option>
                <% DrawListBoxContents Application("ECapture_ArticlesPath"), "" %>
              </select>
            <%
            Else
              Response.Write "<script lanuage=""Javascript"">document.all.txtTopic.value='"& strPath &"';</script>" & Replace(strPath, Application("eCapture_ArticlesPath") & "/", "")
            End If %>
          </td>
        </tr>
        <tr>
          <td colspan="3"><br><br><b>Step 2: </b><%=langChooseDocMethod%><br><br></td>
        </tr>
        <tr>
          <td></td>
          <td colspan="2">
            <font color="#666666"><%=langDocMethod%></font><br>
            <select name="txtMethod" onchange="showMethod(this[this.selectedIndex].value);">
              <option value="upload"><%=langDocUpload%></option>
              <option value="direct"><%=langDocDirect%></option>
              <option value="link"><%=langDocLink%></option>
            </select>
          </td>
        </tr>
        <tr>
          <td colspan="3"><br><br><b>Step 3: </b><%=langDefineDoc%><br><br></td>
        </tr>
        <tr>
          <td></td>
          <td colspan="2">
            <div id="upload">
              <font color="#666666"><%=langDocSelect%></font><br>
              <input type="file" name="binFile" size="50">
            </div>
            <div id="direct" style="display:none;">
              <table>
                <tr>
                  <td><font color="#666666"><%=langDocTitle%></font></td>
                  <td><input type="text" name="txtTitle" size=50 maxlength="30"></td>
                </tr>
                <tr>
                  <td><font color="#666666"><%=langDocContent%></font>&nbsp;&nbsp;</td>
                  <td><input type=checkbox name="blnIsHTML"><%=langDocHTML%>
                <tr>
                  <td colspan=2>
                    <textarea name="txtContent" rows=25 cols=75></textarea>
                  </td>
                </tr>
              </table>
            </div>
            <div id="link" style="display:none;">
              <table>
                <tr>
                  <td><font color="#666666"><%=langDocTitle%></font></td>
                  <td><input type="text" name="txtURLTitle" size=50 maxlength="30"></td>
                </tr>
                <tr>
                  <td><font color="#666666"><%=langDocURL%></font>&nbsp;&nbsp;</td>
                  <td><input type="text" name="txtURL" size=50 maxlength="255"></td>
                </tr>
                <tr>
                  <td></td>
                  <td><input type="checkbox" name="openNew"><%=langDocNewWindow%></td>
              </table>
            </div>
          </td>
        </tr>
        <tr>
          <td colspan="3"><br><br><b>Step 4: </b><%=langYourDocAdd%><br><br></td>
        </tr>
        <tr>
          <td></td>
          <td colspan="2">
            <input type="checkbox" name="chkOverwrite" onclick="if (this.checked) { if (!confirm('Are you sure you want to overwrite the existing file if it exists?', 2)) { this.checked=false; return true;} }"><%=langDocOverwrite%><br>
            <br>
            <input type="submit" value="<%=langAddDocument1%>"></td>
        </tr>
      </table>
    </form>
<% End Select %>
</body>
</html>