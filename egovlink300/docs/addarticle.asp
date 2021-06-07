<!-- #include file="../includes/common.asp" //-->
<!-- #include file="functions.inc" //-->
<%
Server.ScriptTimeout = 180    ' 3 minute timeout

Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
%>

<html>
<head>
 
  <%
  Sub DrawListBoxContents(dirPath)
    Dim newpath, name, sSQL, oRst, i, iLevel
	
    sSql = "EXEC ListFolder " & Session("OrgID") & ", " & Session("UserID") & ", '" & dirPath & "'"
    Set oRst = Server.CreateObject("ADODB.Recordset")
		oRst.Open sSql, Application("DSN"), 3, 1

		If Not oRst.EOF Then
			Do While Not oRst.EOF
				'newpath = Server.URLEncode(oRst("FolderPath"))
				optionpath = oRst("FolderPath")
				name = oRst("FolderName")
        padding = ""
        iLevel = oRst("FolderLevel")*2
				
				If name="root" Then
					Response.Write "<option value=""" & dirPath & "/" & name & "/"" selected>Main Category" & vbcrlf
				Else
          For i = 0 To iLevel
            padding = padding & "&nbsp;"
          Next
					Response.Write "<option value=""" & optionpath & "/"">" & padding & name & vbcrlf
				End If
        oRst.MoveNext
			Loop
			oRst.Close
			Set oRst = Nothing
		End If
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
  <link href="../global.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#ffffff">

<%
Dim strDir, strLink, strContent, oCmd, strTitle, objNewFile, blnNewTarget, blnOverwrite

Select Case Request("task")

  Case "ADD"
    'CREATE AN UPLOADED DOCUMENT
    If Request("method") = "upload" Then
      strTitle = Request("strTitle")

    Else
      ' CREATE A LINKED DOCUMENT
      If Request("txtMethod") = "link" Then
        strTitle = Request.Form("txtURLTitle") & ".htm"
        If Request("chkOverwrite") = "on" Then  blnOverwrite = True      Else blnOverwrite = False
        
        
        If objFSO.FileExists(Server.MapPath(Request.Form("txtTopic") & strTitle)) And Not blnOverwrite Then
          strTitle = ""
        Else
          Set objNewFile = objFSO.CreateTextFile(Server.MapPath(Request.Form("txtTopic") & strTitle), True)
          If Request("openNew") = "on" Then
            objNewFile.Write( "<html><body><script language=""Javascript"">window.open('" & Request.Form("txtURL") & "');</script></body></html>" )
			blnNewTarget = 1
          Else
            objNewFile.Write( "<META HTTP-EQUIV=refresh CONTENT=""0; URL=" & Request.Form("txtURL") & """>" )
			blnNewTarget = 0
          End If
          objNewFile.Close
      
			strDir = Request.Form("txtTopic") & strTitle
			strLink = Request.Form("txtURL")
			'---BEGIN: Update DB fields for Document(LINK) --------------------------------
			Set oCmd = Server.CreateObject("ADODB.Command")
			With oCmd
			.ActiveConnection = Application("DSN")
			.CommandText = "NewDocument"
			.CommandType = adCmdStoredProc
			.Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
			.Parameters.Append oCmd.CreateParameter("CreatorID", adInteger, adParamInput, 4, Session("UserID"))
			.Parameters.Append oCmd.CreateParameter("FolderPath", adVarChar, adParamInput, 300, strDir)
			.Parameters.Append oCmd.CreateParameter("LinkURL", adVarChar, adParamInput, 300, strLink)
			.Parameters.Append oCmd.CreateParameter("LinkTargetsNew", adInteger, adParamInput, 4, blnNewTarget)
			.Execute
			End With
			Set oCmd = Nothing
			'---END: Update DB fields----------------------------------
        End If
      Else
        ' CREATE A NEW DOCUMENT WITH TEXT SUPPLIED
        If Request.Form("txtTitle") <> "" And Request.Form("txtContent") <> "" Then
          If Request.Form("blnIsHTML") = "on" Then
            strTitle = Request.Form("txtTitle") & ".htm"
            strContent = Request.Form("txtContent")
          Else
            strTitle = Request.Form("txtTitle") & ".txt"
            strContent = CRsafe(Request.Form("txtContent"))
          End If

          If Request("chkOverwrite") = "on" Then    blnOverwrite = True      Else blnOverwrite = False

          'If objFSO.FileExists(Server.MapPath(Request.Form("txtTopic") & strTitle)) And Not blnOverwrite Then
          If objFSO.FileExists(Server.MapPath(Request.Form("txtTopic") & strTitle)) And Not blnOverwrite Then
            response.Write request.Form("txtTopic")
            response.End
            
            strTitle = ""
          Else
            Set objNewFile = objFSO.CreateTextFile(Server.MapPath(Request.Form("txtTopic") & strTitle), True)
            objNewFile.Write( Request.Form("txtContent") )
            objNewFile.Close

			strDir = Request.Form("txtTopic") & strTitle
			'---BEGIN: Update DB fields for Document(DIRECT CONTENT) --------------------------------
			Set oCmd = Server.CreateObject("ADODB.Command")
			With oCmd
			.ActiveConnection = Application("DSN")
			.CommandText = "NewDocument"
			.CommandType = adCmdStoredProc
			.Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
			.Parameters.Append oCmd.CreateParameter("CreatorID", adInteger, adParamInput, 4, Session("UserID"))
			.Parameters.Append oCmd.CreateParameter("FolderPath", adVarChar, adParamInput, 300, strDir)
			.Parameters.Append oCmd.CreateParameter("LinkURL", adVarChar, adParamInput, 300, null)
			.Execute
			End With
			Set oCmd = Nothing
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
    <table border="0" cellspacing="0" class="start" width="100%">
      <tr>
        <td style="padding:10px 0px;"><font size="+1"><b><%=langDocuments%>: <%=langAddDocument%></b></font><br><br></td>
      </tr>
      <tr>
        <td width="100%">
          <br>
          <form name="frmAddArticle" action="upload.asp?task=UPLOAD" method="POST" enctype="multipart/form-data">
            <input type="hidden" name="txtTopic">

            <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="main.asp">Cancel</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.frmAddArticle.submit();">Create</a></div>
            <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
              <tr>
                <th align="left" colspan="2">New Folder</th>
              </tr>
              <tr>
                <td width="1%" valign="top">Destination:</td>
                <td>
                  <%
                  Dim strPath
                  strPath = Request("path") & "/"
                  If strPath = "/" Then %>
                    <select name="selTopic" style="width:100%" onchange="document.all.txtTopic.value=this.value;">
                      <% DrawListBoxContents Application("ECapture_ArticlesPath") %>
                    </select>
                  <%
                  Else
                    Response.Write "<script lanuage=""Javascript"">document.all.txtTopic.value='"& strPath &"';</script>" & Replace(strPath, Application("eCapture_ArticlesPath") & "/", "")
                  End If
                  %>
                </td>
              </tr>
              <tr>
                <td valign="top">Method:</td>
                <td>
                  <select name="txtMethod" onchange="showMethod(this[this.selectedIndex].value);">
                    <option value="upload"><%=langDocUpload%></option>
                    <option value="direct"><%=langDocDirect%></option>
                    <option value="link"><%=langDocLink%></option>
                  </select>
                </td>
              </tr>
              <tr>
                <td valign="top" nowrap><%=langDefineDoc%>:&nbsp;</td>
                <td>
                  <div id="upload">
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
                          <textarea name="txtContent" rows=15 cols=75></textarea>
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
                <td nowrap valign="top"><%=langDocOverwrite%>:</td>
                <td><input type="checkbox" class="listCheck" name="chkOverwrite"></td>
              </tr>
            </table>
            <div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="main.asp">Cancel</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.frmAddArticle.submit();">Create</a></div>
          </form>
        </td>
      </tr>
    </table>
<% End Select %>
</body>
</html>

<%
Set objDir = Nothing
Set objFSO = Nothing
%>