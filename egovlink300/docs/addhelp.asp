<!-- #include file="../includes/common.asp" //-->
<!-- #include file="functions.inc" //-->
<%
Dim objFSO, objDir
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

function CreatePath(path)
  Dim ecapturepath, paths, i, temp
  ecapturepath = application("ecapture_articlespath") & "/"

  path = replace(path, ecapturepath, "z.Help/")
  if right(path,1) = "/" then path = left(path,len(path)-1)
  paths = split(path,"/")
  
  for i = 0 to ubound(paths)
    if temp = "" then temp = paths(i) else temp = temp & "/" & paths(i)

    if not objFSO.folderexists(server.mappath(ecapturepath & temp)) then
      objFSO.CreateFolder(server.mappath(ecapturepath & temp))
    end if
  next

  CreatePath = ecapturepath & temp & "/"
end function
%>

<html>
<head>
  <link href="global.css" rel="stylesheet">
   <%
  Sub DrawListBoxContents(dirPath)
    Dim newpath, name, sSQL, oRst, i, iLevel
	
    sSql = "EXEC ListFolder " & Session("OrgID") & ", " & Session("UserID") & ", '" & dirPath & "'"

    Set oRst = Server.CreateObject("ADODB.Recordset")
		oRst.Open sSql, Application("DSN"), 3, 1

		If Not oRst.EOF Then
			Do While Not oRst.EOF
				newpath = Server.URLEncode(oRst("FolderPath"))
				name = oRst("FolderName")
        padding = ""
        iLevel = oRst("FolderLevel")*2
				
				If name="root" Then
					Response.Write "<option value=""" & dirPath & "/" & name & "/"" selected>Main Category"
				Else
          For i = 0 To iLevel
            padding = padding & "&nbsp;"
          Next
					Response.Write "<option value=""" & dirPath & "/" & name & "/"">" & padding & name
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
        upload.style.display = "none";
        direct.style.display = "";  
        frmAddArticle.encoding = "application/x-www-form-urlencoded";
        frmAddArticle.action = "addhelp.asp?task=ADD&method=direct"
      }
      else {
        direct.style.display = "none";
        upload.style.display = "";
        frmAddArticle.encoding = "multipart/form-data";
        frmAddArticle.action = "addhelp.asp?task=ADD&method=upload"
      }
    }
  //-->
  </script>
</head>

<body bgcolor="#ffffff">
<img src="../docs/menu/images/helpdocument.gif">&nbsp;<font size="3"><b><%=langAddHelpDoc%></b></font>
<hr style="height:1px;"><br>

<%
Select Case Request("task")

  Case "ADD"
    If Request("method") = "upload" Then
      RequestBin = Request.BinaryRead(Request.TotalBytes)
      BuildUploadRequest RequestBin

      SetUploadPath CreatePath(GetValue("txtTopic"))
      sFileName = CreateFile("binFile", True)

    Else
      If Request.Form("txtTitle") <> "" And Request.Form("txtContent") <> "" Then
        If Request.Form("blnIsHTML") = "on" Then
          strTitle = Request.Form("txtTitle") & ".htm"
          strContent = Request.Form("txtContent")
        Else
          strTitle = Request.Form("txtTitle") & ".txt"
          strContent = CRsafe(Request.Form("txtContent"))
        End If

        newpath = CreatePath(Request.Form("txtTopic"))
        Set objNewFile = objFSO.CreateTextFile(Server.MapPath(newpath & strTitle))
        objNewFile.Write( Request.Form("txtContent") )
        objNewFile.Close
      End If
      '---BEGIN: Update DB fields--------------------------------
        'Set cnnECDB = Server.CreateObject("ADODB.Connection")
        'cnnECDB.Open Application("ECapture_ConnectionString")

        'strPath = Replace(Request.Form("txtTopic"), "/eclipse/tools/ecapture/Categories/", "z.Help/")

        'sSQL = "INSERT INTO ActionLog (ActionType,Article,Category,ActionBy,ActionByEmail,IPAddress) VALUES ('HELP - ADD'" _
        '                              & ",'" & strTitle & "','" & strPath & "','" & Session("UserName") &  "','" & Session("EmailName") _
        '                              & "','" & Request.ServerVariables("REMOTE_ADDR") & "');"
        'cnnECDB.Execute sSQL
        'cnnECDB.Close
        'Set cnnECDB = Nothing
      '---END: Update DB fields----------------------------------
    End If

    Response.Write "<b>Help document was added successfully.</b>"
    Response.Write "<script language=""Javascript"">parent.fraToc.RefreshPath();</script>"

  Case Else %>
    <form name="frmAddArticle" action="addhelp.asp?task=ADD&method=upload" method="POST" enctype="multipart/form-data">
      <table border=0 cellpadding=2 cellspacing=0 width="90%">
        <tr>
          <td colspan="3"><b>Step 1: </b><%=langChooseTopicHelp%><br><br></td>
        </tr>
        <tr>
          <td width="10"><img src="image/spacer.gif" width="10" height="1" border="0"></td>
          <td colspan="2">
            <font color="#666666"><%=langHelpTopic%></font><br>
            <input type="hidden" name="txtTopic">
            <%
            Dim strPath
            strPath = Request("path") & "/"
            If strPath = "/" Then %>
              <select name="selTopic" style="width:100%" onchange="document.all.txtTopic.value=this.value;">
                <option selected>Choose one...</option>
                <% 'DrawListBoxContents Application("ECapture_ArticlesPath"), "" %>
                <% DrawListBoxContents Application("ECapture_ArticlesPath") %>
              </select>
            <%
            Else
              Response.Write "<script lanuage=""Javascript"">document.all.txtTopic.value='"& strPath &"';</script>" & Replace(strPath, Application("eCapture_ArticlesPath") & "/", "")
            End If %>
          </td>
        </tr>
        <tr>
          <td colspan="3"><br><br><b>Step 2: </b><%=langChooseMethod%><br><br></td>
        </tr>
        <tr>
          <td></td>
          <td colspan="2">
            <font color="#666666"><%=langMethodOption%></font><br>
            <select name="txtMethod" onchange="showMethod(this[this.selectedIndex].value);">
              <option value="upload"><%=langMethodUpload%></option>
              <option value="direct"><%=langMethodDirect%></option>
            </select>
          </td>
        </tr>
        <tr>
          <td colspan="3"><br><br><b>Step 3: </b><%=langDocuDefine%><br><br></td>
        </tr>
        <tr>
          <td></td>
          <td colspan="2">
            <div id="upload">
              <font color="#666666"><%=langSelectFile%></font><br>
              <input type="file" name="binFile" size="50">
            </div>
            <div id="direct" style="display:none;">
              <table>
                <tr>
                  <td><font color="#666666">Title: </font></td>
                  <td><input type="text" name="txtTitle" size=50></td>
                </tr>
                <tr>
                  <td><font color="#666666">Content: </font>&nbsp;&nbsp;</td>
                  <td><input type=checkbox name="blnIsHTML">Is Content HTML?
                <tr>
                  <td colspan=2>
                    <textarea name="txtContent" rows=25 cols=75></textarea>
                  </td>
                </tr>
              </table>
            </div>
          </td>
        </tr>
        <tr>
          <td colspan="3"><br><br><b>Step 4: </b><%=langAddDocument%><br><br></td>
        </tr>
        <tr>
          <td></td>
          <td colspan="2"><input type=submit value="<%=langAddHelpDocument%>"></td>
        </tr>
      </table>
    </form>
<% End Select %>
</body>
</html>

<%
Set objFSO = Nothing
Set objDir = Nothing
%>
