<!-- #include file="../includes/common.asp" -->
<%
Dim objFSO, strPath, fullpath
Dim objDir, padding, strDir,oCmd
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
strPath = Request.Form("txtDir")

Dim sRootPath
sRootPath = Application("ECapture_ArticlesPath") & "/"
%>

<html>
<head>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <%
  Sub DrawListBoxContents(dirPath)
    Dim newpath, name, sSQL, oRst, i, iLevel
	
    sSql = "EXEC ListFolder " & Session("OrgID") & ", " & Session("UserID") & ", '" & dirPath & "'"
    Set oRst = Server.CreateObject("ADODB.Recordset")
		oRst.Open sSql, Application("DSN"), 3, 1

		If Not oRst.EOF Then
			Do While Not oRst.EOF
				curpath = oRst("FolderPath")
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
					Response.Write "<option value=""" & curpath & "/"">" & padding & name
				End If
        oRst.MoveNext
			Loop
			oRst.Close
			Set oRst = Nothing
		End If
  End Sub
  %>
  <script language="Javascript" src="scripts/vbstring.js"></script>
  <script language="Javascript1.2">
  <!--
    function DrawTree(dirPath) {
      var output = "";
      var padding = "";

      path = replace(dirPath, "<%= Application("ECapture_ArticlesPath") & "/" %>", "");
      while ((loc = path.search("/")) != -1) {
        node = path.slice(0,loc);
        path = path.replace(node + "/", "");
        output = output + "<table border=0 cellpadding=0 cellspacing=0><tr><td>" + padding;
        output = output + "<img src='menu/images/folder_open.gif' border=0>&nbsp;</td><td valign=middle nowrap style='color:#666666;'>" + node + "</td></tr></table>";
        padding = padding + "&nbsp;&nbsp;&nbsp;";
      }
      output = output + "<table border=0 cellpadding=0 cellspacing=0><tr><td>" + padding;
      output = output + "<img src='menu/images/folder_closed.gif' width=18 height=18 border=0>&nbsp;</td><td valign=middle nowrap><input type=text name='txtDirName' size=20></td></tr></table>";

      document.all.step2.innerHTML = output;
    }

    function doPicker() {
      window.open('picker/default.asp?type=folder', 'filepicker', 'width=506,height=345,scrollbars=0,toolbars=0,statusbar=0,menubar=0,left=265,top=180');
    }
  //-->
  </script>
</head>

<body>
<%
Select Case Request("task")

  Case "ADD"
    'response.write request("txtDirName") & "<br>"
    'txtDirName = StripCharacters(request("txtDirName"))
    'response.write "BACK <br>"
    'response.write txtDirName
    'response.end
    'strDir = Request.Form("txtDir") & txtDirName
    strDir = Request.Form("txtDir") & request("txtDirName")
    if instr(request("txtdirname"), "'") then
      Response.Write "<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='main.asp' style='font-family: Verdana, Arial, Helvetica; font-weight: bold; font-size: 8pt;' >Back To Documents Main</a><br><br>"
      Session("ECapture_Error") = "The category you attempted to add contains special character(s).<br>Please try again."
      Response.Redirect "addfolder.asp?task=browse"

    elseIf Not objFSO.FolderExists(Server.MapPath(strDir)) Then
      objFSO.CreateFolder(Server.MapPath(strDir))

      '---BEGIN: Update DB fields--------------------------------
      Set oCmd = Server.CreateObject("ADODB.Command")
      With oCmd
      .ActiveConnection = Application("DSN")
      .CommandText = "NewFolder"
      .CommandType = adCmdStoredProc
      .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
      .Parameters.Append oCmd.CreateParameter("CreatorID", adInteger, adParamInput, 4, Session("UserID"))
      .Parameters.Append oCmd.CreateParameter("FolderPath", adVarChar, adParamInput, 300, strDir)
      .Execute
      End With
      Set oCmd = Nothing
      '---END: Update DB fields----------------------------------

       Response.Write "<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='main.asp' style='font-family: Verdana, Arial, Helvetica; font-weight: bold; font-size: 8pt;' >Back To Documents Main</a><br><br>"
      Response.Write "<b>Folder was added successfully.</b>"
      Response.Write "<script language=""Javascript"">parent.fraToc.RefreshPath();</script>"
    Else
	   Response.Write "<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='main.asp' style='font-family: Verdana, Arial, Helvetica; font-weight: bold; font-size: 8pt;' >Back To Documents Main</a><br><br>"
      Session("ECapture_Error") = "The category you attempted to add already exists.<br>Please try again."
      Response.Redirect "addfolder.asp?task=browse"
    End If

  Case Else %>

    <table border="0" cellspacing="0" class="start" width="100%">
      <tr>
        <td style="padding:10px 0px;"><font size="+1"><b><%=langDocuments%>: <%=langAddFolder%></b></font>
	<br>
    	<%If Session("ECapture_Error") <> "" Then
      	Response.Write "<font color=red>" & Session("ECapture_Error") & "</font>"
      	Session("ECapture_Error") = ""
    	End If %>
	
	<br></td>
      </tr>
      <tr>
        <td width="100%">
          <br>
          <form name="NewFolder" method=post action="addfolder.asp" method="post">
            <input type="hidden" name="task" value="ADD">
            <input type="hidden" name="txtDir">

            <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="main.asp">Cancel</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.NewFolder.submit();">Create</a></div>
            <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
              <tr>
                <th align="left" colspan="2">New Folder</th>
              </tr>
              <tr>
                <td width="80">Destination:</td>
                <td>
                  <div id="step1">
                  <select name="selDir" style="width:100%" onchange="document.all.txtDir.value=this.value;DrawTree(this[selectedIndex].value);">
                    <option value="<%= Application("ECapture_ArticlesPath") & "/" %>">Main Category</option>
                    <% DrawListBoxContents Application("ECapture_ArticlesPath") %>
                  </select>
                  </div>
                </td>
              </tr>
              <tr>
                <td width="1" valign="top">Name:&nbsp;</td>
                <td><div id="step2"></div></td>
              </tr>
            </table>
            <div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="main.asp">Cancel</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.NewFolder.submit();">Create</a></div>
          </form>
        </td>
      </tr>
    </table>
<% End Select %>

  <%
  strPath = Request("path") & ""

  If strPath <> "" Then
    Response.Write "<script lanuage=""Javascript"">document.all.txtDir.value='"& strPath &"/';document.all.step1.innerHTML='"& Replace(strPath, sRootPath, "") &"'; DrawTree('"& strPath & "/');</script>"
  Else
    Response.Write "<script lanuage=""Javascript"">document.all.txtDir.value='"& sRootPath &"'; DrawTree('"& sRootPath &"');</script>"
  End If
  %>

</body>
</html>


<%
Set objDir = Nothing
Set objFSO = Nothing
%>
<%
Function StripCharacters(strInput)
	'response.write "HERE" & "<br>"
	'response.write strInput & "<br>"
    Dim iPos, sNew, iTemp
    strInput = Trim(strInput)
    If strInput <> "" Then
        iPos = 1
        iTemp = Len(strInput)
        While iTemp >= iPos
	    'response.write Mid(strInput,iPos,1) & "<br>"
            If IsNumeric(Mid(strInput,iPos,1)) = true then
                sNew = sNew & Mid(strInput,iPos,1)
	    	'response.write sNew & "<br>"
            End If
	    if isAlpha(Mid(strInput,iPos,1)) = true Then
                sNew = sNew & Mid(strInput,iPos,1)
	    	'response.write sNew & "<br>"
            End If
            iPos = iPos + 1
        Wend
    Else
        sNew = ""
    End If
    StripCharacters = sNew
End Function

Function isAlpha(str) 
	bolValid = True 
	if Asc(UCase(str)) < Asc("A") or Asc(UCase(str)) > Asc("Z") then 
		bolValid = False 
	end if
	isAlpha = bolValid 
End Function
%>

