<!-- #include file="../includes/common.asp" //-->

<%
Dim oCmd, oRst, sName, intID, sDesc, sUrl, sUrlType, sAction, sFavIcon, sLinks, bShown

sAction = Request.QueryString("action")
If sAction = "U" Then
  sFavIcon = "newpersonalfav.gif"
Else
  sFavIcon = "newfav.gif"
End If

If Request.Form("_task") <> "" Then

  sUrlType = Request.Form("UrlType")
  sUrl = Request.Form("URL")
  If Left(sUrl, Len(sUrlType)) <> sUrlType Then
    sUrl = sUrlType & sUrl
  End If

  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "UpdateFavorite"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("FavID", adInteger, adParamInput, 4, Request.Form("id"))
    .Parameters.Append oCmd.CreateParameter("Name", adVarChar, adParamInput, 30, Request.Form("Name"))
    .Parameters.Append oCmd.CreateParameter("Url", adVarChar, adParamInput, 250, sUrl)
    .Parameters.Append oCmd.CreateParameter("Description", adVarChar, adParamInput, 250, Request.Form("Description"))
    .Execute
  End With
  Set oCmd = Nothing

  Response.Redirect "../favorites/default.asp?action=" & sAction

Else

  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "GetFavorite"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("FavID", adInteger, adParamInput, 4, Request.QueryString("id"))
  End With
  
  Set oRst = Server.CreateObject("ADODB.Recordset")
  With oRst
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open oCmd
  End With
  Set oCmd = Nothing  
  
  If not oRst.EOF then
    intID = clng(oRst("FavoriteID"))
    sName = oRst("FavoriteName")
    sURL = oRst("FavoriteURL")
    sDesc = oRst("FavoriteDescription")

    If Left(sUrl,7) = "http://" Then
      sUrlType = "http://"
    ElseIf Left(sUrl,8) = "https://" Then
      sUrlType = "https://"
    ElseIf Left(sUrl,7) = "mailto:" Then
      sUrlType = "mailto:"
    ElseIf Left(sUrl,6) = "ftp://" Then
      sUrlType = "ftp://"
    ElseIf Left(sUrl,2) = "\\" Then
      sUrlType = "\\"
    End If
    sUrl = Mid(sUrl, Len(sUrlType)+1)

  End If
  
  If oRst.State=1 then oRst.Close
  Set oRst = Nothing  

End If
%>

<html>
<head>
  <title><%=langBSFavorites%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabHome,1%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_home.jpg"></td>
      <td><font size="+1"><b><%=langFavorites%>: <%=langUpdate%></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langBackToFavoriteList%></a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top" nowrap>

        <!-- START: QUICK LINKS MODULE //-->
        
        <%
        sLinks = "<div style=""padding-bottom:8px;""><b>" & langFavoriteLinks & "</b></div>"

        If HasPermission("CanEditFavorite") Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/" & sFavIcon & """ align=""absmiddle"">&nbsp;<a href=""newfavorite.asp?action=" & sAction & """>" & langNewFavorite & "</a></div>"
          bShown = True
        End If
        
        If HasPermission("CanEditFavorite") Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/" & sFavIcon & """ align=""absmiddle"">&nbsp;<a href=""default.asp?action=" & sAction & """>" & langEditFavorites & "</a></div>"
          bShown = True
        End If
        
        If bShown Then
          Response.Write sLinks & "<br>"
        End If
        %>

        <% Call DrawQuicklinks("", 1) %>
        <!-- END: QUICK LINKS MODULE //-->

      </td>
        <!-- START: UPDATE FAVORITE -->
      <td colspan="2" valign="top">
        <form name="UpdateFavorite" method=post action="updatefavorite.asp?action=<%=sAction%>" method="post">
          <input type="hidden" name="_task" value="update">
          <input type="hidden" name="id" value=<%=intID%>>

          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.UpdateFavorite.submit();"><%=langUpdate%></a></div>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
            <tr>
              <th align="left" colspan="2"><%=langUpdateFavorite%></th>
            </tr>
            <tr>
              <td width="1" valign="top"><%=langName%>:</td>
              <td><input type="text" name="Name" style="width:400px;" maxlength="30" value="<%=sName%>"></td>
            </tr>
            <tr>
              <td valign="top" nowrap><%=langType%>:</td>
              <td> 
                <select name="UrlType" style="font-family:Verdana,Tahoma,Arial; font-size:11px; width:100px;">
                  <option value="http://"<% If sUrlType = "http://" Then Response.Write " selected" %>><%=langTypeWeb%></option>
                  <option value="https://"<% If sUrlType = "https://" Then Response.Write " selected" %>><%=langTypeSecure%></option>
                  <option value="mailto:"<% If sUrlType = "mailto:" Then Response.Write " selected" %>><%=langTypeEmail%></option>
                  <option value="\\"<% If sUrlType = "\\" Then Response.Write " selected" %>><%=langTypeNetwork%></option>
                  <option value="ftp://"<% If sUrlType = "ftp://" Then Response.Write " selected" %>><%=langTypeFTP%></option>
                </select>
              </td>
            </tr>
            <tr>
              <td valign="top"><div id="lblUrl"><%=langURL%>:</div></td>
              <td width="100%"> 
                <input type="text" name="URL" style="width:400px;" maxlength="250" value="<%= sUrl %>">
              </td>
            </tr>
            <tr>
              <td width="1" valign="top"><%=langDescription%>:&nbsp;</td>
              <td><textarea name="Description" rows="3" style="width:400px;"><%=sDesc%></textarea></td>
            </tr>
          </table>
          <div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.UpdateFavorite.submit();"><%=langUpdate%></a></div>
        </form>
      </td>
        <!-- END: UPDATE FAVORITE -->
    </tr>
  </table>
</body>
</html>
