<!-- #include file="../includes/common.asp" //-->
<%
Dim oCmd, sUrl, sUrlType, sAction, iUserID, sFavIcon, sLinks, bShown, iUser

'sAction tells us if the request was for Company Favorites or User Favorites
sAction = Request("action")
If sAction & ""  = "" Then Response.Redirect "default.asp"	

If sAction = "U" then
	iUserID = clng(Session("UserID"))
  sFavIcon = "newpersonalfav.gif"
Else
  iUserID = NULL
  sFavIcon = "newfav.gif"
End If

If Request.Form("_task") = "newfav" Then
  sUrlType = Request.Form("UrlType")
  sUrl		 = Request.Form("URL")
  If Left(sUrl, Len(sUrlType)) <> sUrlType Then
    sUrl = sUrlType & sUrl
  End If

  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "NewFavorite"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
    .Parameters.Append oCmd.CreateParameter("UserID", adInteger, adParamInput, 4, iUserID)
    .Parameters.Append oCmd.CreateParameter("Name", adVarChar, adParamInput, 30, Request.Form("Name"))
    .Parameters.Append oCmd.CreateParameter("Url", adVarChar, adParamInput, 250, sUrl)
    .Parameters.Append oCmd.CreateParameter("Description", adVarChar, adParamInput, 250, Request.Form("Description"))
    .Execute
  End With
  Set oCmd = Nothing

  If sAction = "C" then
	  Response.Redirect "default.asp?action=C"
  Else
	  Response.Redirect "default.asp?action=U"
  End If
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
      <td><font size="+1"><b><%=langFavorites%>: <%=langNew%></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langBackToFavoriteList%></a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top" nowrap>

        <!-- START: QUICK LINKS MODULE //-->
        
        <%
        sLinks = "<div style=""padding-bottom:8px;""><b>" & langFavoriteLinks & "</b></div>"

        If HasPermission("CanEditFavorite") Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/" & sFavIcon & """ align=""absmiddle"">&nbsp;<a href=""../favorites/default.asp?action=" & sAction & """>" & langEditFavorites & "</a></div>"
          bShown = True
        End If
        
        If bShown Then
          Response.Write sLinks & "<br>"
        End If
        %>

        <% Call DrawQuicklinks("", 1) %>
        <!-- END: QUICK LINKS MODULE //-->

      </td>
        <!-- START: NEW FAVORITE -->
      <td colspan="2" valign="top">
        <form name="NewFavorite" method=post action="newfavorite.asp?action=<%=sAction%>">
          <input type=hidden name=action value=<%=sAction%>>
		      <input type="hidden" name="_task" value="newfav">
		      <input type=hidden name=User value=<%=iUser%>>
          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.NewFavorite.submit();"><%=langCreate%></a></div>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
            <tr>
              <th align="left" colspan="2"><%=langNewFavorite%></th>
            </tr>
            <tr>
              <td valign="top"><%=langName%>:</td>
              <td><input type="text" name="Name" style="width:200px;" maxlength="30"></td>
            </tr>
            <tr>
              <td valign="top" nowrap><%=langType%>:</td>
              <td> 
                <select name="UrlType" style="font-family:Verdana,Tahoma,Arial; font-size:11px; width:100px;">
                  <option value="http://"><%=langTypeWeb%></option>
                  <option value="https://"><%=langTypeSecure%></option>
                  <option value="mailto:"><%=langTypeEmail%></option>
                  <option value="\\"><%=langTypeNetwork%></option>
                  <option value="ftp://"><%=langTypeFTP%></option>
                </select>
              </td>
            </tr>
            <tr>
              <td valign="top"><div id="lblUrl"><%=langURL%>:</div></td>
              <td width="100%"> 
                <input type="text" name="URL" style="width:400px;" maxlength="250">
              </td>
            </tr>
            <tr>
              <td valign="top"><%=langDescription%>:&nbsp;</td>
              <td><textarea name="Description" rows="3" style="width:400px;"></textarea></td>
            </tr>
          </table>
          <div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.NewFavorite.submit();"><%=langCreate%></a></div>
        </form>
      </td>
        <!-- END: NEW FAVORITE -->
    </tr>
  </table>
</body>
</html>
