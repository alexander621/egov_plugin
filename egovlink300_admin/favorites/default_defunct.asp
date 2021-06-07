<!-- #include file="../includes/common.asp" //-->

<%
Dim sSql, oRst, sFavorites, index, arrColors(2), iTotal, iCount, sDesc, iUserID, sAction, sFavIcon, sLinks, bShown
Const cCompany = "C"
Const cUser = "U"

sAction = Request.QueryString("action")
If sAction = cUser Then
  iUserID = clng(Session("UserID"))
  sFavIcon = "newpersonalfav.gif"
Else
	iUserID = "NULL"
  sAction = cCompany
  sFavIcon = "newfav.gif"
End If

sSql = "EXEC ListFavorites " & Session("OrgID") & ", " & iUserID

Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .ActiveConnection = Application("DSN")
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open sSql
  .ActiveConnection = Nothing
End With

arrColors(0)="ffffff"
arrColors(1)="eeeeee"
index=0

iCount = 1
iTotal = oRst.RecordCount

Do while not oRst.EOF

  sDesc = oRst("FavoriteDescription")
  If Len(sDesc) > 60 Then sDesc = Left(sDesc,58) & "..."

  sFavorites = sFavorites & "<tr bgcolor='" & arrColors(index) & "'>"
  If (sAction = cUser) OR HasPermission("CanEditFavorites") then	
	sFavorites = sFavorites & "<td><input type ='checkbox' class='listcheck' name='del_" & oRst("FavoriteID") & "'></td>"
	sFavorites = sFavorites & "<td style=""padding:0px;""><img src=""../images/" & sFavIcon & """ border=""0"">&nbsp;</td>"
	sFavorites = sFavorites & "<td valign='top' width=120><a href='updatefavorite.asp?action=" & sAction & "&" & "id=" & oRst("FavoriteID") & "'>" & oRst("FavoriteName") & "</a></td>"
  Else
	sFavorites = sFavorites & "<td valign='top'>&nbsp;</td>"
	sFavorites = sFavorites & "<td style=""padding:0px;""><img src=""../images/" & sFavIcon & """ border=""0"">&nbsp;</td>"
	sFavorites = sFavorites & "<td valign='top' width=120 nowrap>" & oRst("FavoriteName") & "</td>"
  End If
  sFavorites = sFavorites & "<td valign='top' width='60%'>" & sDesc & "&nbsp;</td>"
  sFavorites = sFavorites & "<td valign='top'><div class='link'><a href='" & oRst("FavoriteURL") & "'>" & oRst("FavoriteURL") & "</a></div></td>"
  sFavorites = sFavorites & "</tr>"
  
  index = 1 - index 'flip the index
  iCount = iCount + 1

  oRst.MoveNext
Loop

If sFavorites = "" Then
  If sAction = cCompany Then
    sFavorites = "<tr bgcolor='" & arrColors(index) & "'><td colspan=""5"">No company favorites have been created.</td></tr>"
  Else
    sFavorites = "<tr bgcolor='" & arrColors(index) & "'><td colspan=""5"">No personal favorites have been created.</td></tr>"
  End If
End If

If oRst.State=1 then oRst.Close
Set oRst = Nothing
%>

<html>
<head>
  <title><%=langBSFavorites%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabHome,1%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_home.jpg"></td>
      <%If sAction = cCompany then %><td><font size="+1"><b><%=langFavorites%></b>
	    <%Elseif sAction = cUser then %><td><font size="+1"><b><%=langPersonalFavorites%></b>
	    <%End If %></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="../"><%=langBackToStart%></a></td>
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
        
        If bShown Then
          Response.Write sLinks & "<br>"
        End If
        %>

        <% Call DrawQuicklinks("", 1) %>
        <!-- END: QUICK LINKS MODULE //-->

      </td>
        <!-- START: EDIT FAVORITE -->
      <td colspan="2" valign="top">
		    <form name="DelFavorite" action="deletefavorites.asp" method="post">
		      <input type=hidden name=action value=<%=sAction%>>
          <div style="font-size:10px; padding-bottom:5px;">
		      <%If (sAction = cUser) OR HasPermission("CanEditFavorites") then %>
				    <img src="../images/<%= sFavIcon %>" align="absmiddle">&nbsp;&nbsp;<a href="newfavorite.asp?action=<%= sAction %>"><%=langNewFavorite%></a>
				    &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/small_delete.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.DelFavorite.submit();"><%=langDelete%></a>
		      <%End If %></div>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
            <tr>
              <th align=left>
              <%If (sAction = cUser) OR HasPermission("CanEditFavorites") then%>
              <input class="listCheck" type=checkbox name="chkSelectAll" onClick="selectAll('DelFavorite', this.checked)">
              <%Else%>
              &nbsp;
              <%End If%>
              </th>
              <th>&nbsp;</th>
              <th align="left"><%=langName%></th>
              <th align="left"><%=langDescription%></th>
              <th align="left"><%=langURL%></th>
            </tr>
            <%= sFavorites %>
            
          </table>
        </form>
      </td>
        <!-- END: EDIT FAVORITE -->
    </tr>
  </table>
</body>
</html>
