<!-- #include file="../includes/common.asp" //-->
<%
dim index, arrColors(2), hasShownLink
hasShownLink=false
arrColors(0)="#ffffff"
arrColors(1)="#eeeeee"
index=0
%>

<html>
<head>
  <title><%=langBSAdmin%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <SCRIPT LANGUAGE="JavaScript">
		<!--// This function will open page in new window
		function NewWindow(page) {
			OpenWin = window.open(page, "CtrlWindow", "width=420,height=120,status=no,toolbar=no,menubar=no,top=200,left=300,z-lock=no");
	   	}
		// End -->
    function doPageSize(size) {
      frmPageSize.action = "ChangePageSize.asp?size=" + size;
      frmPageSize.submit();
    }
  </SCRIPT>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabAdmin,1%>

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_admin.jpg"></td>
      <td><font size="+1"><b><%=langAdminLinks%>: <%=langPersonalSettings%></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="../admin"><%=langBackTo%>&nbsp;<%=langAdminLinks%></a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
        <% Call DrawQuicklinks("", 1) %>
      </td>
      <td colspan="2" valign="top">
        <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
          <tr style="height:22px;">
            <th align="left" colspan="2"><%=langPersonalSettings%></th>
          </tr>
          <tr>
            <td width="100">Page Size:</td>
            <td>
              <select onchange="doPageSize(this[this.selectedIndex].value);" style="width:100px;">
                <option value="10" <% If Session("PageSize") = 10 Then Response.Write "selected" %>>10</option>
                <option value="20" <% If Session("PageSize") = 20 Then Response.Write "selected" %>>20</option>
                <option value="30" <% If Session("PageSize") = 30 Then Response.Write "selected" %>>30</option>
                <option value="50" <% If Session("PageSize") = 50 Then Response.Write "selected" %>>50</option>
              </select>
              &nbsp;&nbsp;&nbsp;&nbsp;Determines how many records you will see on each page.
            </td>
          </tr>
          <tr bgcolor="#eeeeee">
            <td width="100">Stock Ticker:</td>
            <td>
              <select onchange="frmTicker.submit();" style="width:100px;">
                <option <% If Session("ShowStockTicker") = 1 Then Response.Write "selected" %>>On</option>
                <option <% If Session("ShowStockTicker") = 0 Then Response.Write "selected" %>>Off</option>
              </select>
              &nbsp;&nbsp;&nbsp;&nbsp;Customizable Stock Ticker on Home page&nbsp;&nbsp;<i>(Requires Internet Explorer).</i>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>

  <form name="frmTicker" action="ToggleStockTicker.asp?redirect=ChangePersonalSettings.asp" method="post"></form>
  <form name="frmPageSize" action="ChangePageSize.asp?size=" method="post"></form>
</body>
</html>