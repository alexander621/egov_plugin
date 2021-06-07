<%@ Page Language="C#" AutoEventWireup="true" CodeFile="user_login.aspx.cs" Inherits="user_login" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<%@ Register TagPrefix="Tbanner" TagName="banner" Src="rd_includes/egov_banner.ascx" %>
<%@ Register TagPrefix="Tnavigation" TagName="navigation" Src="rd_includes/egov_navigation.ascx" %>
<%@ Register TagPrefix="Tfooter" TagName="footer" Src="rd_includes/egov_footer.ascx" %>

<!DOCTYPE html>
<script runat="server">
    static string sOrgID             = common.getOrgId();
    static string sOrgName           = common.getOrgName(sOrgID);
    string sOrgVirtualSiteName       = common.getOrgInfo(sOrgID, "orgVirtualSiteName");
    string sPageTitle                = "E-Gov Services " + sOrgName;
</script>
    
<html>
<head runat="server">
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

  <title><%=sPageTitle%></title>

  <link type="text/css" rel="stylesheet" href="rd_global.css" />
  <%="<link type=\"text/css\" rel=\"stylesheet\" href=\"css/style_" + sOrgID + ".css\" />"%>

  <%="<script type=\"text/javascript\" src=\"http://www.egovlink.com/" + sOrgVirtualSiteName + "/rd_scripts/jquery-1.7.2.min.js\"></script>"%>
  <script type="text/javascript" src="rd_scripts/egov_navigation.js"></script>

</head>
<body>
<div id="wrapper_body">
  <div id="wrapper_header">
    <Tbanner:banner ID="banner" runat="server" />
    <Tnavigation:navigation ID="egov_navigation" runat="server" egovsection="HIDE_SUBMENU" />
  </div>
  <div id="wrapper_content">
    <div id="content">
    login info here
    </div>
  </div>
  <div id="wrapper_footer">
    <Tfooter:footer ID="footer" runat="server" />
  </div>
</div>
</body>
</html>
