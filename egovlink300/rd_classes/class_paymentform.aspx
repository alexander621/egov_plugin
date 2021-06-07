<%@ Page Language="C#" AutoEventWireup="true" CodeFile="class_paymentform.aspx.cs" Inherits="rd_classes_class_paymentform" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<%@ Register TagPrefix="Tbanner" TagName="banner" Src="../rd_includes/egov_banner.ascx" %>
<%@ Register TagPrefix="Tnavigation" TagName="navigation" Src="../rd_includes/egov_navigation.ascx" %>
<%@ Register TagPrefix="Tfooter" TagName="footer" Src="../rd_includes/egov_footer.ascx" %>

<!DOCTYPE html>
<script runat="server">
    //HttpCookie sCookieUserID = Request.Cookies["useridx"];

    static string sOrgID             = common.getOrgId();
    static string sOrgName           = common.getOrgName(sOrgID);
    //static string sSessionID         = "";
    //static string sSessionIDName     = "";  //This is used to identify the column to save the session value to on "egov_aspnet_to_asp_usersessions"
    
    string sOrgVirtualSiteName       = common.getOrgInfo(sOrgID, "orgVirtualSiteName");
    string sPageTitle                = "E-Gov Services " + sOrgName;
    string lcl_isLoggedIn            = "";
    string lcl_checked_isLoggedInYes = "";
    string lcl_checked_isLoggedInNo  = "";
    
    static Int32 iRootCategoryID = classes.getFirstCategory(sOrgID);
    Int32 sCategoryID            = iRootCategoryID;

    Boolean sViewPick      = false;
    Boolean sShowViewPicks = true;
</script>
<%
    if (sOrgID.ToString() == "7")
    {
        sPageTitle = sOrgName;
    }
    
    //Set up variables for common user controls
    egov_navigation.egovsection  = "HIDE_SUBMENU";
    egov_navigation.rootcategory = Convert.ToString(iRootCategoryID);
    egov_navigation.categoryid   = Convert.ToString(sCategoryID);

    //Setup User and Session Variables
    //sCookieUserID = Request.Cookies["useridx"];
    //sSessionID    = HttpContext.Current.Session.SessionID;

    Session["RedirectPage"]    = "rd_classes/class_categories.aspx";
    Session["RedirectLang"]    = "Return to Class Categories";
    Session["LoginDisplayMsg"] = "";
    Session["DisplayMsg"]      = "";
    Session["ManageURL"]       = "";
%>
<html lang="en">
<head id="Head1" runat="server">
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
<script src='https://www.google.com/recaptcha/api.js'></script>

  <title><%=sPageTitle%></title>

  <link type="text/css" rel="stylesheet" href="../rd_global.css" />
  <link type="text/css" rel="stylesheet" href="styles_class.css" />

  <%="<link type=\"text/css\" rel=\"stylesheet\" href=\"../css/style_" + sOrgID + ".css\" />"%>
    
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <%="<script type=\"text/javascript\" src=\"" + common.getBaseURL("") + "/" + sOrgVirtualSiteName + "/rd_scripts/jquery-1.7.2.min.js\"></script>"%>
  <script type="text/javascript" src="../rd_scripts/egov_navigation.js"></script>

<script type="text/javascript">
//    $(document).ready(function() {

//    });

    var eGovLink = eGovLink || {};

    eGovLink.Class = (function() {
        var processPayment = function() {
            $('#sjname').val($('#firstname').val() + ' ' + $('#lastname').val());

            $('#COMPLETE_PAYMENT').prop('disabled', true);

            $('#paymentForm').submit()
        }

        return {
            processPayment: processPayment
        };
    } ());

</script>

</head>
<body>
<div id="wrapper_body">
  <div id="wrapper_header">
    <Tbanner:banner ID="banner" runat="server" />
    <Tnavigation:navigation ID="egov_navigation" runat="server" egovsection="" rootcategory="" categoryid="" />
  </div>
  <div id="wrapper_content">
    <div id="content">
      <% displayPaymentForm(Convert.ToInt32(sOrgID)); %>
    </div>
  </div>
  <div id="wrapper_footer">
    <Tfooter:footer ID="footer" runat="server" />
  </div>
</div>
</body>
</html>
