<%@ Page Language="C#" AutoEventWireup="true" CodeFile="class_categories.aspx.cs" Inherits="rd_classes_class_categories" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<%@ Register TagPrefix="Tbanner" TagName="banner" Src="../rd_includes/egov_banner.ascx" %>
<%@ Register TagPrefix="Tnavigation" TagName="navigation" Src="../rd_includes/egov_navigation.ascx" %>
<%@ Register TagPrefix="Tfooter" TagName="footer" Src="../rd_includes/egov_footer.ascx" %>

<%@ Register TagPrefix="classes_memberWarning" TagName="classesMemberWarning" Src="../rd_includes/egov_classes_memberwarning.ascx" %>

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

    if (Request["viewpick"] != null)
    {
        try
        {
            if (Convert.ToInt32(Request["viewpick"]) == 1)
            {
                sViewPick = true;
            }
        }
        catch
        {
            sViewPick = false;
        }
    }
    
    sShowViewPicks = false;
    
    //Set up variables for common user controls
    egov_navigation.egovsection  = "CLASSES";
    egov_navigation.rootcategory = Convert.ToString(iRootCategoryID);
    egov_navigation.categoryid   = Convert.ToString(sCategoryID);

    //Set up variables for feature specific user controls
    egov_classes_memberwarning.orgid = sOrgID;

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
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

  <title><%=sPageTitle%></title>

  <link type="text/css" rel="stylesheet" href="../rd_global.css" />
  <link type="text/css" rel="stylesheet" href="styles_class.css" />

  <%="<link type=\"text/css\" rel=\"stylesheet\" href=\"../css/style_" + sOrgID + ".css\" />"%>
    
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <%="<script type=\"text/javascript\" src=\"" + common.getBaseURL("") + "/" + sOrgVirtualSiteName + "/rd_scripts/jquery-1.7.2.min.js\"></script>"%>
  <script type="text/javascript" src="../rd_scripts/egov_navigation.js"></script>


<script type="text/javascript">
    $(document).ready(function() {
        $('#dropdown_viewpick').change(function() {
            //document.getElementById('frmSearch').action = 'class_categories.aspx';
            //$('#frmSearch').submit();
            $('#reorderList').submit();
        });
    });

    function clickToRegister() {
        alert('continue registration process...');
    }

    function viewCategoryList(iCategoryID) {
        location.href = 'class_categories.aspx';
    }

    function viewCategoryClassList(iCategoryID) {
        location.href = 'class_list.aspx?categoryid=' + iCategoryID;
    }
</script>

</head>
<body>
<!--
<h1>Uploading Files</h1>
<form id="form1" enctype="multipart/form-data" runat="server">
    <p>
      File to Upload: <br />
      <asp:FileUpload type="file" id="FileUpload1" runat="server" />
    </p>
   
    <h2>File Information</h2>
    <p>
      <b>Status:</b>     <asp:label id="Label1" runat="server"></asp:label><br />
      <b>File Name:</b>  <asp:label id="fileName" runat="server"></asp:label><br />
      <b>File Size:</b>  <asp:label id="fileSize" runat="server"></asp:label><br />
      <b>File Type:</b>  <asp:label id="fileType" runat="server"></asp:label><br />
      <b>String Path</b> <asp:label id="fileStringPath" runat="server"></asp:label><br />
    </p>
</form>
-->
<div id="wrapper_body">
  <div id="wrapper_header">
    <Tbanner:banner ID="banner" runat="server" />
    <Tnavigation:navigation ID="egov_navigation" runat="server" egovsection="" rootcategory="" categoryid="" />
  </div>
  <div id="wrapper_content">
    <div id="content">
      <classes_memberWarning:classesMemberWarning id="egov_classes_memberwarning" runat="server" orgid="" />
	<input type="button"  style="margin-left:40px;" onclick="window.location='class_list.aspx'" value="Advanced Search" />
<%    
      listCategories(Convert.ToInt32(sOrgID),
                     sCategoryID);
%>
    </div>
  </div>
  <div id="wrapper_footer">
    <Tfooter:footer ID="footer" runat="server" />
  </div>
</div>
</body>
</html>
