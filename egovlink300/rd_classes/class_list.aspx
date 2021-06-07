<%@ Page Language="C#" AutoEventWireup="true" CodeFile="class_list_test.aspx.cs" Inherits="rd_classes_class_list_test" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<%@ Register TagPrefix="Tbanner" TagName="banner" Src="../rd_includes/egov_banner.ascx" %>
<%@ Register TagPrefix="Tnavigation" TagName="navigation" Src="../rd_includes/egov_navigation.ascx" %>
<%@ Register TagPrefix="Tfooter" TagName="footer" Src="../rd_includes/egov_footer.ascx" %>

<%@ Register TagPrefix="classes_memberWarning" TagName="classesMemberWarning" Src="../rd_includes/egov_classes_memberwarning.ascx" %>

<!DOCTYPE html>
<script runat="server">
    static string sOrgID             = common.getOrgId();
    static string sOrgName           = common.getOrgName(sOrgID);
    string sOrgVirtualSiteName       = common.getOrgInfo(sOrgID, "orgVirtualSiteName");
    string sPageTitle                = "E-Gov Services " + sOrgName;
    
    static Int32 iRootCategoryID = classes.getFirstCategory(sOrgID);
    Int32 sCategoryID            = iRootCategoryID;

    Boolean sViewPick      = false;
    Boolean sShowViewPicks = true;
    Boolean sDisplayImage  = true;
    Boolean sDisplayDesc   = true;
</script>
<%
    if (!String.IsNullOrEmpty(Request["categoryid"]))
    {
        try
        {
            sCategoryID = Convert.ToInt32(Request["categoryid"]);
        }
        catch
        {
            //Response.Redirect("class_categories.aspx");
        }
    }
    else
    {
        if (String.IsNullOrEmpty(Request["keywordSearch"]) && String.IsNullOrEmpty(Request["categoryid"]) && String.IsNullOrEmpty(Request["season"]) && String.IsNullOrEmpty(Request["sort"])) {
            //Response.Redirect("class_categories.aspx");
        }
    }

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
    
    if (sCategoryID == iRootCategoryID)
    {
        sShowViewPicks = false;
        sDisplayImage  = true;
        sDisplayDesc   = false;
    }
    
    //Set up variables for common user controls
    egov_navigation.egovsection   = "CLASSES";
    egov_navigation.rootcategory  = Convert.ToString(iRootCategoryID);
    egov_navigation.categoryid    = Convert.ToString(sCategoryID);
    
    //Set up variables for feature specific user controls
    egov_classes_memberwarning.orgid = sOrgID;

    //Setup User and Session Variables
    //sCookieUserID = Request.Cookies["useridx"];
    //sSessionID    = HttpContext.Current.Session.SessionID;

    Session["RedirectPage"]    = "rd_classes/class_list.aspx?categoryid=" + sCategoryID.ToString();
    Session["RedirectLang"]    = "Return to Class List";
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
  <!-- <script type="text/javascript" src="../rd_scripts/egov_navigation_classcategories.js"></script> -->

<script type="text/javascript">
    $(document).ready(function() {
        $('#dropdown_viewpick').change(function() {
            //document.getElementById('frmSearch').action = 'class_list.aspx';
            //$('#frmSearch').submit();
            $('#reorderList').submit();
        });
    });

    function viewCategoryList(iCategoryID) {
        location.href = 'class_categories.aspx';
    }

    function viewCategoryClassList(iCategoryID) {
        location.href = 'class_list.aspx?categoryid=' + iCategoryID;
    }

    //function viewClassDetails(iClassID, iCategoryID, iCategoryTitle) {
    function viewClassDetails(iClassID, iCategoryID) {
        var lcl_url = '';

        lcl_url = 'class_details.aspx';
        lcl_url += '?classid='    + iClassID;
        lcl_url += '&categoryid=' + iCategoryID;
        //lcl_url += '&categorytitle=' + iCategoryTitle;

        location.href = lcl_url;
    }
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
      <classes_memberWarning:classesMemberWarning id="egov_classes_memberwarning" runat="server" orgid="" />

      <div style="margin-left:41px; border: 1px solid black;padding: 10px; display:inline-block; margin-bottom:10px;">
      		<h4 style="margin:5px;">Advanced Search:</h4>
		<form action="#" method="get">
		<div style="float:left;margin:5px;">
      			Keyword <input type="text" name="keywordSearch" value="<%=Request["keywordSearch"]%>" />
		</div>
		<div style="float:left;margin:5px;">
			Category
			<select name="categoryid">
				<option value="">All</option>
				<% displayCategories(Convert.ToInt32(sOrgID), sCategoryID); %>
			</select>
		</div>
		<div style="float:left;margin:5px;">
			Season
			<select name="season">
				<option value="">All</option>
				<% displaySeasons(Convert.ToInt32(sOrgID), Request["season"]); %>
			</select>
		</div>
		<div style="float:left;margin:5px;">
			Sort
			<select name="sort">
				<option value="">Class Name</option>
				<option value="date" <% if (Request["sort"] == "date") { Response.Write(" selected");}%>>Class Start Date</option>
			</select>
		</div>
		<div style="float:left;margin:5px;">
			<input type="submit" value="Search" />
		</div>
		</form>

      </div>
<%
    if (!String.IsNullOrEmpty(Request["categoryid"]))
    {
        
        displayCategoryInfo(Convert.ToInt32(sOrgID), sCategoryID);
    }

    displayClassList(Convert.ToInt32(sOrgID), sCategoryID, sShowViewPicks, sViewPick);
%>
    </div>
  </div>
  <div id="wrapper_footer">
    <Tfooter:footer ID="footer" runat="server" />
  </div>
</div>
</body>
</html>
