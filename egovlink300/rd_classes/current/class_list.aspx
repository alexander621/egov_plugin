<%@ Page Language="C#" AutoEventWireup="true" CodeFile="class_list.aspx.cs" Inherits="rd_classes_class_list" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<%@ Register TagPrefix="Tbanner" TagName="banner" Src="../rd_includes/egov_banner.ascx" %>
<%@ Register TagPrefix="Tnavigation" TagName="navigation" Src="../rd_includes/egov_navigation.ascx" %>
<!-- <%@ Register TagPrefix="Twelcomeinfo" TagName="welcomeinfo" Src="../rd_includes/egov_welcomeinfo.ascx" %> -->
<%@ Register TagPrefix="Tfooter" TagName="footer" Src="../rd_includes/egov_footer.ascx" %>

<!DOCTYPE HTML>
<script runat="server">
    static string sOrgID             = common.getOrgId();
    static string sOrgName           = common.getOrgName(sOrgID);
    string sOrgVirtualSiteName       = common.getOrgInfo(sOrgID, "orgVirtualSiteName");
    string sPageTitle                = "E-Gov Services " + sOrgName;
    string lcl_isLoggedIn            = "";
    string lcl_checked_isLoggedInYes = "";
    string lcl_checked_isLoggedInNo  = "";
    
    static Int32 iRootCategoryID = getFirstCategory(sOrgID);
    Int32 sCategoryID            = iRootCategoryID;

    Boolean sViewPick      = false;
    Boolean sShowViewPicks = true;
    Boolean sDisplayImage  = true;
    Boolean sDisplayDesc   = true;
</script>
<%
    if (Request["categoryid"] != null)
    {
        try
        {
            sCategoryID = Convert.ToInt32(Request["categoryid"]);
        }
        catch
        {
            Response.Redirect("class_list.aspx");
        }
    }

    if (sOrgID.ToString() == "7")
    {
        sPageTitle = sOrgName;
    }

    lcl_isLoggedIn = Request["loggedin"];

    if (lcl_isLoggedIn == "Y")
    {
       
        lcl_checked_isLoggedInYes = " checked=\"checked\"";
        Response.Cookies["userid"].Value = "1159";
    }
    else
    {
        lcl_checked_isLoggedInNo = " checked=\"checked\"";

        Response.Cookies["userid"].Value = "";
        Response.Cookies["userid"].Expires = DateTime.Now.AddDays(-1);
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
        sDisplayImage  = false;
        sDisplayDesc   = false;
    }

    egov_navigation.egovsection   = "CLASSES";
    egov_navigation.rootcategory  = Convert.ToString(iRootCategoryID);
    egov_navigation.showviewpicks = Convert.ToString(sShowViewPicks);
    egov_navigation.viewpick      = Convert.ToString(sViewPick);
    egov_navigation.categoryid    = Convert.ToString(sCategoryID);
    
%>
<html lang="en">
<head id="Head1" runat="server">
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

  <title><%=sPageTitle%></title>

  <link type="text/css" rel="stylesheet" href="../global.css" />
  
  <!--[if lte IE 8]>
<!--    <link type="text/css" rel="stylesheet" href="styles_class_ie_old.css" /> -->
<!--  <![endif]-->

  <!--[if !(lte IE 8)]>
<!--    <link type="text/css" rel="stylesheet" href="styles_class.css" /> -->
<!--  <![endif]-->

  <%="<link type=\"text/css\" rel=\"stylesheet\" href=\"../css/style_" + sOrgID + ".css\" />"%>
  
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <%="<script type=\"text/javascript\" src=\"http://www.egovlink.com/" + sOrgVirtualSiteName + "/rd_scripts/jquery-1.7.2.min.js\"></script>"%>
  <script type="text/javascript" src="../rd_scripts/egov_navigation.js"></script>
<!--  <script type="text/javascript" src="../rd_scripts/egov_navigation_classcategories.js"></script> -->

<script type="text/javascript">
    $(document).ready(function() {
        $('#isLoggedIn_yes').click(function() {
            location.href = 'class_list.aspx?loggedin=Y';
        });

        $('#isLoggedIn_no').click(function() {
            location.href = 'class_list.aspx?loggedin=N';
        });

        $('#searchButton').click(function() {
            var lcl_searchphrase = '';

            if ($('#txtsearchphrase')) {
                lcl_searchphrase = $('#txtsearchphrase').val();
            }

            if (lcl_searchphrase == '') {
                $('#txtsearchphrase').focus();
                inlineMsg(document.getElementById('searchButton').id, '<strong>Required Field Missing: </strong> Please enter a value to search.', 10, 'searchButton');
                return false;
            } else {
                clearMsg('searchButton');
                $('#frmSearch').submit();
            }
        });

        $('#txtsearchphrase').change(function() {
            clearMsg('searchButton');
        });

        $('#dropdown_viewpick').change(function() {
            //document.getElementById('frmSearch').action = 'class_list.aspx';
            //$('#frmSearch').submit();
            $('#reorderList').submit();
        });

        $('#submenu_quicklinks').click(function() {
            expand_submenu('QUICKLINKS');
        });

        $('#submenu_categories').click(function() {
            expand_submenu('CATEGORIES');
        });

        $('#submenu_search').click(function() {
            expand_submenu('SEARCH');
        });
    });

    function expand_submenu(iSubMenuOption) {
        var lcl_list_show  = '';
        var lcl_list_hide1 = '';
        var lcl_list_hide2 = '';

        if (iSubMenuOption == 'CATEGORIES') {
            lcl_list_show  = 'submenu_categories_list';
            lcl_list_hide1 = 'submenu_quicklinks_list';
            lcl_list_hide2 = 'submenu_search_box';
        } else if (iSubMenuOption == 'SEARCH') {
            lcl_list_show  = 'submenu_search_box';
            lcl_list_hide1 = 'submenu_quicklinks_list';
            lcl_list_hide2 = 'submenu_categories_list';
        } else {
            lcl_list_show  = 'submenu_quicklinks_list';
            lcl_list_hide1 = 'submenu_categories_list';
            lcl_list_hide2 = 'submenu_search_box';
        }

        if ($('#' + lcl_list_show).css('display') == 'block') {
            $('#submenu_lists').slideUp('slow');
            $('#submenu_quicklinks_list').css('display', 'none');
            $('#submenu_categories_list').css('display', 'none');
            $('#submenu_search_box').css('display', 'none');
        } else {
            if ($('#submenu_lists').css('display') == 'block') {
                $('#' + lcl_list_hide1).css('display', 'none');
                $('#' + lcl_list_hide2).css('display', 'none');
                $('#submenu_lists').slideUp('slow', function() {
                    $('#submenu_lists').slideDown('slow');
                    $('#' + lcl_list_show).css('display', 'block');
                });
            } else {
                $('#' + lcl_list_hide1).css('display', 'none');
                $('#' + lcl_list_hide2).css('display', 'none');
                $('#' + lcl_list_show).css('display', 'block');
                $('#submenu_lists').slideDown('slow');
            }
        }
    }

    function clickToRegister() {
        alert('continue registration process...');
    }
</script>

</head>
<body>
<div id="wrapper_body">
  <div id="wrapper_header">
    <Tbanner:banner ID="banner" runat="server" />
    <Tnavigation:navigation ID="egov_navigation" runat="server" egovsection="" rootcategory="" showviewpicks="" viewpick="" categoryid="" />
  </div>
  <div id="wrapper_content">
    <!-- <Twelcomeinfo:welcomeinfo ID="egov_welcomeinfo" runat="server" /> -->
    <div id="content">
<%
    showMemberWarning(sOrgID);
    
    if (sCategoryID == iRootCategoryID)
    {
        listCategories2(Convert.ToInt32(sOrgID),
                        sCategoryID,
                        sDisplayImage,
                        sDisplayDesc);
    }
    else
    {
        Response.Write("show category title + class listings");
    }
            
%>
    </div>
  </div>
  
  <p>
    Display screen as Logged In: 
    <input type="radio" name="isLoggedin" id="isLoggedIn_yes" value="Y"<%=lcl_checked_isLoggedInYes%> />Yes
    <input type="radio" name="isLoggedIn" id="isLoggedIn_no" value="N"<%=lcl_checked_isLoggedInNo%> />No
  </p>

  <div id="wrapper_footer">
    <Tfooter:footer ID="footer" runat="server" />
  </div>
</div>
</body>
</html>
