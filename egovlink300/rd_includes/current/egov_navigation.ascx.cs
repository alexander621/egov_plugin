using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_includes_egovnavigation : System.Web.UI.UserControl
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    public void showNavigationMenu()
    {
        string sOrgID         = common.getOrgId();
        string sblnMenuOn     = common.getOrgInfo(sOrgID, "orgDisplayMenu");
        string sblnCustomMenu = common.getOrgInfo(sOrgID, "orgCustomMenu");
        string sLabelCity     = "";
        string sLabelEgov     = "";
        string sURLCity       = "";
        string sURLEgov       = "";

        Boolean sDisplayMenu            = Convert.ToBoolean(sblnMenuOn);
        Boolean sCustomMenu             = Convert.ToBoolean(sblnCustomMenu);
        Boolean sMenuOptionEnabled_City = common.checkMenuOptionEnabled(sOrgID, "CITY");
        Boolean sMenuOptionEnabled_Egov = common.checkMenuOptionEnabled(sOrgID, "EGOV");

        if (sDisplayMenu)
        {
            if (sCustomMenu)
            {
                Response.Write("<nav id=\"nav-wrap\">");
                Response.Write("<div id=\"menu-icon\">Main Menu</div>");
                Response.Write("  <ul id=\"navmenu\">");

                if (sMenuOptionEnabled_City)
                {
                    sLabelCity = common.getMenuOptionLabel(sOrgID, "CITY");
                    sURLCity   = common.getOrgInfo(sOrgID, "OrgPublicWebsiteURL");

                    Response.Write("<li><a href=\"" + sURLCity + "\">" + sLabelCity + "</a></li>");
                }

                if (sMenuOptionEnabled_Egov)
                {
                    sLabelEgov = common.getMenuOptionLabel(sOrgID, "EGOV");
                    sURLEgov   = common.getOrgInfo(sOrgID, "OrgEgovWebsiteURL");

                    Response.Write("<li><a href=\"" + sURLEgov + "\">" + sLabelEgov + "</a></li>");
                }

                showPublicDropDownMenu(sOrgID);

                Response.Write("</ul>");
                Response.Write("</nav>");
            }
        }
    }

    public void showPublicDropDownMenu(string iOrgID)
    {
        string sSQL     = "";
        string sMenuURL = "";

        sSQL  = "SELECT o.orgEgovWebsiteURL, ";
        sSQL += " isnull(fo.publicURL,f.publicURL) as publicURL, ";
        sSQL += " isnull(fo.featurename,f.featurename) as featurename ";
        sSQL += " FROM organizations o, ";
        sSQL +=      " egov_organizations_to_features fo, ";
        sSQL +=      " egov_organization_features f ";
        sSQL += " WHERE fo.publiccanview = 1 ";
        sSQL += " AND f.haspublicview = 1 ";
        sSQL += " AND o.orgid = fo.orgid ";
        sSQL += " AND fo.featureid = f.featureid ";
        sSQL += " AND o.orgid = " + iOrgID;
        sSQL += " ORDER BY fo.publicdisplayorder, f.publicdisplayorder ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                sMenuURL = myReader["publicURL"].ToString();

                if (!sMenuURL.ToUpper().StartsWith("HTTP"))
                {
                    sMenuURL = myReader["orgEgovWebsiteURL"].ToString();
                    sMenuURL += "/";
                    sMenuURL += myReader["publicURL"].ToString();
                }

                Response.Write("<li><a href=\"" + sMenuURL + "\">" + myReader["featurename"] + "</a></li>");
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

    }

    private string lcl_egovsection;
    private string lcl_rootcategory;
    private string lcl_showviewpicks;
    private string lcl_viewpick;
    private string lcl_categoryid;

    public string egovsection
    {
        get { return lcl_egovsection; }
        set { lcl_egovsection = value; }
    }

    public string rootcategory
    {
        get { return lcl_rootcategory; }
        set { lcl_rootcategory = value; }
    }

    public string showviewpicks
    {
        get { return lcl_showviewpicks; }
        set { lcl_showviewpicks = value; }
    }

    public string viewpick
    {
        get { return lcl_viewpick; }
        set { lcl_viewpick = value; }
    }

    public string categoryid
    {
        get { return lcl_categoryid; }
        set { lcl_categoryid = value; }
    }

    public void showSubMenu()
    {
        string sOrgID = common.getOrgId();

        Response.Write("<div id=\"submenu_nav\">");
        Response.Write("  <ul id=\"submenu_options\">");
        Response.Write("    <li id=\"submenu_quicklinks\">Quick Links</li>");

        if (lcl_egovsection == "CLASSES")
        {
            Response.Write("<li id=\"submenu_categories\">Categories</li>");
            Response.Write("<li id=\"submenu_search\">Search</li>");
        }

        Response.Write("  </ul>");
        Response.Write("  <div id=\"submenu_nav_lists\">");
        Response.Write("    <div id=\"submenu_lists\">");
        Response.Write("      <div id=\"submenu_quicklinks_list\">");
                                showLoggedInLinks();
        Response.Write("      </div>");

        if (lcl_egovsection == "CLASSES")
        {
            Response.Write("      <div id=\"submenu_categories_list\">");
            displaySubCategoryMenu2(sOrgID,
                                    Convert.ToInt32(lcl_rootcategory),
                                    Convert.ToBoolean(lcl_showviewpicks),
                                    Convert.ToBoolean(lcl_viewpick),
                                    Convert.ToInt32(lcl_categoryid));
            Response.Write("      </div>");
            Response.Write("      <div id=\"submenu_search_box\">");
                                    displayClassesSearchBox(Convert.ToInt32(lcl_categoryid));
            Response.Write("      </div>");
        }

        Response.Write("    </div>");
        Response.Write("  </div>");
        Response.Write("</div>");
    }

    public void showLoggedInLinks()
    {
        string sOrgID = common.getOrgId();
        string sOrgURL = "";
        string sProtocol = "http://";
        string sCookieUserID = "";

        Boolean lcl_canViewPeddlers = false;
        Boolean lcl_canViewSolicitors = false;

        Boolean lcl_orghasfeature_payments = common.orgHasFeature(sOrgID, "payments");
        Boolean lcl_orghasfeature_action_line = common.orgHasFeature(sOrgID, "action line");
        Boolean lcl_orghasfeature_activities = common.orgHasFeature(sOrgID, "activities");
        Boolean lcl_orghasfeature_facilities = common.orgHasFeature(sOrgID, "facilities");
        Boolean lcl_orghasfeature_memberships = common.orgHasFeature(sOrgID, "memberships");
        Boolean lcl_orghasfeature_gifts = common.orgHasFeature(sOrgID, "gifts");
        Boolean lcl_orghasfeature_bid_postings = common.orgHasFeature(sOrgID, "bid_postings");
        Boolean lcl_orghasfeature_donotknock = common.orgHasFeature(sOrgID, "donotknock");

        Boolean lcl_publicCanViewFeature_payments = common.publicCanViewFeature(sOrgID, "payments");
        Boolean lcl_publicCanViewFeature_action_line = common.publicCanViewFeature(sOrgID, "action line");
        Boolean lcl_publicCanViewFeature_activities = common.publicCanViewFeature(sOrgID, "activities");
        Boolean lcl_publicCanViewFeature_facilities = common.publicCanViewFeature(sOrgID, "facilities");
        Boolean lcl_publicCanViewFeature_memberships = common.publicCanViewFeature(sOrgID, "memberships");
        Boolean lcl_publicCanViewFeature_gifts = common.publicCanViewFeature(sOrgID, "gifts");
        Boolean lcl_publicCanViewFeature_bid_postings = common.publicCanViewFeature(sOrgID, "bid_postings");
        Boolean lcl_publicCanViewFeature_donotknock = common.publicCanViewFeature(sOrgID, "donotknock");

        //Setup the OrgURL
        if (HttpContext.Current.Request.ServerVariables["HTTPS"].ToUpper() == "ON")
        {
            sProtocol = "https://";
        }

        sOrgURL = sProtocol;
        sOrgURL += HttpContext.Current.Request.ServerVariables["server_name"].ToLower();
        sOrgURL += "/";
        sOrgURL += common.GetVirtualDirectyName(HttpContext.Current.Request.ServerVariables["URL"].ToLower());

        Response.Write("<div id=\"loggedinlinks\">");
        Response.Write("  <ul id=\"loggedinlinks_list\">");

        //Manage Account Link
        Response.Write("    <li><a href=\"" + sOrgURL + "/manage_account.asp\">Manage Account</a></li>");

        //View Standard EGov Payments Link
        if (lcl_orghasfeature_payments && lcl_publicCanViewFeature_payments)
        {
            Response.Write("    <li><a href=\"" + sOrgURL + "/user_home.asp?trantype=1\">View Payments</a></li>");
        }

        //View Submitted Action Line Requests Link
        if (lcl_orghasfeature_action_line && lcl_publicCanViewFeature_action_line)
        {
            Response.Write("    <li><a href=\"" + sOrgURL + "/user_home.asp?trantype=0\">View Requests</a></li>");
        }

        //View Shopping Cart (Purchases) Link
        if (lcl_orghasfeature_activities && lcl_publicCanViewFeature_activities)
        {
            Response.Write("    <li><a href=\"" + sOrgURL + "/classes/class_cart.asp\">View Cart</a></li>");
        }

        if ((lcl_orghasfeature_facilities && lcl_publicCanViewFeature_facilities)
           || (lcl_orghasfeature_activities && lcl_publicCanViewFeature_activities)
           || (lcl_orghasfeature_memberships && lcl_publicCanViewFeature_memberships)
           || (lcl_orghasfeature_gifts && lcl_publicCanViewFeature_gifts))
        {
            Response.Write("    <li><a href=\"" + sOrgURL + "/purchases_report/purchases_list.asp\">View Purchases</a></li>");
        }

        //View Bids (Bid Postings) Link
        if (lcl_orghasfeature_bid_postings && lcl_publicCanViewFeature_bid_postings)
        {
            Response.Write("    <li><a href=\"" + sOrgURL + "/view_bids.asp\">View Bids</a></li>");
        }

        //Do Not Knock List Link
        if (Request.Cookies["userid"] != null)
        {
            sCookieUserID = Request.Cookies["userid"].Value;
        }

        lcl_canViewPeddlers = common.checkAccessToList(sCookieUserID, sOrgID, "peddlers");
        lcl_canViewSolicitors = common.checkAccessToList(sCookieUserID, sOrgID, "solicitors");

        if (lcl_orghasfeature_donotknock &&
           lcl_publicCanViewFeature_donotknock &&
           (lcl_canViewPeddlers || lcl_canViewSolicitors))
        {
            Response.Write("    <li><a href=\"" + sOrgURL + "/view_donotknock.asp\">View \"Do Not Knock\" List</a></li>");

        }

        //Logout Link
        Response.Write("    <li><a href=\"" + sOrgURL + "/logout.asp\">Log Out</a></li>");

        Response.Write("  </ul>");
        Response.Write("</div>");
    }

    public void displaySubCategoryMenu2(string iOrgID,
                                   Int32 iRootCategoryID,
                                   Boolean iShowViewPicks,
                                   Boolean iViewPick,
                                   Int32 iCategoryID)
    {
        string sSQL = "";
        Int32 sOrgID = 0;
        Int32 sLineCount = 0;

        if (iOrgID != null)
        {
            try
            {
                sOrgID = Convert.ToInt32(iOrgID);
            }
            catch
            {
                sOrgID = 0;
            }
        }

        sSQL = "SELECT categorytitle, ";
        sSQL += " subcategoryid, ";
        sSQL += " subcategorytitle ";
        sSQL += " FROM class_categories ";
        sSQL += " WHERE orgid = " + sOrgID;
        sSQL += " AND categoryid = " + iRootCategoryID;
        sSQL += " ORDER BY sequenceid, subcategorytitle";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                sLineCount = sLineCount + 1;

                if (sLineCount == 1)
                {
                    //Response.Write("<div id=\"subcategorymenu_new\" style=\"display:none\">");
                    Response.Write("<div id=\"subcategorymenu_new\">");
                    Response.Write("  <ul id=\"subcategorymenu_list\">");
                    Response.Write("    <li><a id=\"subcategorymenu_rootoption\" href=\"class_list.aspx?categoryid=" + iRootCategoryID.ToString() + "\">" + myReader["categorytitle"].ToString() + "</a></li>");
                }

                Response.Write("    <li><a href=\"class_list.aspx?categoryid=" + myReader["subcategoryid"].ToString() + "\">" + myReader["subcategorytitle"].ToString() + "</a>");
            }

            Response.Write("    </li>");
            Response.Write("  </ul>");
            Response.Write("</div>");
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public void displayClassesSearchBox(Int32 iCategoryID)
    {
        Response.Write("<div id=\"classesSearchBox\">");
        Response.Write("<form name=\"frmSearch\" id=\"frmSearch\" method=\"post\" action=\"class_search_results.asp\">");
        Response.Write("  <input type=\"hidden\" name=\"categoryid\" id=\"categoryid\" value=\"" + iCategoryID.ToString() + "\" />");
        Response.Write("  <strong>Search: </strong>");
        Response.Write("  <input type=\"text\" name=\"txtsearchphrase\" id=\"txtsearchphrase\" value=\"STILL NEED TO CREATE SEARCH RESULTS PAGE\" />");
        Response.Write("  <input type=\"button\" name=\"searchButton\" id=\"searchButton\" value=\"Find\" class=\"button\" />");
        Response.Write("</form>");
        Response.Write("</div>");
    }
}
