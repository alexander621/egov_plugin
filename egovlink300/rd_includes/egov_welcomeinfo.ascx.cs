using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_includes_egovwelcomeinfo : System.Web.UI.UserControl
{
    public void showRegisteredUser()
    {
//        HttpCookie sCookieUserID = Request.Cookies["useridx"];
        HttpCookie sCookieUserID = Request.Cookies["userid"];

        string sOrgID                = common.getOrgId();
        string sImgBaseURL           = common.getOrgInfo(sOrgID, "orgEgovWebsiteURL");
        string lcl_isOrgRegistration = common.getOrgInfo(sOrgID, "orgRegistration");
        string sSQL                  = "";
        string sUserName             = "";
        string sWelcomeMsg           = "";

        Boolean sOrgRegistration = Convert.ToBoolean(lcl_isOrgRegistration);
        Boolean sShowCurrentDate = true;

        if (HttpContext.Current.Request.ServerVariables["HTTPS"].ToUpper() == "ON")
        {
            sImgBaseURL = sImgBaseURL.Replace("http://www.egovlink.com", "https://secure.egovlink.com");
        }

        Response.Write("<div id=\"welcome_info\">");

        if (sOrgRegistration)
        {
            if((sCookieUserID.Value != "") && (sCookieUserID.Value != "-1"))
            {
                //We hide the current date because we are instead showing the
                //welcome message and logged-in links.
                sShowCurrentDate = false;

                //Get the username from "userid" cookie
                sSQL = "SELECT userfname + ' ' + userlname AS username ";
                sSQL += " FROM egov_users ";
                sSQL += " WHERE userid = " + sCookieUserID.Value;

                SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
                sqlConn.Open();

                SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
                SqlDataReader myReader;
                myReader = myCommand.ExecuteReader();

                if (myReader.HasRows)
                {
                    while (myReader.Read())
                    {
                        sUserName = myReader["username"].ToString();
                    }
                }

                myReader.Close();
                sqlConn.Close();
                myReader.Dispose();
                sqlConn.Dispose();

                //Build the Welcome Message
                sWelcomeMsg = "Welcome";

                if (sUserName.Trim() != "")
                {
                    sWelcomeMsg += ", <strong>" + sUserName.ToUpper() +"</strong>!<br />";
                }

                Response.Write("  <div id=\"accountmenu\">");
                Response.Write("    <img id=\"img_welcomeinfo\" class=\"accountmenu\" src=\"" + sImgBaseURL + "/images/accountmenu.jpg\" />" + sWelcomeMsg);
                                    showLoggedInLinks();
                                    //showLoggedInLinks2();
                Response.Write("  </div>");
            }
        }

        //Show the current date if the user is not logged in.
        if (sShowCurrentDate)
        {
            Response.Write("<div id=\"datetagline\">");
            Response.Write("<font class=\"datetagline\">Today is " + DateTime.Now.ToString("dddd, MMMM dd, yyyy") + ".</font>");
            Response.Write("</div>");
        }

        Response.Write("</div>");
    }

    public void showLoggedInLinks()
    {
//        HttpCookie sCookieUserID = Request.Cookies["useridx"];
        HttpCookie sCookieUserID = Request.Cookies["userid"];

        string sOrgID = common.getOrgId();
        string sOrgURL = "";
        string sProtocol = "http://";

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

        Response.Write("<div id=\"loggedinlinks_menu\">");
        Response.Write("QUICK LINKS<img id=\"quicklinks_menu-icon\" src=\"../rd_classes/menu-icon_white.png\" align=\"right\" />");
        Response.Write("</div>");
        Response.Write("<div id=\"loggedinlinks\">");

        //Manage Account Link
        Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/manage_account.asp\">MANAGE ACCOUNT</a>");
        Response.Write("<span class=\"accountmenu_seperator\"> | </span>");

        //View Standard EGov Payments Link
        if (lcl_orghasfeature_payments && lcl_publicCanViewFeature_payments)
        {
            Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/user_home.asp?trantype=1\">VIEW PAYMENTS</a>");
            Response.Write("<span class=\"accountmenu_seperator\"> | </span>");
        }

        //View Submitted Action Line Requests Link
        if (lcl_orghasfeature_action_line && lcl_publicCanViewFeature_action_line)
        {
            Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/user_home.asp?trantype=0\">VIEW REQUESTS</a>");
            Response.Write("<span class=\"accountmenu_seperator\"> | </span>");
        }

        //View Shopping Cart (Purchases) Link
        if (lcl_orghasfeature_activities && lcl_publicCanViewFeature_activities)
        {
            Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/classes/class_cart.asp\">VIEW CART</a>");
            Response.Write("<span class=\"accountmenu_seperator\"> | </span>");
        }

        if ((lcl_orghasfeature_facilities && lcl_publicCanViewFeature_facilities)
           || (lcl_orghasfeature_activities && lcl_publicCanViewFeature_activities)
           || (lcl_orghasfeature_memberships && lcl_publicCanViewFeature_memberships)
           || (lcl_orghasfeature_gifts && lcl_publicCanViewFeature_gifts))
        {
            Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/purchases_report/purchases_list.asp\">VIEW PURCHASES</a>");
            Response.Write("<span class=\"accountmenu_seperator\"> | </span>");
        }

        //View Bids (Bid Postings) Link
        if (lcl_orghasfeature_bid_postings && lcl_publicCanViewFeature_bid_postings)
        {
            Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/view_bids.asp\">VIEW BIDS</a>");
            Response.Write("<span class=\"accountmenu_seperator\"> | </span>");
        }

        //Do Not Knock List Link
        lcl_canViewPeddlers   = common.checkAccessToList(sCookieUserID.Value, 
                                                         sOrgID, 
                                                         "peddlers");

        lcl_canViewSolicitors = common.checkAccessToList(sCookieUserID.Value, 
                                                         sOrgID, 
                                                         "solicitors");

        if (lcl_orghasfeature_donotknock &&
           lcl_publicCanViewFeature_donotknock &&
           (lcl_canViewPeddlers || lcl_canViewSolicitors))
        {
            Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/view_donotknock.asp\">VIEW \"DO NOT KNOCK LIST\"</a>");
            Response.Write("<span class=\"accountmenu_seperator\"> | </span>");
        }

        //Logout Link
        Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/logout.asp\">LOG OUT</a>");

        Response.Write("</div>");
    }

    public void showLoggedInLinks2()
    {
//        HttpCookie sCookieUserID = Request.Cookies["useridx"];
        HttpCookie sCookieUserID = Request.Cookies["userid"];

        string sOrgID = common.getOrgId();
        string sOrgURL = "";
        string sProtocol = "http://";

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

        //Manage Account Link
        Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/manage_account.asp\">MANAGE ACCOUNT</a>");
        Response.Write("<span class=\"accountmenu_seperator\"> - </span>");

        //View Standard EGov Payments Link
        if (lcl_orghasfeature_payments && lcl_publicCanViewFeature_payments)
        {
            Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/user_home.asp?trantype=1\">VIEW PAYMENTS</a>");
            Response.Write("<span class=\"accountmenu_seperator\"> - </span>");
        }

        //View Submitted Action Line Requests Link
        if (lcl_orghasfeature_action_line && lcl_publicCanViewFeature_action_line)
        {
            Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/user_home.asp?trantype=0\">VIEW REQUESTS</a>");
            Response.Write("<span class=\"accountmenu_seperator\"> - </span>");
        }

        //View Shopping Cart (Purchases) Link
        if (lcl_orghasfeature_activities && lcl_publicCanViewFeature_activities)
        {
            Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/classes/class_cart.asp\">VIEW CART</a>");
            Response.Write("<span class=\"accountmenu_seperator\"> - </span>");
        }

        if ((lcl_orghasfeature_facilities && lcl_publicCanViewFeature_facilities)
           || (lcl_orghasfeature_activities && lcl_publicCanViewFeature_activities)
           || (lcl_orghasfeature_memberships && lcl_publicCanViewFeature_memberships)
           || (lcl_orghasfeature_gifts && lcl_publicCanViewFeature_gifts))
        {
            Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/purchases_report/purchases_list.asp\">VIEW PURCHASES</a>");
            Response.Write("<span class=\"accountmenu_seperator\"> - </span>");
        }

        //View Bids (Bid Postings) Link
        if (lcl_orghasfeature_bid_postings && lcl_publicCanViewFeature_bid_postings)
        {
            Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/view_bids.asp\">VIEW BIDS</a>");
            Response.Write("<span class=\"accountmenu_seperator\"> - </span>");
        }

        //Do Not Knock List Link
        lcl_canViewPeddlers   = common.checkAccessToList(sCookieUserID.Value, 
                                                         sOrgID, 
                                                         "peddlers");

        lcl_canViewSolicitors = common.checkAccessToList(sCookieUserID.Value, 
                                                         sOrgID, 
                                                         "solicitors");

        if (lcl_orghasfeature_donotknock &&
           lcl_publicCanViewFeature_donotknock &&
           (lcl_canViewPeddlers || lcl_canViewSolicitors))
        {
            Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/view_donotknock.asp\">VIEW \"DO NOT KNOCK LIST\"</a>");
            Response.Write("<span class=\"accountmenu_seperator\"> - </span>");
        }

        //Logout Link
        Response.Write("<a class=\"accountmenu\" href=\"" + sOrgURL + "/logout.asp\">LOG OUT</a>");
    }
}
