using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_includes_egovfooter : System.Web.UI.UserControl
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    public void buildFooter()
    {
        showFadeLines();

        Response.Write("  <div id=\"footer\">");

                            showBottomNavigation();
                            showCopyrightInfo();
                            showGoogleTranslator();

        Response.Write("  </div>");
        //Response.Write("<script>if (window.top!=window.self) { var height = document.body.scrollHeight; parent.postMessage(height,\"*\"); } </script> ");
	Response.Write("<script> function onElementHeightChange(elm, callback){ var lastHeight = elm.clientHeight, newHeight; (function run(){ newHeight = elm.clientHeight; if( lastHeight != newHeight ) callback(); lastHeight = newHeight; if( elm.onElementHeightChangeTimer ) clearTimeout(elm.onElementHeightChangeTimer); elm.onElementHeightChangeTimer = setTimeout(run, 200); })(); } if (window.top!=window.self) { onElementHeightChange(document.body, function(){ var height = document.body.scrollHeight; parent.postMessage({event_id: 'heightchange',data: { heightval: height, initial: false }},\"*\") }); var height = document.body.scrollHeight; parent.postMessage({event_id: 'heightchange',data: { heightval: height, initial: true }},\"*\") } </script> ");
    }

    public void showFadeLines()
    {
        Response.Write("<div id=\"footer_fadeline_top\"></div>");
        Response.Write("<div id=\"footer_fadeline_bottom\"></div>");
    }

    public void showBottomNavigation()
    {
        string sOrgID     = common.getOrgId();
        string sSQL       = "";
        string sLabelCity = "";
        string sLabelEgov = "";
        string sURLCity   = "";
        string sURLEgov   = "";
        string sMenuURL   = "";

        Int32 sTotalCount = 0;

        Boolean sMenuOptionEnabled_City          = common.checkMenuOptionEnabled(sOrgID, "CITY");
        Boolean sMenuOptionEnabled_Egov          = common.checkMenuOptionEnabled(sOrgID, "EGOV");
        Boolean sOrgHasFeature_administratorLink = common.orgHasFeature(sOrgID, "AdministrationLink");
        Boolean sOrgHasDisplay_privacyPolicy     = common.orgHasDisplay(sOrgID, "privacy policy");
        Boolean sOrgHasDisplay_refundPolicy      = common.orgHasDisplay(sOrgID, "refund policy");


        Response.Write("<div id=\"bottom_navigation\">");

        //Menu Option: City Home -------------------------------------------------
        if (sMenuOptionEnabled_City)
        {
            sLabelCity  = common.getMenuOptionLabel(sOrgID, "CITY");
            sURLCity    = common.getOrgInfo(sOrgID, "OrgPublicWebsiteURL");
            sTotalCount = sTotalCount + 1;

            Response.Write("  <a href=\"" + sURLCity + "\" target=\"_top\">" + sLabelCity + "</a>");
        }

        //Menu Option: E-Gov Home ------------------------------------------------
        if (sMenuOptionEnabled_Egov)
        {
            sLabelEgov  = common.getMenuOptionLabel(sOrgID, "EGOV");
            sURLEgov    = common.getOrgInfo(sOrgID, "OrgEgovWebsiteURL");
            sTotalCount = sTotalCount + 1;

            if (sMenuOptionEnabled_City)
            {
                Response.Write("<span class=\"bottom_menu_seperator\"> | </span>");
            }

            Response.Write("<a href=\"" + sURLEgov + "\" target=\"_top\">" + sLabelEgov + "</a>");
        }

        //Menu Option: Features --------------------------------------------------
        sSQL  = "SELECT o.orgEgovWebsiteURL, ";
        sSQL += " isnull(fo.publicURL, f.publicURL) as publicURL, ";
        sSQL += " isnull(fo.featurename, f.featurename) as featurename ";
        sSQL += " FROM organizations o, ";
        sSQL +=      " egov_organizations_to_features fo, ";
        sSQL +=      " egov_organization_features f ";
        sSQL += " WHERE fo.publicCanView = 1 ";
        sSQL += " AND f.hasPublicView = 1 ";
        sSQL += " AND o.orgid = fo.orgid ";
        sSQL += " AND fo.featureid = f.featureid ";
        sSQL += " AND o.orgid = " + sOrgID;
        sSQL += " ORDER BY fo.publicDisplayOrder, f.publicDisplayOrder ";
        
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();
        
        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();
        
        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                sTotalCount = sTotalCount + 1;

                Response.Write("<span class=\"bottom_menu_seperator\"> | </span>");

                sMenuURL = myReader["publicURL"].ToString();

                if (! sMenuURL.ToUpper().StartsWith("HTTP"))
                {
                    sMenuURL  = myReader["orgEgovWebsiteURL"].ToString();
                    sMenuURL += "/";
                    sMenuURL += myReader["publicURL"].ToString();
                }

                Response.Write("<a href=\"" + sMenuURL + "\" target=\"_top\">" + myReader["featurename"] + "</a>");
            }
        }
        
        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        //BEGIN: New Line for Non-Feature menu options ---------------------------
        Response.Write("<br />");
        sTotalCount = 0;

        //Menu Option: Privacy Policy --------------------------------------------
        if (sOrgHasDisplay_privacyPolicy)
        {
            Response.Write("<a href=\"" + sURLEgov + "/privacy_policy_display.asp\" target=\"_top\"><strong>Privacy Policy</strong></a>");

            sTotalCount = sTotalCount + 1;
        }

        //Menu Option: Refund Policy ---------------------------------------------
        if (sOrgHasDisplay_refundPolicy)
        {
            if (sTotalCount > 0)
            {
                Response.Write("<span class=\"bottom_menu_seperator\"> | </span>");
            }

            Response.Write("<a href=\"" + sURLEgov + "/refund_policy.asp\" target=\"_top\">Refund Policy</a>");
            
            sTotalCount = sTotalCount + 1;
        }

        //Menu Option: Log-In and Register ---------------------------------------
        if (sTotalCount > 0)
        {
            Response.Write("<span class=\"bottom_menu_seperator\"> | </span>");
        }

        //Response.Write("  <a href=\"" + sURLEgov + "/user_login.asp\" target=\"_top\">Log-In</a>");
        Response.Write("  <a href=\"" + sURLEgov + "/rd_user_login.aspx\" target=\"_top\">Log-In</a>");
        Response.Write("  <span class=\"bottom_menu_seperator\"> | </span>");
        Response.Write("  <a href=\"" + sURLEgov + "/register.asp?from=ASPX\" target=\"_top\">Register</a>");

        //Menu Option: Admin Link ------------------------------------------------
        if (sOrgHasFeature_administratorLink)
        {
            Response.Write("<span class=\"bottom_menu_seperator\"> | </span>");
            Response.Write("  <a href=\"" + sURLEgov + "/admin\" target=\"_new\"><strong>Administrator</strong></a>");
        }
        
        Response.Write("</div>");
    }

    public void showCopyrightInfo()
    {
        Response.Write("<div id=\"footer_copyright\">");
        Response.Write("  Copyright &copy;2004-" + DateTime.Now.ToString("yyyy") + ".  ");
        Response.Write("  Electronic Commerce Link, Inc.</em>");
        Response.Write("</div>");
    }

    public void showGoogleTranslator()
    {
        string sOrgID = common.getOrgId();
        
        Boolean sOrgHasFeature_googleTranslator = common.orgHasFeature(sOrgID, "google_translator");

        if (Application["environment"] == "PROD")
        {

        }
        
        Response.Write("<div id=\"google_translator\">");
        Response.Write("  <div id=\"google_translate_element\"></div>");
        Response.Write("  <script>");
        Response.Write("    function googleTranslateElementInit() {");
        Response.Write("      new google.translate.TranslateElement({");
        Response.Write("        pageLanguage: 'en'");
        Response.Write("      }, 'google_translate_element');");
        Response.Write("    }");
        Response.Write("  </script>");

        if (HttpContext.Current.Request.ServerVariables["HTTPS"].ToUpper() != "ON")
        {
            Response.Write("<script src=\"http://translate.google.com/translate_a/element.js?cb=googleTranslateElementInit\"></script>");
        }

        Response.Write("</div>");
    }

}
