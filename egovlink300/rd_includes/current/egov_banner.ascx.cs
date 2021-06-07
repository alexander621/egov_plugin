using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_includes_egovbanner : System.Web.UI.UserControl
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    public void showOrgBanner()
    {
        string sOrgID                = common.getOrgId();
        string sImgBaseURL           = common.getOrgInfo(sOrgID,"orgEgovWebsiteURL");
        string sTopGraphicLeftURL    = common.getOrgInfo(sOrgID,"orgTopGraphicLeftURL");
        string sTopGraphicRightURL   = common.getOrgInfo(sOrgID,"orgTopGraphicRightURL");
        string sHeaderSize           = common.getOrgInfo(sOrgID,"orgHeaderSize");
        string sOrgPublicWebsiteURL  = common.getOrgInfo(sOrgID,"orgPublicWebsiteURL");
        string lcl_banner            = "";
        string lcl_style_banner      = "";
        string lcl_style_banner_logo = "";
        
        if (HttpContext.Current.Request.ServerVariables["HTTPS"].ToUpper() == "ON") {
            sImgBaseURL = sImgBaseURL.Replace("http://www.egovlink.com","https://secure.egovlink.com");
        }

        if (sTopGraphicLeftURL != "")
        {
            lcl_style_banner  = "background-image:url('" + sTopGraphicRightURL + "');";
            //lcl_style_banner_logo = " style=\"height:" + sHeaderSize + "px;\"";

            lcl_banner  = "<a href=\"" + sOrgPublicWebsiteURL + "\">";
            lcl_banner += "<img src=\"" + sTopGraphicLeftURL + "\" name=\"City Logo\" id=\"City Logo\" border=\"0\" title=\"Click here to return to the E-Gov Services start page\"" + lcl_style_banner_logo + " />";
            lcl_banner += "</a>";

            Response.Write("<div id=\"header\">");
            Response.Write("  <div class=\"topbanner\" style=\"" + lcl_style_banner + "\">");
            Response.Write("    " + lcl_banner);
            Response.Write("  </div>");
            Response.Write("</div>");
        }
    }
}
