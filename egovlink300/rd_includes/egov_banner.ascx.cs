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

	 Response.Write("<script> if (window.top!=window.self) { document.body.classList.add(\"iframeformat\"); }</script> ");
	Response.Write("<script>window.addEventListener('load', function () { if (navigator.userAgent.indexOf('Safari') != -1 && navigator.userAgent.indexOf('Chrome') == -1 && navigator.userAgent.indexOf('CriOS') == -1) { if ( window.location !== window.parent.location ) {  document.getElementsByTagName('body')[0].innerHTML =  '<center><input type=\"button\" class=\"reserveformbutton\" style=\"width:auto;text-align:center;\" name=\"continue\" id=\"continueButton\" value=\"Safari Users must click here to continue\" onclick=\"safNewWin();\" /></center>'; } } }); function safNewWin() {window.open(window.location, \"_blank\")} </script>");
        if (sTopGraphicLeftURL != "")
        {

            sImgBaseURL         = common.getBaseURL(sImgBaseURL);
            sTopGraphicLeftURL  = common.getBaseURL(sTopGraphicLeftURL);
            sTopGraphicRightURL = common.getBaseURL(sTopGraphicRightURL);

            lcl_style_banner  = "background-image:url('" + sTopGraphicRightURL + "');";
            //lcl_style_banner_logo = " style=\"height:" + sHeaderSize + "px;\"";

            lcl_banner  = "<a href=\"" + sOrgPublicWebsiteURL + "\">";
            lcl_banner += "<img src=\"" + sTopGraphicLeftURL + "\" name=\"City Logo\" id=\"City Logo\" border=\"0\" title=\"Click here to return to the E-Gov Services start page\"" + lcl_style_banner_logo + " />";
            lcl_banner += "</a>";

	    Response.Headers.Add("P3P", "CP=This is not a P3P privacy policy!  Read the privacy policy here: " + sImgBaseURL + "/privacy_policy.asp");

	    Response.Write("<div id=\"iframenav\" style=\"display:none;\">");
	    Response.Write("<div class=\"iframenavlink iframenavbutton\"><a href=\"" + sImgBaseURL + "/rd_classes/class_categories.aspx\">Classes and Events</a></div>");
	    Response.Write("<div class=\"iframenavlink iframenavbutton\"><a href=\"" + sImgBaseURL + "/rentals/rentalcategories.asp\">Rentals</a></div>");
	    Response.Write("<div class=\"iframenavlink iframenavbutton\"><a href=\"" + sImgBaseURL + "/user_login.asp\">Login</a></div>");
 	    Response.Write("<div class=\"searchMenuDiv\">    ");
	 	    Response.Write("<div class=\"searchBoxText iframenavbutton\"><span>Search</span></div>    ");
	 	    Response.Write("<div class=\"searchBox\">      ");
		 	    Response.Write("<div class=\"classesSearchBox\">");
			 	    Response.Write("<input type=\"text\" name=\"txtsearchphrase\" class=\"txtsearchphrase\" value=\"\" size=\"40\" />        ");
			 	    Response.Write("<input type=\"button\" name=\"searchButton\" class=\"searchButton\" value=\"Find\" />      ");
		 	    Response.Write("</div>    ");
	 	    Response.Write("</div>  ");
 	    Response.Write("</div>");
	    Response.Write("</div>");
	    Response.Write("<div id=\"footerbug\" style=\"display:none;\">");
	    Response.Write("<a href=\"http://www.egovlink.com\" target=\"_top\">Powered By EGovLink</a>");
	    Response.Write("</div>");
            Response.Write("<div id=\"header\">");
            Response.Write("  <div class=\"topbanner\" style=\"" + lcl_style_banner + "\">" + lcl_banner + "</div>");
            Response.Write("</div>");
        }
    }
}
