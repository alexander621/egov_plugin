using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_classes_rd_user_login : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
    }

    public void displayUserLogin(Int32 iOrgID)
    {
        string sSessionLoginDisplayMsg      = "";
        string sReturnButtonActionLine      = "";
        string sReturnButtonLabelActionLine = "";
        string sDNKLoginMessage             = "";
        string sUserEmail                   = "";

        Boolean sOrgHasFeature_showActionLineLinks    = common.orgHasFeature(iOrgID.ToString(), "show actionline links");
        Boolean sOrgHasDisplay_doNotKnockLoginMessage = common.orgHasDisplay(iOrgID.ToString(), "donotknock_login_message");

        if (Session["LoginDisplayMsg"] != null && Session["LoginDisplayMsg"] != "")
        {
            sSessionLoginDisplayMsg = "<div id=\"loginDisplayMsg\">" + Session["LoginDisplayMsg"] + "</div>";
        }

        if (sOrgHasFeature_showActionLineLinks)
        {
            sReturnButtonLabelActionLine = common.getFeatureName(iOrgID.ToString(), "action line");
            sReturnButtonActionLine      = "<input type=\"button\" name=\"returnButtonActionLine\" id=\"returnButtonActionLine\" value=\"Return to " + sReturnButtonLabelActionLine + "\" onclick=\"location.href='action.asp';\" />";
        }

        //Determine if the user is wanting to signup on the "Do Not Knock" list.
        if (Request["p"] == "dnk")
        {
            if (sOrgHasDisplay_doNotKnockLoginMessage)
            {
                sDNKLoginMessage = common.getOrgDisplay(iOrgID.ToString(), "donotknock_login_message");
            }
            else
            {
                sDNKLoginMessage  = "<div id=\"loginMsg_dnk_title\">DO NOT KNOCK REGISTRATION</div>";
                sDNKLoginMessage += "<div id=\"loginMsg_dnk_msg\">";
                sDNKLoginMessage += "  Log into an existing, or register a new, account to";
                sDNKLoginMessage += "  add yourself to the \"Do Not Knock\" list(s)";
                sDNKLoginMessage += "</div>";
            }

            sDNKLoginMessage = "<div id=\"loginMsg_doNotKnock\">" + sDNKLoginMessage + "</div>";
        }

        Response.Write(sReturnButtonActionLine);

        Response.Write("<div id=\"loginContainer\">\n");
        Response.Write("  <div id=\"loginErrorMsg\">&nbsp;</div>\n");
        Response.Write(sSessionLoginDisplayMsg);
        Response.Write(sDNKLoginMessage);

        Response.Write("<fieldset class=\"login_fieldset\">\n");
        Response.Write("  <legend>Sign In</legend>\n");
        Response.Write("  <form name=\"userLogin\" id=\"userLogin\" method=\"post\" action=\"rd_user_login_action.aspx\">\n");
        Response.Write("    <input type=\"hidden\" name=\"orgid\" id=\"orgid\" value=\"" + iOrgID.ToString() + "\" />\n");
        Response.Write("  <table id=\"userLoginTable\">\n");
        Response.Write("    <tr>\n");
        Response.Write("        <th>Email:</th>\n");
        Response.Write("        <td><input type=\"text\" name=\"email\" id=\"email\" value=\"" + sUserEmail + "\" size=\"30\" maxlength=\"100\" /></td>\n");
        Response.Write("    </tr>\n");
        Response.Write("    <tr>\n");
        Response.Write("        <th>Password:</th>\n");
        Response.Write("        <td><input type=\"password\" name=\"password\" id=\"password\" value=\"\" size=\"30\" maxlength=\"100\" /></td>\n");
        Response.Write("    </tr>\n");
        Response.Write("    <tr>\n");
        Response.Write("        <td colspan=\"2\" align=\"center\">\n");
        Response.Write("            <input type=\"button\" name=\"signInButton\" id=\"signInButton\" value=\"Sign In\" />\n");
        Response.Write("        </td>\n");
        Response.Write("    </tr>\n");
	/*
	Response.Write("    <script src=\"https://apis.google.com/js/platform.js\" async defer></script>\n");
	Response.Write("    <meta name=\"google-signin-client_id\" content=\"1087616019738-p41a8s5a4hd9k7b6r4j27sto7d1e760d.apps.googleusercontent.com\">\n");
	Response.Write("    <tr>\n");
	Response.Write("    <td colspan=\"2\">\n");
	Response.Write("    <br />\n");
	Response.Write("    <br />\n");
	Response.Write("    <div id=\"g-signin\" onclick=\"setClickTrue();\" data-onsuccess=\"onSignIn\"></div>\n");
	Response.Write("    <a id=\"g-signout\" href=\"#\" onclick=\"signOut();\" style=\"display:none;\">Sign out</a>\n");
	Response.Write("    </td>\n");
	Response.Write("    </tr>\n");
	Response.Write("    <script>\n");
    	Response.Write("    function renderButton() {\n");
      	Response.Write("    gapi.signin2.render('g-signin', {\n");
        	Response.Write("    'scope': 'profile email',\n");
        	Response.Write("    'width': 285,\n");
        	Response.Write("    'height': 50,\n");
        	Response.Write("    'longtitle': true,\n");
        	Response.Write("    'theme': 'dark',\n");
        	Response.Write("    'onsuccess': onSignIn,\n");
        	Response.Write("    'onfailure': null\n");
      	Response.Write("    });\n");
    	Response.Write("    }\n");
	Response.Write("    var click = false;\n");
  	Response.Write("    function onSignIn(googleUser) {\n");
        	Response.Write("    // The ID token you need to pass to your backend:\n");
        	Response.Write("    var id_token = googleUser.getAuthResponse().id_token;\n");
	Response.Write("    \n");
		Response.Write("    document.getElementById('g-signin').style.display = 'none';\n");
	Response.Write("    \n");
	Response.Write("    \n");
		Response.Write("    //IF NOT ALREAY LOGGED IN!\n");
		Response.Write("    if (readCookie(\"userid\") == \"\" || readCookie(\"userid\") == \"-1\" || click)\n");
		Response.Write("    {\n");
			Response.Write("    window.location = 'test_gauth.asp?id_token=' + id_token;\n");
		Response.Write("    }\n");
	Response.Write("    \n");
	Response.Write("    }\n");
	Response.Write("    function setClickTrue()\n");
	Response.Write("    {\n");
		Response.Write("    click = true;\n");
	Response.Write("    }\n");
	Response.Write("    \n");
	Response.Write("    (function(){\n");
    	Response.Write("    var cookies;\n");
	Response.Write("    \n");
    	Response.Write("    function readCookie(name,c,C,i){\n");
        	Response.Write("    if(cookies){ return cookies[name]; }\n");
	Response.Write("    \n");
        	Response.Write("    c = document.cookie.split('; ');\n");
        	Response.Write("    cookies = {};\n");
	Response.Write("    \n");
        	Response.Write("    for(i=c.length-1; i>=0; i--){\n");
           	Response.Write("    C = c[i].split('=');\n");
           	Response.Write("    cookies[C[0]] = C[1];\n");
        	Response.Write("    }\n");
	Response.Write("    \n");
        	Response.Write("    return cookies[name];\n");
    	Response.Write("    }\n");
	Response.Write("    \n");
    	Response.Write("    window.readCookie = readCookie; // or expose it however you want\n");
	Response.Write("    })();\n");
	Response.Write("    </script>\n");
  	Response.Write("    <script src=\"https://apis.google.com/js/platform.js?onload=renderButton\" async defer></script>\n");
	*/
        Response.Write("  </table>\n");
        Response.Write("  </form>");

        //Other Login Links
        Response.Write("  <div class=\"userLoginOtherLink\" onclick=\"location.href='rd_forgot_password.aspx'\">Forgot your password?</div>");
        //Response.Write("  <div class=\"userLoginOtherLink\" onclick=\"location.href='rd_register.aspx'\">Not registered yet?</div>");
        Response.Write("  <div class=\"userLoginOtherLink\" onclick=\"location.href='register.asp?from=ASPX'\">Not registered yet?</div>");
        Response.Write("</fieldset>");
        Response.Write("</div>");

        //Internal Only Fields
        Response.Write("<div id=\"problemTextDiv\">");
        //Response.Write("<div>");
        Response.Write("  Internal Use Only, Leave Blank:");
        Response.Write("  <input type=\"text\" name=\"frmsubjecttext\" id=\"problemtextinput\" value=\"\" size=\"6\" maxlength=\"6\" /><br />");
        Response.Write("  <strong>Please leave this field blank and remove any values that have been populated for it.</strong>");
        Response.Write("</div>");
    }
}
