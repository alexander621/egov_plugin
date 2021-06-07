using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

public partial class rd_forgot_password : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    public void displayForgotPassword()
    {
        Response.Write("<div id=\"forgotpwd_error\"></div>");
        Response.Write("<fieldset class=\"forgotpwd_fieldset\">");
        Response.Write("  <legend>Password Assistance</legend>");
        //Response.Write("  <form name=\"lookupemail\" id=\"lookupemail\" method=\"post\" action=\"forgot_password_action.aspx\" onsubmit=\"return checkform();\">");
        Response.Write("  <div id=\"passwordMsg\">Please enter the email address that you used to register your account.</div>");
        Response.Write("  <div id=\"forgotPassword\">");
        Response.Write("    <strong>Email: </strong>");
        Response.Write("    <input type=\"text\" name=\"email\" id=\"email\" value=\"\" />");
        Response.Write("    <input type=\"button\" name=\"lookupButton\" id=\"lookupButton\" value=\"Lookup\" />");
        Response.Write("  </div>");
        //Response.Write("  </form>");
        Response.Write("</fieldset>");
        Response.Write("<div>");
        Response.Write("  <input type=\"button\" name=\"loginButton\" id=\"loginButton\" value=\"Return to Login\" />");
        Response.Write("</div>");
    }
}
