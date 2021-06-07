using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_logout : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //Because ASPX and ASP cookies do NOT work well together, it was decided that a new "userid" cookie be used
        //for ASPX pages.  When a cookie is created in ASP or ASPX and "destroyed" and then a cookie with the same 
        //name is attempted to be created in the other language as second cookie is actually created.  Therefore, TWO
        //cookies with the same name are created and the system doesn't know which one to use.
        //For example, if a user logs in via ASP attempts to log in via ASPX (i.e. as a different user/account)
        //  then there would be two "userid" cookies.
        //NOTE: When we log out we need to "destory" both userid cookies in ASP and ASPX.
        //To do this we first come here and then redirect to the logout.asp.
//        HttpCookie sCookieUserID = Request.Cookies["useridx"];
        HttpCookie sCookieUserID = Request.Cookies["userid"];

        //string sOrgID = common.getOrgId();
        //string sEGovDefaultPage = common.getEGovDefaultPage(Convert.ToInt32(sOrgID));

        //if (sCookieUserID != null)
        //{
        if (sCookieUserID != null)
        {
            sCookieUserID.Value = "";
            sCookieUserID.Expires = DateTime.Now.AddDays(-1);

            string appName = Page.ResolveUrl("~"); //Gets the application name
            sCookieUserID.Path = appName.Substring(0, appName.Length - 1); //Trims the trailing slash to match the cookie path created by Classic ASP

            Response.Cookies.Add(sCookieUserID);
        }
        //}

        //Response.Redirect(sEGovDefaultPage);
        Response.Redirect("logout.asp?from=ASPX");
    }
}
