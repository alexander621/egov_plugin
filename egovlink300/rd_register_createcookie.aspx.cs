using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_register_createcookie : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //Because ASPX and ASP cookies do NOT work well together, it was decided that a new "userid" cookie be used
        //for ASPX pages.  When a cookie is created in ASP or ASPX and "destroyed" and then a cookie with the same 
        //name is attempted to be created in the other language as second cookie is actually created.  Therefore, TWO
        //cookies with the same name are created and the system doesn't know which one to use.
        //For example, if a user logs in via ASP attempts to log in via ASPX (i.e. as a different user/account)
        //  then there would be two "userid" cookies.
        HttpCookie sCookieUserID = new HttpCookie("useridx");

        string sLoginUserID = "";

        if (Request["userid"] != "")
        {
            sLoginUserID = Request["userid"];
            sLoginUserID = common.dbSafe(sLoginUserID);
        }

        //Add the UserID cookie.
        sCookieUserID.Value   = sLoginUserID;
        sCookieUserID.Expires = DateTime.Now.AddYears(1);

        Response.Cookies.Add(sCookieUserID);

        Response.Redirect("rd_user_login.aspx");
    }
}
