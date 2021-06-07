using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class cookie_final : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
    }

    public void displayCookieData()
    {
			Response.Write("the dupname cookie value is: " + Request.Cookies["dupname"].Value + "<hr>");
            foreach (string cookiekey in Request.Cookies.Keys)
            {
                Response.Write("<br /><b>" + cookiekey + "</b>: " + Request.Cookies[cookiekey].Value);
            }
			
		
    }
}
