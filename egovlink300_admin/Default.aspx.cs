using System;
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

public partial class _Default : System.Web.UI.Page 
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //Label1.Text = common.GetVirtualDirectyName(Request.ServerVariables["URL"]);
        //Label1.Text = Request.ServerVariables["HTTPS"];
        //Label1.Text = Request.ServerVariables["server_name"];
        string sProtocol = "http://";
        if (Request.ServerVariables["HTTPS"].ToUpper() == "ON")
            sProtocol = "https://";
        Label1.Text = sProtocol + Request.ServerVariables["server_name"] + "/" + common.GetVirtualDirectyName(Request.ServerVariables["URL"]);
        if (Session["orgid"] == null)
            common.setOrganizationSessionVariables();
        Label2.Text = Session["orgid"].ToString();
        //Session["orgid"] = "";
    }
    protected void Button1_Click(object sender, EventArgs e)
    {
        int BadNo = 3;
        int zero = 0;
        int Valuebad = BadNo / zero;
    }
}
