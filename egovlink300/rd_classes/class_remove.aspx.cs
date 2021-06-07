using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class class_remove : System.Web.UI.Page
{
    double startCounter = 0.00;

    static string sOrgID = common.getOrgId();
    static string sOrgName = common.getOrgName(sOrgID);

    protected void Page_PreInit(object sender, EventArgs e)
    {
        // This is the earliest thing the page does, so set the start time here.
        startCounter = DateTime.Now.TimeOfDay.TotalSeconds;
    }

    protected void Page_PreRender(object sender, EventArgs e)
    {
        //THIS STILL NEEDS FIXED!!!!
        common.logThePageVisit(startCounter, "class_remove.aspx", "public");
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        Int32 sCartID = 0;
        Int32 sTimeID = 0;

        string sBuyOrWait = "";

        try
        {
            sCartID = Convert.ToInt32(Request["cartid"]);
        }
        catch
        {
            sCartID = 0;
        }

        try
        {
            sTimeID = Convert.ToInt32(Request["timeid"]);
        }
        catch
        {
            sTimeID = 0;
        }

        if (Request["buyorwait"] != "")
        {
            sBuyOrWait = Request["buyorwait"];
            sBuyOrWait = common.dbSafe(sBuyOrWait);
        }

        classes.removeItemFromCart(sCartID,
                                   sTimeID,
                                   sBuyOrWait);

        classes.resetCartPrices();
        classes.determineDiscounts();

        Response.Redirect("class_cart.aspx");
    }
}
