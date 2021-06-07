using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_classes_ProcessFailure : System.Web.UI.Page
{
    double startCounter = 0.00;

    static string sOrgID   = common.getOrgId();
    static string sOrgName = common.getOrgName(sOrgID);

    protected void Page_PreInit(object sender, EventArgs e)
    {
        // This is the earliest thing the page does, so set the start time here.
        startCounter = DateTime.Now.TimeOfDay.TotalSeconds;
    }

    protected void Page_PreRender(object sender, EventArgs e)
    {
        //THIS STILL NEEDS FIXED!!!!
        common.logThePageVisit(startCounter, "ProcessFailure.aspx", "public");
    }

    protected void Page_Load(object sender, EventArgs e)
    {
    }

    public void displayProcessFailure(Int32 iOrgID)
    {
        Int32 sProcessingErrorID = 0;

        string sErrorMsg       = "";
        string sSQL            = "";

        try
        {
            sProcessingErrorID = Convert.ToInt32(Request["p"]);
        }
        catch
        {
            sProcessingErrorID = 0;
        }

        //Build the error message
        sErrorMsg = common.getProcessingPaymentErrorMsg(iOrgID, sProcessingErrorID);

        Response.Write("<fieldset class=\"processFailureFieldset\">");
        Response.Write("  <legend>Payment Processing Failure</legend>");
        Response.Write("  <div class=\"processFailureErrorMsg\">We are sorry, but we cannot process your payment at this time due to the following error...</div>");
        Response.Write(   sErrorMsg);
        Response.Write("</fieldset>");
    }

}
