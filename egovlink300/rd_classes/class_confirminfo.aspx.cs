using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class class_confirminfo : System.Web.UI.Page
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
        common.logThePageVisit(startCounter, "class_confirminfo.aspx", "public");
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        Int32 sOrgID          = 0;
        Int32 sFamilyMemberID = 0;

        string sEmergencyInfo = "";

        try
        {
            sOrgID = Convert.ToInt32(Request.Form["orgid"]);
        }
        catch
        {
            sOrgID = 0;
        }

        try
        {
            sFamilyMemberID = Convert.ToInt32(Request.Form["familymemberid"]);
        }
        catch
        {
            sFamilyMemberID = 0;
        }

        sEmergencyInfo = classes.showEmergencyInfo(sOrgID,
                                                   sFamilyMemberID);
        
        Response.Write(sEmergencyInfo);


    }
}
