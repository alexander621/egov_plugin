using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class check_cart_purchases : System.Web.UI.Page
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
        common.logThePageVisit(startCounter, "check_cart_purchases.aspx", "public");
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        string sSessionID = "";
        string sSQL       = "";
        string lcl_return = "WAITLISTONLY";

        if (Request["sessionid"] != "")
        {
            sSessionID = Request["sessionid"];
            sSessionID = common.dbSafe(sSessionID);
        }

        sSessionID = "'" + sSessionID + "'";

        //We could just look for the B's or do a DISTINCT, but this is the fastest pull of the data
        sSQL  = "SELECT buyorwait ";
        sSQL += " FROM egov_class_cart ";
        sSQL += " WHERE sessionid_csharp = " + sSessionID;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                //We just need to find one.
                if (Convert.ToString(myReader["buyorwait"]).ToUpper() == "B")
                {
                    lcl_return = "PURCHASES";
                    break;
                }
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        Response.Write(lcl_return);
    }
}
