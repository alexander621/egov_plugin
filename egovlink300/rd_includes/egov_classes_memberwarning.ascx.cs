using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_includes_egovclasses_memberwarning : System.Web.UI.UserControl
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    private string lcl_orgid;

    public string orgid
    {
        get { return lcl_orgid; }
        set { lcl_orgid = value; }
    }

    public void showMemberWarning()
    {
        Int32 sOrgID = 0;
        string sOrgDisplay = "";

        if (lcl_orgid != null)
        {
            try
            {
                sOrgID = Convert.ToInt32(lcl_orgid);
            }
            catch
            {
                sOrgID = 0;
            }
        }

        if (common.orgHasDisplay(sOrgID.ToString(), "classdetailsnotice"))
        {
            sOrgDisplay = common.getOrgDisplay(sOrgID.ToString(), "classdetailsnotice");

            Response.Write("<div id=\"classdetailsnotice\">" + sOrgDisplay + "</div>");
        }
    }
}
