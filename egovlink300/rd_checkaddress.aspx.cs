using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_checkaddress : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

        Int32 sOrgID = 0;

        string lcl_streetnumber = "";
        string lcl_stname       = "";
        string lcl_addresstype  = "";
        string lcl_returntype   = "CHECK";
        string lcl_return       = "";

        if (Request["orgid"] != "")
        {
            try
            {
                sOrgID = Convert.ToInt32(Request["orgid"]);
            }
            catch
            {
                sOrgID = 0;
            }
        }

        if (Request["stnumber"] != "")
        {
            lcl_streetnumber = Request["stnumber"];
            lcl_streetnumber = common.dbSafe(lcl_streetnumber);
            lcl_streetnumber = "'" + lcl_streetnumber + "'";
        }

        if (Request["stname"] != "")
        {
            lcl_stname = Request["stname"];
            //lcl_stname = common.dbSafe(lcl_stname);
            lcl_stname = "'" + lcl_stname + "'";
        }

        if (Request["addresstype"] != "")
        {
            lcl_addresstype = Request["addresstype"];
            lcl_addresstype = lcl_addresstype.ToUpper();
            lcl_addresstype = common.dbSafe(lcl_addresstype);
        }

        //Determine if we are "CHECKING" to see if the address exists
        //or "DISPLAYING" a list of valid addresses.
        if (Request["returntype"] != "")
        {
            lcl_returntype = Request["returntype"];
            lcl_returntype = lcl_returntype.ToUpper();
            lcl_returntype = common.dbSafe(lcl_returntype);
        }

        if (lcl_returntype == "DISPLAY_OPTIONS")
        {
            lcl_return = buildAddressOptions(sOrgID,
                                             lcl_stname,
                                             lcl_addresstype);
        }
        else
        {
            lcl_return = checkAddressOptions(sOrgID,
                                             lcl_streetnumber,
                                             lcl_stname,
                                             lcl_addresstype);
        }
 
        Response.Write(lcl_return);
    }

    public static string buildAddressOptions(Int32 iOrgID,
                                             string iSTName,
                                             string iAddressType)
    {
        string lcl_display_options = "";
        string sSQL                = "";
        string sOptionValue        = "";

        sSQL  = "SELECT DISTINCT residentstreetnumber, ";
        sSQL += " residentstreetname, ";
        sSQL += " CAST(residentstreetnumber AS INT) AS ordernumb, ";
        sSQL += " ISNULL(residentstreetprefix,'') AS residentstreetprefix, ";
        sSQL += " ISNULL(streetsuffix,'') AS streetsuffix, ";
        sSQL += " ISNULL(streetdirection,'') AS streetdirection ";
        sSQL += " FROM egov_residentaddresses ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();

        if (iAddressType == "LARGE")
        {
            sSQL += " AND (residentstreetname = " + iSTName;
            sSQL += " OR residentstreetname + ' ' + streetsuffix = " + iSTName;
            sSQL += " OR residentstreetname + ' ' + streetdirection = " + iSTName;
            sSQL += " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = " + iSTName;
            sSQL += " OR residentstreetprefix + ' ' + residentstreetname = " + iSTName;
            sSQL += " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = " + iSTName;
            sSQL += " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = " + iSTName;
            sSQL += " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = " + iSTName;
            sSQL += " ) ";
        }
        else
        {
            sSQL += " AND residentaddressid = " + iSTName;
        }

        sSQL += " AND excludefromactionline = 0 ";
        sSQL += " ORDER BY 2, 5, 6, 4, 3, 1 ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            lcl_display_options  = "<strong>Valid Address Choices</strong><br />";
            lcl_display_options += "<select name=\"stnumber\" id=\"stnumber\" size=\"10\">";

            while (myReader.Read())
            {
                sOptionValue = Convert.ToString(myReader["residentstreetnumber"]) + " " + Convert.ToString(myReader["residentstreetprefix"]) + " " + Convert.ToString(myReader["residentstreetname"]) + " " + Convert.ToString(myReader["streetsuffix"]) + " " + Convert.ToString(myReader["streetdirection"]);
                sOptionValue = sOptionValue.Trim();

                lcl_display_options += "<option value=\"" + Convert.ToString(myReader["residentstreetnumber"]) + "\">" + sOptionValue + "</option>";
            }

            lcl_display_options += "</select>";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
        
        return lcl_display_options;

    }

    public static string checkAddressOptions(Int32 iOrgID,
                                             string iStreetNumber,
                                             string iSTName,
                                             string iAddressType)
    {
        string sSQL       = "";
        string lcl_return = "NOT FOUND";

        sSQL = "SELECT count(residentaddressid) as hits ";
        sSQL += " FROM egov_residentaddresses ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();
        sSQL += " AND residentstreetnumber = " + iStreetNumber;
        sSQL += " AND (residentstreetname = " + iSTName;
        sSQL += " OR residentstreetname + ' ' + streetsuffix = " + iSTName;
        sSQL += " OR residentstreetname + ' ' + streetdirection = " + iSTName;
        sSQL += " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = " + iSTName;
        sSQL += " OR residentstreetprefix + ' ' + residentstreetname = " + iSTName;
        sSQL += " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = " + iSTName;
        sSQL += " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = " + iSTName;
        sSQL += " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = " + iSTName;
        sSQL += ")";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToInt32(myReader["hits"]) > 0)
            {
                lcl_return = "FOUND CHECK";
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }
}
